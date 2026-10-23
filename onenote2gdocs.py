#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "msal>=1.28",
#   "requests>=2.31",
#   "beautifulsoup4>=4.12",
#   "google-api-python-client>=2.130",
#   "google-auth-oauthlib>=1.2",
#   "rich>=13.7",
# ]
# ///
"""
OneNote → Google Docs — one Doc per section, pages as document tabs.

Usage:
    uv run onenote2gdocs.py export  [--only "Notebook/Section"]
    uv run onenote2gdocs.py import  [--only ...] [--dry-run] [--force] [--drive-parent FOLDER_URL]
                                    [--drive-folder NAME]
    uv run onenote2gdocs.py check-google [--quota-project ID]
    uv run onenote2gdocs.py status
    uv run onenote2gdocs.py convert export/.../content.html     (debug: print parsed blocks)

Two phases, so the slow, quota-limited Google writes can resume without re-downloading:
    export  Microsoft Graph → ./export/    raw page HTML, images, attachments, manifest.json
    import  ./export/ → Google Drive       folders per notebook / section group,
                                           a Doc per section, a tab per page,
                                           subpages as child tabs (OneNote levels 1-2)

Mapping notes:
    • Google Docs caps a doc at 100 tabs, so big sections split into "Section (part N)".
    • Images are never uploaded or shared. Each one becomes a highlighted
      "[Image N: paste here]" placeholder, and images_todo.md lists every placeholder with
      a link to its tab and the local image file under ./export/ to paste in by hand.
    • Attachments are uploaded privately to "<Section> files" and linked from the page.
    • Not carried over: handwriting/ink (Graph doesn't return it), links between OneNote
      pages (left as plain text), fonts and sizes other than headings and monospace.

Sign-in is ephemeral and needs no app registration on either side:
    Microsoft  Microsoft's own public "Microsoft Graph Command Line Tools" client (the one
               Connect-MgGraph uses). Tokens stay in memory.
    Google     gcloud's built-in sign-in, run against a throwaway CLOUDSDK_CONFIG that is
               deleted once the token is in memory; the grant is revoked when the run ends.
               Your normal gcloud config directory is never read or changed.

Config (flags override env):
    ONENOTE_CLIENT_ID     optional: your own Entra public client instead of Microsoft's
    ONENOTE_TENANT        tenant ID or domain (default: organizations)
    GOOGLE_QUOTA_PROJECT  Cloud project with Drive + Docs APIs enabled (needed with gcloud)
    GOOGLE_CLIENT_FILE    only with --google-auth client (default: ./credentials.json)
"""

import argparse
import hashlib
import json
import mimetypes
import os
import random
import re
import sys
import time
from dataclasses import dataclass, field, replace
from pathlib import Path

import requests
from bs4 import BeautifulSoup, Comment, Declaration, Doctype, NavigableString, ProcessingInstruction, Tag
from rich.console import Console
from rich.progress import BarColumn, MofNCompleteColumn, Progress, TextColumn, TimeElapsedColumn, TimeRemainingColumn
from rich.style import Style
from rich.text import Text

BASE = Path(__file__).resolve().parent
EXPORT_DIR = BASE / "export"
MANIFEST = EXPORT_DIR / "manifest.json"
STATE_FILE = BASE / "import_state.json"
IMAGES_TODO = BASE / "images_todo.md"

GRAPH = "https://graph.microsoft.com/v1.0"
MS_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph Command Line Tools
MS_SCOPES = ["Notes.Read", "User.Read"]
G_DRIVE_FILE = "https://www.googleapis.com/auth/drive.file"   # only files this script creates
G_DRIVE_FULL = "https://www.googleapis.com/auth/drive"        # needed to write into an existing folder
G_DOCS = "https://www.googleapis.com/auth/documents"
G_GCLOUD_SCOPES = ["openid", "https://www.googleapis.com/auth/userinfo.email",
                   "https://www.googleapis.com/auth/cloud-platform"]   # gcloud's ADC login defaults
DOC_MIME = "application/vnd.google-apps.document"
FOLDER_MIME = "application/vnd.google-apps.folder"

MAX_TABS = 100          # Google Docs hard limit per document (tabs + subtabs)
MAX_DEPTH = 2           # 0-based; Docs allows 3 tab levels, same as OneNote pages
DOC_WRITES_PER_MIN = 55  # Docs API write quota is 60/min/user; stay under it
DOC_READS_PER_MIN = 250  # read quota is 300/min/user
FLUSH_AT = 300           # requests per batchUpdate
MAX_ATTEMPTS = 8         # per API call, retrying 429/5xx with exponential backoff
PLACEHOLDER_HIGHLIGHT = (1, .93, .35)


# ── Terminal output ──────────────────────────────────────────────────────────
# rich handles color and progress bars on macOS, Linux, and Windows, and drops both
# when output is piped to a file. Messages are printed with markup off: page titles
# can contain [brackets], which rich would otherwise read as style tags.

console = Console(highlight=False)


def say(msg: str, style: str | None = None):
    console.print(msg, style=style, markup=False)


def ok(msg: str):
    say(f"✓ {msg}", "green")


def warn(msg: str):
    say(f"! {msg}", "yellow")


def say_link(prefix: str, url: str, label: str | None = None, style: str | None = None):
    """Print `prefix` then a clickable link (OSC 8) in terminals that support it.

    The label defaults to the URL itself, so it stays readable and copyable where links aren't supported.
    """
    console.print(Text.assemble((prefix, style or ""),
                                (label or url, Style(color="cyan", underline=True, link=url))))


def progress_bar() -> Progress:
    # Disabled when piped, so a log file gets only the ✓ / ! lines, not bar redraws.
    return Progress(TextColumn("{task.description}", markup=False), BarColumn(), MofNCompleteColumn(),
                    TimeElapsedColumn(), TextColumn("eta"), TimeRemainingColumn(),
                    console=console, transient=True, disable=not console.is_terminal)


def clip(s: str, n: int = 40) -> str:
    return s if len(s) <= n else s[:n - 1] + "…"


def plural(n: int, word: str) -> str:
    return f"{n} {word}" + ("" if n == 1 else "s")


# ── Small helpers ────────────────────────────────────────────────────────────

def short(s: str) -> str:
    return hashlib.sha1(s.encode()).hexdigest()[:8]


def safe(name: str, limit: int = 80) -> str:
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name or "").strip(" .")
    return (name or "untitled")[:limit]


def u16(s: str) -> int:
    """Docs API indexes in UTF-16 code units, so emoji count as 2."""
    return len(s.encode("utf-16-le")) // 2


# Every text file is read and written as UTF-8 explicitly. Python's default is the
# locale encoding, which on Windows (cp1252) can't hold emoji or most non-Latin text.
UTF8 = "utf-8"


def load_json(path: Path, default=None):
    try:
        return json.loads(path.read_text(encoding=UTF8))
    except FileNotFoundError:
        return default


def save_json(path: Path, data):
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding=UTF8)
    tmp.replace(path)


def matches(path: list, only: list | None) -> bool:
    if not only:
        return True
    joined = "/".join(path).lower()
    return any(joined.startswith(o.lower().strip("/")) for o in only)


def ymd(iso: str | None) -> str:
    return (iso or "")[:10]


# ── Microsoft Graph (export side) ────────────────────────────────────────────

def ms_token_fn(client_id: str, tenant: str, device_code: bool):
    import msal
    # No token_cache argument: MSAL keeps tokens in memory, so they die with the process.
    app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant}")

    def token() -> str:
        accounts = app.get_accounts()
        res = app.acquire_token_silent(MS_SCOPES, account=accounts[0]) if accounts else None
        if not res and device_code:
            flow = app.initiate_device_flow(scopes=MS_SCOPES)
            if "user_code" not in flow:
                sys.exit(f"Microsoft device-code sign-in failed: {flow.get('error_description', flow)}")
            say(flow["message"], "bold cyan")
            res = app.acquire_token_by_device_flow(flow)
        elif not res:
            say("Opening your browser to sign in to Microsoft…", "cyan")
            res = app.acquire_token_interactive(MS_SCOPES, prompt="select_account")
        if "access_token" not in res:
            sys.exit(f"Microsoft sign-in failed: {res.get('error_description', res)}")
        return res["access_token"]

    return token


class Graph:
    def __init__(self, token_fn):
        self.token_fn = token_fn
        self.s = requests.Session()

    def get(self, url: str, params: dict | None = None, raw: bool = False):
        if not url.startswith("http"):
            url = GRAPH + url
        for attempt in range(MAX_ATTEMPTS):
            r = self.s.get(url, params=params, timeout=120,
                           headers={"Authorization": f"Bearer {self.token_fn()}"})
            if r.status_code in (429, 500, 502, 503, 504) and attempt < MAX_ATTEMPTS - 1:
                wait = r.headers.get("Retry-After")
                time.sleep(min(int(wait) if wait and wait.isdigit() else 2 ** attempt, 120))
                continue
            if r.status_code >= 400:
                hint = ""
                if "10008" in r.text:
                    hint = ("\n  (Graph error 10008: your OneDrive library has >5,000 OneNote items. "
                            "Microsoft's workaround is to move some notebooks out of that library.)")
                raise RuntimeError(f"Graph {r.status_code} on {url}: {r.text[:300]}{hint}")
            return r if raw else r.json()

    def get_all(self, url: str, params: dict | None = None) -> list:
        items, data = [], self.get(url, params)
        while True:
            items += data.get("value", [])
            nxt = data.get("@odata.nextLink")
            if not nxt:
                return items
            data = self.get(nxt)


def collect_sections(graph: Graph, base: str, path: list, out: list):
    for sec in graph.get_all(f"{base}/sections", {"$select": "id,displayName"}):
        out.append({"id": sec["id"], "name": sec["displayName"], "path": path + [sec["displayName"]]})
    for grp in graph.get_all(f"{base}/sectionGroups", {"$select": "id,displayName"}):
        collect_sections(graph, f"/me/onenote/sectionGroups/{grp['id']}", path + [grp["displayName"]], out)


def resource_urls(html: str) -> set:
    soup = BeautifulSoup(html, "html.parser")
    urls = set()
    for img in soup.find_all("img"):
        urls.add(img.get("data-fullres-src") or img.get("src"))
    for obj in soup.find_all("object"):
        urls.add(obj.get("data"))
    return {u for u in urls if u and u.startswith("https://graph.microsoft.com/")}


def export_page(graph: Graph, page: dict, pdir: Path):
    html = graph.get(f"/me/onenote/pages/{page['id']}/content", raw=True).text
    rdir = pdir / "resources"
    rdir.mkdir(parents=True, exist_ok=True)
    resources = {}
    for url in sorted(resource_urls(html)):
        r = graph.get(url, raw=True)
        mime = r.headers.get("Content-Type", "application/octet-stream").split(";")[0]
        name = short(url) + (mimetypes.guess_extension(mime) or ".bin")
        (rdir / name).write_bytes(r.content)
        resources[url] = {"file": name, "mime": mime}
    (pdir / "resources.json").write_text(json.dumps(resources, indent=2), encoding=UTF8)
    (pdir / "content.html").write_text(html, encoding=UTF8)


def export_cmd(args):
    graph = Graph(ms_token_fn(args.ms_client_id, args.ms_tenant, args.device_code))
    me = graph.get("/me", {"$select": "userPrincipalName"})
    ok(f"Signed in to Microsoft as {me.get('userPrincipalName')}")

    old = load_json(MANIFEST, {"sections": []})
    old_by_id = {s["id"]: s for s in old["sections"]}
    old_pages = {p["id"]: p for s in old["sections"] for p in s["pages"]}

    sections = []
    for nb in graph.get_all("/me/onenote/notebooks", {"$select": "id,displayName", "$orderby": "displayName"}):
        collect_sections(graph, f"/me/onenote/notebooks/{nb['id']}", [nb["displayName"]], sections)
    sections = [s for s in sections if matches(s["path"], args.only)]
    say(f"Found {plural(len(sections), 'section')} to export", "bold")
    EXPORT_DIR.mkdir(exist_ok=True)
    done = set()

    def save_manifest():
        # Saved after every section so an interrupted export resumes where it stopped.
        # Sections not reached yet (or outside --only) keep their previous entry.
        listed = {s["id"] for s in sections}
        merged = [s if s["id"] in done else old_by_id[s["id"]]
                  for s in sections if s["id"] in done or s["id"] in old_by_id]
        merged += [s for s in old["sections"] if s["id"] not in listed]
        save_json(MANIFEST, {"exported": time.strftime("%Y-%m-%dT%H:%M:%S"), "sections": merged})

    with progress_bar() as bar:
        all_task = bar.add_task("Sections", total=len(sections))
        page_task = bar.add_task("", total=None)
        for sec in sections:
            name = " / ".join(sec["path"])
            sdir = EXPORT_DIR.joinpath(*[safe(p) for p in sec["path"][:-1]]) / f"{safe(sec['name'])}__{short(sec['id'])}"
            sec["dir"] = str(sdir.relative_to(EXPORT_DIR))
            bar.update(page_task, description=f"  {clip(name)}", completed=0, total=None)
            pages = graph.get_all(f"/me/onenote/sections/{sec['id']}/pages", {"pagelevel": "true", "$top": "100"})
            pages.sort(key=lambda p: p.get("order", 0))
            bar.update(page_task, total=len(pages))
            sec["pages"], unchanged = [], 0
            for page in pages:
                pdir = sdir / short(page["id"])
                rec = {"id": page["id"], "title": page.get("title") or "Untitled page",
                       "level": page.get("level", 0), "order": page.get("order", 0),
                       "created": page.get("createdDateTime"), "modified": page.get("lastModifiedDateTime"),
                       "dir": str(pdir.relative_to(EXPORT_DIR))}
                prev = old_pages.get(page["id"])
                fresh = prev and prev.get("modified") == rec["modified"] and not prev.get("error")
                if fresh and (pdir / "content.html").exists():
                    unchanged += 1
                else:
                    try:
                        export_page(graph, page, pdir)
                    except Exception as e:  # keep going; import marks the page as failed
                        rec["error"] = str(e)[:500]
                        warn(f"{rec['title']}: {rec['error']}")
                sec["pages"].append(rec)
                bar.advance(page_task)
            done.add(sec["id"])
            save_manifest()
            ok(f"{name} — {plural(len(pages), 'page')}" + (f" ({unchanged} unchanged)" if unchanged else ""))
            bar.advance(all_task)

    errors = sum(1 for s in sections for p in s["pages"] if p.get("error"))
    summary = (f"Export complete: {plural(sum(len(s['pages']) for s in sections), 'page')}, "
               f"{plural(errors, 'error')} → ")
    if errors:
        say_link(f"! {summary}", EXPORT_DIR.as_uri(), str(EXPORT_DIR), "yellow")
    else:
        say_link(f"✓ {summary}", EXPORT_DIR.as_uri(), str(EXPORT_DIR), "green")


# ── OneNote HTML → block model ───────────────────────────────────────────────

@dataclass(frozen=True)
class Fmt:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    code: bool = False
    link: str | None = None
    highlight: tuple | None = None
    color: tuple | None = None


@dataclass
class Run:
    text: str
    fmt: Fmt = Fmt()


@dataclass
class Para:
    runs: list
    style: str = "NORMAL_TEXT"
    list_kind: str | None = None   # "bullet" | "number"
    level: int = 0


@dataclass
class Table:
    rows: list                       # rows → cells → runs
    header: bool = False
    merges: list = field(default_factory=list)  # (row, col, rowspan, colspan)


@dataclass
class Image:
    src: str


@dataclass
class Attachment:
    src: str
    name: str


BLOCK_TAGS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "pre", "blockquote", "cite", "dd", "dt"}
SKIP_TAGS = {"script", "style", "head", "title", "meta"}
TAG_ICONS = {"to-do": "☐ ", "to-do:completed": "☑ ", "important": "⭐ ", "question": "❓ "}
NAMED_COLORS = {"yellow": (1, 1, 0), "lime": (0, 1, 0), "aqua": (0, 1, 1), "cyan": (0, 1, 1),
                "fuchsia": (1, 0, 1), "magenta": (1, 0, 1), "red": (1, 0, 0), "blue": (0, 0, 1),
                "green": (0, .5, 0), "gray": (.5, .5, .5), "grey": (.5, .5, .5),
                "silver": (.75, .75, .75), "orange": (1, .65, 0), "purple": (.5, 0, .5),
                "teal": (0, .5, .5), "navy": (0, 0, .5), "maroon": (.5, 0, 0), "olive": (.5, .5, 0)}
WS = re.compile(r"[ \t\r\n\f]+")
PARA0 = {"style": "NORMAL_TEXT", "list_kind": None, "level": 0}


def parse_style(s: str | None) -> dict:
    out = {}
    for part in (s or "").split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip().lower()
    return out


def parse_color(v: str | None) -> tuple | None:
    if not v:
        return None
    v = v.strip().lower()
    if v in NAMED_COLORS:
        return NAMED_COLORS[v]
    m = re.fullmatch(r"#([0-9a-f]{3}|[0-9a-f]{6})", v)
    if m:
        h = m.group(1)
        h = "".join(c * 2 for c in h) if len(h) == 3 else h
        return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    m = re.fullmatch(r"rgba?\((\d+),\s*(\d+),\s*(\d+).*\)", v)
    if m:
        return tuple(int(x) / 255 for x in m.groups())
    return None


def is_plain(c: tuple | None, background: bool) -> bool:
    """Default colors we don't copy: white/none backgrounds, black text."""
    if c is None:
        return True
    return all(x > .97 for x in c) if background else all(x < .05 for x in c)


def pos_key(tag: Tag):
    st = parse_style(tag.get("style"))
    num = lambda k: float(re.sub(r"[^\d.]", "", st.get(k, "0")) or 0)
    return (num("top"), num("left"))


def span_attr(td: Tag, attr: str) -> int:
    value = str(td.get(attr, ""))
    return max(1, int(value)) if value.isdigit() else 1


def tag_icon(tag: Tag) -> str:
    tags = [t.strip() for t in (tag.get("data-tag") or "").split(",")]
    return "".join(TAG_ICONS.get(t, "") for t in tags)


class Converter:
    def __init__(self):
        self.blocks = []
        self.runs = []
        self.cur = PARA0

    def flush(self):
        runs = []
        for r in self.runs:
            if runs and runs[-1].text.endswith(" ") and r.text.startswith(" "):
                r = Run(r.text.lstrip(" "), r.fmt)
            if not r.text:
                continue
            if runs and runs[-1].fmt == r.fmt:
                runs[-1] = Run(runs[-1].text + r.text, r.fmt)
            else:
                runs.append(r)
        if runs:
            runs[0] = Run(runs[0].text.lstrip(), runs[0].fmt)
            runs[-1] = Run(runs[-1].text.rstrip(), runs[-1].fmt)
            runs = [r for r in runs if r.text]
        if any(r.text.strip() for r in runs):
            self.blocks.append(Para(runs, **self.cur))
        self.runs = []

    def fmt_for(self, node: Tag, fmt: Fmt) -> Fmt:
        n, st = node.name, parse_style(node.get("style"))
        f = fmt
        if n in ("b", "strong") or st.get("font-weight") in ("bold", "bolder", "600", "700", "800", "900"):
            f = replace(f, bold=True)
        if st.get("font-weight") == "normal":
            f = replace(f, bold=False)
        if n in ("i", "em") or st.get("font-style") == "italic":
            f = replace(f, italic=True)
        deco = st.get("text-decoration", "")
        if n == "u" or "underline" in deco:
            f = replace(f, underline=True)
        if n in ("s", "strike", "del") or "line-through" in deco:
            f = replace(f, strike=True)
        if n in ("code", "kbd", "pre") or any(m in st.get("font-family", "") for m in ("consolas", "courier", "mono")):
            f = replace(f, code=True)
        if n == "a" and node.get("href", "").startswith(("http://", "https://", "mailto:")):
            f = replace(f, link=node["href"])
        bg = parse_color(st.get("background-color") or st.get("background"))
        if not is_plain(bg, background=True):
            f = replace(f, highlight=bg)
        fg = parse_color(st.get("color"))
        if fg and not is_plain(fg, background=False) and n != "a":
            f = replace(f, color=fg)
        return f

    def children(self, node: Tag, fmt: Fmt, para: dict):
        kids = list(node.children)
        if node.name == "body" or node.get("data-absolute-enabled"):
            # OneNote places outline containers absolutely; read them top-to-bottom, left-to-right.
            placed = [k for k in kids if isinstance(k, Tag) and "absolute" in parse_style(k.get("style")).get("position", "")]
            rest = [k for k in kids if k not in placed]
            kids = rest + sorted(placed, key=pos_key)
        for k in kids:
            self.walk(k, fmt, para)

    def walk(self, node, fmt: Fmt, para: dict):
        if isinstance(node, (Comment, Doctype, Declaration, ProcessingInstruction)):
            return
        if isinstance(node, NavigableString):
            text = WS.sub(" ", str(node)).replace("\xa0", " ")
            if text:
                self.runs.append(Run(text, fmt))
            return
        name = node.name
        if name in SKIP_TAGS:
            return
        if name == "br":
            self.flush()
            return
        if name in ("ul", "ol"):
            self.flush()
            level = min(para["level"] + 1, 8) if para["list_kind"] else 0
            sub = {"style": "NORMAL_TEXT", "list_kind": "number" if name == "ol" else "bullet", "level": level}
            self.cur = sub
            self.children(node, fmt, sub)
            self.flush()
            self.cur = para
            return
        if name == "table":
            self.flush()
            self.table(node, fmt)
            self.cur = para
            return
        if name == "img":
            self.flush()
            src = node.get("data-fullres-src") or node.get("src")
            if src:
                self.blocks.append(Image(src))
            return
        if name == "object":
            self.flush()
            if node.get("data"):
                self.blocks.append(Attachment(node["data"], node.get("data-attachment") or "attachment"))
            return
        if name == "iframe":
            url = node.get("data-original-src") or node.get("src") or ""
            if url.startswith("http"):
                self.flush()
                self.blocks.append(Para([Run("▶ " + url, replace(fmt, link=url))], **para))
            return
        f = self.fmt_for(node, fmt)
        if name in BLOCK_TAGS:
            self.flush()
            p = dict(para)
            if re.fullmatch(r"h[1-6]", name):
                p["style"] = f"HEADING_{name[1]}"
            if name == "li" and not p["list_kind"]:
                p.update(list_kind="bullet", level=0)
            self.cur = p
            if tag_icon(node):
                self.runs.append(Run(tag_icon(node), Fmt()))
            self.children(node, f, p)
            self.flush()
            self.cur = para
            return
        if tag_icon(node):
            self.runs.append(Run(tag_icon(node), Fmt()))
        self.children(node, f, para)

    def table(self, node: Tag, fmt: Fmt):
        trs = [tr for grp in [node] + node.find_all(["thead", "tbody", "tfoot"], recursive=False)
               for tr in grp.find_all("tr", recursive=False)]
        grid, taken, merges, extra = {}, set(), [], []
        header = bool(trs) and all(c.name == "th" for c in trs[0].find_all(["td", "th"], recursive=False))
        for r, tr in enumerate(trs):
            c = 0
            for td in tr.find_all(["td", "th"], recursive=False):
                while (r, c) in taken:
                    c += 1
                rs, cs = span_attr(td, "rowspan"), span_attr(td, "colspan")
                runs, more = self.cell(td, fmt)
                grid[(r, c)] = runs
                extra += more
                taken |= {(r + i, c + j) for i in range(rs) for j in range(cs)}
                if rs > 1 or cs > 1:
                    merges.append((r, c, rs, cs))
                c += cs
        if not taken:
            return
        nrows = max(r for r, _ in taken) + 1
        ncols = max(c for _, c in taken) + 1
        rows = [[grid.get((r, c), []) for c in range(ncols)] for r in range(nrows)]
        self.blocks.append(Table(rows, header, merges))
        self.blocks += extra   # images/attachments found in cells follow the table

    def cell(self, td: Tag, fmt: Fmt):
        sub = Converter()
        sub.children(td, fmt, PARA0)
        sub.flush()
        runs, extra = [], []
        for b in sub.blocks:
            if isinstance(b, Para):
                if runs:
                    runs.append(Run("\n"))
                if b.list_kind:
                    runs.append(Run("  " * b.level + ("• " if b.list_kind == "bullet" else "- ")))
                runs += b.runs
            elif isinstance(b, Table):   # nested table → flattened text
                for row in b.rows:
                    if runs:
                        runs.append(Run("\n"))
                    runs.append(Run(" | ".join("".join(r.text for r in cell) for cell in row)))
            else:
                extra.append(b)
        return runs, extra


def html_to_blocks(html: str) -> list:
    soup = BeautifulSoup(html, "html.parser")
    conv = Converter()
    conv.children(soup.body or soup, Fmt(), PARA0)
    conv.flush()
    return conv.blocks


# ── Google Docs / Drive (import side) ────────────────────────────────────────

class Throttle:
    def __init__(self, per_min: int):
        self.interval, self.last = 60 / per_min, 0.0

    def wait(self):
        delay = self.last + self.interval - time.monotonic()
        if delay > 0:
            time.sleep(delay)
        self.last = time.monotonic()


GCLOUD_NOISE = ("Credentials saved to file", "These credentials will be used by any library",
                "Cannot find a quota project", "set-quota-project", "Updates are available",
                "components update", "please run:")


def gcloud_noise(line: str) -> bool:
    s = line.strip()
    # A bare "WARNING:" header precedes the quota-project note; other warnings still show.
    return s == "WARNING:" or s.startswith("$ gcloud") or any(n in s for n in GCLOUD_NOISE)


def google_scopes(args) -> list:
    return [G_DRIVE_FULL if args.drive_parent else G_DRIVE_FILE, G_DOCS]


def drive_folder_id(value: str) -> str:
    m = re.search(r"/folders/([\w-]+)", value)
    return m.group(1) if m else value.strip()


def gcloud_credentials(api_scopes: list, quota_project: str | None, keep_grant: bool):
    """Sign in through gcloud's built-in ADC client without touching your gcloud config.

    gcloud always writes the credential to disk, so it runs against a throwaway
    CLOUDSDK_CONFIG directory that is deleted as soon as the token is read into memory.
    The grant is revoked at exit unless keep_grant is set.
    """
    import atexit
    import shutil
    import subprocess
    import tempfile
    from google.oauth2.credentials import Credentials

    gcloud = shutil.which("gcloud") or sys.exit(
        "gcloud not found. Install the Google Cloud CLI: https://cloud.google.com/sdk/docs/install")
    scopes = G_GCLOUD_SCOPES + api_scopes
    for stale in Path(tempfile.gettempdir()).glob("o2g-gcloud-*"):   # left by a killed earlier run
        shutil.rmtree(stale, ignore_errors=True)
    tmp = tempfile.mkdtemp(prefix="o2g-gcloud-")   # readable only by you (mode 700 / per-user temp)
    try:
        say("Opening your browser to sign in to Google (gcloud)…", "cyan")
        # Quota project is set in Python, not by gcloud, so its quota warning and its
        # "credentials saved to file" notes (about the temp dir) are only noise here.
        proc = subprocess.Popen(
            [gcloud, "auth", "application-default", "login", "--quiet", "--disable-quota-project",
             "--scopes=" + ",".join(scopes)],
            env={**os.environ, "CLOUDSDK_CONFIG": tmp, "CLOUDSDK_COMPONENT_MANAGER_DISABLE_UPDATE_CHECK": "1"},
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1,
            encoding=UTF8, errors="replace")
        blank = True
        for line in proc.stdout:
            if gcloud_noise(line):
                continue
            if not line.strip():
                if blank:
                    continue
                blank = True
            else:
                blank = False
            say(line.rstrip("\n"), "dim")
        proc.wait()
        cred_file = Path(tmp) / "application_default_credentials.json"
        if proc.returncode or not cred_file.exists():
            sys.exit("gcloud sign-in failed or was cancelled. If Google showed \"Access blocked\", trust the "
                     "\"Google Auth Library\" app in the Admin console (see README.md).")
        info = json.loads(cred_file.read_text(encoding=UTF8))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    creds = Credentials.from_authorized_user_info(info, scopes)
    if quota_project:
        creds = creds.with_quota_project(quota_project)
    if not keep_grant:
        def revoke(token=info["refresh_token"]):
            try:
                revoked = requests.post("https://oauth2.googleapis.com/revoke", data={"token": token}, timeout=15).ok
            except requests.RequestException:
                revoked = False
            if revoked:
                ok("Revoked the Google sign-in for this run.")
            else:
                warn("Could not revoke the Google sign-in; remove it at https://myaccount.google.com/permissions")
        atexit.register(revoke)
    return creds


def google_preflight(drive, docs, quota_project: str | None):
    """One cheap Drive call and one Docs call, so quota-project problems surface before any writes."""
    from googleapiclient.errors import HttpError

    def fail(api: str, e: HttpError):
        body = e.content.decode(errors="replace") if isinstance(e.content, bytes) else str(e.content)
        hint = ""
        if any(s in body for s in ("SERVICE_DISABLED", "has not been used in project", "quota project",
                                   "accessNotConfigured")):
            hint = ("\n  The call was billed to a project without this API enabled. Pass --quota-project "
                    "<project-id> for a project where the Drive API and Docs API are enabled"
                    + (f" (currently {quota_project})." if quota_project else "."))
        sys.exit(f"Google {api} API check failed ({e.resp.status}): {body[:400]}{hint}")

    try:
        who = drive.about().get(fields="user(emailAddress)").execute()["user"]["emailAddress"]
    except HttpError as e:
        fail("Drive", e)
    try:
        docs.documents().get(documentId="preflight-nonexistent").execute()
    except HttpError as e:
        if e.resp.status not in (400, 404):   # 400/404 = API reachable, fake doc just doesn't exist
            fail("Docs", e)
    billing = f"API usage billed to project {quota_project}" if quota_project else "no quota project set"
    ok(f"Signed in to Google as {who}. Drive and Docs APIs reachable ({billing}).")


def check_parent(drive, value: str) -> str:
    from googleapiclient.errors import HttpError
    folder_id = drive_folder_id(value)
    try:
        f = drive.files().get(fileId=folder_id, supportsAllDrives=True,
                              fields="name,mimeType,capabilities(canAddChildren)").execute()
    except HttpError as e:
        sys.exit(f"Can't open Drive folder {folder_id} ({e.resp.status}). Check the link or ID.")
    if f["mimeType"] != FOLDER_MIME:
        sys.exit(f"{f['name']} is not a folder.")
    if not f.get("capabilities", {}).get("canAddChildren"):
        sys.exit(f"You can't add files to the folder {f['name']}.")
    say_link("✓ Target folder: ", f"https://drive.google.com/drive/folders/{folder_id}", f["name"], "green")
    return folder_id


def google_services(args):
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    # Credentials stay in memory for this run only.
    scopes = google_scopes(args)
    if args.google_auth == "gcloud":
        creds = gcloud_credentials(scopes, args.quota_project, args.keep_google_grant)
    else:
        if not Path(args.google_client_file).exists():
            sys.exit(f"Google OAuth client file not found: {args.google_client_file} (see README.md).")
        creds = InstalledAppFlow.from_client_secrets_file(args.google_client_file, scopes).run_local_server(port=0)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    docs = build("docs", "v1", credentials=creds, cache_discovery=False)
    google_preflight(drive, docs, args.quota_project)
    args.root_id = check_parent(drive, args.drive_parent) if args.drive_parent else "root"
    return drive, docs


def tab_fields(depth: int = MAX_DEPTH + 1) -> str:
    inner = ("tabProperties(tabId),documentTab(body(content(startIndex,endIndex,"
             "table(tableRows(tableCells(content(startIndex)))))))")
    s = inner
    for _ in range(depth):
        s = f"{inner},childTabs({s})"
    return f"tabs({s})"


def find_tab(tabs: list, tab_id: str):
    for t in tabs:
        if t["tabProperties"]["tabId"] == tab_id:
            return t
        hit = find_tab(t.get("childTabs", []), tab_id)
        if hit:
            return hit
    return None


class GoogleTarget:
    def __init__(self, drive, docs, state: dict, root_id: str = "root"):
        from googleapiclient.errors import HttpError
        self.HttpError = HttpError
        self.drive, self.docs, self.state, self.root_id = drive, docs, state, root_id
        self.writes, self.reads = Throttle(DOC_WRITES_PER_MIN), Throttle(DOC_READS_PER_MIN)

    def run(self, req, throttle: Throttle | None = None):
        for attempt in range(MAX_ATTEMPTS):
            if throttle:
                throttle.wait()
            try:
                return req.execute()
            except self.HttpError as e:
                if e.resp.status in (429, 500, 502, 503, 504) and attempt < MAX_ATTEMPTS - 1:
                    time.sleep(min(2 ** attempt + random.random(), 64))
                    continue
                raise

    def batch(self, doc_id: str, reqs: list) -> list:
        replies = []
        for i in range(0, len(reqs), FLUSH_AT):
            body = {"requests": reqs[i:i + FLUSH_AT]}
            res = self.run(self.docs.documents().batchUpdate(documentId=doc_id, body=body), self.writes)
            replies += res.get("replies", [])
        return replies

    def tab(self, doc_id: str, tab_id: str) -> dict:
        doc = self.run(self.docs.documents().get(documentId=doc_id, includeTabsContent=True,
                                                 fields=tab_fields()), self.reads)
        return find_tab(doc["tabs"], tab_id)

    def folder_key(self, parts: list) -> str:
        return f"{self.root_id}:" + "/".join(parts)

    def folder(self, parts: list) -> str:
        cache = self.state.setdefault("folders", {})
        parent = self.root_id
        for i in range(len(parts)):
            key = self.folder_key(parts[:i + 1])
            if key not in cache:
                body = {"name": parts[i], "mimeType": FOLDER_MIME, "parents": [parent]}
                cache[key] = self.run(self.drive.files().create(body=body, fields="id", supportsAllDrives=True))["id"]
                save_json(STATE_FILE, self.state)
            parent = cache[key]
        return parent

    def create_doc(self, title: str, folder_id: str) -> str:
        body = {"name": title, "mimeType": DOC_MIME, "parents": [folder_id]}
        return self.run(self.drive.files().create(body=body, fields="id", supportsAllDrives=True))["id"]

    def upload(self, path: Path, name: str, mime: str, folder_id: str) -> str:
        """Upload privately (no sharing) and return the file's Drive link."""
        from googleapiclient.http import MediaFileUpload
        media = MediaFileUpload(str(path), mimetype=mime, resumable=path.stat().st_size > 5_000_000)
        f = self.run(self.drive.files().create(body={"name": name, "parents": [folder_id]}, supportsAllDrives=True,
                                               media_body=media, fields="webViewLink"))
        return f["webViewLink"]

    def trash(self, file_id: str):
        try:
            self.run(self.drive.files().update(fileId=file_id, body={"trashed": True}, supportsAllDrives=True))
        except self.HttpError as e:   # e.g. already deleted by hand; a leftover is harmless, so carry on
            warn(f"Couldn't move {file_id} to trash ({e.resp.status}); delete it by hand if it's still there")


RESET_FIELDS = ("bold,italic,underline,strikethrough,smallCaps,link,weightedFontFamily,"
                "backgroundColor,foregroundColor,fontSize,baselineOffset")


def rgb(c: tuple) -> dict:
    return {"color": {"rgbColor": {"red": c[0], "green": c[1], "blue": c[2]}}}


def text_style(f: Fmt):
    ts = {}
    if f.bold:
        ts["bold"] = True
    if f.italic:
        ts["italic"] = True
    if f.underline:
        ts["underline"] = True
    if f.strike:
        ts["strikethrough"] = True
    if f.code:
        ts["weightedFontFamily"] = {"fontFamily": "Roboto Mono"}
    if f.link:
        ts["link"] = {"url": f.link}
    if f.highlight:
        ts["backgroundColor"] = rgb(f.highlight)
    if f.color:
        ts["foregroundColor"] = rgb(f.color)
    return ts, ",".join(ts)


class TabWriter:
    """Appends blocks to one document tab, tracking the insertion index locally."""

    def __init__(self, target: GoogleTarget, doc_id: str, tab_id: str, page_dir: Path,
                 resources: dict, assets, warn, log_image):
        self.t, self.doc_id, self.tab_id = target, doc_id, tab_id
        self.page_dir, self.resources, self.assets, self.warn = page_dir, resources, assets, warn
        self.log_image = log_image
        self.idx, self.reqs, self.n_images, self.last_text = 1, [], 0, ""

    def loc(self, i):
        return {"index": i, "tabId": self.tab_id}

    def rng(self, a, b):
        return {"startIndex": a, "endIndex": b, "tabId": self.tab_id}

    def flush(self):
        if self.reqs:
            reqs, self.reqs = self.reqs, []
            self.t.batch(self.doc_id, reqs)

    def resync(self):
        """Point idx at the tab's trailing empty paragraph (the end of the body)."""
        self.idx = self.t.tab(self.doc_id, self.tab_id)["documentTab"]["body"]["content"][-1]["endIndex"] - 1

    def styled_runs(self, start: int, runs: list):
        pos = start
        for r in runs:
            n = u16(r.text)
            ts, fields = text_style(r.fmt)
            if fields and n:
                self.reqs.append({"updateTextStyle": {"range": self.rng(pos, pos + n),
                                                      "textStyle": ts, "fields": fields}})
            pos += n

    def para(self, runs: list, style: str = "NORMAL_TEXT", tabs: int = 0):
        text = "\t" * tabs + "".join(r.text for r in runs) + "\n"
        start, end = self.idx, self.idx + u16(text)
        self.reqs += [
            {"insertText": {"location": self.loc(start), "text": text}},
            {"updateTextStyle": {"range": self.rng(start, end), "textStyle": {}, "fields": RESET_FIELDS}},
            {"updateParagraphStyle": {"range": self.rng(start, end),
                                      "paragraphStyle": {"namedStyleType": style}, "fields": "namedStyleType"}},
        ]
        self.styled_runs(start + tabs, runs)
        self.idx = end
        self.last_text = text.strip() or self.last_text
        if len(self.reqs) >= FLUSH_AT:
            self.flush()

    def list_group(self, paras: list):
        # createParagraphBullets reads leading tabs as nesting depth and deletes them,
        # so it must run in the same batch, after the text styles are applied.
        start, tabs = self.idx, 0
        for p in paras:
            self.para(p.runs, "NORMAL_TEXT", tabs=p.level)
            tabs += p.level
        preset = "NUMBERED_DECIMAL_ALPHA_ROMAN" if paras[0].list_kind == "number" else "BULLET_DISC_CIRCLE_SQUARE"
        self.reqs.append({"createParagraphBullets": {"range": self.rng(start, self.idx), "bulletPreset": preset}})
        self.idx -= tabs
        self.flush()

    def write(self, blocks: list):
        i = 0
        while i < len(blocks):
            b = blocks[i]
            if isinstance(b, Para) and b.list_kind:
                j = i + 1
                while (j < len(blocks) and isinstance(blocks[j], Para) and blocks[j].list_kind
                       and (blocks[j].level > 0 or blocks[j].list_kind == b.list_kind)):
                    j += 1
                self.list_group(blocks[i:j])
                i = j
                continue
            if isinstance(b, Para):
                self.para(b.runs, b.style)
            elif isinstance(b, Table):
                self.table(b)
            elif isinstance(b, Image):
                self.image(b)
            elif isinstance(b, Attachment):
                self.attachment(b)
            i += 1

    def table(self, t: Table):
        self.flush()
        nrows, ncols = len(t.rows), len(t.rows[0])
        try:
            self.t.batch(self.doc_id, [{"insertTable": {"rows": nrows, "columns": ncols, "location": self.loc(self.idx)}}])
        except self.t.HttpError as e:
            self.warn(f"table {nrows}x{ncols} inserted as text ({e.resp.status})")
            for row in t.rows:
                self.para([r for k, cell in enumerate(row) for r in ([Run(" | ")] if k else []) + cell])
            return
        body = self.t.tab(self.doc_id, self.tab_id)["documentTab"]["body"]["content"]
        el = next(e for e in body if "table" in e and e["startIndex"] >= self.idx)
        cells = [(r, c, cell["content"][0]["startIndex"])
                 for r, row in enumerate(el["table"]["tableRows"])
                 for c, cell in enumerate(row["tableCells"])]
        reqs = []
        for r, c, start in reversed(cells):   # fill last cell first so earlier indices stay valid
            runs = t.rows[r][c]
            text = "".join(x.text for x in runs)
            if not text:
                continue
            end = start + u16(text)
            reqs += [
                {"insertText": {"location": self.loc(start), "text": text}},
                {"updateTextStyle": {"range": self.rng(start, end), "textStyle": {}, "fields": RESET_FIELDS}},
            ]
            if t.header and r == 0:
                reqs.append({"updateTextStyle": {"range": self.rng(start, end), "textStyle": {"bold": True}, "fields": "bold"}})
            saved, self.reqs = self.reqs, []
            self.styled_runs(start, runs)
            reqs += self.reqs
            self.reqs = saved
        for r, c, rs, cs in t.merges:
            reqs.append({"mergeTableCells": {"tableRange": {
                "tableCellLocation": {"tableStartLocation": self.loc(el["startIndex"]), "rowIndex": r, "columnIndex": c},
                "rowSpan": rs, "columnSpan": cs}}})
        if reqs:
            self.t.batch(self.doc_id, reqs)
        self.resync()
        self.last_text = "table: " + " | ".join("".join(r.text for r in cell) for cell in t.rows[0])

    def resource(self, src: str):
        info = self.resources.get(src)
        if not info:
            return None
        path = self.page_dir / "resources" / info["file"]
        return (path, info["mime"]) if path.exists() else None

    def image(self, img: Image):
        """Images never leave this machine: write a placeholder and log where it is."""
        self.n_images += 1
        found = self.resource(img.src)
        before = self.last_text
        label = f"[Image {self.n_images}: paste here]"
        self.para([Run("🖼 "), Run(label, Fmt(bold=True, highlight=PLACEHOLDER_HIGHLIGHT))])
        self.last_text = before
        if not found:
            self.warn(f"image {self.n_images} missing from export")
        self.log_image({"image": self.n_images, "placeholder": label, "after": before[:80],
                        "file": str(found[0]) if found else None})

    def attachment(self, a: Attachment):
        found = self.resource(a.src)
        if not found:
            self.warn(f"attachment missing from export: {a.name}")
            self.para([Run(f"📎 {a.name} (missing from export)", Fmt(italic=True))])
            return
        path, mime = found
        link = self.t.upload(path, a.name, mime, self.assets())
        self.para([Run("📎 "), Run(a.name, Fmt(link=link))])


def page_tree(pages: list):
    """(parent index, depth) per page from OneNote levels, clamped to Docs' 3 tab levels."""
    parents, depth = [], []
    for i, p in enumerate(pages):
        lvl, j = p.get("level", 0), i - 1
        while j >= 0 and pages[j].get("level", 0) >= lvl:
            j -= 1
        while j >= 0 and depth[j] >= MAX_DEPTH:
            j = parents[j] if parents[j] is not None else -1
        parent = j if lvl > 0 and j >= 0 else None
        parents.append(parent)
        depth.append(0 if parent is None else depth[parent] + 1)
    return parents, depth


def split_pages(pages: list, limit: int = MAX_TABS) -> list:
    """Pack page trees (a top-level page + its subpages) into docs of at most `limit` tabs."""
    trees = []
    for p in pages:
        if p.get("level", 0) == 0 or not trees:
            trees.append([p])
        else:
            trees[-1].append(p)
    chunks, cur = [], []
    for tree in trees:
        if cur and len(cur) + len(tree) > limit:
            chunks.append(cur)
            cur = []
        cur += tree
        while len(cur) > limit:   # one tree bigger than a whole doc: hard split
            chunks.append(cur[:limit])
            cur = cur[limit:]
    if cur:
        chunks.append(cur)
    return chunks


def tab_title(page: dict) -> str:
    return (page.get("title") or "Untitled page").strip()[:100] or "Untitled page"


def unique_titles(pages: list, parents: list) -> list:
    """Docs rejects duplicate tab titles anywhere in a document, even under different parents
    (e.g. a "03-13-26" subpage under each person's 1:1 page). Later duplicates get their
    parent page's title, else their created date, then a counter."""
    seen, out = set(), []
    for p, parent in zip(pages, parents):
        base = tab_title(p)
        title = base
        if title.casefold() in seen and parent is not None:
            title = f"{base[:60]} ({tab_title(pages[parent])[:36]})"
        if title.casefold() in seen and ymd(p.get("created")):
            title = f"{base[:86]} ({ymd(p['created'])})"
        n = 2
        while title.casefold() in seen:
            title = f"{base[:92]} ({n})"
            n += 1
        seen.add(title.casefold())
        out.append(title)
    return out


def create_tabs(t: GoogleTarget, doc_id: str, pages: list) -> list:
    parents, depth = page_tree(pages)
    doc =t.run(t.docs.documents().get(documentId=doc_id, includeTabsContent=True,
                                       fields="tabs(tabProperties(tabId))"), t.reads)
    ids = [None] * len(pages)
    ids[0] = doc["tabs"][0]["tabProperties"]["tabId"]
    titles = unique_titles(pages, parents)
    sibling, seen = [], {}
    for p in parents:
        sibling.append(seen.get(p, 0))
        seen[p] = sibling[-1] + 1
    for d in range(MAX_DEPTH + 1):
        reqs, order = [], []
        if d == 0:
            reqs.append({"updateDocumentTabProperties": {
                "tabProperties": {"tabId": ids[0], "title": titles[0]}, "fields": "title"}})
            order.append(None)
        for i in range(1, len(pages)):
            if depth[i] == d:
                props = {"title": titles[i], "index": sibling[i]}
                if parents[i] is not None:
                    props["parentTabId"] = ids[parents[i]]
                reqs.append({"addDocumentTab": {"tabProperties": props}})
                order.append(i)
        if not reqs:
            continue
        for i, reply in zip(order, t.batch(doc_id, reqs)):
            if i is not None:
                ids[i] = reply["addDocumentTab"]["tabProperties"]["tabId"]
    return ids


def page_blocks(page: dict, header: bool):
    pdir = EXPORT_DIR / page["dir"]
    head = []
    if header:
        head = [Para([Run(page["title"])], "TITLE"),
                Para([Run(f"Created {ymd(page['created'])} · Modified {ymd(page['modified'])} · from OneNote",
                          Fmt(italic=True, color=(.45, .45, .45)))])]
    if page.get("error") or not (pdir / "content.html").exists():
        return head + [Para([Run(f"⚠ This page failed to export from OneNote: {page.get('error', 'missing')}",
                                 Fmt(italic=True))])], pdir, {}
    resources = load_json(pdir / "resources.json", {})
    return head + html_to_blocks((pdir / "content.html").read_text(encoding=UTF8)), pdir, resources


def import_section(t: GoogleTarget, sec: dict, args, state: dict, bar: Progress, task):
    folder_path = ([args.drive_folder] if args.drive_folder else []) + sec["path"][:-1]
    folder_id = t.folder(folder_path)
    rec = state["sections"][sec["id"]] = {"name": " / ".join(sec["path"]), "status": "in_progress",
                                          "docs": [], "warnings": [], "images": []}
    save_json(STATE_FILE, state)
    assets_id = None

    def assets():
        nonlocal assets_id
        if not assets_id:
            rec["assets_key"] = t.folder_key(folder_path + [f"{sec['name']} files"])
            assets_id = t.folder(folder_path + [f"{sec['name']} files"])
        return assets_id

    # An empty section still gets a doc; the stand-in page (no id) only names its one tab.
    pages = sec["pages"] or [{"id": "", "title": "Empty section", "level": 0, "dir": "", "created": "", "modified": ""}]
    chunks = split_pages(pages)
    for n, chunk in enumerate(chunks, 1):
        title = sec["name"] if len(chunks) == 1 else f"{sec['name']} (part {n})"
        doc_id = t.create_doc(title, folder_id)
        rec["docs"].append(doc_id)
        save_json(STATE_FILE, state)
        tab_ids = create_tabs(t, doc_id, chunk)
        for page, tab_id in zip(chunk, tab_ids):
            if not page["id"]:
                continue
            bar.update(task, description=f"{clip(sec['name'], 24)} · {clip(page['title'])}")
            record_warning = lambda msg, p=page: rec["warnings"].append(f"{p['title']}: {msg}")
            tab_url = f"https://docs.google.com/document/d/{doc_id}/edit?tab={tab_id}"
            log_image = lambda item, p=page, u=tab_url: rec["images"].append({"page": p["title"], "url": u, **item})
            blocks, pdir, resources = page_blocks(page, not args.no_header)
            w = TabWriter(t, doc_id, tab_id, pdir, resources, assets, record_warning, log_image)
            try:
                w.write(blocks)
                w.flush()
            except Exception as e:
                record_warning(f"conversion stopped partway: {str(e)[:300]}")
                warn(f"{page['title']}: {str(e)[:200]}")
                try:
                    w.reqs = []
                    w.resync()
                    w.para([Run("⚠ The rest of this page could not be converted. See import_state.json.",
                                Fmt(italic=True))])
                    w.flush()
                except Exception:
                    pass   # best effort: the failure is already recorded in the section's warnings
            bar.advance(task)
        say_link(f"✓ {' / '.join(sec['path'][:-1] + [title])} → ",
                 f"https://docs.google.com/document/d/{doc_id}/edit", style="green")
    rec["status"] = "done"
    save_json(STATE_FILE, state)
    write_images_todo(state)


def write_images_todo(state: dict):
    """Checklist of every image placeholder: where it is in Google and where the file is locally."""
    secs = [s for s in state["sections"].values() if s.get("images")]
    total = sum(len(s["images"]) for s in secs)
    lines = ["# Images to copy by hand", "",
             f"{total} images. In each tab, the placeholder is highlighted yellow; search the tab for it, "
             "then drag the local file in (or Insert > Image > Upload from computer) and delete the placeholder.", ""]
    for s in secs:
        lines += [f"## {s['name']}", ""]
        for im in s["images"]:
            where = f' after “{im["after"]}”' if im.get("after") else " at the top of the page"
            file = f"[local file]({Path(im['file']).as_uri()})" if im.get("file") else "**local file missing from export**"
            lines.append(f"- [ ] **{im['page']}**: `{im['placeholder']}`{where}. "
                         f"[Open tab]({im['url']}) · {file}")
        lines.append("")
    IMAGES_TODO.write_text("\n".join(lines), encoding=UTF8)


def dry_run(sections: list):
    total = {"pages": 0, "tables": 0, "images": 0, "files": 0, "writes": 0}
    for sec in sections:
        counts = {"pages": len(sec["pages"]), "tables": 0, "images": 0, "files": 0}
        for page in sec["pages"]:
            blocks, _, _ = page_blocks(page, True)
            counts["tables"] += sum(isinstance(b, Table) for b in blocks)
            counts["images"] += sum(isinstance(b, Image) for b in blocks)
            counts["files"] += sum(isinstance(b, Attachment) for b in blocks)
        docs = len(split_pages(sec["pages"])) if sec["pages"] else 1
        # 1 text flush per page + 2 per table, plus tab setup per doc
        counts["writes"] = counts["pages"] + 2 * counts["tables"] + 3 * docs
        for k in total:
            total[k] += counts[k]
        split = f", {docs} docs" if docs > 1 else ""
        say(f"{' / '.join(sec['path'])}: {counts['pages']} pages, {counts['tables']} tables, "
            f"{counts['images']} images, {counts['files']} attachments{split}")
    say(f"\nTotal: {total['pages']} pages, {total['tables']} tables, {total['images']} images, "
        f"{total['files']} attachments", "bold")
    say(f"Estimated Docs API writes: ~{total['writes']} (≈{total['writes'] / DOC_WRITES_PER_MIN:.0f} min at quota)")


def import_cmd(args):
    manifest = load_json(MANIFEST)
    if not manifest:
        sys.exit(f"No export found at {MANIFEST}. Run the export command first.")
    sections = [s for s in manifest["sections"] if matches(s["path"], args.only)]
    if args.dry_run:
        dry_run(sections)
        return
    state = load_json(STATE_FILE, {"folders": {}, "sections": {}})
    t = GoogleTarget(*google_services(args), state, args.root_id)
    todo = [s for s in sections if args.force or state["sections"].get(s["id"], {}).get("status") != "done"]
    if len(todo) < len(sections):
        say(f"Skipping {plural(len(sections) - len(todo), 'section')} already imported "
            "(use --force to redo them)", "dim")
    with progress_bar() as bar:
        task = bar.add_task("Starting", total=sum(len(s["pages"]) for s in todo))
        for sec in todo:
            prev = state["sections"].get(sec["id"])
            for doc_id in (prev or {}).get("docs", []):   # partial or forced: start the section over
                t.trash(doc_id)
            if prev and prev.get("assets_key") in state["folders"]:
                t.trash(state["folders"].pop(prev["assets_key"]))
            import_section(t, sec, args, state, bar, task)
    status_cmd(args)


def status_cmd(args):
    state = load_json(STATE_FILE, {"sections": {}})
    manifest = load_json(MANIFEST, {"sections": []})
    by_status = {}
    for s in state["sections"].values():
        by_status[s["status"]] = by_status.get(s["status"], 0) + 1
    incomplete = by_status.get("in_progress", 0)
    say(f"\nSections exported: {len(manifest['sections'])}  imported: {by_status.get('done', 0)}  "
        f"incomplete: {incomplete}", "bold yellow" if incomplete else "bold")
    images = sum(len(s.get("images", [])) for s in state["sections"].values())
    if images:
        say_link(f"Images to paste by hand: {images} — checklist: ", IMAGES_TODO.as_uri(), str(IMAGES_TODO))
    for s in state["sections"].values():
        if s["warnings"]:
            say(f"\n{s['name']} — {len(s['warnings'])} warnings", "bold yellow")
            for w in s["warnings"][:10]:
                say(f"  • {w}", "yellow")
            if len(s["warnings"]) > 10:
                say(f"  … {len(s['warnings']) - 10} more in {STATE_FILE.name}", "yellow")


def convert_cmd(args):
    for b in html_to_blocks(Path(args.html).read_text(encoding=UTF8)):
        if isinstance(b, Para):
            tag = f"{b.style}" + (f" {b.list_kind}[{b.level}]" if b.list_kind else "")
            text = "".join(f"[{r.text}]" if r.fmt != Fmt() else r.text for r in b.runs)
            print(f"{tag:<28} {text}")
        else:
            print(b)


def format_options(catalog: dict, col: int = 30) -> str:
    """Render every command's optional flags for the top-level -h (argparse lists only commands)."""
    import shutil
    import textwrap
    width = max(col + 30, min(100, shutil.get_terminal_size().columns - 2))   # follow the terminal, like argparse
    out = []
    for section, entries in catalog.items():
        out.append(f"{section}:")
        for label, text in entries:
            body = textwrap.wrap(text, width - col, break_on_hyphens=False) or [""]   # keep --flag-names whole
            if len(label) + 2 <= col - 2:
                out.append(f"  {label:<{col - 2}}{body[0]}")
            else:   # long flag: help starts on the next line, as argparse does
                out += [f"  {label}", " " * col + body[0]]
            out += [" " * col + line for line in body[1:]]
        out.append("")
    out.append("Run 'onenote2gdocs.py <command> -h' for one command's usage.")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description="Migrate OneNote notebooks to Google Docs (section → Doc, page → tab).",
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True, title="commands")
    catalog = {}   # top-level -h section → [(flag label, help)], filled as flags are defined

    def opt(target, section: str, *flags, **kw):
        action = target.add_argument(*flags, **kw)
        label = ", ".join(action.option_strings)
        if action.nargs != 0:   # takes a value (store_true flags have nargs 0)
            choices = "{" + ",".join(action.choices) + "}" if action.choices else None
            label += " " + (action.metavar or choices or action.dest.upper())
        entries = catalog.setdefault(section, [])
        if (label, action.help) not in entries:   # Google options are shared by two commands
            entries.append((label, action.help))

    ex = sub.add_parser("export", help="Download notebooks from Microsoft Graph into ./export")
    ex_opts = "export options"
    opt(ex, ex_opts, "--only", action="append", metavar="PATH",
        help='Only sections under this path, e.g. "Work" or "Work/Projects" (repeatable)')
    opt(ex, ex_opts, "--device-code", action="store_true",
        help="Sign in with a device code instead of a browser window")
    opt(ex, ex_opts, "--ms-tenant", metavar="TENANT", default=os.environ.get("ONENOTE_TENANT", "organizations"),
        help="Entra tenant ID or domain (default: organizations; env ONENOTE_TENANT)")
    opt(ex, ex_opts, "--ms-client-id", metavar="ID", default=os.environ.get("ONENOTE_CLIENT_ID", MS_CLI_CLIENT_ID),
        help="Your own Entra public client ID instead of Microsoft Graph Command Line Tools (env ONENOTE_CLIENT_ID)")
    ex.set_defaults(fn=export_cmd)

    im = sub.add_parser("import", help="Create Google Docs from ./export")
    im_opts = "import options"
    opt(im, im_opts, "--dry-run", action="store_true", help="Parse everything and print counts; no Google calls")
    opt(im, im_opts, "--only", action="append", metavar="PATH",
        help="Only sections under this path (repeatable)")
    opt(im, im_opts, "--drive-folder", metavar="NAME", default="OneNote Import",
        help='Folder created to hold the notebooks (default: "OneNote Import"); '
             'pass "" to put notebook folders directly in the target')
    opt(im, im_opts, "--force", action="store_true",
        help="Re-import sections already marked done (trashes their old docs)")
    opt(im, im_opts, "--no-header", action="store_true",
        help="Skip the title and created/modified line at the top of each tab")
    im.set_defaults(fn=import_cmd)

    ck = sub.add_parser("check-google", help="Sign in to Google and test Drive + Docs access; writes nothing")
    ck.set_defaults(fn=google_services)

    g_opts = "Google sign-in options (import, check-google)"
    for p in (im, ck):
        g = p.add_argument_group("Google sign-in")
        opt(g, g_opts, "--quota-project", metavar="PROJECT_ID", default=os.environ.get("GOOGLE_QUOTA_PROJECT"),
            help="Cloud project with the Drive and Docs APIs enabled, billed for API usage "
                 "(env GOOGLE_QUOTA_PROJECT)")
        opt(g, g_opts, "--drive-parent", metavar="FOLDER", default=os.environ.get("GOOGLE_DRIVE_PARENT"),
            help="Existing Drive folder (link or ID) to import into instead of My Drive. Requests full "
                 "Drive access for the run instead of drive.file (env GOOGLE_DRIVE_PARENT)")
        opt(g, g_opts, "--keep-google-grant", action="store_true",
            help="Don't revoke the gcloud sign-in when the run ends")
        opt(g, g_opts, "--google-auth", choices=["gcloud", "client"], default="gcloud",
            help="gcloud: gcloud's built-in sign-in (default); client: your own OAuth client JSON")
        opt(g, g_opts, "--google-client-file", metavar="PATH",
            default=os.environ.get("GOOGLE_CLIENT_FILE", str(BASE / "credentials.json")),
            help="OAuth desktop client JSON for --google-auth client (default: ./credentials.json; "
                 "env GOOGLE_CLIENT_FILE)")

    sub.add_parser("status", help="Summarize import progress, warnings, and images to paste").set_defaults(fn=status_cmd)

    cv = sub.add_parser("convert", help="Debug: print the block model for one exported content.html")
    cv.add_argument("html", help="Path to an exported content.html")
    cv.set_defaults(fn=convert_cmd)

    ap.epilog = format_options(catalog)

    # Progress lines use ✓, →, and page titles with emoji. Output piped to a file uses the locale
    # encoding; substitute instead of crashing.
    for stream in (sys.stdout, sys.stderr):
        if (stream.encoding or "").lower().replace("-", "") != "utf8":
            stream.reconfigure(errors="replace")

    args = ap.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
