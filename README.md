# OneNote to Google Docs

`onenote2gdocs.py` copies your OneNote notebooks into Google Drive:

| OneNote                  | Google                                  |
| ------------------------ | --------------------------------------- |
| Notebook, section group  | Drive folder                            |
| Section                  | Google Doc                              |
| Page                     | Document tab                            |
| Subpage (level 1 and 2)  | Child tab                               |
| Images                   | Highlighted placeholder, listed in `images_todo.md` |
| File attachments         | Private files in `<Section> files`, linked from the page |

The script runs in two phases. `export` downloads everything from Microsoft Graph into `./export/`. `import` builds the Google Docs from that local copy. The local copy also works as an offline backup, and you can rerun `import` without downloading again.

This is intended to serve as a migration aid from OneNote to Google Docs, and is not written as a bilateral sync tool.

## Privacy

- **Images:** images never leave your computer. Each image becomes a yellow `[Image N: paste here]` placeholder in its tab. `images_todo.md` lists every placeholder with the page, the text it follows, a link to the tab, and a link to the image file in `./export/`.

    Images are manual because the Docs API inserts an image only from a public URL, which Google fetches at insert time. Automating it would mean sharing each image as "anyone with the link," even if only briefly, so the script leaves image insertion to you.
- **Attachments:** attachments are uploaded to your Drive with no sharing.
- **Tokens:** sign-in tokens stay in memory for one run, so you sign in again each run. Neither side needs an app registration.

## Before you begin

You need [uv](https://docs.astral.sh/uv/). It installs Python 3.10+ and the script's dependencies on first run. You also need the [Google Cloud CLI](https://cloud.google.com/sdk/docs/install) (`gcloud`).

The script runs on macOS and Windows. The examples use macOS and Linux shell syntax. In PowerShell, set environment variables with `$env:NAME = "value"` instead of `export NAME=value`.
### Microsoft: no setup needed

By default the script signs in through **Microsoft Graph Command Line Tools**, Microsoft's own public client. It's the same one `Connect-MgGraph` in Microsoft Graph PowerShell uses, so you don't register an app. On first sign-in, Microsoft asks you to let it read your notebooks (`Notes.Read`). If your tenant blocks user consent, an admin approves it once.

That consent shows up in the Entra admin center under **Enterprise applications** > **Microsoft Graph Command Line Tools**. To revoke it after the migration, open that entry and go to **Permissions**.

To use your own app registration instead, set `ONENOTE_CLIENT_ID`.

### Google: gcloud sign-in, no OAuth client

The script signs in through gcloud's built-in Application Default Credentials client, which shows as **Google Auth Library** on the consent screen. You don't create an OAuth client or consent screen.

The script runs `gcloud auth application-default login` against a temporary `CLOUDSDK_CONFIG` folder. It deletes that folder as soon as it reads the token into memory and revokes the sign-in when the run ends. Your normal gcloud configuration (`~/.config/gcloud`, or `%APPDATA%\gcloud` on Windows) is never read or changed.

Revoking may also sign out any other Application Default Credentials session for your account. If that happens, run `gcloud auth application-default login` again. To skip revoking, add `--keep-google-grant`.

Requested scopes are `drive.file` (only files the script creates), `documents`, and gcloud's required `openid`, `userinfo.email`, and `cloud-platform`.

1. If Google shows **Access blocked**, trust the client in the [Admin console](https://admin.google.com): go to **Security** > **Access and data control** > **API controls** > **Manage third-party app access** > **Configure new app**, search for client ID `764086051850-6qr4p6gpi6hn506pt8ejuq83di341hur.apps.googleusercontent.com`, and set it to **Trusted**.
1. Pick a quota project. Calls through gcloud's shared client are billed to a Google-owned project where the Drive and Docs APIs are off, so the script sends API usage to a project of yours. The project needs both APIs enabled. It doesn't need an OAuth client or consent screen.

    ```sh
    gcloud services enable drive.googleapis.com docs.googleapis.com --project <project-id>
    export GOOGLE_QUOTA_PROJECT=<project-id>
    ```

    Your account needs `serviceusage.services.use` on that project. The Owner, Editor, and Service Usage Consumer roles include it.

1. Test sign-in and API access. This step writes nothing to Drive.

    ```sh
    uv run onenote2gdocs.py check-google
    ```

If gcloud sign-in can't work in your tenant, `--google-auth client --google-client-file credentials.json` uses your own Desktop OAuth client instead.

## Run the migration

1. Export one notebook first. A browser window opens for Microsoft sign-in. If your browser can't reach it, add `--device-code`.

    ```sh
    uv run onenote2gdocs.py export --only "Notebook"
    ```

1. Check what the import will do. This step makes no Google calls.

    ```sh
    uv run onenote2gdocs.py import --dry-run
    ```

1. Import one section and review it in Drive:

    ```sh
    uv run onenote2gdocs.py import --only "Notebook/SpecialStuff"
    ```

1. When the result looks right, export and import everything:

    ```sh
    uv run onenote2gdocs.py export
    uv run onenote2gdocs.py import
    ```

1. Open `images_todo.md` and work through the checklist. For each image, open the tab, search the tab for the placeholder, drag the local file in, and delete the placeholder.

To check progress, per-page warnings, and the image count, run `uv run onenote2gdocs.py status`.

## Choose where the import goes

By default, the import creates `OneNote Import` at the top of My Drive and puts one folder per notebook inside it.

- **Rename the wrapper folder:** `--drive-folder "Archive - OneNote"`.
- **Skip the wrapper folder:** `--drive-folder ""` puts the notebook folders directly in the target.
- **Import into an existing folder:** pass its link or ID with `--drive-parent`. Shared drive folders work too.

    ```sh
    uv run onenote2gdocs.py import --drive-parent "https://drive.google.com/drive/folders/<id>"
    ```

    The default `drive.file` scope can only reach files the script created, so `--drive-parent` asks for full Drive access (`drive`) for that run instead. The sign-in is still in memory only and revoked at the end. To keep `drive.file`, import with the default and then move `OneNote Import` into place in Drive. The move works even partway through a migration.

`check-google --drive-parent <link>` confirms that the folder exists and that you can add files to it.

Use the same `--drive-parent` and `--drive-folder` on every run of one migration. Changing either one starts a new folder tree.

## Resume and rerun

- If `export` stops partway, run it again. It saves progress after each section and downloads again only pages whose last-modified time changed.
- If `import` stops partway, run it again. It skips finished sections and rebuilds any half-written section from scratch, moving the partial doc to the trash.
- To rebuild sections that are already imported, for example after re-exporting edited notes, add `--force`.

## Limits

- **Speed:** the Docs API allows 60 writes per minute per user. The script uses 55, about one page per second plus two extra writes per table. `--dry-run` prints an estimate.
- **100 tabs per doc:** sections with more pages split into `Section (part 2)` and so on, always between top-level pages.
- **Not carried over:** handwriting and ink (Graph doesn't return them), links between OneNote pages (kept as plain text), page layout positions (outlines are stacked top to bottom), and custom fonts and sizes. Headings, bold, italic, underline, strikethrough, highlights, text color, links, monospace, lists, to-do tags, and tables are kept.
- **Notebook location:** `export` reads notebooks in your own OneDrive (`/me/onenote`). Notebooks stored in a SharePoint site or shared with you by someone else aren't included.

## Files this script writes

| Path                | Contents                                              |
| ------------------- | ----------------------------------------------------- |
| `export/`           | Raw page HTML, images, attachments, `manifest.json`   |
| `import_state.json` | Drive folder and Doc IDs, import progress, warnings, image log |
| `images_todo.md`    | Checklist of image placeholders to fill by hand       |

The script writes no credentials. The gcloud temporary folder exists only for the seconds between sign-in and reading the token. If a run is killed in that window, the next run deletes the folder.

## License

[MIT](LICENSE). The script is provided as is, with no warranty. You use it at your own risk.
