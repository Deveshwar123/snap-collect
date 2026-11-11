# snap-collect

A small mobile-first web form that collects finger-snap recordings from volunteers (shared via an Instagram bio link). Static page on GitHub Pages (`https://deveshwar123.github.io/snap-collect/`) plus a Google Apps Script backend that saves clips to Drive and logs one row per submission to a Sheet.

- `index.html` is self-contained (no frameworks, no external scripts).
- `apps-script/Code.gs` is the backend.

## Deploy the Apps Script

1. Go to script.google.com, create a New project, paste in `apps-script/Code.gs`.
2. Select the `setup` function and click Run. Approve the authorization prompt. The log prints the Drive folder and Sheet URLs (created once, IDs stored in script properties).
3. Deploy > New deployment > type: Web app. Execute as: **Me**. Who has access: **Anyone**. Deploy and copy the `/exec` URL.
4. After any later code edit, use Deploy > Manage deployments > edit > New version, or the live URL keeps running the old code.

## Connect the page

In `index.html`, replace the placeholder in the line near the top of the script:

```js
const ENDPOINT = "PASTE_APPS_SCRIPT_WEB_APP_URL_HERE";
```

Then publish the folder via GitHub Pages.

## Test

- Health check: open the `/exec` URL in a browser; you should see `{"ok":true,"service":"snap-collect"}`.
- Locally: `python3 -m http.server` in this folder, open `http://localhost:8000` (microphone needs https or localhost).
- Fill in name, tick both boxes, record 3 snaps (or use the upload fallback), submit. Check the Sheet row and the per-submission subfolder in Drive.
- Test on iPhone Safari and Android Chrome, and open the link inside Instagram to confirm the in-app browser banner and upload fallback appear.

## Notes

- Clips are limited to 1.5 MB each; uploaded files to 15 s. The server re-checks size and clip count (max 6).
- The honeypot field `website` is hidden; if filled, the page pretends success and sends nothing.
- Anyone with the URL can post, so the Drive folder will grow only as fast as people submit; delete rows/folders on request (consent text promises deletion by DM).
