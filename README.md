# sxl-releases

Public release feed and hosted web app for **sxl** (the Windows-first Excel
agent with a durable, reviewable mutation ledger).

- **Update feed:** [`latest.json`](https://bwestlund17.github.io/sxl-releases/latest.json)
- **Web app:** <https://bwestlund17.github.io/sxl-releases/>
- The desktop app checks the feed on startup (override with `SXL_UPDATE_FEED_URL`).
- Installers are published under this repo's **Releases**.

Source of truth for the web app is `packages/platform-server/public/` in the
private `shortcut-xl` repository; this repo only hosts the built static files and
the update manifest.
