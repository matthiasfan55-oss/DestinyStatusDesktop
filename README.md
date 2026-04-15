# DestinyStatusDesktop

Small Windows desktop status app for monitoring configured Kick channels and switching the indicator image based on which channel group is online.

## Repo layout

- `src/DestinyStatusDesktop.cs`: WinForms app source
- `src/EditDestinyStatusConfig.ps1`: settings editor
- `src/config.json`: fallback default config for first launch
- `build.ps1`: builds the app, creates the distributable package, and regenerates `update.json`
- `dist/`: built executable and update package that the app downloads for self-updates

## Auto-update flow

The app checks:

- `https://raw.githubusercontent.com/matthiasfan55-oss/DestinyStatusDesktop/main/update.json`

If `update.json` contains a newer version than the running build, the app downloads:

- `dist/DestinyStatusDesktop-package.zip`

It then replaces the local install folder and restarts itself.

## Releasing an update

1. Make source changes.
2. Bump `version.json`.
3. Run `Build Destiny Status Desktop.cmd`.
4. Commit and push:
   - source changes
   - `dist/DestinyStatusDesktop.exe`
   - `dist/DestinyStatusDesktop-package.zip`
   - `update.json`

Once those files are on `main`, installed copies will update themselves automatically.
