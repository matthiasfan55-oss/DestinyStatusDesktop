# DestinyStatusDesktop

Repository for the Windows Kick status square build.

## Repo layout

- `src/KickStatusApp.cs`: WinForms app source
- `src/EditKickStatusConfig.ps1`: settings editor
- `src/config.json`: clean fallback defaults for first launch
- `build.ps1`: builds the app, creates the distributable package, and regenerates `update.json`
- `dist/`: built executable and update package that installed copies download

## Local settings

Installed copies store the real user configuration under `%APPDATA%\KickStatusAppData`.
That keeps saved channels and image paths outside the repo and outside the install folder.

## Auto-update flow

The app checks:

- `https://raw.githubusercontent.com/matthiasfan55-oss/DestinyStatusDesktop/main/update.json`

If `update.json` contains a newer version than the running build, the app downloads:

- `dist/KickStatusSquare-package.zip`

It then replaces the local install folder and restarts itself.

## Releasing an update

1. Make source changes.
2. Bump `version.json`.
3. Run `Build Destiny Status Desktop.cmd`.
4. Commit and push:
   - source changes
   - `dist/KickStatusSquare.exe`
   - `dist/KickStatusSquare-package.zip`
   - `update.json`

Once those files are on `main`, installed copies will update themselves automatically.
