# Building YourOfficeSuite

## What you need

- Git
- Node.js 24 or newer (the editor engine's build needs it; the app itself runs on the Electron version in `package.json`)
- Python 3 with `python-docx`, `openpyxl`, `python-pptx` and `Pillow` if you want to regenerate the templates
- About 3 GB of free disk space
- Linux, macOS, or Windows with WSL or Git Bash

## 1. Get the code and the engine

```bash
git clone https://github.com/aglr-afk/YourOfficeSuite.git
cd YourOfficeSuite
npm install
npm run engine
```

`npm run engine` runs `scripts/get-engine.sh`. It clones the editor project at a pinned commit into `.engine-src/`, builds it, copies the result to `engine/`, and runs `scripts/patch-engine.py` to apply this project's changes. The changes are listed in [NOTICE.md](../NOTICE.md).

## 2. Run from source

```bash
npm start
```

## 3. Make the Windows package

```bash
npm run pack:win
cp shell/licenses.txt release/win-unpacked/Licences.txt
```

This creates `release/win-unpacked/` with `YourOfficeSuite.exe`. The `afterPack` hook sets the icon and version details on the exe without needing Wine.

## 4. Make the installer

The installer is an [NSIS](https://nsis.sourceforge.io/) script, `build/installer.nsi`. It packs `release/win-unpacked/` into `release/YourOfficeSuite-Setup-<version>.exe`.

On Windows, install NSIS 3 and run:

```bat
makensis build\installer.nsi
```

On Linux, any NSIS 3 `makensis` works:

```bash
(cd build && makensis installer.nsi)
```

Compression takes several minutes. The version number is set in two places: `package.json` and the `VER` line at the top of `build/installer.nsi`.

### Signing

The published installer is not code-signed. If you have a code-signing certificate, sign `release/win-unpacked/YourOfficeSuite.exe` before step 4 and the setup file after it, for example with `signtool` on Windows.

## Regenerating the templates

```bash
npm run templates
```

This rewrites the files under `shell/templates/`. The preview pictures in `shell/templates/thumbs/` were made by converting each template to PDF inside the app and rendering the first page.

## Running the tests

The suite needs Playwright for Node and a few Python packages:

```bash
npm install --no-save playwright
pip install python-docx openpyxl python-pptx pypdf pikepdf pypdfium2 reportlab Pillow
```

Then, from the `qa/` folder:

```bash
cd qa
python3 fixtures.py            # creates test files in qa/od and resets qa/home
xvfb-run -a -s "-screen 0 1440x900x24" node qa.js     # on a desktop, plain `node qa.js` also works
```

Set `APP_BIN` to a packaged binary (for example `release-linux/linux-unpacked/yourofficesuite` after `npm run pack:linux`) to test the packaged app instead of the source tree.

Notes:

- The screen must be 1440 x 900, because some checks click at fixed positions inside the editors.
- The suite was written for Linux. It has not been run on Windows.
- Run `fixtures.py` before every run; the tests change the files they work on.
