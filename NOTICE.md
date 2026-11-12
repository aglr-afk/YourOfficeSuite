# Third-party software and changes

YourOfficeSuite is licensed under the GNU Affero General Public License v3.0 (see [LICENSE](LICENSE)). It includes or builds on the projects below, each under its own licence.

## Editors

| Project | Licence | Source |
| --- | --- | --- |
| ONLYOFFICE sdkjs | AGPL-3.0, Copyright (c) Ascensio System SIA | https://github.com/ONLYOFFICE/sdkjs |
| ONLYOFFICE web-apps | AGPL-3.0, Copyright (c) Ascensio System SIA | https://github.com/ONLYOFFICE/web-apps |
| ranuts/document (browser build of the editors) | AGPL-3.0 | https://github.com/ranuts/document |
| onlyoffice-x2t-wasm (format converter) | AGPL-3.0 | https://github.com/cryptpad/onlyoffice-x2t-wasm |

The engine is built from `ranuts/document` at commit `d1c9b03263aabe14789297f5c717fefff2fe4f0d` by `scripts/get-engine.sh`. That project's own `LICENSE` and `NOTICE` files are shipped unchanged inside the installed app under `resources/engine/`.

### Changes made to the engine

These are applied by `scripts/patch-engine.py`:

1. The converter WebAssembly file is stored uncompressed, because the app serves it over plain local HTTP.
2. Font table rows for Calibri, Andale Mono and Comic Sans MS are given their bold and italic faces.
3. A colour layer (`build/brand/editor-theme.css`) is appended to each editor's stylesheet. It also hides the editors' own About and Help entries; the same legal information is shown in the app under About > Open-source licences.
4. Loading screens use the app's colours.
5. The spreadsheet cell cursor colour is changed.

At run time the app also replaces the engine's caching service worker with a pass-through one, and removes the browser file-picker functions inside editor frames so that the editors' Save button goes through the app.

## PDF Bench and file tools

| Project | Licence | Source |
| --- | --- | --- |
| pdf-lib (Cantoo fork) | MIT | https://github.com/cantoo-scribe/pdf-lib |
| PDF.js | Apache-2.0 | https://github.com/mozilla/pdf.js |
| fflate | MIT | https://github.com/101arrowz/fflate |

## App shell

| Project | Licence | Source |
| --- | --- | --- |
| Electron | MIT | https://github.com/electron/electron |
| Inter typeface | SIL OFL 1.1 | https://github.com/rsms/inter |
| Caveat typeface | SIL OFL 1.1 | https://github.com/googlefonts/caveat |

Document fonts shipped with the engine (Liberation, Carlito, Caladea, DejaVu, Noto and others) are under the SIL Open Font License or similar free licences; see the engine's own notices.

## Trademarks

ONLYOFFICE is a trademark of Ascensio System SIA. Microsoft, Word, Excel, PowerPoint, OneDrive and Windows are trademarks of Microsoft. Adobe and Acrobat are trademarks of Adobe. YourOfficeSuite is an independent project and is not affiliated with or endorsed by any of them.
