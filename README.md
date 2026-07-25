# Open Office Excel

A free, local-first desktop spreadsheet app — a clean alternative to Apple Numbers and other tools that convert or corrupt Excel files.

Open Office Excel opens, edits, and saves real **`.xlsx`** workbooks. It never silently changes the file type. Built for everyday spreadsheet work with an Excel-like grid.

**Stack:** Electron · React · TypeScript · [Univer](https://univer.ai/) · [ExcelJS](https://github.com/exceljs/exceljs)

---

## Features

### Files
- **New / Open / Save / Save As** for `.xlsx` only
- Multi-sheet tabs
- Unsaved-changes warning on close, New, or Open
- Shortcuts: `Cmd/Ctrl+N`, `O`, `S`, `Shift+S`

### Editing
- Edit cell values
- Insert / delete rows and columns
- Resize rows and columns (saved with the file)
- Undo / redo, cut / copy / paste
- Formula bar

### Formatting
- Bold, italic, font size, text and fill color, alignment
- Number and date formats (Excel serial dates round-trip correctly)

### Data tools
- **Filter** — AutoFilter on a selection or the data range; filter via column dropdowns
- **Freeze header** — pin the first row
- **Freeze to this row** — pin rows above the selection (toolbar or right-click)
- **Unfreeze** — clear frozen panes
- **Date picker** — calendar dropdown on selected cells

### What it is not (yet)
Charts, pivot tables, VBA, collaborative editing, and legacy `.xls` are out of scope for v1.

---

## Requirements

- [Node.js](https://nodejs.org/) **18+** (20+ recommended)
- macOS, Windows, or Linux

---

## Run for development

```bash
git clone https://github.com/rakesh113/open-office-excel.git
cd open-office-excel
npm install
npm run dev
```

This starts Electron with hot reload. The app window should open automatically.

### Troubleshooting

If Electron fails with `Cannot read properties of undefined (reading 'whenReady')`, make sure `ELECTRON_RUN_AS_NODE` is unset. Some IDEs set this. The `npm run dev` script clears it automatically.

---

## Install for direct use

You can run a production build locally, or package a desktop app.

### Option A — Run the production build (no installer)

```bash
npm install
npm run build
npm run preview
```

This launches the packaged renderer + Electron main process without the Vite dev server.

### Option B — Build a desktop app (recommended for daily use)

```bash
npm install
npm run dist
```

Artifacts are written to the `release/` folder:

| Platform | Output |
|----------|--------|
| macOS | `.dmg` and `.app` under `release/` |
| Windows | `.exe` installer under `release/` |
| Linux | `.AppImage` under `release/` |

On macOS, open the `.dmg` and drag **Open Office Excel** into Applications, or open the `.app` directly.

> **Note:** Unsigned local builds may show a Gatekeeper warning on macOS. Use **System Settings → Privacy & Security** to allow the app, or right-click → **Open**.

### Option C — Install from a GitHub Release

If this repository publishes Releases, download the installer for your OS from the **Releases** page, then install as usual.

---

## Useful scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Development mode with hot reload |
| `npm run build` | Compile main, preload, and renderer |
| `npm run preview` | Run the production build |
| `npm run dist` | Build + package installers into `release/` |
| `npm run typecheck` | TypeScript check |
| `npm run test:roundtrip` | Smoke-test `.xlsx` open/save conversion |

---

## Sample file

[`fixtures/sample.xlsx`](fixtures/sample.xlsx) includes text, a number, a date, basic formatting, and custom row/column sizes — useful for a quick open/edit/save check.

---

## Project layout

```
src/
  main/        # Electron main process, menus, dialogs, ExcelJS I/O
  preload/     # Secure bridge between main and renderer
  renderer/    # React UI + Univer spreadsheet
  shared/      # IPC channels and workbook types
fixtures/      # Sample .xlsx files
```

---

## Design notes

- Files are always written as `.xlsx` (OOXML). The app does not convert to Numbers, CSV, or other formats on save.
- Univer’s official import/export is server/Pro-oriented. This app uses **ExcelJS** in the Electron main process and maps to Univer’s in-memory workbook snapshot, so everything stays local and free.

---

## License

[MIT](LICENSE)
