# Floor Plan Progress Tracker

A static GitHub Pages-ready app for marking rooms on a floor plan and shading them by percentage complete.

## Use Locally

Open `index.html` in a browser. The app runs fully in the browser and autosaves to local storage.

## Basic Workflow

1. Click **Plan** and choose a PDF, PNG, JPG, WebP, or SVG floor plan.
2. Click **Draw Room**.
3. Enter the room code, click around the room boundary, then click **Finish**.
4. Select a room and adjust its percent complete.
5. Use **Save** to download a project JSON backup.
6. Use **Import CSV** to update percentages later.
7. Use **Export XLSM** to download a macro-enabled Excel workbook with the floor plan, zone shapes, and editable room progress table.
8. Use **Install Macro** once if you have a macro template workbook and want exports from this browser to include the live Excel refresh macro.

CSV files can use either the simple progress-only format:

```csv
room_id,percent
A101,25
A102,70
```

or:

```csv
Room,Progress
A101,25%
A102,70%
```

Exports from the app include overlay points too:

```csv
room_id,percent,points,plan_width,plan_height
A101,25,"[[72,72],[357,72],[357,258],[72,258]]",1200,800
```

When that CSV is imported again, existing room percentages and overlay points are updated. If a room code does not exist yet but the CSV has points, the room is created.

## Publishing To GitHub Pages

Upload these files to a GitHub repository:

- `index.html`
- `.nojekyll`
- `styles.css`
- `app.js`
- `assets/sample-floor.svg`
- `vendor/pdf.min.js`
- `vendor/pdf.worker.min.js`
- `vendor/README.md`
- `vendor/excel/vbaProject.bin`
- `vendor/excel/README.md`
- `vendor/excel/ThisWorkbookCode.bas`
- `vendor/excel/RefreshZoneColours.bas`
- `vendor/excel/ProgressSheetChange.bas`
- `tools/extract_vba_project.py`
- `tools/install_excel_macro_template.ps1`
- `tools/run_checks.ps1`
- `tools/smoke_test.js`
- `tools/browser_smoke_test.js`
- `README.md`

Then enable GitHub Pages for the repository branch. No build step is required.

## Local Checks

Run the full local check before publishing changes:

```powershell
powershell -ExecutionPolicy Bypass -File tools\run_checks.ps1
```

Run only the JavaScript smoke test when you want a quicker check:

```powershell
node tools\smoke_test.js
```

In this Codex workspace, use the bundled Node runtime if regular `node` is not installed:

```powershell
& 'C:\Users\imnot\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' tools\smoke_test.js
```

`tools/browser_smoke_test.js` is a browser-context check for the loaded page. It is intended for the browser console or automation and confirms the sample overlays, labels, macro status, PDF support, and export controls are present.

## Excel Macro Setup

If Excel is installed on the Windows machine, the helper below can create the macro template and install the compiled VBA project into the repo:

```powershell
powershell -ExecutionPolicy Bypass -File tools\install_excel_macro_template.ps1
```

If Excel blocks VBA project access, enable **Trust access to the VBA project object model** in Excel Trust Center and run it again. Without that compiled VBA project, `.xlsm` exports still open as snapshots, but Excel cannot recolour shapes when column D changes.

## Notes

- PDF plans are rendered in the browser using the first page you choose during upload.
- Room boundaries are traced once, then saved in the project JSON.
- The original drawing remains visible under the semi-transparent progress overlay.
- Excel export creates an `.xlsm` package with editable room data on the `Progress` sheet, zone drawing shapes on the `Plan` sheet, and separate white text labels layered over each zone. `Percent Complete` is exported as a numeric 0-100 value in column D.
- Live Excel-side recolouring needs a compiled VBA project. The browser can embed that file, but it cannot compile `.bas` source files itself. Use **Install Macro** to load a macro-enabled template locally in the browser, or replace `vendor/excel/vbaProject.bin` once for every copy of the site. See `vendor/excel/README.md`.
- Installed macro templates are stored only in your browser's local storage. Floor plans, PDFs, and templates are not uploaded by this static page.
