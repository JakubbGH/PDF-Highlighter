# Floor Plan Progress Tracker

A static GitHub Pages-ready app for marking rooms on a floor plan and shading them by percentage complete.

## Use Locally

Open `index.html` in a browser. The app runs fully in the browser and autosaves to local storage.

## Basic Workflow

1. Click **Plans** and choose one or more PDF, PNG, JPG, WebP, or SVG floor plans.
   Multi-page PDFs and multiple selected PDFs are loaded into separate page tabs.
2. Use **Rename** to name each page tab, such as `Ground Floor` or `Level 02`.
3. Click **Draw Room**.
4. Enter the room code, click around the room boundary, then click **Finish**.
5. Select a room and adjust its percent complete.
6. Use **Save** to download a project JSON backup.
7. Use **Import CSV** to update percentages later.
8. Use **Export XLSM** to download a macro-enabled Excel workbook with the floor plan, zone shapes, and editable room progress table.
9. Use **Copy VBA** and **Install Macro** once if you want exports from this browser to include the live Excel refresh macro.

For drawings with room IDs containing `ZL`, click **Auto ZL** after loading the plan. The app first reads any PDF text layer locally; if that finds nothing, it tries browser-native text detection on the rendered drawing image. It places pins on unmapped ZL labels and tries to create box-like room boundaries by scanning for straight wall lines around each label. Door gaps are tolerated, but browser image-text detection support varies, and the result should still be reviewed and adjusted where the drawing is broken or busy.

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
page,page_name,room_id,percent,points,plan_width,plan_height
1,Plan 1,A101,25,"[[72,72],[357,72],[357,258],[72,258]]",1200,800
```

When that CSV is imported again, existing room percentages, overlay points, and page mappings are updated. If a room code does not exist yet but the CSV has points, the room is created on the matching page. Older CSVs without a page column still import into the active tab.

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

## Optional Developer Checks

You do not need these commands to use the app, export workbooks, or publish the static site to GitHub Pages. They are only quick checks for a development machine that can run scripts.

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

Recommended no-PowerShell route:

1. Open Excel and create a macro-enabled workbook.
2. In the web app, click **Copy VBA**. If the browser cannot copy to the clipboard, it will download or open `ThisWorkbookCode.bas` instead.
3. Press `Alt+F11` in Excel, open `ThisWorkbook`, paste the copied VBA code, and save as `.xlsm`.
4. Return to the web app, click **Install Macro**, and choose that `.xlsm`.
5. Click **Macro Bin** to download `vbaProject.bin`.
6. Replace `vendor/excel/vbaProject.bin` in the repo with that downloaded file.

If Excel is installed on a Windows machine that can run scripts, the helper below can create the macro template and install the compiled VBA project into the repo automatically:

```powershell
powershell -ExecutionPolicy Bypass -File tools\install_excel_macro_template.ps1
```

Then run the strict publish check:

```powershell
powershell -ExecutionPolicy Bypass -File tools\run_checks.ps1 -RequireCompiledMacro
```

If Excel blocks VBA project access, enable **Trust access to the VBA project object model** in Excel Trust Center and run it again. Without that compiled VBA project, `.xlsm` exports still open as snapshots, but Excel cannot refresh labels, colours, or opacity when columns A, B, E, or F change.

## Notes

- PDF plans are rendered in the browser. Multi-page PDFs become page tabs in the app; renamed page tabs become the plan worksheet names in Excel.
- Room boundaries are traced once, then saved in the project JSON.
- The original drawing remains visible under the semi-transparent progress overlay.
- Excel export creates an `.xlsm` package with editable room data on the `Progress` sheet, zone drawing shapes on the named plan sheet for each page, and separate white text labels layered over each zone. `Percent Complete` is exported as a numeric 0-100 value in column E. `Progress` column B stores the exact plan worksheet name used by the macro.
- Live Excel-side recolouring needs a compiled VBA project. The browser can embed that file, but it cannot compile `.bas` source files itself. Use **Copy VBA** to prepare the Excel template, **Install Macro** to load it locally in the browser, or replace `vendor/excel/vbaProject.bin` once for every copy of the site. See `vendor/excel/README.md`.
- Installed macro templates are stored only in your browser's local storage. Floor plans, PDFs, and templates are not uploaded by this static page.
