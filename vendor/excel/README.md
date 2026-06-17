# Excel Macro Template

`vbaProject.bin` is the compiled macro project embedded into `.xlsm` exports. Excel shape fills and shape text cannot be changed by worksheet formulas alone, so the live recolouring depends on this binary containing the zone refresh macro.

The current `vbaProject.bin` is a placeholder sample project. It makes the exported workbook structurally macro-enabled, but it does not contain the project-specific sheet-change macro yet. Until this file is replaced, exported workbooks will show the floor plan, zones, and labels, but changing `Progress!D:D` will not recolour the boxes or refresh the label text.

## Option A: Install A Template In The Browser

1. Open Excel and create a new macro-enabled workbook.
2. Press `Alt+F11` to open the VBA editor.
3. Open `ThisWorkbook` in the template workbook.
4. Paste in the full contents of `ThisWorkbookCode.bas`.
5. Save the workbook as `.xlsm`, for example `floor-plan-macro-template.xlsm`.
6. Open the local website and click **Install Macro**.
7. Choose `floor-plan-macro-template.xlsm`.

The site extracts only `xl/vbaProject.bin` from the workbook and stores that compiled VBA project in browser local storage. It is not uploaded anywhere. Future exports from that browser will use the installed VBA project.

## Option B: Build And Replace The Repo Template Automatically

If Excel is installed on the Windows machine, run this from the project root:

```powershell
powershell -ExecutionPolicy Bypass -File tools\install_excel_macro_template.ps1
```

The script creates a temporary `.xlsm`, inserts `ThisWorkbookCode.bas`, extracts the compiled `xl/vbaProject.bin`, and replaces `vendor/excel/vbaProject.bin`. Excel may require **Trust access to the VBA project object model** to be enabled in Trust Center.

## Option C: Replace The Repo Template Manually

1. Open Excel and create a new macro-enabled workbook.
2. Press `Alt+F11` to open the VBA editor.
3. Open `ThisWorkbook` in the template workbook.
4. Paste in the full contents of `ThisWorkbookCode.bas`.
5. Save the workbook as `.xlsm`, for example `floor-plan-macro-template.xlsm`.
6. From the project root, run:

```powershell
py -3 tools\extract_vba_project.py floor-plan-macro-template.xlsm
```

The script extracts `xl/vbaProject.bin` from the template and replaces `vendor/excel/vbaProject.bin`. Refresh the website, export a new workbook, enable macros in Excel, and then edits to `Progress` column `D` or `E` will update the zone colour, transparency, and white overlay label text.

## Source Files

- `ThisWorkbookCode.bas` is the recommended all-in-one workbook event macro.
- `RefreshZoneColours.bas` contains the refresh routine as a standard module, useful if you prefer to keep event code and refresh code separate.
- `ProgressSheetChange.bas` contains the older worksheet-level event version; paste it into the `Progress` worksheet code module only if you are building a sheet-specific template.
