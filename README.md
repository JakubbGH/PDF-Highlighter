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

CSV files can use either:

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

## Publishing To GitHub Pages

Upload these files to a GitHub repository:

- `index.html`
- `styles.css`
- `app.js`
- `assets/sample-floor.svg`
- `vendor/pdf.min.js`
- `vendor/pdf.worker.min.js`
- `vendor/README.md`
- `vendor/excel/vbaProject.bin`
- `README.md`

Then enable GitHub Pages for the repository branch. No build step is required.

## Notes

- PDF plans are rendered in the browser using the first page you choose during upload.
- Room boundaries are traced once, then saved in the project JSON.
- The original drawing remains visible under the semi-transparent progress overlay.
- Excel export creates an `.xlsm` package with editable room data on the `Progress` sheet and zone drawing shapes on the `Plan` sheet.
