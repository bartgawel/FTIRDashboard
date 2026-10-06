# FTIR Dashboard — Excel Content Add-in (embedded in worksheet)

This package is for the **embedded HTML dashboard** version, like the project dashboards.

## What this version does
- Inserts the FTIR Plotly dashboard **inside a worksheet** as an embedded content add-in.
- Reads FTIR processed sheets from the current workbook.
- Shows interactive Plotly charts directly in the embedded dashboard:
  - all spectra for a selected stage
  - QC chart for one selected sample
- Includes controls for:
  - stage selector
  - sample selection
  - QC sample selector
  - auto-refresh after `FTIR_Status` changes
  - button to create a native Excel dashboard sheet if needed

## GitHub hosting
The hosted page URL used by this package is:

    https://bartgawel.github.io/FTIRDashboard/

Upload/update these files in the root of your `FTIRDashboard` GitHub Pages repository:

    index.html
    taskpane.js
    styles.css
    assets/
    .nojekyll

## How to use in Excel
1. Update GitHub Pages with the latest files.
2. In Excel, open the workbook with your FTIR processed sheets.
3. Insert the add-in using `manifest_content.xml`.
4. Excel will place the dashboard on a worksheet as an embedded object.
5. Resize it on the sheet as needed.
6. The dashboard should read the workbook data and show the Plotly charts.

## Expected workbook sheets
- Log
- Baseline
- Corrected
- Smoothed
- Normalized
- FTIR_Status

## Notes
- This is the **worksheet-embedded HTML version**, not the task pane version.
- It needs the hosted files to be reachable online (GitHub Pages).
- You can keep both versions if you want:
  - task pane manifest: `manifest.xml`
  - worksheet embedded manifest: `manifest_content.xml`
