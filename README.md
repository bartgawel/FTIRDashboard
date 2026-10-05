# FTIR Dashboard — Plotly + Native Excel

GitHub Pages URL:

    https://bartgawel.github.io/FTIRDashboard

## What changed
- Add-in button is placed on the Excel **Add-ins** tab.
- Task pane contains interactive **Plotly** FTIR charts.
- Plot all spectra or select only specific spectra.
- Plotly supports hover, zoom, pan, legend toggling, and scroll zoom.
- QC plot shows Log / Baseline / Corrected / Smoothed for one sample.
- `Build native Excel dashboard` creates a normal Excel chart on the `Dashboard` sheet.
- Auto-refresh watches `FTIR_Status`, which is written last by the xlwings Lite processing script.

## GitHub
Upload these files to the root of the existing FTIRDashboard repository:

    index.html
    taskpane.js
    styles.css
    assets/
    .nojekyll

`manifest.xml` is normally sideloaded into Excel rather than served from GitHub,
but keeping a copy in the repository is useful.

After uploading, wait for GitHub Pages to redeploy and verify:

    https://bartgawel.github.io/FTIRDashboard/

Then remove the old add-in if needed and sideload the new manifest.

## Expected workbook sheets

    Log
    Baseline
    Corrected
    Smoothed
    Normalized
    FTIR_Status

The Python analysis script from the previous package can stay unchanged.
