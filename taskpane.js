/* global Office, Excel */

let statusSheet = null;
let statusHandlerRegistered = false;
let refreshTimer = null;
let dashboardBusy = false;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  document.getElementById("reload").addEventListener("click", async () => {
    await loadSamples();
  });

  document.getElementById("build").addEventListener("click", async () => {
    await loadSamples();
    await buildDashboard();
  });

  document.getElementById("stage").addEventListener("change", buildDashboard);
  document.getElementById("qcSample").addEventListener("change", buildDashboard);
  document.getElementById("showQc").addEventListener("change", buildDashboard);

  await loadSamples();
  await registerAutoRefresh();

  // Build once when the task pane opens if processed data already exist.
  await buildDashboard();
});

function setStatus(message, kind = "") {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = `status ${kind}`.trim();
}

function scheduleRefresh() {
  if (!document.getElementById("autoRefresh").checked) return;

  if (refreshTimer) clearTimeout(refreshTimer);

  refreshTimer = setTimeout(async () => {
    await loadSamples();
    await buildDashboard();
  }, 700);
}

async function registerAutoRefresh() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItemOrNullObject("FTIR_Status");
      ws.load("isNullObject");
      await context.sync();

      if (ws.isNullObject) {
        setStatus(
          'Waiting for FTIR_Status. Run the Python process_ftir script once.',
          ""
        );
        return;
      }

      statusSheet = ws;

      if (!statusHandlerRegistered) {
        ws.onChanged.add(() => {
          scheduleRefresh();
        });
        statusHandlerRegistered = true;
        await context.sync();
      }
    });
  } catch (error) {
    console.error(error);
  }
}

async function loadSamples() {
  try {
    await Excel.run(async (context) => {
      const logSheet = context.workbook.worksheets.getItemOrNullObject("Log");
      logSheet.load("isNullObject");
      await context.sync();

      if (logSheet.isNullObject) {
        throw new Error('Sheet "Log" is not available yet.');
      }

      const used = logSheet.getUsedRange();
      used.load("columnCount,values");
      await context.sync();

      const headers = used.values[0];
      const samples = [];

      for (let c = 1; c < used.columnCount; c++) {
        const v = headers[c];
        if (v !== null && v !== "") samples.push(String(v));
      }

      const select = document.getElementById("qcSample");
      const previous = select.value;
      select.innerHTML = "";

      samples.forEach((name) => {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        select.appendChild(opt);
      });

      if (samples.includes(previous)) {
        select.value = previous;
      }

      setStatus(`Ready: ${samples.length} processed spectra found.`, "ok");
    });
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  }
}

async function buildDashboard() {
  if (dashboardBusy) return;

  const stageName = document.getElementById("stage").value;
  const qcSample = document.getElementById("qcSample").value;
  const showQc = document.getElementById("showQc").checked;

  dashboardBusy = true;

  try {
    setStatus("Refreshing native Excel dashboard…");

    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheets = workbook.worksheets;

      const stageSheet = sheets.getItemOrNullObject(stageName);
      stageSheet.load("isNullObject");

      const logSheet = sheets.getItemOrNullObject("Log");
      logSheet.load("isNullObject");

      const baselineSheet = sheets.getItemOrNullObject("Baseline");
      baselineSheet.load("isNullObject");

      const correctedSheet = sheets.getItemOrNullObject("Corrected");
      correctedSheet.load("isNullObject");

      const smoothedSheet = sheets.getItemOrNullObject("Smoothed");
      smoothedSheet.load("isNullObject");

      await context.sync();

      if (stageSheet.isNullObject) {
        throw new Error(`Sheet "${stageName}" was not found.`);
      }
      if (logSheet.isNullObject || baselineSheet.isNullObject ||
          correctedSheet.isNullObject || smoothedSheet.isNullObject) {
        throw new Error("One or more FTIR processing sheets are missing.");
      }

      const used = stageSheet.getUsedRange();
      used.load("rowCount,columnCount,values");

      const logUsed = logSheet.getUsedRange();
      logUsed.load("columnCount,values");

      await context.sync();

      const nPoints = used.rowCount - 1;
      const headers = used.values[0];

      if (nPoints < 2 || used.columnCount < 2) {
        throw new Error(`No usable spectra found on "${stageName}".`);
      }

      const samples = [];
      for (let c = 1; c < used.columnCount; c++) {
        if (headers[c] !== null && headers[c] !== "") {
          samples.push({ name: String(headers[c]), columnIndex: c });
        }
      }

      if (!samples.length) {
        throw new Error("No sample columns were found.");
      }

      const logHeaders = logUsed.values[0].map((v) => String(v ?? ""));
      const qcIndex = logHeaders.indexOf(qcSample);

      if (showQc && qcIndex < 1) {
        throw new Error(`QC sample "${qcSample}" was not found.`);
      }

      // Dashboard sheet
      let dashboard = sheets.getItemOrNullObject("Dashboard");
      dashboard.load("isNullObject");
      await context.sync();

      if (dashboard.isNullObject) {
        dashboard = sheets.add("Dashboard");
      }

      // Delete prior charts.
      const charts = dashboard.charts;
      charts.load("items");
      await context.sync();
      charts.items.forEach((chart) => chart.delete());

      // Clear a compact dashboard info area only.
      dashboard.getRange("A1:B8").clear();

      dashboard.getRange("A1:B6").values = [
        ["FTIR Dashboard", ""],
        ["Processing stage", stageName],
        ["Spectra plotted", samples.length],
        ["QC sample", showQc ? qcSample : "Not shown"],
        ["Auto refresh", document.getElementById("autoRefresh").checked ? "On" : "Off"],
        ["Source", "xlwings Lite FTIR processing"]
      ];

      dashboard.getRange("A1").format.font.bold = true;
      dashboard.getRange("A1").format.font.size = 18;
      dashboard.getRange("A2:A6").format.font.bold = true;
      dashboard.getRange("A:B").format.autofitColumns();

      // Dummy source only to instantiate chart objects.
      const dummy = dashboard.getRange("Z1:AA2");
      dummy.values = [["X", "Y"], [0, 0]];

      // MAIN CHART -------------------------------------------------------------
      const mainChart = dashboard.charts.add(
        Excel.ChartType.xyscatterLinesNoMarkers,
        dummy,
        Excel.ChartSeriesBy.columns
      );

      mainChart.name = "FTIR_AllSpectra";
      mainChart.title.text = `${stageName} — all FTIR spectra`;
      mainChart.setPosition("D2", "O25");

      mainChart.series.load("items");
      await context.sync();
      mainChart.series.items.forEach((s) => s.delete());

      const xRange = stageSheet.getRangeByIndexes(1, 0, nPoints, 1);

      for (const sample of samples) {
        const series = mainChart.series.add(sample.name);
        const yRange = stageSheet.getRangeByIndexes(
          1, sample.columnIndex, nPoints, 1
        );
        series.setXAxisValues(xRange);
        series.setValues(yRange);
      }

      mainChart.axes.categoryAxis.title.text = "Wavenumber (cm⁻¹)";
      mainChart.axes.categoryAxis.title.visible = true;
      mainChart.axes.categoryAxis.reversePlotOrder = true;

      mainChart.axes.valueAxis.title.text =
        stageName === "Normalized" ? "Normalized intensity" : "Intensity";
      mainChart.axes.valueAxis.title.visible = true;
      mainChart.legend.visible = true;

      // QC CHART ---------------------------------------------------------------
      if (showQc) {
        const qcChart = dashboard.charts.add(
          Excel.ChartType.xyscatterLinesNoMarkers,
          dummy,
          Excel.ChartSeriesBy.columns
        );

        qcChart.name = "FTIR_ProcessingQC";
        qcChart.title.text = `Processing QC — ${qcSample}`;
        qcChart.setPosition("D27", "O50");

        qcChart.series.load("items");
        await context.sync();
        qcChart.series.items.forEach((s) => s.delete());

        const sources = [
          { name: "-log(R)", sheet: logSheet },
          { name: "Baseline", sheet: baselineSheet },
          { name: "Corrected", sheet: correctedSheet },
          { name: "Smoothed", sheet: smoothedSheet }
        ];

        for (const src of sources) {
          const u = src.sheet.getUsedRange();
          u.load("rowCount");
          await context.sync();

          const rows = Math.min(nPoints, u.rowCount - 1);
          const sx = src.sheet.getRangeByIndexes(1, 0, rows, 1);
          const sy = src.sheet.getRangeByIndexes(1, qcIndex, rows, 1);

          const series = qcChart.series.add(src.name);
          series.setXAxisValues(sx);
          series.setValues(sy);
        }

        qcChart.axes.categoryAxis.title.text = "Wavenumber (cm⁻¹)";
        qcChart.axes.categoryAxis.title.visible = true;
        qcChart.axes.categoryAxis.reversePlotOrder = true;
        qcChart.axes.valueAxis.title.text = "Intensity";
        qcChart.axes.valueAxis.title.visible = true;
        qcChart.legend.visible = true;
      }

      dashboard.activate();
      await context.sync();

      setStatus(
        `Dashboard refreshed: ${samples.length} spectra from ${stageName}.`,
        "ok"
      );
    });
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    dashboardBusy = false;
  }
}
