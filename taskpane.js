/* global Office, Excel, Plotly */

let spectraCache = {};
let qcCache = {};
let sampleNames = [];
let refreshTimer = null;
let busy = false;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  document.getElementById("stage").addEventListener("change", async () => {
    await readWorkbookData();
    renderAll();
  });

  document.getElementById("qcSample").addEventListener("change", renderQc);

  document.getElementById("selectAll").addEventListener("click", () => {
    document.querySelectorAll(".sample-toggle").forEach(x => x.checked = true);
    renderMain();
  });

  document.getElementById("clearAll").addEventListener("click", () => {
    document.querySelectorAll(".sample-toggle").forEach(x => x.checked = false);
    renderMain();
  });

  document.getElementById("refresh").addEventListener("click", async () => {
    await readWorkbookData();
    renderAll();
  });

  document.getElementById("native").addEventListener("click", buildNativeDashboard);

  document.getElementById("autoOpen").addEventListener("change", async (e) => {
    try {
      await setAutoOpen(e.target.checked);
    } catch (error) {
      console.error(error);
      setStatus(error.message || String(error), "error");
    }
  });

  await loadAutoOpenState();
  await readWorkbookData();
  renderAll();
  await registerStatusWatcher();
});

function setStatus(message, kind = "") {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = `status ${kind}`.trim();
}

function selectedSamples() {
  return Array.from(document.querySelectorAll(".sample-toggle"))
    .filter(x => x.checked)
    .map(x => x.value);
}

async function readWorkbookData() {
  if (busy) return;
  busy = true;

  const stage = document.getElementById("stage").value;

  try {
    setStatus(`Reading ${stage} data…`);

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;

      const stageSheet = sheets.getItemOrNullObject(stage);
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

      if (stageSheet.isNullObject) throw new Error(`Sheet "${stage}" not found.`);
      if (logSheet.isNullObject || baselineSheet.isNullObject ||
          correctedSheet.isNullObject || smoothedSheet.isNullObject) {
        throw new Error("Required FTIR processing sheets are missing.");
      }

      const stageUsed = stageSheet.getUsedRange();
      stageUsed.load("values,rowCount,columnCount");

      const logUsed = logSheet.getUsedRange();
      logUsed.load("values,rowCount,columnCount");

      const baseUsed = baselineSheet.getUsedRange();
      baseUsed.load("values,rowCount,columnCount");

      const corrUsed = correctedSheet.getUsedRange();
      corrUsed.load("values,rowCount,columnCount");

      const smoothUsed = smoothedSheet.getUsedRange();
      smoothUsed.load("values,rowCount,columnCount");

      await context.sync();

      const v = stageUsed.values;
      const headers = v[0];
      const x = v.slice(1).map(r => Number(r[0]));

      spectraCache = {};
      sampleNames = [];

      for (let c = 1; c < stageUsed.columnCount; c++) {
        const name = String(headers[c] ?? "").trim();
        if (!name) continue;
        sampleNames.push(name);
        spectraCache[name] = {
          x,
          y: v.slice(1).map(r => Number(r[c]))
        };
      }

      function convertQc(usedRange) {
        const vals = usedRange.values;
        const h = vals[0].map(z => String(z ?? "").trim());
        const xx = vals.slice(1).map(r => Number(r[0]));
        const out = {};
        for (let c = 1; c < usedRange.columnCount; c++) {
          const name = h[c];
          if (!name) continue;
          out[name] = {
            x: xx,
            y: vals.slice(1).map(r => Number(r[c]))
          };
        }
        return out;
      }

      qcCache = {
        Log: convertQc(logUsed),
        Baseline: convertQc(baseUsed),
        Corrected: convertQc(corrUsed),
        Smoothed: convertQc(smoothUsed)
      };
    });

    rebuildSelectors();
    setStatus(`Loaded ${sampleNames.length} spectra from ${stage}.`, "ok");

  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  } finally {
    busy = false;
  }
}

function rebuildSelectors() {
  const list = document.getElementById("sampleList");
  const previouslySelected = new Set(selectedSamples());
  list.innerHTML = "";

  sampleNames.forEach((name, i) => {
    const row = document.createElement("label");
    row.className = "sample-row";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "sample-toggle";
    cb.value = name;
    cb.checked = previouslySelected.size ? previouslySelected.has(name) : true;
    cb.addEventListener("change", renderMain);

    const text = document.createElement("span");
    text.textContent = name;

    row.appendChild(cb);
    row.appendChild(text);
    list.appendChild(row);
  });

  const qc = document.getElementById("qcSample");
  const prev = qc.value;
  qc.innerHTML = "";

  sampleNames.forEach(name => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    qc.appendChild(opt);
  });

  if (sampleNames.includes(prev)) qc.value = prev;
}

function renderAll() {
  renderMain();
  renderQc();
}

function renderMain() {
  const stage = document.getElementById("stage").value;
  const chosen = selectedSamples();

  const traces = chosen.map(name => ({
    x: spectraCache[name]?.x || [],
    y: spectraCache[name]?.y || [],
    type: "scatter",
    mode: "lines",
    name
  }));

  const layout = {
    margin: { l: 55, r: 20, t: 35, b: 50 },
    title: { text: `${stage} FTIR spectra`, font: { size: 15 } },
    xaxis: {
      title: "Wavenumber (cm⁻¹)",
      autorange: "reversed"
    },
    yaxis: {
      title: stage === "Normalized" ? "Normalized intensity" : "Intensity"
    },
    legend: { orientation: "h" },
    hovermode: "x unified"
  };

  Plotly.react("mainPlot", traces, layout, {
    responsive: true,
    displaylogo: false,
    scrollZoom: true
  });
}

function renderQc() {
  const sample = document.getElementById("qcSample").value;
  if (!sample) return;

  const names = ["Log", "Baseline", "Corrected", "Smoothed"];
  const labelMap = {
    Log: "-log(R)",
    Baseline: "Baseline",
    Corrected: "Corrected",
    Smoothed: "Smoothed"
  };

  const traces = names
    .filter(stage => qcCache[stage]?.[sample])
    .map(stage => ({
      x: qcCache[stage][sample].x,
      y: qcCache[stage][sample].y,
      type: "scatter",
      mode: "lines",
      name: labelMap[stage]
    }));

  const layout = {
    margin: { l: 55, r: 20, t: 35, b: 50 },
    title: { text: `Processing QC — ${sample}`, font: { size: 15 } },
    xaxis: {
      title: "Wavenumber (cm⁻¹)",
      autorange: "reversed"
    },
    yaxis: { title: "Intensity" },
    legend: { orientation: "h" },
    hovermode: "x unified"
  };

  Plotly.react("qcPlot", traces, layout, {
    responsive: true,
    displaylogo: false,
    scrollZoom: true
  });
}

async function buildNativeDashboard() {
  const stage = document.getElementById("stage").value;
  const chosen = selectedSamples();
  const qcSample = document.getElementById("qcSample").value;

  if (!chosen.length) {
    setStatus("Select at least one spectrum before creating native charts.", "error");
    return;
  }

  try {
    setStatus("Building native Excel dashboard…");

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const stageSheet = sheets.getItem(stage);

      const used = stageSheet.getUsedRange();
      used.load("rowCount,columnCount,values");
      await context.sync();

      const nPoints = used.rowCount - 1;
      const headers = used.values[0].map(v => String(v ?? "").trim());

      let dashboard = sheets.getItemOrNullObject("Dashboard");
      dashboard.load("isNullObject");
      await context.sync();
      if (dashboard.isNullObject) dashboard = sheets.add("Dashboard");

      const charts = dashboard.charts;
      charts.load("items");
      await context.sync();
      charts.items.forEach(c => c.delete());

      dashboard.getRange("A1:B6").clear();
      dashboard.getRange("A1:B5").values = [
        ["FTIR Dashboard", ""],
        ["Processing stage", stage],
        ["Spectra plotted", chosen.length],
        ["QC sample", qcSample],
        ["Source", "FTIR Plotly add-in"]
      ];

      dashboard.getRange("A1").format.font.bold = true;
      dashboard.getRange("A1").format.font.size = 18;
      dashboard.getRange("A2:A5").format.font.bold = true;
      dashboard.getRange("A:B").format.autofitColumns();

      const dummy = dashboard.getRange("Z1:AA2");
      dummy.values = [["X","Y"],[0,0]];

      const chart = dashboard.charts.add(
        Excel.ChartType.xyscatterLinesNoMarkers,
        dummy,
        Excel.ChartSeriesBy.columns
      );

      chart.name = "FTIR_SelectedSpectra";
      chart.title.text = `${stage} — selected FTIR spectra`;
      chart.setPosition("D2", "O25");

      chart.series.load("items");
      await context.sync();
      chart.series.items.forEach(s => s.delete());

      const xRange = stageSheet.getRangeByIndexes(1, 0, nPoints, 1);

      for (const name of chosen) {
        const idx = headers.indexOf(name);
        if (idx < 1) continue;
        const series = chart.series.add(name);
        series.setXAxisValues(xRange);
        series.setValues(stageSheet.getRangeByIndexes(1, idx, nPoints, 1));
      }

      chart.axes.categoryAxis.title.text = "Wavenumber (cm⁻¹)";
      chart.axes.categoryAxis.title.visible = true;
      chart.axes.categoryAxis.reversePlotOrder = true;
      chart.axes.valueAxis.title.text = stage === "Normalized" ? "Normalized intensity" : "Intensity";
      chart.axes.valueAxis.title.visible = true;
      chart.legend.visible = true;

      dashboard.activate();
      await context.sync();
    });

    setStatus("Native Excel dashboard updated.", "ok");
  } catch (error) {
    console.error(error);
    setStatus(error.message || String(error), "error");
  }
}

function setAutoOpen(enabled) {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(
        "Office.AutoShowTaskpaneWithDocument",
        enabled
      );

      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus(
            enabled
              ? "Auto-open enabled for this workbook. Save the workbook to keep this setting."
              : "Auto-open disabled for this workbook. Save the workbook to keep this setting.",
            "ok"
          );
          resolve();
        } else {
          reject(new Error(result.error?.message || "Could not save the auto-open setting."));
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

async function loadAutoOpenState() {
  try {
    const current = Office.context.document.settings.get(
      "Office.AutoShowTaskpaneWithDocument"
    );
    document.getElementById("autoOpen").checked = current === true;
  } catch (error) {
    console.error(error);
  }
}

async function registerStatusWatcher() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItemOrNullObject("FTIR_Status");
      ws.load("isNullObject");
      await context.sync();

      if (ws.isNullObject) {
        setStatus("Plotly dashboard ready. Run Python processing to enable auto-refresh.");
        return;
      }

      ws.onChanged.add(() => {
        if (!document.getElementById("autoRefresh").checked) return;
        if (refreshTimer) clearTimeout(refreshTimer);

        refreshTimer = setTimeout(async () => {
          await readWorkbookData();
          renderAll();
        }, 700);
      });

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
