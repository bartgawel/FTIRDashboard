/* global Office, Excel, Plotly */

let spectraCache = {};
let qcCache = {};
let sampleNames = [];
let fitCurveCache = {};
let peakDetails = [];
let peakAreas = [];
let fitQuality = [];
let timer = null;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(x => x.classList.remove("active"));
      document.querySelectorAll(".tabpage").forEach(x => x.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.tab).classList.add("active");
      setTimeout(() => {
        Plotly.Plots.resize("mainPlot");
        Plotly.Plots.resize("qcPlot");
        Plotly.Plots.resize("fitPlot");
        Plotly.Plots.resize("areaPlot");
      }, 50);
    });
  });

  document.getElementById("stage").addEventListener("change", async () => {
    await readSpectra();
    renderSpectra();
  });
  document.getElementById("qcSample").addEventListener("change", renderQc);
  document.getElementById("fitSample").addEventListener("change", renderFit);
  document.getElementById("showComponents").addEventListener("change", renderFit);
  document.getElementById("areaMetric").addEventListener("change", renderAreas);

  document.getElementById("selectAll").addEventListener("click", () => {
    document.querySelectorAll(".sample-toggle").forEach(x => x.checked = true);
    renderMain();
  });
  document.getElementById("clearAll").addEventListener("click", () => {
    document.querySelectorAll(".sample-toggle").forEach(x => x.checked = false);
    renderMain();
  });

  document.getElementById("refresh").addEventListener("click", async () => {
    await readSpectra();
    await readFitResults();
    renderAll();
  });
  document.getElementById("reloadFit").addEventListener("click", async () => {
    await readFitResults();
    renderFit();
    renderAreas();
  });

  await readSpectra();
  await readFitResults();
  renderAll();
  await registerWatchers();
});

function status(msg, kind="") {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = `status ${kind}`.trim();
}

function selectedSamples() {
  return Array.from(document.querySelectorAll(".sample-toggle"))
    .filter(x => x.checked).map(x => x.value);
}

async function readSpectra() {
  const stage = document.getElementById("stage").value;

  try {
    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      const names = [stage, "Log", "Baseline", "Corrected", "Smoothed"];
      const ws = {};
      names.forEach(n => {
        ws[n] = sheets.getItemOrNullObject(n);
        ws[n].load("isNullObject");
      });
      await ctx.sync();

      for (const n of names) if (ws[n].isNullObject) throw new Error(`Sheet "${n}" not found.`);

      const used = {};
      names.forEach(n => {
        used[n] = ws[n].getUsedRange();
        used[n].load("values,rowCount,columnCount");
      });
      await ctx.sync();

      const main = used[stage].values;
      const headers = main[0];
      const x = main.slice(1).map(r => Number(r[0]));
      spectraCache = {};
      sampleNames = [];

      for (let c=1; c<used[stage].columnCount; c++) {
        const name = String(headers[c] ?? "").trim();
        if (!name) continue;
        sampleNames.push(name);
        spectraCache[name] = {x, y: main.slice(1).map(r => Number(r[c]))};
      }

      function conv(u) {
        const vals=u.values, h=vals[0].map(v=>String(v??"").trim());
        const xx=vals.slice(1).map(r=>Number(r[0]));
        const out={};
        for (let c=1;c<u.columnCount;c++) {
          if (!h[c]) continue;
          out[h[c]]={x:xx,y:vals.slice(1).map(r=>Number(r[c]))};
        }
        return out;
      }

      qcCache = {
        Log:conv(used.Log),
        Baseline:conv(used.Baseline),
        Corrected:conv(used.Corrected),
        Smoothed:conv(used.Smoothed)
      };
    });

    rebuildSpectraSelectors();
    status(`Loaded ${sampleNames.length} spectra.`, "ok");
  } catch(e) {
    console.error(e);
    status(e.message || String(e), "error");
  }
}

function rebuildSpectraSelectors() {
  const list=document.getElementById("sampleList");
  const previous=new Set(selectedSamples());
  list.innerHTML="";

  sampleNames.forEach(name=>{
    const row=document.createElement("label");
    row.className="sample-row";
    const cb=document.createElement("input");
    cb.type="checkbox"; cb.className="sample-toggle"; cb.value=name;
    cb.checked=previous.size ? previous.has(name) : true;
    cb.addEventListener("change",renderMain);
    const span=document.createElement("span"); span.textContent=name;
    row.append(cb,span); list.appendChild(row);
  });

  const qc=document.getElementById("qcSample");
  const prev=qc.value; qc.innerHTML="";
  sampleNames.forEach(name=>{
    const o=document.createElement("option"); o.value=name; o.textContent=name; qc.appendChild(o);
  });
  if (sampleNames.includes(prev)) qc.value=prev;
}

async function readFitResults() {
  try {
    await Excel.run(async ctx=>{
      const sheets=ctx.workbook.worksheets;
      const cn=["Fit_Curves","Peak_Details","Peak_Areas","Fit_Quality"];
      const ws={};
      cn.forEach(n=>{ ws[n]=sheets.getItemOrNullObject(n); ws[n].load("isNullObject"); });
      await ctx.sync();

      if (ws.Fit_Curves.isNullObject) {
        fitCurveCache={}; peakDetails=[]; peakAreas=[]; fitQuality=[];
        return;
      }

      const used={};
      cn.forEach(n=>{
        if (!ws[n].isNullObject) {
          used[n]=ws[n].getUsedRange();
          used[n].load("values,rowCount,columnCount");
        }
      });
      await ctx.sync();

      // Fit curves.
      const vals=used.Fit_Curves.values;
      const headers=vals[0].map(v=>String(v??""));
      const x=vals.slice(1).map(r=>Number(r[0]));
      fitCurveCache={};

      for (let c=1;c<headers.length;c++) {
        const parts=headers[c].split("|");
        if (parts.length!==2) continue;
        const sample=parts[0], kind=parts[1];
        if (!fitCurveCache[sample]) fitCurveCache[sample]={x,series:{}};
        fitCurveCache[sample].series[kind]=vals.slice(1).map(r=>Number(r[c]));
      }

      peakDetails = used.Peak_Details ? tableObjects(used.Peak_Details.values) : [];
      peakAreas = used.Peak_Areas ? tableObjects(used.Peak_Areas.values) : [];
      fitQuality = used.Fit_Quality ? tableObjects(used.Fit_Quality.values) : [];
    });

    rebuildFitSelector();
  } catch(e) {
    console.error(e);
  }
}

function tableObjects(values) {
  if (!values || !values.length) return [];
  const h=values[0].map(v=>String(v??""));
  return values.slice(1)
    .filter(r=>r.some(v=>v!==null && v!==""))
    .map(r=>{
      const o={}; h.forEach((k,i)=>o[k]=r[i]); return o;
    });
}

function rebuildFitSelector() {
  const sel=document.getElementById("fitSample");
  const prev=sel.value; sel.innerHTML="";
  Object.keys(fitCurveCache).forEach(name=>{
    const o=document.createElement("option"); o.value=name; o.textContent=name; sel.appendChild(o);
  });
  if (fitCurveCache[prev]) sel.value=prev;
}

function renderAll() {
  renderSpectra();
  renderFit();
  renderAreas();
}

function renderSpectra() { renderMain(); renderQc(); }

function renderMain() {
  const stage=document.getElementById("stage").value;
  const traces=selectedSamples().map(name=>({
    x:spectraCache[name]?.x||[], y:spectraCache[name]?.y||[],
    type:"scatter",mode:"lines",name
  }));
  Plotly.react("mainPlot",traces,{
    margin:{l:55,r:20,t:38,b:50},
    title:{text:`${stage} FTIR spectra`,font:{size:15}},
    xaxis:{title:"Wavenumber (cm⁻¹)",autorange:"reversed"},
    yaxis:{title:stage==="Normalized"?"Normalized intensity":"Intensity"},
    legend:{orientation:"h"},hovermode:"x unified"
  },{responsive:true,displaylogo:false,scrollZoom:true});
}

function renderQc() {
  const sample=document.getElementById("qcSample").value;
  if (!sample) return;
  const order=["Log","Baseline","Corrected","Smoothed"];
  const labels={Log:"-log(R)",Baseline:"Baseline",Corrected:"Corrected",Smoothed:"Smoothed"};
  const traces=order.filter(k=>qcCache[k]?.[sample]).map(k=>({
    x:qcCache[k][sample].x,y:qcCache[k][sample].y,type:"scatter",mode:"lines",name:labels[k]
  }));
  Plotly.react("qcPlot",traces,{
    margin:{l:55,r:20,t:38,b:50},
    title:{text:`Processing QC — ${sample}`,font:{size:15}},
    xaxis:{title:"Wavenumber (cm⁻¹)",autorange:"reversed"},
    yaxis:{title:"Intensity"},legend:{orientation:"h"},hovermode:"x unified"
  },{responsive:true,displaylogo:false,scrollZoom:true});
}

function renderFit() {
  const sample=document.getElementById("fitSample").value;
  const obj=fitCurveCache[sample];
  if (!obj) {
    Plotly.react("fitPlot",[],{title:"Run fit_ftir_peaks in Python to generate fitting results."});
    document.getElementById("peakTable").innerHTML="";
    return;
  }

  const s=obj.series;
  const traces=[];
  if (s.Observed) traces.push({x:obj.x,y:s.Observed,type:"scatter",mode:"lines",name:"Observed",line:{width:2}});
  if (s.TotalFit) traces.push({x:obj.x,y:s.TotalFit,type:"scatter",mode:"lines",name:"Total fit",line:{width:3}});
  if (document.getElementById("showComponents").checked) {
    Object.keys(s).filter(k=>!["Observed","TotalFit","Residual"].includes(k)).forEach(k=>{
      traces.push({x:obj.x,y:s[k],type:"scatter",mode:"lines",name:k});
    });
  }

  Plotly.react("fitPlot",traces,{
    margin:{l:55,r:20,t:42,b:50},
    title:{text:`Gaussian deconvolution — ${sample}`,font:{size:15}},
    xaxis:{title:"Wavenumber (cm⁻¹)",autorange:"reversed"},
    yaxis:{title:"Intensity"},hovermode:"x unified"
  },{responsive:true,displaylogo:false,scrollZoom:true});

  const rows=peakDetails.filter(r=>String(r.Sample)===sample);
  document.getElementById("peakTable").innerHTML=makeTable(
    rows,
    ["PeakLabel","Center_cm-1","Height","FWHM_cm-1","Area_FitWindow","Area_FullGaussian"]
  );
}

function renderAreas() {
  if (!peakAreas.length) {
    Plotly.react("areaPlot",[],{title:"Run fit_ftir_peaks in Python to generate peak areas."});
    document.getElementById("qualityTable").innerHTML="";
    return;
  }

  const mode=document.getElementById("areaMetric").value;
  const keys=Object.keys(peakAreas[0]).filter(k=>k.endsWith("_Area"));
  const traces=keys.map(k=>{
    const label=k.replace("_Area","");
    const y=peakAreas.map(r=>{
      const val=Number(r[k]||0);
      if (mode==="absolute") return val;
      const total=Number(r.Total_Component_Area||0);
      return total ? 100*val/total : 0;
    });
    return {type:"bar",name:label,x:peakAreas.map(r=>String(r.Sample)),y};
  });

  Plotly.react("areaPlot",traces,{
    barmode:"stack",
    margin:{l:60,r:20,t:42,b:80},
    title:{text:mode==="absolute"?"Gaussian peak areas":"Peak area fractions",font:{size:15}},
    xaxis:{title:"Sample"},
    yaxis:{title:mode==="absolute"?"Integrated area":"Area fraction (%)"}
  },{responsive:true,displaylogo:false});

  document.getElementById("qualityTable").innerHTML=makeTable(
    fitQuality,
    ["Sample","Success","R2","RMSE","MAE","NFEV"]
  );
}

function makeTable(rows, cols) {
  if (!rows.length) return "<p>No data.</p>";
  const fmt=v=>{
    if (typeof v==="number") return Number.isFinite(v) ? v.toPrecision(6) : "";
    return String(v??"");
  };
  return `<table><thead><tr>${cols.map(c=>`<th>${c}</th>`).join("")}</tr></thead>
    <tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${fmt(r[c])}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}

async function registerWatchers() {
  try {
    await Excel.run(async ctx=>{
      const sheets=ctx.workbook.worksheets;
      const names=["FTIR_Status","Fit_Status"];
      for (const name of names) {
        const ws=sheets.getItemOrNullObject(name);
        ws.load("isNullObject");
        await ctx.sync();
        if (!ws.isNullObject) {
          ws.onChanged.add(()=>{
            if (!document.getElementById("autoRefresh").checked) return;
            if (timer) clearTimeout(timer);
            timer=setTimeout(async()=>{
              await readSpectra();
              await readFitResults();
              renderAll();
            },700);
          });
        }
      }
      await ctx.sync();
    });
  } catch(e) { console.error(e); }
}
