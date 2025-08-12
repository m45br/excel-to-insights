/* global XLSX, vegaEmbed */
const { useState, useRef, useEffect } = React;

function prettyNum(n){
  if (n === null || n === undefined || Number.isNaN(n)) return '—';
  const abs = Math.abs(n);
  if (abs >= 1e9) return (n/1e9).toFixed(1)+'B';
  if (abs >= 1e6) return (n/1e6).toFixed(1)+'M';
  if (abs >= 1e3) return (n/1e3).toFixed(1)+'K';
  return String(n);
}

function inferType(values){
  const lower = v => (typeof v === 'string' ? v.trim().toLowerCase() : v);
  const isBool = values.every(v => v === true || v === false || ['true','false','yes','no','y','n','1','0'].includes(lower(v)));
  if (isBool) return 'boolean';
  const dateCount = values.slice(0, 200).filter(v => {
    if (v == null || v === '') return false;
    const d = new Date(v);
    return !Number.isNaN(d.getTime());
  }).length;
  if (dateCount > Math.min(10, Math.ceil(values.length * 0.2))) return 'date';
  const numCount = values.slice(0, 200).filter(v => {
    if (v == null || v === '') return false;
    if (typeof v === 'number') return true;
    if (typeof v === 'string'){
      const s = v.replace(/[%,$₹€£\s]/g, '');
      return s !== '' && !isNaN(Number(s));
    }
    return false;
  }).length;
  if (numCount > Math.min(10, Math.ceil(values.length * 0.4))) return 'number';
  return 'string';
}

function coerce(values, type){
  if (type === 'boolean'){
    return values.map(v => {
      const s = (typeof v === 'string' ? v.trim().toLowerCase() : v);
      if (s === true || s === 'true' || s === 'yes' || s === 'y' || s === '1') return true;
      if (s === false || s === 'false' || s === 'no' || s === 'n' || s === '0') return false;
      return null;
    });
  }
  if (type === 'date'){
    return values.map(v => {
      const d = new Date(v);
      return isNaN(d) ? null : d.toISOString();
    });
  }
  if (type === 'number'){
    return values.map(v => {
      if (typeof v === 'number') return v;
      if (typeof v === 'string'){
        const s = v.replace(/[%,$₹€£\s]/g, '');
        const num = Number(s);
        return isNaN(num) ? null : num;
      }
      return null;
    });
  }
  return values.map(v => (v == null || v === '' ? null : String(v)));
}

function summarizeColumn(values, type){
  const clean = values.filter(v => v != null);
  const summary = { type, count: values.length, missing: values.length - clean.length };
  if (type === 'number'){
    const nums = clean;
    const min = Math.min(...nums), max = Math.max(...nums);
    const mean = nums.reduce((a,b)=>a+b,0)/nums.length;
    summary.min = min; summary.max = max; summary.mean = mean;
  }
  if (type === 'string' || type === 'boolean'){
    const freq = new Map();
    clean.forEach(v => freq.set(v, (freq.get(v)||0)+1));
    const top = [...freq.entries()].sort((a,b)=>b[1]-a[1]).slice(0,3);
    summary.top = top;
  }
  if (type === 'date'){
    const dates = clean.map(d=>new Date(d).getTime());
    summary.min = new Date(Math.min(...dates)).toISOString();
    summary.max = new Date(Math.max(...dates)).toISOString();
  }
  return summary;
}

function dataProfile(table){
  const profile = { rows: table.length, cols: 0, columns: {} };
  if (!table.length) return profile;
  const cols = Object.keys(table[0]);
  profile.cols = cols.length;
  cols.forEach(col => {
    const values = table.map(r=>r[col]);
    const type = inferType(values);
    const coerced = coerce(values, type);
    for(let i=0;i<table.length;i++){ table[i][col] = coerced[i]; }
    profile.columns[col] = summarizeColumn(coerced, type);
  });
  return profile;
}

function detectVisuals(profile){
  const cols = Object.entries(profile.columns);
  const dates = cols.filter(([k,v]) => v.type === 'date').map(([k])=>k);
  const nums = cols.filter(([k,v]) => v.type === 'number').map(([k])=>k);
  const cats = cols.filter(([k,v]) => v.type === 'string' || v.type === 'boolean').map(([k])=>k);

  const visuals = [];
  if (dates.length && nums.length){
    visuals.push({
      kind:'timeseries',
      title: `${nums[0]} over time`,
      spec: (dataName) => ({
        $schema: 'https://vega.github.io/schema/vega-lite/v5.json',
        width: 'container', height: 260,
        data: { name: dataName },
        mark: { type: 'line', interpolate: 'monotone' },
        encoding: {
          x: { field: dates[0], type: 'temporal', title: dates[0] },
          y: { field: nums[0], type: 'quantitative', title: nums[0] }
        }
      }),
      meta: `Auto: detected date (${dates[0]}) + metric (${nums[0]}).`
    });
  }
  if (cats.length && nums.length){
    visuals.push({
      kind:'bar',
      title: `${nums[0]} by ${cats[0]}`,
      spec: (dataName) => ({
        $schema: 'https://vega.github.io/schema/vega-lite/v5.json',
        width: 'container', height: 260,
        data: { name: dataName },
        mark: { type: 'bar' },
        encoding: {
          x: { field: cats[0], type: 'nominal', sort: '-y', title: cats[0] },
          y: { aggregate: 'sum', field: nums[0], type: 'quantitative', title: `Sum of ${nums[0]}` }
        }
      }),
      meta: `Auto: detected category (${cats[0]}) + metric (${nums[0]}).`
    });
  }
  if (nums.length){
    visuals.push({
      kind:'histogram',
      title: `Distribution of ${nums[0]}`,
      spec: (dataName) => ({
        $schema: 'https://vega.github.io/schema/vega-lite/v5.json',
        width: 'container', height: 260,
        data: { name: dataName },
        mark: 'bar',
        encoding: {
          x: { bin: true, field: nums[0], type: 'quantitative' },
          y: { aggregate: 'count', type: 'quantitative', title: 'Count' }
        }
      }),
      meta: `Auto: numeric distribution.`
    });
  }
  if (nums.length >= 2){
    visuals.push({
      kind:'scatter',
      title: `${nums[1]} vs ${nums[0]}`,
      spec: (dataName) => ({
        $schema: 'https://vega.github.io/schema/vega-lite/v5.json',
        width: 'container', height: 260,
        data: { name: dataName },
        mark: { type: 'point' },
        encoding: {
          x: { field: nums[0], type: 'quantitative' },
          y: { field: nums[1], type: 'quantitative' }
        }
      }),
      meta: `Auto: relationship between two metrics.`
    });
  }
  return visuals.slice(0, 6);
}

function ExportButtons({ visuals, vegaViews, fileStem, hasPptx }){
  const exportPPT = async () => {
    if (!hasPptx) {
      alert('PowerPoint export is unavailable (PptxGenJS could not load). PDF export still works.');
      return;
    }
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    pptx.addSlide().addText(
      [{ text: 'Insights', options: { fontSize: 28, bold: true } },
       { text: `\\n${fileStem}`, options: { fontSize: 16, color: '666666' } },
       { text: `\\nGenerated ${new Date().toLocaleString()}`, options: { fontSize: 12, color: '888888' } }],
      { x:0.5, y:1.5, w:9, h:3 }
    );
    for (let i=0;i<visuals.length;i++){
      const slide = pptx.addSlide();
      const title = visuals[i].title || `Visual ${i+1}`;
      slide.addText(title, { x:0.6, y:0.4, w:9, h:0.6, fontSize: 20, bold: true });
      const view = vegaViews[i];
      if (!view) continue;
      const dataUrl = await view.toImageURL('png', 2);
      slide.addImage({ data: dataUrl, x:0.6, y:1.1, w:8.8, h:4.6 });
      slide.addText(`Source: ${fileStem}`, { x:0.6, y:6, w:8.8, h:0.3, fontSize: 10, color: '666666' });
    }
    const ts = new Date();
    const pad = (v) => String(v).padStart(2,'0');
    const fname = `${fileStem}-Insights-${ts.getFullYear()}${pad(ts.getMonth()+1)}${pad(ts.getDate())}-${pad(ts.getHours())}${pad(ts.getMinutes())}.pptx`;
    await pptx.writeFile({ fileName: fname });
  };

  const exportPDF = () => {
    const w = window.open('', '_blank');
    const head = `
      <title>${fileStem} – PDF Export</title>
      <style>
        body { font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial; color:#111; margin:0; }
        .page { page-break-after: always; padding: 36px; }
        h1 { font-size: 22px; margin: 0 0 8px 0; }
        .byline { font-size: 12px; color:#555; margin-bottom: 12px; }
        .chart { width: 100%; height: auto; }
      </style>
    `;
    let body = '';
    visuals.forEach((vis, i) => {
      const el = document.getElementById(`vis-${i}`);
      const svg = el?.querySelector('svg');
      const svgMarkup = svg ? svg.outerHTML : '<div>Chart not available</div>';
      body += `
        <section class="page">
          <h1>${vis.title || `Visual ${i+1}`}</h1>
          <div class="byline">Source: ${fileStem} • Exported ${new Date().toLocaleString()}</div>
          ${svgMarkup}
        </section>
      `;
    });
    w.document.write(`<!doctype html><html><head>${head}</head><body>${body}</body></html>`);
    w.document.close();
    setTimeout(() => w.print(), 350);
  };

  return React.createElement('div', { className: 'toolbar-right no-print', role: 'group', 'aria-label': 'Export options' },
    React.createElement('button', { className: 'btn primary', onClick: exportPPT }, hasPptx ? '⤓ Export PPT' : '⤓ Export PPT (offline)'),
    React.createElement('button', { className: 'btn', onClick: exportPDF }, '⤓ Export PDF')
  );
}

function App(){
  const [fileName, setFileName] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [activeSheet, setActiveSheet] = useState(null);
  const [profile, setProfile] = useState(null);
  const [visuals, setVisuals] = useState([]);
  const [vegaViews, setVegaViews] = useState([]);
  const [hasPptx, setHasPptx] = useState(!!window.PptxGenJS);
  const [banner, setBanner] = useState(null);

  const inputRef = useRef(null);
  const dropRef = useRef(null);

  useEffect(()=>{
    // Try to load PptxGenJS asynchronously. If blocked, we proceed without it.
    if (!window.PptxGenJS) {
      window.ensurePptx && window.ensurePptx((ok)=>{
        setHasPptx(!!ok);
        if (!ok) setBanner('PowerPoint export is disabled: PptxGenJS could not be loaded from CDNs. PDF export still works.');
      });
    }
  }, []);

  useEffect(()=>{
    const dz = dropRef.current;
    if (!dz) return;
    function prevent(e){ e.preventDefault(); e.stopPropagation(); }
    function enter(){ dz.classList.add('drag'); }
    function leave(){ dz.classList.remove('drag'); }
    dz.addEventListener('dragover', prevent);
    dz.addEventListener('dragenter', enter);
    dz.addEventListener('dragleave', leave);
    dz.addEventListener('drop', prevent);
    return () => {
      dz.removeEventListener('dragover', prevent);
      dz.removeEventListener('dragenter', enter);
      dz.removeEventListener('dragleave', leave);
      dz.removeEventListener('drop', prevent);
    };
  }, []);

  function handleFiles(files){
    if (!files || !files.length) return;
    const file = files[0];
    if (file.size > 50 * 1024 * 1024){
      alert('File is larger than 50MB. Please use a smaller file.');
      return;
    }
    setFileName(file.name.replace(/\.[^.]+$/, ''));
    const reader = new FileReader();
    reader.onload = (e) => {
      try{
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const names = workbook.SheetNames || [];
        setSheetNames(names);
        const first = names[0];
        setActiveSheet(first);
        const ws = workbook.Sheets[first];
        const json = XLSX.utils.sheet_to_json(ws, { defval: null });
        hydrate(json);
      }catch(err){
        console.error(err);
        alert('We could not parse this file. Please check the format.');
      }
    };
    if (/\.(csv)$/i.test(file.name)){
      const readerText = new FileReader();
      readerText.onload = (e2) => {
        const text = e2.target.result;
        const ws = XLSX.read(text, { type: 'string' }).Sheets.Sheet1;
        const json = XLSX.utils.sheet_to_json(ws, { defval: null });
        hydrate(json);
      };
      readerText.readAsText(file);
      return;
    }
    reader.readAsArrayBuffer(file);
  }

  function hydrate(json){
    const clone = json.map(r=>({ ...r }));
    const prof = dataProfile(clone);
    setProfile(prof);
    const vis = detectVisuals(prof);
    setVisuals(vis);
    setTimeout(()=> renderVisuals(vis, clone), 0);
  }

  async function renderVisuals(vis, data){
    const views = [];
    for (let i=0;i<vis.length;i++){
      const el = document.getElementById(`vis-${i}`);
      if (!el) continue;
      el.innerHTML = '';
      const spec = vis[i].spec('t');
      const embedSpec = { ...spec, data: { name: 't' }, autosize: { type: 'fit', contains: 'padding' } };
      const result = await vegaEmbed(el, embedSpec, { actions: false, renderer: 'svg' });
      result.view.insert('t', data).run();
      views.push(result.view);
    }
    setVegaViews(views);
  }

  function onSheetChange(e){
    const name = e.target.value;
    setActiveSheet(name);
    const file = inputRef.current.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const ws = workbook.Sheets[name];
      const json = XLSX.utils.sheet_to_json(ws, { defval: null });
      hydrate(json);
    };
    reader.readAsArrayBuffer(file);
  }

  const fileStem = fileName || 'Workbook';

  return (
    <div className="container">
      {banner && <div className="banner" role="status">{banner}</div>}
      <div className="brand no-print" aria-hidden="true">
        <div className="logo" aria-hidden="true"></div>
        <h1>Excel → Insights</h1>
      </div>

      {!profile && (
        <div className="upload-card" role="region" aria-label="Upload data">
          <div
            id="dropzone"
            ref={dropRef}
            className="dropzone"
            tabIndex={0}
            onDrop={(e)=>{ e.preventDefault(); handleFiles(e.dataTransfer.files); }}
            onKeyDown={(e)=>{ if (e.key === 'Enter') inputRef.current?.click(); }}
            aria-label="Drag and drop your Excel or CSV file here"
          >
            <p style={{fontSize:'18px', fontWeight:600, marginBottom:'8px'}}>Drop your Excel or CSV</p>
            <p className="muted" style={{marginTop:0}}>or</p>
            <div style={{marginTop:'12px'}}>
              <label className="file-label btn primary" htmlFor="fileInput" aria-label="Upload Excel (.xlsx/.xls) or CSV">
                ⬆ Upload Excel (.xlsx/.xls/.csv)
              </label>
              <input
                id="fileInput"
                ref={inputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e)=>handleFiles(e.target.files)}
              />
            </div>
            <p className="muted" style={{marginTop:'16px'}}>Max 50MB • Client-side parsing • <span title="We do not store your data.">Privacy-first</span></p>
          </div>
        </div>
      )}

      {profile && (
        <>
          <div className="toolbar no-print" role="toolbar" aria-label="Main controls">
            <div className="toolbar-left">
              <span className="chip" role="status" aria-live="polite">Rows: <strong>&nbsp;{prettyNum(profile.rows)}</strong></span>
              <span className="chip">Columns: <strong>&nbsp;{prettyNum(profile.cols)}</strong></span>
              {sheetNames.length > 1 && (
                <label className="chip" aria-label="Sheet selector">
                  Sheet:&nbsp;
                  <select onChange={onSheetChange} value={activeSheet} aria-label="Choose sheet">
                    {sheetNames.map(n => <option key={n} value={n}>{n}</option>)}
                  </select>
                </label>
              )}
            </div>
            <ExportButtons visuals={visuals} vegaViews={vegaViews} fileStem={fileStem} hasPptx={hasPptx} />
          </div>

          <section className="card" aria-labelledby="profile-title">
            <h3 id="profile-title">Quick data profile</h3>
            <div className="muted meta">We inferred types, handled missing values, and profiled columns.</div>
            <div style={{overflowX:'auto', marginTop:'12px'}}>
              <table className="summary-table" role="table" aria-label="Data profile summary">
                <thead>
                  <tr>
                    <th>Column</th><th>Type</th><th>Missing</th><th>Min</th><th>Max</th><th>Mean / Top</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(profile.columns).map(([name, s]) => (
                    <tr key={name}>
                      <td>{name}</td>
                      <td>{s.type}</td>
                      <td>{prettyNum(s.missing)}</td>
                      <td>{s.min ? (s.type==='number' ? prettyNum(s.min) : new Date(s.min).toLocaleString()) : '—'}</td>
                      <td>{s.max ? (s.type==='number' ? prettyNum(s.max) : new Date(s.max).toLocaleString()) : '—'}</td>
                      <td>
                        {s.type==='number' ? (s.mean ? prettyNum(s.mean) : '—')
                          : (s.top ? s.top.map(([v,c])=>`${v} (${c})`).join(', ') : '—')}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section className="grid" aria-label="Suggested visuals">
            {visuals.map((v, i) => (
              <article key={i} className="card" aria-labelledby={`title-${i}`}>
                <h3 id={`title-${i}`}>{v.title}</h3>
                <div className="muted meta">{v.meta}</div>
                <div id={`vis-${i}`} className="vis-container" role="img" aria-label={v.title}></div>
              </article>
            ))}
          </section>
        </>
      )}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
