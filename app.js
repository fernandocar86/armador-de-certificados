/* Interfaz: conecta la página con Certs (generación) y PptxPdf (conversión a PDF). */
(function () {
  const $ = s => document.querySelector(s);
  let cats = [], headers = [], rows = [], uid = 0, generated = [], pending = [];

  const tick = () => new Promise(r => setTimeout(r, 0));
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  }
  async function zipBlob(files) {
    const z = new JSZip();
    files.forEach(f => z.file(f.path, f.data));
    return z.generateAsync({ type: 'blob' });
  }

  /* ---------- 1. categorías ---------- */
  function addCat() {
    const c = { id: uid++, name: '', zip: null, vars: [], map: {} };
    cats.push(c);
    const d = document.createElement('div');
    d.className = 'row';
    d.innerHTML = '<input type="text" placeholder="Nombre de la categoría"><input type="file" accept=".pptx"><span class="vars"></span><button title="Quitar">✕</button>';
    const [n, f, v, x] = d.children;
    n.oninput = () => { c.name = n.value; renderMaps(); };
    f.onchange = async () => {
      try {
        c.zip = await JSZip.loadAsync(f.files[0]);
        c.vars = await Certs.detectVars(c.zip);
        v.innerHTML = c.vars.length ? `<span class="ok">Variables: ${c.vars.join(', ')}</span>` : '<span class="er">Sin variables xxx/yyy/zzz</span>';
      } catch (e) { c.zip = null; c.vars = []; v.innerHTML = '<span class="er">Archivo inválido</span>'; }
      renderMaps();
    };
    x.onclick = () => { cats = cats.filter(k => k !== c); d.remove(); renderMaps(); };
    $('#cats').append(d);
  }
  $('#add').onclick = addCat;
  addCat();

  /* ---------- 2. Excel, asignación y generación ---------- */
  $('#xl').onchange = async e => {
    const wb = XLSX.read(await e.target.files[0].arrayBuffer());
    const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', raw: false });
    headers = (data[0] || []).map((h, i) => String(h).trim() || `Columna ${i + 1}`);
    rows = data.slice(1).filter(r => r.some(c => String(c).trim()));
    $('#xlinfo').textContent = `${rows.length} filas, ${headers.length} columnas (${headers.join(' | ')})`;
    $('#catcol').innerHTML = headers.map((h, i) => `<option value="${i}">${h}</option>`).join('');
    $('#catcol').selectedIndex = Math.min(1, headers.length - 1);
    $('#s2').classList.remove('hide');
    renderMaps();
  };
  $('#catcol').onchange = renderMaps;

  function renderMaps() {
    const box = $('#maps');
    if (!headers.length) return;
    const cc = +$('#catcol').value, others = headers.map((h, i) => i).filter(i => i !== cc);
    box.innerHTML = '';
    cats.filter(c => c.vars.length).forEach(c => {
      const prev = c.map || {};
      c.map = {};
      let html = `<b>${c.name || '(sin nombre)'}</b><div class="map">`;
      c.vars.forEach((k, j) => {
        const def = prev[k] !== undefined && prev[k] !== cc ? prev[k] : (others[j] ?? others[0]);
        c.map[k] = def;
        html += `<span>${k} →</span><select data-c="${c.id}" data-v="${k}">${headers.map((h, i) => `<option value="${i}" ${i === def ? 'selected' : ''}>${h}</option>`).join('')}</select>`;
      });
      box.insertAdjacentHTML('beforeend', html + '</div>');
    });
    box.querySelectorAll('select').forEach(s => { s.onchange = () => { cats.find(c => c.id == s.dataset.c).map[s.dataset.v] = +s.value; }; });
  }

  $('#go').onclick = async () => {
    const log = $('#log'); log.textContent = '';
    const say = m => { log.textContent += m + '\n'; };
    const ready = cats.filter(c => c.name.trim() && c.zip && c.vars.length);
    if (!ready.length) return say('Agregá al menos una categoría con nombre y un modelo con variables.');
    $('#go').disabled = true; $('#dl').classList.add('hide');
    try {
      const res = await Certs.generate({ cats: ready, rows, catCol: +$('#catcol').value,
        onProgress: (i, n) => { log.dataset.p = `Procesando ${i}/${n}…`; } });
      generated = res.files;
      say(`✅ Certificados generados: ${generated.length}`);
      if (res.skipped.length) say(`⚠️ Sin categoría reconocida (${res.skipped.length}):\n - ` + res.skipped.join('\n - '));
      if (generated.length) $('#dl').classList.remove('hide');
    } catch (e) { say('Error: ' + e.message); }
    $('#go').disabled = false;
  };
  $('#dl-pptx').onclick = async () => download(await zipBlob(generated), 'certificados_pptx.zip');
  $('#dl-pdf').onclick = async () => {
    const b = $('#dl-pdf'); b.disabled = true;
    await convertAll(generated, $('#log'), 'certificados_pdf.zip');
    b.disabled = false;
  };

  /* ---------- 3. PPTX -> PDF ---------- */
  async function collect(fileList) {
    const out = [];
    for (const f of fileList) {
      const rel = (f.webkitRelativePath || f.name).replace(/^[^/]+\//, f.webkitRelativePath ? '' : '');
      if (/\.pptx$/i.test(f.name)) out.push({ path: rel || f.name, data: new Uint8Array(await f.arrayBuffer()) });
      else if (/\.zip$/i.test(f.name)) {
        const z = await JSZip.loadAsync(f);
        for (const name of Object.keys(z.files))
          if (/\.pptx$/i.test(name) && !z.files[name].dir && !/(^|\/)(~\$|__MACOSX)/.test(name))
            out.push({ path: name, data: await z.files[name].async('uint8array') });
      }
    }
    return out;
  }
  async function pick(e) {
    pending = await collect(e.target.files);
    $('#convinfo').textContent = pending.length ? `${pending.length} archivo(s) .pptx listos para convertir.` : 'No se encontraron archivos .pptx.';
    $('#conv').disabled = !pending.length;
  }
  $('#folder').onchange = pick; $('#files').onchange = pick;

  async function convertAll(items, logEl, zipName, bar) {
    logEl.textContent = ''; const errors = [], pdfs = [];
    if (bar) { bar.classList.remove('hide'); bar.max = items.length; bar.value = 0; }
    for (let i = 0; i < items.length; i++) {
      logEl.textContent = `Convirtiendo ${i + 1}/${items.length}: ${items[i].path}`;
      if (bar) bar.value = i;
      await tick();
      try { pdfs.push({ path: items[i].path.replace(/\.pptx$/i, '.pdf'), data: await PptxPdf.convert(items[i].data) }); }
      catch (e) { errors.push(`${items[i].path}: ${e.message}`); }
    }
    if (bar) bar.value = items.length;
    logEl.textContent = `✅ PDF generados: ${pdfs.length}` + (errors.length ? `\n⚠️ Con error (${errors.length}):\n - ` + errors.join('\n - ') : '');
    if (pdfs.length) download(await zipBlob(pdfs), zipName);
  }
  $('#conv').onclick = async () => {
    $('#conv').disabled = true;
    await convertAll(pending, $('#convlog'), 'certificados_pdf.zip', $('#bar'));
    $('#conv').disabled = false;
  };
})();
