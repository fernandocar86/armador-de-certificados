/* Lógica de generación de certificados .pptx (sin interfaz). Expone window.Certs */
(function () {
  const VARS = ['xxx', 'yyy', 'zzz'];

  const norm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const safe = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '_').slice(0, 80) || 'sin_nombre';
  // La plantilla suele envolver el título entre comillas dobles: las del valor pasan a simples.
  const fixQuotes = t => t.replace(/[“”]/g, m => (m === '“' ? '‘' : '’')).replace(/"/g, "'");

  const slidePaths = zip => Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p));
  const parse = s => new DOMParser().parseFromString(s, 'application/xml');
  const paraText = p => [...p.getElementsByTagName('a:t')].map(t => t.textContent).join('');

  /** Devuelve las variables (xxx/yyy/zzz) presentes en el modelo. */
  async function detectVars(zip) {
    const found = new Set();
    for (const path of slidePaths(zip)) {
      const doc = parse(await zip.file(path).async('string'));
      for (const p of doc.getElementsByTagName('a:p')) {
        const t = paraText(p);
        VARS.forEach(k => t.includes(k) && found.add(k));
      }
    }
    return VARS.filter(k => found.has(k));
  }

  /** Reemplaza variables en el XML de una diapositiva conservando el formato. */
  function fillXml(xml, vals) {
    const doc = parse(xml);
    const apply = s => { for (const k in vals) s = s.split(k).join(vals[k]); return s; };
    for (const p of doc.getElementsByTagName('a:p')) {
      const ts = [...p.getElementsByTagName('a:t')];
      ts.forEach(t => { t.textContent = apply(t.textContent); });
      // variable partida entre varios runs: se unifica el párrafo en el primero
      if (ts.length > 1 && VARS.some(k => k in vals && paraText(p).includes(k))) {
        ts[0].textContent = apply(paraText(p));
        ts.slice(1).forEach(t => { t.textContent = ''; });
      }
    }
    const out = new XMLSerializer().serializeToString(doc);
    return out.startsWith('<?xml') ? out : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out;
  }

  /**
   * cats: [{name, zip, vars, map:{xxx:colIdx,...}}], rows: array de arrays.
   * Devuelve {files:[{path,data}], skipped:[string]}.
   */
  async function generate({ cats, rows, catCol, onProgress }) {
    const files = [], skipped = [], used = {};
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const parts = String(row[catCol]).split(/[;,\/]|\sy\s/).map(norm).filter(Boolean);
      const hits = cats.filter(c => parts.some(p => p.includes(norm(c.name))));
      if (!hits.length) { skipped.push(`${row[0]} (“${row[catCol]}”)`); continue; }
      for (const c of hits) {
        const vals = {};
        c.vars.forEach(k => { vals[k] = fixQuotes(String(row[c.map[k]] ?? '').trim()); });
        const z = await JSZip.loadAsync(await c.zip.generateAsync({ type: 'uint8array' }));
        for (const p of slidePaths(z)) z.file(p, fillXml(await z.file(p).async('string'), vals));
        let base = `${safe(row[c.map[c.vars[0]]])}_${safe(c.name)}`;
        used[base] = (used[base] || 0) + 1;
        if (used[base] > 1) base += '_' + used[base];
        files.push({ path: `${safe(c.name)}/${base}.pptx`, data: await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) });
      }
      onProgress && onProgress(i + 1, rows.length);
    }
    return { files, skipped };
  }

  window.Certs = { VARS, norm, safe, detectVars, generate };
})();
