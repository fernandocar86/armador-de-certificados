/*
 * Conversor PPTX -> PDF en el navegador. Expone window.PptxPdf.convert(data) -> Uint8Array.
 * Dibuja cada diapositiva en un canvas (imágenes, textos con formato, formas y líneas)
 * y arma un PDF con jsPDF. Es una conversión aproximada pensada para certificados.
 */
(function () {
  const DPI = 192, JPEG_Q = 0.92;
  const E = v => v / 12700;                       // EMU -> puntos
  const parse = s => new DOMParser().parseFromString(s, 'application/xml');
  const kids = (el, n) => (el ? [...el.children].filter(c => c.tagName === n) : []);
  const kid = (el, n) => kids(el, n)[0] || null;
  const num = (el, a, d) => (el && el.getAttribute(a) != null ? +el.getAttribute(a) : d);

  const FONTS = {
    'times new roman': '"Times New Roman",Tinos,"Liberation Serif",serif',
    cambria: 'Cambria,Caladea,serif',
    georgia: 'Georgia,"Times New Roman",Tinos,serif',
    calibri: 'Calibri,Carlito,"Segoe UI",sans-serif',
    arial: 'Arial,Arimo,"Liberation Sans",sans-serif',
    helvetica: 'Helvetica,Arial,Arimo,sans-serif',
  };
  const family = f => FONTS[(f || '').toLowerCase()] || `"${f}",Arimo,Arial,sans-serif`;

  async function rels(zip, path) {
    const dir = path.replace(/[^/]*$/, ''), name = path.split('/').pop();
    const f = zip.file(`${dir}_rels/${name}.rels`), map = {};
    if (!f) return map;
    for (const r of parse(await f.async('string')).getElementsByTagName('Relationship')) {
      let t = r.getAttribute('Target');
      if (r.getAttribute('TargetMode') !== 'External') {
        const parts = (t.startsWith('/') ? t.slice(1) : dir + t).split('/'), out = [];
        for (const p of parts) p === '..' ? out.pop() : p !== '.' && out.push(p);
        t = out.join('/');
      }
      map[r.getAttribute('Id')] = t;
    }
    return map;
  }

  function themeInfo(doc) {
    const colors = {}, fonts = { major: 'Arial', minor: 'Arial' };
    if (!doc) return { colors, fonts };
    const cs = doc.getElementsByTagName('a:clrScheme')[0];
    if (cs) for (const c of cs.children) {
      const v = c.firstElementChild;
      colors[c.tagName.replace('a:', '')] = v && (v.getAttribute('val') !== 'windowText' && v.getAttribute('val') !== 'window' ? v.getAttribute('lastClr') || v.getAttribute('val') : v.getAttribute('lastClr'));
    }
    const mj = doc.getElementsByTagName('a:majorFont')[0], mn = doc.getElementsByTagName('a:minorFont')[0];
    if (mj) fonts.major = kid(mj, 'a:latin')?.getAttribute('typeface') || 'Arial';
    if (mn) fonts.minor = kid(mn, 'a:latin')?.getAttribute('typeface') || 'Arial';
    return { colors, fonts };
  }

  function colorOf(fill, theme) {      // fill: elemento a:solidFill
    const c = fill && fill.firstElementChild;
    if (!c) return null;
    let hex = c.tagName === 'a:srgbClr' ? c.getAttribute('val') : theme.colors[{ tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[c.getAttribute('val')] || c.getAttribute('val')];
    if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return '#000000';
    let [r, g, b] = [0, 2, 4].map(i => parseInt(hex.substr(i, 2), 16));
    const mod = num(kid(c, 'a:lumMod'), 'val', 100000) / 100000, off = num(kid(c, 'a:lumOff'), 'val', 0) / 100000;
    if (mod !== 1 || off) [r, g, b] = [r, g, b].map(v => Math.max(0, Math.min(255, Math.round(v * mod + 255 * off))));
    return `rgb(${r},${g},${b})`;
  }

  async function bitmap(zip, path) {
    const f = zip.file(path);
    if (!f) return null;
    try {
      const ext = path.split('.').pop().toLowerCase();
      const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' }[ext];
      if (!type) return null;
      return await createImageBitmap(new Blob([await f.async('uint8array')], { type }));
    } catch (e) { return null; }
  }

  function drawBitmap(ctx, bmp, w, h, blipFill) {
    const sr = kid(blipFill, 'a:srcRect');
    const l = num(sr, 'l', 0) / 1e5, t = num(sr, 't', 0) / 1e5, r = num(sr, 'r', 0) / 1e5, b = num(sr, 'b', 0) / 1e5;
    ctx.drawImage(bmp, bmp.width * l, bmp.height * t, bmp.width * (1 - l - r), bmp.height * (1 - t - b), 0, 0, w, h);
  }

  /* ---------- texto ---------- */
  function runStyle(rPr, base, theme) {
    const s = { ...base };
    if (rPr) {
      if (rPr.getAttribute('sz')) s.sz = +rPr.getAttribute('sz') / 100;
      if (rPr.getAttribute('b') != null) s.b = rPr.getAttribute('b') === '1';
      if (rPr.getAttribute('i') != null) s.i = rPr.getAttribute('i') === '1';
      if (rPr.getAttribute('u')) s.u = rPr.getAttribute('u') !== 'none';
      if (rPr.getAttribute('strike')) s.strike = rPr.getAttribute('strike') !== 'noStrike';
      const c = colorOf(kid(rPr, 'a:solidFill'), theme); if (c) s.color = c;
      const lat = kid(rPr, 'a:latin')?.getAttribute('typeface');
      if (lat) s.face = lat === '+mj-lt' ? theme.fonts.major : lat === '+mn-lt' ? theme.fonts.minor : lat;
    }
    return s;
  }
  const fontStr = s => `${s.i ? 'italic ' : ''}${s.b ? 'bold ' : ''}${s.sz}px ${family(s.face)}`;

  function layoutPara(ctx, p, availW, theme, defStyle) {
    const pPr = kid(p, 'a:pPr');
    const algn = pPr?.getAttribute('algn') || 'l';
    const marL = E(num(pPr, 'marL', 0)), indent = E(num(pPr, 'indent', 0));
    const ln = kid(pPr, 'a:lnSpc'), lnPct = kid(ln, 'a:spcPct'), lnPts = kid(ln, 'a:spcPts');
    const spcPts = n => num(kid(kid(pPr, n), 'a:spcPts'), 'val', 0) / 100;
    const base = runStyle(null, defStyle, theme);
    const endSt = runStyle(kid(p, 'a:endParaRPr'), base, theme);
    const items = [];
    for (const c of p.children) {
      if (c.tagName === 'a:br') items.push({ br: true, st: runStyle(kid(c, 'a:rPr'), base, theme) });
      else if (c.tagName === 'a:r' || c.tagName === 'a:fld') {
        const st = runStyle(kid(c, 'a:rPr'), base, theme), txt = kid(c, 'a:t')?.textContent || '';
        for (const tok of txt.match(/\s+|[^\s]+/g) || []) items.push({ tok, st });
      }
    }
    const lines = [{ pieces: [], w: 0 }];
    const cur = () => lines[lines.length - 1];
    const lineAvail = () => availW - marL - (lines.length === 1 ? indent : 0);
    for (const it of items) {
      if (it.br) { lines.push({ pieces: [], w: 0, st: it.st }); continue; }
      ctx.font = fontStr(it.st);
      const w = ctx.measureText(it.tok).width, ws = /^\s/.test(it.tok);
      if (!ws && cur().w + w > lineAvail() && cur().pieces.length) lines.push({ pieces: [], w: 0 });
      if (ws && !cur().pieces.length && lines.length > 1) continue;     // espacios al inicio de línea envuelta
      cur().pieces.push({ ...it, w }); cur().w += w;
    }
    for (const l of lines) {                                            // métricas y ancho sin espacios finales
      while (l.pieces.length && /^\s+$/.test(l.pieces[l.pieces.length - 1].tok)) l.w -= l.pieces.pop().w;
      let asc = 0, desc = 0;
      const sts = l.pieces.length ? l.pieces.map(x => x.st) : [l.st || endSt];
      for (const st of sts) {
        ctx.font = fontStr(st);
        const m = ctx.measureText('Hg');
        asc = Math.max(asc, m.fontBoundingBoxAscent ?? st.sz * 0.9);
        desc = Math.max(desc, m.fontBoundingBoxDescent ?? st.sz * 0.25);
      }
      let h = asc + desc;
      if (lnPct) h *= num(lnPct, 'val', 100000) / 1e5; else if (lnPts) h = num(lnPts, 'val', 1200) / 100;
      l.h = h; l.desc = desc;
    }
    return { lines, algn, marL, indent, before: spcPts('a:spcBef'), after: spcPts('a:spcAft') };
  }

  function drawText(ctx, txBody, w, h, theme) {
    const bp = kid(txBody, 'a:bodyPr');
    const L = E(num(bp, 'lIns', 91440)), T = E(num(bp, 'tIns', 45720)), R = E(num(bp, 'rIns', 91440)), B = E(num(bp, 'bIns', 45720));
    const anchor = bp?.getAttribute('anchor') || 't';
    const availW = bp?.getAttribute('wrap') === 'none' ? 1e6 : w - L - R;
    const def = { sz: 18, b: false, i: false, u: false, strike: false, color: '#000000', face: theme.fonts.minor };
    const paras = kids(txBody, 'a:p').map(p => layoutPara(ctx, p, availW, theme, def));
    let total = 0;
    paras.forEach((p, i) => { p.top = total + (i ? p.before : 0); total = p.top + p.lines.reduce((s, l) => s + l.h, 0) + p.after; });
    let y0 = anchor === 'ctr' ? T + (h - T - B - total) / 2 : anchor === 'b' ? h - B - total : T;
    ctx.textBaseline = 'alphabetic';
    for (const p of paras) {
      let y = y0 + p.top;
      p.lines.forEach((l, li) => {
        const x0 = p.marL + (li === 0 ? p.indent : 0), room = availW - x0;
        let x = L + x0 + (p.algn === 'ctr' ? (room - l.w) / 2 : p.algn === 'r' ? room - l.w : 0);
        const base = y + l.h - l.desc;
        for (const pc of l.pieces) {
          ctx.font = fontStr(pc.st); ctx.fillStyle = pc.st.color;
          ctx.fillText(pc.tok, x, base);
          if (pc.st.u || pc.st.strike) {
            ctx.fillRect(x, pc.st.u ? base + pc.st.sz * 0.1 : base - pc.st.sz * 0.3, pc.w, Math.max(0.6, pc.st.sz / 18));
          }
          x += pc.w;
        }
        y += l.h;
      });
    }
  }

  /* ---------- formas ---------- */
  function xfrmOf(spPr) {
    const x = kid(spPr, 'a:xfrm');
    if (!x) return null;
    const off = kid(x, 'a:off'), ext = kid(x, 'a:ext');
    return { x: E(num(off, 'x', 0)), y: E(num(off, 'y', 0)), w: E(num(ext, 'cx', 0)), h: E(num(ext, 'cy', 0)),
      rot: num(x, 'rot', 0) / 60000 * Math.PI / 180, fh: x.getAttribute('flipH') === '1', fv: x.getAttribute('flipV') === '1' };
  }

  async function drawTree(ctx, tree, S) {
    for (const el of tree.children) {
      const tag = el.tagName;
      if (!['p:sp', 'p:pic', 'p:cxnSp', 'p:grpSp'].includes(tag)) continue;
      if (S.skipPh && el.getElementsByTagName('p:ph').length && tag === 'p:sp') continue;   // placeholders de layout/master
      const isGrp = tag === 'p:grpSp';
      const spPr = kid(el, isGrp ? 'p:grpSpPr' : 'p:spPr');
      let xf = xfrmOf(spPr);
      if (!xf && tag === 'p:sp') {                                       // hereda posición del layout
        const ph = kid(kid(kid(el, 'p:nvSpPr'), 'p:nvPr'), 'p:ph');
        if (ph && S.layout) {
          for (const lsp of S.layout.getElementsByTagName('p:sp')) {
            const lph = kid(kid(kid(lsp, 'p:nvSpPr'), 'p:nvPr'), 'p:ph');
            if (lph && ((ph.getAttribute('type') && lph.getAttribute('type') === ph.getAttribute('type')) || (ph.getAttribute('idx') && lph.getAttribute('idx') === ph.getAttribute('idx')))) { xf = xfrmOf(kid(lsp, 'p:spPr')); break; }
          }
        }
      }
      if (!xf) continue;
      ctx.save();
      ctx.translate(xf.x + xf.w / 2, xf.y + xf.h / 2);
      ctx.rotate(xf.rot);
      ctx.translate(-xf.w / 2, -xf.h / 2);
      if (isGrp) {
        const gx = kid(spPr, 'a:xfrm'), co = kid(gx, 'a:chOff'), ce = kid(gx, 'a:chExt');
        const cw = E(num(ce, 'cx', 0)) || xf.w, ch = E(num(ce, 'cy', 0)) || xf.h;
        ctx.scale(xf.w / cw, xf.h / ch);
        ctx.translate(-E(num(co, 'x', 0)), -E(num(co, 'y', 0)));
        await drawTree(ctx, el, S);
      } else if (tag === 'p:pic') {
        const bf = kid(el, 'p:blipFill'), id = kid(bf, 'a:blip')?.getAttribute('r:embed');
        const bmp = id && S.rels[id] ? await bitmap(S.zip, S.rels[id]) : null;
        if (bmp) {
          ctx.save();
          if (xf.fh || xf.fv) { ctx.translate(xf.fh ? xf.w : 0, xf.fv ? xf.h : 0); ctx.scale(xf.fh ? -1 : 1, xf.fv ? -1 : 1); }
          drawBitmap(ctx, bmp, xf.w, xf.h, bf); ctx.restore();
        }
      } else {
        const geom = kid(spPr, 'a:prstGeom')?.getAttribute('prst') || 'rect';
        const fill = colorOf(kid(spPr, 'a:solidFill'), S.theme);
        const ln = kid(spPr, 'a:ln'), lnCol = colorOf(kid(ln, 'a:solidFill'), S.theme);
        const isLine = /line|Connector/i.test(geom) || tag === 'p:cxnSp';
        ctx.save();
        if (xf.fh || xf.fv) { ctx.translate(xf.fh ? xf.w : 0, xf.fv ? xf.h : 0); ctx.scale(xf.fh ? -1 : 1, xf.fv ? -1 : 1); }
        const cust = kid(spPr, 'a:custGeom');
        if (cust) {                                               // formas libres (marcos, ornamentos)
          for (const path of kids(kid(cust, 'a:pathLst'), 'a:path')) {
            const sx = xf.w / (E(num(path, 'w', 0)) || xf.w), sy = xf.h / (E(num(path, 'h', 0)) || xf.h);
            const pt = e => [E(num(e, 'x', 0)) * sx, E(num(e, 'y', 0)) * sy];
            ctx.beginPath();
            for (const cmd of path.children) {
              const pts = kids(cmd, 'a:pt').map(pt), t = cmd.tagName;
              if (t === 'a:moveTo') ctx.moveTo(...pts[0]);
              else if (t === 'a:lnTo') ctx.lineTo(...pts[0]);
              else if (t === 'a:cubicBezTo') ctx.bezierCurveTo(...pts[0], ...pts[1], ...pts[2]);
              else if (t === 'a:quadBezTo') ctx.quadraticCurveTo(...pts[0], ...pts[1]);
              else if (t === 'a:close') ctx.closePath();
            }
            if (fill && path.getAttribute('fill') !== 'none') { ctx.fillStyle = fill; ctx.fill(); }
            if (lnCol && path.getAttribute('stroke') !== '0') { ctx.strokeStyle = lnCol; ctx.lineWidth = Math.max(0.25, E(num(ln, 'w', 12700))); ctx.stroke(); }
          }
          ctx.restore();
          ctx.restore();
          continue;
        }
        ctx.beginPath();
        if (isLine) { ctx.moveTo(0, 0); ctx.lineTo(xf.w, xf.h); }
        else if (geom === 'ellipse') ctx.ellipse(xf.w / 2, xf.h / 2, xf.w / 2, xf.h / 2, 0, 0, 7);
        else if (geom === 'roundRect') ctx.roundRect(0, 0, xf.w, xf.h, Math.min(xf.w, xf.h) / 6);
        else ctx.rect(0, 0, xf.w, xf.h);
        if (fill && !isLine) { ctx.fillStyle = fill; ctx.fill(); }
        if (lnCol) { ctx.strokeStyle = lnCol; ctx.lineWidth = Math.max(0.25, E(num(ln, 'w', 12700))); ctx.stroke(); }
        ctx.restore();
        const tb = kid(el, 'p:txBody');
        if (tb && tag === 'p:sp') drawText(ctx, tb, xf.w, xf.h, S.theme);
      }
      ctx.restore();
    }
  }

  async function drawBackground(ctx, W, H, docs, S) {
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    for (const { doc, rels: rl } of docs) {
      const bg = doc && doc.getElementsByTagName('p:bg')[0];
      if (!bg) continue;
      const bgPr = kid(bg, 'p:bgPr');
      const c = colorOf(kid(bgPr, 'a:solidFill'), S.theme);
      if (c) { ctx.fillStyle = c; ctx.fillRect(0, 0, W, H); return; }
      const bf = kid(bgPr, 'a:blipFill'), id = kid(bf, 'a:blip')?.getAttribute('r:embed');
      const bmp = id && rl[id] ? await bitmap(S.zip, rl[id]) : null;
      if (bmp) { drawBitmap(ctx, bmp, W, H, bf); return; }
    }
  }

  async function preloadFonts(docs, theme) {
    const want = new Set([theme.fonts.minor, theme.fonts.major]);
    docs.forEach(d => d && [...d.getElementsByTagName('a:latin')].forEach(l => want.add(l.getAttribute('typeface'))));
    const jobs = [];
    for (const f of want) if (f && !f.startsWith('+')) for (const w of ['', 'bold ']) for (const i of ['', 'italic ']) jobs.push(document.fonts.load(`${i}${w}16px ${family(f)}`, 'AaÁéñ“”').catch(() => 0));
    await Promise.all(jobs);
  }

  /** data: Uint8Array | ArrayBuffer | JSZip -> Uint8Array con el PDF */
  async function convert(data) {
    const zip = data instanceof JSZip ? data : await JSZip.loadAsync(data);
    const pres = parse(await zip.file('ppt/presentation.xml').async('string'));
    const sz = pres.getElementsByTagName('p:sldSz')[0];
    const wPt = E(+sz.getAttribute('cx')), hPt = E(+sz.getAttribute('cy'));
    const presRels = await rels(zip, 'ppt/presentation.xml');
    const slidePaths = [...pres.getElementsByTagName('p:sldId')].map(s => presRels[s.getAttribute('r:id')]).filter(Boolean);
    const themeFile = zip.file('ppt/theme/theme1.xml');
    const theme = themeInfo(themeFile ? parse(await themeFile.async('string')) : null);

    const { jsPDF } = window.jspdf;
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(wPt * DPI / 72); canvas.height = Math.round(hPt * DPI / 72);
    const ctx = canvas.getContext('2d');
    const pdf = new jsPDF({ unit: 'pt', format: [wPt, hPt], orientation: wPt >= hPt ? 'l' : 'p', compress: true });

    for (let i = 0; i < slidePaths.length; i++) {
      const sPath = slidePaths[i], sDoc = parse(await zip.file(sPath).async('string')), sRels = await rels(zip, sPath);
      const lPath = Object.values(sRels).find(t => /slideLayouts\//.test(t));
      const lDoc = lPath && zip.file(lPath) ? parse(await zip.file(lPath).async('string')) : null;
      const lRels = lPath ? await rels(zip, lPath) : {};
      const mPath = Object.values(lRels).find(t => /slideMasters\//.test(t));
      const mDoc = mPath && zip.file(mPath) ? parse(await zip.file(mPath).async('string')) : null;
      const mRels = mPath ? await rels(zip, mPath) : {};
      const S = { zip, theme, rels: sRels, layout: lDoc };
      if (i === 0) await preloadFonts([sDoc, lDoc, mDoc], theme);
      ctx.setTransform(DPI / 72, 0, 0, DPI / 72, 0, 0);
      await drawBackground(ctx, wPt, hPt, [{ doc: sDoc, rels: sRels }, { doc: lDoc, rels: lRels }, { doc: mDoc, rels: mRels }], S);
      const showL = sDoc.documentElement.getAttribute('showMasterSp') !== '0';
      const showM = showL && lDoc && lDoc.documentElement.getAttribute('showMasterSp') !== '0';
      for (const [doc, rl, on] of [[mDoc, mRels, showM], [lDoc, lRels, showL]]) {   // formas fijas del master y del layout
        if (doc && on) await drawTree(ctx, doc.getElementsByTagName('p:spTree')[0], { ...S, rels: rl, skipPh: true });
      }
      await drawTree(ctx, sDoc.getElementsByTagName('p:spTree')[0], S);
      if (i) pdf.addPage([wPt, hPt], wPt >= hPt ? 'l' : 'p');
      pdf.addImage(canvas.toDataURL('image/jpeg', JPEG_Q), 'JPEG', 0, 0, wPt, hPt, undefined, 'FAST');
    }
    return new Uint8Array(pdf.output('arraybuffer'));
  }

  window.PptxPdf = { convert };
})();
