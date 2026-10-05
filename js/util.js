/* Utilidades compartidas — Control de Proyectos ENEL */
(function (root) {
  const U = {};

  U.uid = (p = 'id') => p + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

  U.norm = (s) => String(s ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();

  U.esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* Fechas: se manejan como texto ISO yyyy-mm-dd para evitar corrimientos de zona horaria */
  U.todayISO = () => U.toISO(new Date());
  U.toISO = (d) => {
    if (!(d instanceof Date) || isNaN(d)) return null;
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  };
  U.fromISO = (s) => {
    if (!s) return null;
    const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
  };
  U.addDays = (iso, n) => { const d = U.fromISO(iso); if (!d) return null; d.setDate(d.getDate() + n); return U.toISO(d); };
  U.daysBetween = (a, b) => {
    const da = U.fromISO(a), db = U.fromISO(b);
    if (!da || !db) return null;
    return Math.round((db - da) / 86400000);
  };
  U.mondayOf = (iso) => { const d = U.fromISO(iso); if (!d) return null; const w = (d.getDay() + 6) % 7; d.setDate(d.getDate() - w); return U.toISO(d); };
  const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  U.MESES = MESES;
  U.fmtDate = (iso) => { const d = U.fromISO(iso); return d ? d.getDate() + ' ' + MESES[d.getMonth()] + ' ' + String(d.getFullYear()).slice(2) : '—'; };

  /* Dinero y porcentajes en formato colombiano */
  U.money = (v) => (v === null || v === undefined || v === '' || isNaN(v)) ? '—' : '$' + Math.round(+v).toLocaleString('es-CO');
  U.moneyShort = (v) => {
    if (v === null || v === undefined || isNaN(v)) return '—';
    const a = Math.abs(v), s = v < 0 ? '-' : '';
    if (a >= 1e6) return s + '$' + (a / 1e6).toLocaleString('es-CO', { maximumFractionDigits: a >= 1e8 ? 0 : 1 }) + ' M';
    return s + '$' + Math.round(a).toLocaleString('es-CO');
  };
  U.pct = (v, dec = 0) => (v === null || v === undefined || isNaN(v)) ? '—' : (v * 100).toLocaleString('es-CO', { maximumFractionDigits: dec }) + '%';
  U.num = (v) => {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isNaN(v) ? null : v;
    let s = String(v).replace(/[$\s]/g, '');
    if (/^-?[\d.]+,\d+$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');      // 1.234,56
    else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');             // 1.234.567
    else s = s.replace(/,/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  };
  U.clamp01 = (v) => Math.max(0, Math.min(1, +v || 0));

  U.sum = (arr, f) => arr.reduce((a, x) => a + (+(f ? f(x) : x) || 0), 0);
  U.groupBy = (arr, f) => arr.reduce((m, x) => { const k = f(x); (m[k] = m[k] || []).push(x); return m; }, {});
  U.debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

  /* ── Coincidencia de nombres de proyecto entre archivos ──
     Los archivos nombran el mismo proyecto de forma distinta
     ("SE TENJO" / "CUBIERTA - SE TENJO" / "Cambio de Cubierta SE Tenjo Corte 1").
     Se usa similitud coseno con pesos IDF sobre tokens normalizados. */
  const STOP = new Set(('se de del la el los las y en a al para con por no n nº acta corte estimado estimados cambio ' +
    'subestacion sub estacion sede sedes proyecto fase ed edf edificio piso construccion instalacion suministro ' +
    'adecuacion adecuaciones general generales mejora mejoramiento obra obras civil civiles trabajo comercial reembolsable avenida').split(' '));
  const SYN = { cl: 'calle', cll: 'calle', clle: 'calle', santalibrada: 'santa librada', cds: 'centro atencion', ene: 'enel' };
  U.tokens = (s) => {
    let t = U.norm(s).replace(/acta\s*(no\.?|n[°º.]?)?\s*\d+/g, ' ').replace(/corte\s*\d+/g, ' ')
      .replace(/([a-z])(\d)/g, '$1 $2').replace(/(\d)([a-z])/g, '$1 $2')
      .replace(/[^a-z0-9 ]/g, ' ');
    const out = [];
    t.split(' ').forEach(w => {
      if (!w) return;
      let x = SYN[w] !== undefined ? SYN[w] : w;
      x.split(' ').forEach(y => {
        if (!y) return;
        if (y.length > 4 && y.endsWith('es') && !/[aeiou]es$/.test(y)) y = y.slice(0, -2);
        else if (y.length > 4 && y.endsWith('s')) y = y.slice(0, -1);
        if (!STOP.has(y)) out.push(y);
      });
    });
    return [...new Set(out)];
  };

  const CATS = ['cubierta', 'cerramiento', 'diseno', 'cortafuego'];
  U.makeMatcher = (items, textOf) => {
    const docs = items.map(it => ({ it, toks: [...new Set(textOf(it).flatMap(U.tokens))] }));
    const df = {};
    docs.forEach(d => d.toks.forEach(t => { df[t] = (df[t] || 0) + 1; }));
    const N = Math.max(docs.length, 1);
    const idf = t => Math.log(1 + (N + 1) / ((df[t] || 0) + 0.5));
    const norm = toks => Math.sqrt(toks.reduce((a, t) => a + idf(t) ** 2, 0)) || 1;
    docs.forEach(d => { d.n = norm(d.toks); d.set = new Set(d.toks); });
    return {
      score(text, extra = []) {
        const q = [...new Set([...U.tokens(text), ...extra.flatMap(U.tokens).filter(t => CATS.includes(t))])];
        const qn = norm(q);
        return docs.map(d => {
          let s = 0;
          q.forEach(t => { if (d.set.has(t)) s += idf(t) ** 2; });
          // Mezcla de coseno (similitud global) y cobertura de la consulta (¿qué tanto del texto buscado aparece en el proyecto?)
          return { it: d.it, score: 0.45 * (s / (qn * d.n)) + 0.55 * (s / (qn * qn)) };
        }).filter(x => x.score > 0).sort((a, b) => b.score - a.score);
      }
    };
  };

  /* Si el proyecto declara un tipo de obra distinto al de la fila (cubierta vs diseño), se penaliza */
  U.catPenalty = (q, it) => {
    const qc = [q.text || '', ...(q.extra || [])].flatMap(U.tokens).filter(t => CATS.includes(t));
    if (!qc.length) return 1;
    const pc = [it.nombre, it.descripcion || '', it.categoria || ''].flatMap(U.tokens).filter(t => CATS.includes(t));
    return pc.length && !pc.some(t => qc.includes(t)) ? 0.5 : 1;
  };

  /* Asignación uno-a-uno: mayores puntajes primero */
  U.greedyAssign = (queries, matcher, threshold = 0.33, oneToOne = true) => {
    const pairs = [];
    queries.forEach((q, qi) => matcher.score(q.text, q.extra || []).slice(0, 6).forEach(r => pairs.push({ qi, it: r.it, score: r.score * U.catPenalty(q, r.it) })));
    pairs.sort((a, b) => b.score - a.score);
    const usedQ = new Set(), usedIt = new Set(), res = new Array(queries.length).fill(null);
    pairs.forEach(p => {
      if (p.score < threshold || usedQ.has(p.qi) || (oneToOne && usedIt.has(p.it))) return;
      usedQ.add(p.qi); usedIt.add(p.it); res[p.qi] = { it: p.it, score: p.score };
    });
    return res;
  };

  root.U = U;
})(typeof window !== 'undefined' ? window : globalThis);
