/* Importador de los formatos Excel actuales de Mínima para ENEL.
   Reconoce cada archivo por su contenido (no por el nombre) y fusiona en la base:
   1. Seguimiento de Liquidaciones  → proyectos (código, residente, valores, liquidación), pendientes y compras
   2. Seguimiento semanal (cronograma de ejecución) → avances por componente, estado, fechas
   3. Cronograma de proyectos (plan semanal) → fechas de inicio/fin y fecha de corte
   4. Estatus de facturación → actas facturadas, asignadas al proyecto por similitud de nombre
   Las hojas de nómina (PERSONAL ...) se ignoran a propósito: no deben quedar en la herramienta. */
(function (root) {
  const U = root.U;

  function rowsOf(ws) {
    return root.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  }
  const H = (v) => U.norm(v).replace(/[^a-z0-9% ]/g, ' ').replace(/\s+/g, ' ').trim();

  function findHeader(rows, required, maxScan = 15) {
    for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
      const cells = (rows[i] || []).map(H);
      if (required.every(req => cells.some(c => c === req || c.startsWith(req)))) return i;
    }
    return -1;
  }
  function colIndex(header, ...names) {
    const cells = header.map(H);
    for (const n of names) {
      const i = cells.findIndex(c => c === n);
      if (i >= 0) return i;
    }
    for (const n of names) {
      const i = cells.findIndex(c => c && c.startsWith(n));
      if (i >= 0) return i;
    }
    return -1;
  }
  const cell = (r, i) => (i >= 0 && r ? r[i] : null);
  const txt = (v) => (v === null || v === undefined) ? '' : String(v).trim();

  function toISO(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      const d = root.XLSX.SSF.parse_date_code(v);
      return d ? d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0') : null;
    }
    if (v instanceof Date) return U.toISO(v);
    const s = String(v);
    let m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
    m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[0] : null;
  }
  const pct = (v) => { const n = U.num(v); return n === null ? null : (n > 1.0001 ? n / 100 : n); };
  const codeNorm = (c) => {
    let s = txt(c).toUpperCase().replace(/\s+/g, '');
    s = s.replace(/-ENE-/, '-ENEL-');
    return s;
  };
  const ESTADOS = ['PRELIMINARES', 'PREPARACIÓN', 'EN EJECUCIÓN', 'STAND BY', 'POST VENTAS', 'EN LIQUIDACIÓN', 'FINALIZADO'];
  const estadoNorm = (v) => {
    const n = U.norm(v);
    if (!n) return null;
    return ESTADOS.find(e => U.norm(e) === n) || txt(v).toUpperCase();
  };
  const splitItems = (text) => txt(text)
    .split(/\n|(?:^|\s)\*/).map(s => s.replace(/^[\s*•\-–]+/, '').trim())
    .filter(s => s.length > 3);

  /* ── Detección de tipo de archivo ── */
  function detect(wb) {
    const names = wb.SheetNames;
    for (const n of names) {
      const rows = rowsOf(wb.Sheets[n]);
      if (findHeader(rows, ['codigo', 'proyecto', 'estado liquidacion']) >= 0) return 'liquidaciones';
      if (findHeader(rows, ['acta', 'facturado']) >= 0) return 'facturacion';
      if (findHeader(rows, ['proyecto', 'inicio', 'origen inicio', 'fin']) >= 0) return 'plan';
      if (findHeader(rows, ['proyecto', 'avance obra', 'avance dossier']) >= 0) return 'seguimiento';
    }
    return null;
  }

  /* ── Contexto de importación sobre la base ── */
  function ctx(db, stamp) {
    const live = (c) => db[c].filter(r => !r.deleted);
    const report = { creados: [], actualizados: [], vinculados: [], actasNuevas: 0, actasSinProyecto: 0, pendientes: 0, avisos: [] };
    function projectMatcher() {
      return U.makeMatcher(live('proyectos'), p => [p.nombre, p.descripcion || '', ...(p.alias || [])]);
    }
    function findByAlias(name) {
      const n = U.norm(name);
      return live('proyectos').find(p => U.norm(p.nombre) === n || (p.alias || []).some(a => U.norm(a) === n));
    }
    function addAlias(p, name) {
      p.alias = p.alias || [];
      if (name && U.norm(name) !== U.norm(p.nombre) && !p.alias.some(a => U.norm(a) === U.norm(name))) p.alias.push(txt(name));
    }
    function newProject(fields) {
      const p = stamp(Object.assign({
        id: U.uid('p'), codigo: '', nombre: '', descripcion: '', categoria: '', anio: null, gestor: '', residente: '',
        contrato: '', centroCostos: '', estado: 'PRELIMINARES', fechaInicio: null, fechaFin: null,
        avances: { informe: 0, presupuesto: 0, preop: 0, obra: 0, liquidacion: 0, dossier: 0 },
        valorContrato: null, valorLiquidacion: null, estadoLiquidacion: '', obsCliente: '', obsInterna: '',
        costosOtros: null, linkCarpeta: '', alias: [], hseChecklist: {}
      }, fields));
      db.proyectos.push(p);
      report.creados.push(p.nombre);
      return p;
    }
    function setIf(p, k, v) { if (v !== null && v !== undefined && v !== '') p[k] = v; }
    return { live, report, projectMatcher, findByAlias, addAlias, newProject, setIf };
  }

  /* 1 ── Seguimiento de liquidaciones */
  function importLiquidaciones(wb, db, C) {
    const sheetName = wb.SheetNames.find(n => U.norm(n) === 'seguimiento') ||
      wb.SheetNames.find(n => findHeader(rowsOf(wb.Sheets[n]), ['codigo', 'proyecto', 'estado liquidacion']) >= 0);
    const rows = rowsOf(wb.Sheets[sheetName]);
    const hi = findHeader(rows, ['codigo', 'proyecto', 'estado liquidacion']);
    const h = rows[hi];
    const c = {
      cod: colIndex(h, 'codigo'), nom: colIndex(h, 'proyecto'), res: colIndex(h, 'responsable'), ges: colIndex(h, 'gestor'),
      av: colIndex(h, '% avance obra', 'avance obra'), est: colIndex(h, 'estado proyecto'), fen: colIndex(h, 'fecha de entrega'),
      eliq: colIndex(h, 'estado liquidacion'), val: colIndex(h, 'pagado'), vliq: colIndex(h, 'valor liquidacion'),
      ocl: colIndex(h, 'observaciones cliente', 'observaciones'), oin: colIndex(h, 'observaciones internas'),
      pen: colIndex(h, 'pendientes'), com: colIndex(h, 'compras'), car: colIndex(h, 'link carpeta')
    };
    // Datos complementarios de la hoja RESPONSABLES (centro de costos, nombre largo)
    const extra = {};
    const rs = wb.SheetNames.find(n => U.norm(n) === 'responsables');
    if (rs) {
      const rr = rowsOf(wb.Sheets[rs]); const rh = findHeader(rr, ['centro de costos', 'codigo']);
      if (rh >= 0) {
        const hh = rr[rh]; const ic = colIndex(hh, 'centro de costos'), ik = colIndex(hh, 'codigo'), ip = colIndex(hh, 'proyecto'), ict = colIndex(hh, 'contrato');
        rr.slice(rh + 1).forEach(r => {
          const k = codeNorm(cell(r, ik)); if (!k) return;
          extra[k] = { cc: txt(cell(r, ic)), desc: txt(cell(r, ip)).replace(/\u00a0/g, ' ').trim(), contrato: txt(cell(r, ict)) };
        });
      }
    }
    let n = 0;
    rows.slice(hi + 1).forEach(r => {
      const codigo = codeNorm(cell(r, c.cod)); const nombre = txt(cell(r, c.nom));
      if (!codigo || !nombre || !/ENEL/.test(codigo)) return;
      let p = C.live('proyectos').find(x => codeNorm(x.codigo) === codigo);
      if (!p) p = C.newProject({ codigo, nombre: nombre.toUpperCase() });
      else C.report.actualizados.push(p.nombre);
      C.addAlias(p, nombre);
      C.setIf(p, 'residente', txt(cell(r, c.res)));
      C.setIf(p, 'gestor', txt(cell(r, c.ges)));
      const est = estadoNorm(cell(r, c.est)); if (est) p.estado = est;
      const fen = toISO(cell(r, c.fen)); if (fen) p.fechaFin = p.fechaFin || fen;
      C.setIf(p, 'estadoLiquidacion', txt(cell(r, c.eliq)).toUpperCase());
      const v = U.num(cell(r, c.val)); if (v !== null) p.valorContrato = v;
      const vl = U.num(cell(r, c.vliq)); if (vl !== null) p.valorLiquidacion = vl;
      C.setIf(p, 'obsCliente', txt(cell(r, c.ocl)));
      C.setIf(p, 'obsInterna', txt(cell(r, c.oin)));
      const av = pct(cell(r, c.av)); if (av !== null && !p.avances.obra) p.avances.obra = U.clamp01(av);
      const car = txt(cell(r, c.car)); if (/^https?:/.test(car)) p.linkCarpeta = car;
      const ex = extra[codigo];
      if (ex) { C.setIf(p, 'centroCostos', ex.cc); if (ex.desc && !p.descripcion) p.descripcion = ex.desc; }
      // Pendientes y compras: una línea por viñeta
      [['pen', 'Obra'], ['com', 'Compra']].forEach(([k, tipo]) => {
        splitItems(cell(r, c[k])).forEach(d => {
          const exists = C.live('pendientes').some(x => x.proyectoId === p.id && U.norm(x.descripcion) === U.norm(d));
          if (exists) return;
          db.pendientes.push(C_stamp({ id: U.uid('pe'), proyectoId: p.id, descripcion: d, tipo, responsable: p.residente || '', fechaCompromiso: null, prioridad: 'Media', estado: 'Abierto', origen: 'Importado liquidaciones' }));
          C.report.pendientes++;
        });
      });
      n++;
    });
    return { hoja: sheetName, filas: n };
  }
  let C_stamp = (x) => x;

  /* 2 ── Seguimiento semanal (avance por componentes) */
  function importSeguimiento(wb, db, C) {
    // Si hay varias hojas de seguimiento, se toma la de fechas más recientes
    let best = null;
    wb.SheetNames.forEach(n => {
      const rows = rowsOf(wb.Sheets[n]); const hi = findHeader(rows, ['proyecto', 'avance obra', 'avance dossier']);
      if (hi < 0) return;
      // Mediana de las fechas de entrega: la hoja vigente tiene las fechas más recientes
      const fi = colIndex(rows[hi], 'fecha de entrega');
      const ds = rows.slice(hi + 1).map(r => typeof cell(r, fi) === 'number' ? toISO(cell(r, fi)) : null).filter(Boolean).sort();
      const maxD = ds.length ? ds[Math.floor(ds.length / 2)] : '';
      if (!best || maxD > best.maxD) best = { n, rows, hi, maxD };
    });
    if (!best) return { filas: 0 };
    const { rows, hi } = best; const h = rows[hi];
    const c = {
      cat: colIndex(h, 'categoria'), anio: colIndex(h, 'ano', 'a o', 'año'), ges: colIndex(h, 'gestor'), nom: colIndex(h, 'proyecto'),
      est: colIndex(h, 'estado'), inf: colIndex(h, 'avance de informe', 'avance informe'), pre: colIndex(h, 'avance presupuesto'),
      pop: colIndex(h, 'avance preopracional', 'avance preoperacional'), obra: colIndex(h, 'avance proyecto'),
      liq: colIndex(h, 'avance liquidacion'), dos: colIndex(h, 'avance dossier'), ini: colIndex(h, 'fecha de inicio'),
      fin: colIndex(h, 'fecha de entrega'), obs: colIndex(h, 'observaciones')
    };
    // Pesos de avance (fila superior al encabezado)
    const wrow = rows[hi - 1] || [];
    const w = [c.inf, c.pre, c.pop, c.obra, c.liq, c.dos].map(i => U.num(wrow[i]));
    if (w.every(x => x !== null) && Math.abs(U.sum(w) - 1) < 0.01) {
      db.meta.pesos = { informe: w[0], presupuesto: w[1], preop: w[2], obra: w[3], liquidacion: w[4], dossier: w[5] };
    }
    const data = rows.slice(hi + 1).filter(r => txt(cell(r, c.nom)) && txt(cell(r, c.cat)));
    const pending = [];
    data.forEach(r => {
      const nombre = txt(cell(r, c.nom));
      const p = C.findByAlias(nombre);
      if (p) apply(p, r); else pending.push(r);
    });
    const matcher = C.projectMatcher();
    const asg = U.greedyAssign(pending.map(r => ({ text: txt(cell(r, c.nom)), extra: [txt(cell(r, c.cat)).replace(/S$/i, '')] })), matcher);
    pending.forEach((r, i) => {
      const nombre = txt(cell(r, c.nom));
      let p = asg[i] && asg[i].it;
      if (p) { C.report.vinculados.push(nombre + ' → ' + (p.codigo ? p.codigo + ' ' : '') + p.nombre); C.addAlias(p, nombre); }
      else p = C.newProject({ nombre: nombre.toUpperCase() });
      apply(p, r);
    });
    function apply(p, r) {
      C.setIf(p, 'categoria', txt(cell(r, c.cat)).toUpperCase());
      const an = U.num(cell(r, c.anio)); if (an) p.anio = an;
      C.setIf(p, 'gestor', txt(cell(r, c.ges)));
      const est = estadoNorm(cell(r, c.est)); if (est) p.estado = est;
      const map = { informe: c.inf, presupuesto: c.pre, preop: c.pop, obra: c.obra, liquidacion: c.liq, dossier: c.dos };
      Object.entries(map).forEach(([k, i]) => { const v = pct(cell(r, i)); if (v !== null) p.avances[k] = U.clamp01(v); });
      const obs = txt(cell(r, c.obs));
      const ini = toISO(cell(r, c.ini)) || (/fecha de inicio/i.test(obs) ? toISO(obs) : null);
      if (ini) p.fechaInicio = ini;
      const fin = toISO(cell(r, c.fin)); if (fin) p.fechaFin = fin;
      const nota = obs.replace(/fecha de inicio:?\s*[\d/]*\s*/i, '').trim();
      if (nota && !(p.obsInterna || '').includes(nota)) p.obsInterna = [p.obsInterna, nota].filter(Boolean).join('\n');
      p.updatedAt = Date.now();
    }
    return { hoja: best.n, filas: data.length };
  }

  /* 3 ── Cronograma (plan con inicio/fin y fecha de corte) */
  function importPlan(wb, db, C) {
    const n = wb.SheetNames.find(s => findHeader(rowsOf(wb.Sheets[s]), ['proyecto', 'inicio', 'origen inicio', 'fin']) >= 0);
    const rows = rowsOf(wb.Sheets[n]); const hi = findHeader(rows, ['proyecto', 'inicio', 'origen inicio', 'fin']); const h = rows[hi];
    rows.slice(0, hi).forEach(r => { const i = (r || []).findIndex(v => /fecha de corte/i.test(txt(v))); if (i >= 0) { const d = r.slice(i + 1).map(toISO).find(Boolean); if (d) db.meta.corte = d; } });
    const c = { nom: colIndex(h, 'proyecto'), cat: colIndex(h, 'categoria'), ges: colIndex(h, 'gestor'), est: colIndex(h, 'estado'), ini: colIndex(h, 'inicio'), ori: colIndex(h, 'origen inicio'), fin: colIndex(h, 'fin') };
    const data = rows.slice(hi + 1).filter(r => txt(cell(r, c.nom)) && txt(cell(r, c.cat)) && toISO(cell(r, c.ini)));
    const pending = [];
    data.forEach(r => { const p = C.findByAlias(txt(cell(r, c.nom))); if (p) apply(p, r); else pending.push(r); });
    const asg = U.greedyAssign(pending.map(r => ({ text: txt(cell(r, c.nom)), extra: [txt(cell(r, c.cat)).replace(/S$/i, '')] })), C.projectMatcher());
    pending.forEach((r, i) => {
      const nombre = txt(cell(r, c.nom)); let p = asg[i] && asg[i].it;
      if (p) { C.report.vinculados.push(nombre + ' → ' + (p.codigo ? p.codigo + ' ' : '') + p.nombre); C.addAlias(p, nombre); }
      else p = C.newProject({ nombre: nombre.toUpperCase(), categoria: txt(cell(r, c.cat)).toUpperCase() });
      apply(p, r);
    });
    function apply(p, r) {
      const ini = toISO(cell(r, c.ini)), fin = toISO(cell(r, c.fin));
      // "Fecha corte" / "Sin fecha" en ORIGEN INICIO = el archivo no conoce el inicio real
      const real = U.norm(cell(r, c.ori)) === 'fuente';
      if (ini && real) { p.fechaInicio = ini; p.inicioEstimado = false; }
      else if (ini && !p.fechaInicio) { p.fechaInicio = ini; p.inicioEstimado = true; }
      if (fin) p.fechaFin = fin;
      const est = estadoNorm(cell(r, c.est)); if (est) p.estado = est;
      C.setIf(p, 'categoria', txt(cell(r, c.cat)).toUpperCase());
      p.updatedAt = Date.now();
    }
    return { hoja: n, filas: data.length };
  }

  /* 4 ── Estatus de facturación (actas) */
  function importFacturacion(wb, db, C) {
    const n = wb.SheetNames.find(s => findHeader(rowsOf(wb.Sheets[s]), ['acta', 'facturado']) >= 0);
    const rows = rowsOf(wb.Sheets[n]); const hi = findHeader(rows, ['acta', 'facturado']); const h = rows[hi];
    const c = {
      fc: colIndex(h, 'fecha de conciliacion'), acta: colIndex(h, 'acta'), con: colIndex(h, 'contrato'), val: colIndex(h, 'facturado'),
      lin: colIndex(h, 'linea de negocio'), ped: colIndex(h, 'pedido'), conf: colIndex(h, 'conformidad'), fr: colIndex(h, 'fecha de radicacion', 'fecha radicacion'),
      fac: colIndex(h, 'factura'), dr: colIndex(h, 'id drape'), ges: colIndex(h, 'gestor')
    };
    const data = rows.slice(hi + 1).filter(r => {
      const a = txt(cell(r, c.acta)); const v = U.num(cell(r, c.val));
      return a && v && !/^total|^cierre|^vr total/i.test(a);
    });
    // Las actas pertenecen a proyectos contractuales: primero se buscan entre los que tienen código
    const conCodigo = U.makeMatcher(C.live('proyectos').filter(p => p.codigo), p => [p.nombre, p.descripcion || '', ...(p.alias || [])]);
    const todos = C.projectMatcher();
    let nuevas = 0, sin = 0;
    data.forEach(r => {
      const acta = txt(cell(r, c.acta)).replace(/\s+/g, ' ');
      const valor = U.num(cell(r, c.val));
      const key = U.norm(acta) + '|' + Math.round(valor);
      if (C.live('actas').some(x => U.norm(x.acta) + '|' + Math.round(x.valor) === key)) return;
      const pick = (m) => {
        const l = m.score(acta).map(r => ({ it: r.it, score: r.score * U.catPenalty({ text: acta }, r.it) })).sort((a, b) => b.score - a.score);
        return l[0] ? Object.assign(l[0], { margen: l[0].score - (l[1] ? l[1].score : 0) }) : null;
      };
      let best = pick(conCodigo);
      if (!best || best.score < 0.36) best = pick(todos);
      const proyectoId = best && best.score >= 0.36 ? best.it.id : null;
      if (!proyectoId) sin++;
      const factura = txt(cell(r, c.fac));
      db.actas.push(C_stamp({
        id: U.uid('ac'), proyectoId, revisarAsignacion: !!proyectoId && (best.score < 0.6 || best.margen < 0.12),
        acta, contrato: txt(cell(r, c.con)), fechaConciliacion: toISO(cell(r, c.fc)), valor,
        lineaNegocio: txt(cell(r, c.lin)).replace('Enel Colombia S.A. ESP - ', ''), pedido: txt(cell(r, c.ped)),
        conformidad: txt(cell(r, c.conf)).replace(/\n/g, ' · '), fechaRadicacion: toISO(cell(r, c.fr)),
        factura: /^MA/i.test(factura) ? factura : '', estadoFactura: /^MA/i.test(factura) ? 'Facturada' : (factura || 'Sin factura'),
        idDrape: txt(cell(r, c.dr)), gestor: txt(cell(r, c.ges)), estimado: /estimad/i.test(acta)
      }));
      nuevas++;
    });
    C.report.actasNuevas += nuevas; C.report.actasSinProyecto += sin;
    return { hoja: n, filas: data.length };
  }

  const ORDER = ['liquidaciones', 'seguimiento', 'plan', 'facturacion'];
  const LABEL = { liquidaciones: 'Seguimiento de liquidaciones', seguimiento: 'Seguimiento semanal de ejecución', plan: 'Cronograma de proyectos', facturacion: 'Estatus de facturación' };
  const FN = { liquidaciones: importLiquidaciones, seguimiento: importSeguimiento, plan: importPlan, facturacion: importFacturacion };

  /* files: [{name, data:ArrayBuffer}] */
  function run(files, db, stamp) {
    C_stamp = stamp;
    const C = ctx(db, stamp);
    const parsed = files.map(f => {
      const wb = root.XLSX.read(f.data, { type: 'array', cellDates: false });
      return { name: f.name, wb, tipo: detect(wb) };
    });
    const archivos = [];
    parsed.filter(p => !p.tipo).forEach(p => C.report.avisos.push(`No se reconoció el formato de “${p.name}”.`));
    ORDER.forEach(t => parsed.filter(p => p.tipo === t).forEach(p => {
      const r = FN[t](p.wb, db, C);
      archivos.push({ archivo: p.name, tipo: LABEL[t], ...r });
    }));
    if (parsed.some(p => p.wb.SheetNames.some(s => /^personal/i.test(s))))
      C.report.avisos.push('Las hojas de nómina (PERSONAL …) se omitieron: la herramienta no almacena salarios.');
    C.report.archivos = archivos;
    C.report.creados = [...new Set(C.report.creados)];
    C.report.actualizados = [...new Set(C.report.actualizados)];
    return C.report;
  }

  root.Importer = { run, detect, LABEL };
})(typeof window !== 'undefined' ? window : globalThis);
