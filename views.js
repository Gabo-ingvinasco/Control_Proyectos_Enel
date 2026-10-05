/* Vistas y navegación */
(function (root) {
  const U = root.U, S = root.Store, A = root.A;
  const $ = (s) => document.querySelector(s);
  const F = {};            // filtros por vista (se conservan mientras la pestaña esté abierta)
  const esc = U.esc;

  /* ═══════════ Router ═══════════ */
  const NAV = [
    { r: 'tablero', l: 'Tablero', i: 'M3 13h8V3H3v10zm0 8h8v-6H3v6zm10 0h8V11h-8v10zm0-18v6h8V3h-8z' },
    { r: 'proyectos', l: 'Proyectos', i: 'M4 4h16v4H4zM4 10h16v4H4zM4 16h16v4H4z' },
    { r: 'cronograma', l: 'Cronograma', i: 'M3 5h10v3H3zM7 10.5h12v3H7zM5 16h8v3H5z' },
    { r: 'pendientes', l: 'Pendientes de obra', i: 'M9 6h11M9 12h11M9 18h11M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2' },
    { r: 'hse', l: 'HSE', i: 'M12 2l8 3v6c0 5-3.4 9.3-8 11-4.6-1.7-8-6-8-11V5l8-3z' },
    { r: 'financiero', l: 'Financiero', i: 'M4 20V10M10 20V4M16 20v-7M22 20H2' },
    { r: 'contratistas', l: 'Contratistas', i: 'M16 11a4 4 0 1 0-8 0 4 4 0 0 0 8 0zM4 21c0-4 3.6-6 8-6s8 2 8 6' },
    { r: 'datos', l: 'Datos y conexión', i: 'M12 3c4.4 0 8 1.3 8 3s-3.6 3-8 3-8-1.3-8-3 3.6-3 8-3zM4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3' }
  ];
  A.route = () => {
    const h = location.hash.replace(/^#\/?/, '').split('/');
    return { r: h[0] || 'tablero', id: h[1], tab: h[2] };
  };
  A.go = (h) => { location.hash = h; };

  A.render = () => {
    if (A.modalOpen()) { A._pendingRender = true; return; }
    const { r, id, tab } = A.route();
    document.querySelectorAll('.nav a').forEach(a => a.classList.toggle('active', a.dataset.r === r || (r === 'proyecto' && a.dataset.r === 'proyectos')));
    const v = VIEWS[r] || VIEWS.tablero;
    const main = $('#view');
    const y = main.scrollTop;
    main.innerHTML = v(id, tab);
    if (A._keepScroll) { main.scrollTop = y; A._keepScroll = false; }
    if (A._after) { const f = A._after; A._after = null; f(); }
  };
  const rerender = () => { A._keepScroll = true; A.render(); };

  function buildNav() {
    $('#nav').innerHTML = NAV.map(n => `<a href="#/${n.r}" data-r="${n.r}"><svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="${n.i}"/></svg><span>${n.l}</span></a>`).join('');
  }

  function statusChip() {
    const el = $('#sync');
    const t = S.lastSaved ? S.lastSaved.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' }) : '';
    const map = {
      loading: ['Cargando…', 'busy'], saving: ['Guardando…', 'busy'], dirty: ['Cambios por guardar', 'busy'],
      saved: [S.mode === 'm365' ? `Guardado en SharePoint${t ? ' · ' + t : ''}` : 'Modo local (solo este navegador)', S.mode === 'm365' ? 'ok' : 'local'],
      error: ['Error al guardar', 'err'], offline: ['Sin conexión · cambios en espera', 'err'], idle: ['Sin conectar', 'local']
    };
    const [l, c] = map[S.status] || map.idle;
    el.className = 'sync ' + c; el.textContent = l; el.title = S.error || '';
    $('#user').textContent = S.user.name || (S.mode === 'local' ? 'Usuario local' : '');
    $('#logout').hidden = S.mode !== 'm365';
  }

  /* ═══════════ Encabezado de página ═══════════ */
  const head = (title, sub = '', actions = '') => `<div class="page-head"><div><h1>${title}</h1>${sub ? `<p class="sub">${sub}</p>` : ''}</div><div class="actions">${actions}</div></div>`;
  const kpis = (items, extra = '') => { const n = items.length; const cols = n <= 6 ? n : Math.ceil(n / 2); return `<section class="kpis ${extra}" style="--n:${cols}">${items.join('')}</section>`; };
  const kpi = (n, l, cls = '', hint = '') => `<div class="kpi ${cls}" ${hint ? `title="${esc(hint)}"` : ''}><div class="kpi-n">${n}</div><div class="kpi-l">${l}</div></div>`;
  const sel = (id, label, opts, val, onch) => `<label class="flt"><span>${label}</span><select id="${id}" onchange="${onch}"><option value="">Todos</option>${opts.map(o => `<option ${o === val ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select></label>`;
  const empty = (msg, cta = '') => `<div class="empty-state"><p>${msg}</p>${cta}</div>`;

  const VIEWS = {};
  const alertLi = (a) => `<li>${A.dot(a.nivel)}<a href="#/proyecto/${a.p.id}/${a.tab}"><strong>${esc(a.p.nombre)}</strong><span>${esc(a.txt)}</span></a></li>`;

  /* ═══════════ TABLERO ═══════════ */
  VIEWS.tablero = () => {
    const ps = S.list('proyectos');
    if (!ps.length) return head('Tablero ENEL') + empty('Aún no hay proyectos. Importa los Excel actuales (liquidaciones, seguimiento semanal, cronograma y facturación) o crea el primer proyecto.',
      `<div class="row-gap"><a class="btn primary" href="#/datos">Importar Excel</a><button class="btn" onclick="A.formProyecto()">Nuevo proyecto</button></div>`);
    const act = ps.filter(A.ACTIVO);
    const ejec = ps.filter(p => p.estado === 'EN EJECUCIÓN');
    // También proyectos terminados cuya liquidación no está aprobada: ahí se queda la plata
    const alerts = ps.flatMap(p => { const al = A.alertas(p); return (A.ACTIVO(p) || p.estadoLiquidacion !== 'APROBADA') ? al : al.filter(a => a.nivel === 3); }).sort((a, b) => b.nivel - a.nivel);
    const principales = alerts.filter(a => a.nivel >= 2), menores = alerts.filter(a => a.nivel < 2);
    const hoy = U.todayISO();
    const pend = S.list('pendientes', x => x.estado !== 'Cerrado');
    const pendV = pend.filter(x => x.fechaCompromiso && x.fechaCompromiso < hoy);
    const valorAct = U.sum(act, p => p.valorContrato);
    const avPond = valorAct ? U.sum(act, p => (p.valorContrato || 0) * A.pond(p)) / valorAct : (act.length ? U.sum(act, A.pond) / act.length : 0);
    const actas = S.list('actas');
    const fact = U.sum(actas.filter(a => a.estadoFactura === 'Facturada'), a => a.valor);
    const sinFact = U.sum(actas.filter(a => a.estadoFactura !== 'Facturada'), a => a.valor);
    const liqCnt = U.groupBy(ps.filter(p => p.estadoLiquidacion), p => p.estadoLiquidacion);
    const proximas = act.filter(p => { const d = A.diasFin(p); return d !== null && d >= 0 && d <= 21 && p.estado !== 'STAND BY'; }).sort((a, b) => a.fechaFin.localeCompare(b.fechaFin));
    const hseEv = S.list('hse', h => h.tipo === 'Incidente' || h.tipo === 'Accidente').map(h => h.fecha).filter(Boolean).sort();
    const diasSin = hseEv.length ? U.daysBetween(hseEv[hseEv.length - 1], hoy) : null;
    const hseReqPend = U.sum(ps.filter(A.EN_OBRA), p => A.hseResumen(p).pend);

    const altas = alerts.filter(a => a.nivel === 3).length;
    const porGestor = U.groupBy(act, p => p.gestor || 'Sin gestor');

    return head('Tablero ENEL', `Fecha de corte <input type="date" class="inline-date" value="${A.corte()}" onchange="A.setCorte(this.value)"> · ${act.length} proyectos activos de ${ps.length}`,
      `<button class="btn" onclick="A.exportExcel()">Exportar Excel</button><button class="btn primary" onclick="A.formProyecto()">Nuevo proyecto</button>`) +
      `${kpis([
kpi(ejec.length, 'en ejecución'),
kpi(U.pct(avPond), 'avance ponderado de la cartera activa', '', 'Ponderado por valor proyectado de cada proyecto'),
kpi(altas, 'alertas altas', altas ? 'k-bad' : ''),
kpi(pendV.length + ' / ' + pend.length, 'pendientes vencidos / abiertos', pendV.length ? 'k-warn' : ''),
kpi(U.moneyShort(fact), 'facturado a ENEL'),
kpi(U.moneyShort(sinFact), 'conciliado sin factura', sinFact ? 'k-warn' : ''),
kpi(diasSin === null ? '—' : diasSin, 'días sin incidentes registrados'),
kpi(hseReqPend, 'requisitos HSE pendientes en obra', hseReqPend ? 'k-warn' : '')
], '')}
      <div class="grid-2">
        <section class="panel">
          <div class="panel-h"><h2>Requiere atención</h2><span class="muted">${principales.length} alertas altas y medias</span></div>
          <ul class="alert-list">${principales.map(alertLi).join('') || '<li class="muted">Sin alertas altas o medias con la información registrada.</li>'}</ul>
          ${menores.length ? `<details class="menores"><summary>${menores.length} avisos menores (HSE documental, datos por completar)</summary><ul class="alert-list">${menores.map(alertLi).join('')}</ul></details>` : ''}
        </section>
        <div class="stack">
          <section class="panel">
            <div class="panel-h"><h2>Entregas en las próximas 3 semanas</h2></div>
            ${proximas.length ? `<ul class="due-list">${proximas.map(p => `<li><a href="#/proyecto/${p.id}"><span class="due-d">${U.fmtDate(p.fechaFin)}</span><span class="due-n">${esc(p.nombre)}</span>${A.bar(p.avances.obra)}<span class="due-p">${U.pct(p.avances.obra)}</span></a></li>`).join('')}</ul>` : '<p class="muted">No hay entregas programadas en ese periodo.</p>'}
          </section>
          <section class="panel">
            <div class="panel-h"><h2>Liquidaciones</h2><a href="#/financiero" class="small">Ver detalle</a></div>
            <div class="pipeline">${A.ESTADOS_LIQ.filter(e => liqCnt[e]).map(e => `<div class="pipe"><b>${liqCnt[e].length}</b><span>${esc(e.toLowerCase())}</span></div>`).join('') || '<p class="muted">Sin estados de liquidación registrados.</p>'}</div>
          </section>
          <section class="panel">
            <div class="panel-h"><h2>Carga por gestor ENEL</h2></div>
            <table class="mini"><thead><tr><th>Gestor</th><th class="num">Activos</th><th class="num">Alertas altas</th><th class="num">Valor</th></tr></thead><tbody>
            ${Object.entries(porGestor).sort((a, b) => b[1].length - a[1].length).map(([g, l]) => `<tr><td>${esc(g)}</td><td class="num">${l.length}</td><td class="num">${l.flatMap(A.alertas).filter(a => a.nivel === 3).length}</td><td class="num">${U.moneyShort(U.sum(l, p => p.valorContrato))}</td></tr>`).join('')}
            </tbody></table>
          </section>
        </div>
      </div>`;
  };
  A.setCorte = (v) => { if (v) { S.setMeta({ corte: v }); rerender(); } };

  /* ═══════════ PROYECTOS ═══════════ */
  VIEWS.proyectos = () => {
    const f = F.proy = F.proy || { q: '', cat: '', est: '', ges: '', soloActivos: true };
    let ps = S.list('proyectos');
    const q = U.norm(f.q);
    ps = ps.filter(p => (!f.soloActivos || A.ACTIVO(p)) && (!f.cat || p.categoria === f.cat) && (!f.est || p.estado === f.est) && (!f.ges || p.gestor === f.ges)
      && (!q || U.norm([p.codigo, p.nombre, p.residente, ...(p.alias || [])].join(' ')).includes(q)));
    const cols = [
      { k: 'codigo', l: 'Código', h: p => `<span class="code">${esc(p.codigo || '—')}</span>`, w: '110px' },
      { k: 'nombre', l: 'Proyecto', h: p => `<strong>${esc(p.nombre)}</strong><div class="muted small">${[p.categoria, p.residente].filter(Boolean).map(esc).join(' · ')}</div>` },
      { k: 'gestor', l: 'Gestor ENEL' },
      { k: 'estado', l: 'Estado', h: p => A.badge(p.estado) },
      { k: 'pond', l: 'Avance', v: A.pond, h: p => `<div class="av">${A.bar(A.pond(p))}<span>${U.pct(A.pond(p))}</span></div><div class="muted small">obra ${U.pct(p.avances.obra)}${A.programado(p) !== null ? ' · prog. ' + U.pct(A.programado(p)) : ''}</div>`, w: '150px' },
      { k: 'fechaFin', l: 'Entrega', h: p => { const d = A.diasFin(p); return `${U.fmtDate(p.fechaFin)}${d !== null && A.ACTIVO(p) ? `<div class="small ${d < 0 ? 'neg' : d <= 14 ? 'warn' : 'muted'}">${d < 0 ? 'hace ' + -d : 'en ' + d} d</div>` : ''}`; } },
      { k: 'valorContrato', l: 'Valor proyectado', cls: 'num', h: p => U.moneyShort(p.valorContrato) },
      { k: 'estadoLiquidacion', l: 'Liquidación', h: p => `<span class="small">${esc((p.estadoLiquidacion || '—').toLowerCase())}</span>` },
      { k: 'al', l: 'Alertas', v: p => A.alertas(p).reduce((s, a) => s + a.nivel * 10, 0), h: p => { const al = A.alertas(p); return al.length ? `<span class="al-cnt">${al.slice(0, 4).map(a => A.dot(a.nivel)).join('')}</span>` : ''; }, w: '80px' }
    ];
    return head('Proyectos', `${ps.length} de ${S.list('proyectos').length} proyectos`, `<button class="btn primary" onclick="A.formProyecto()">Nuevo proyecto</button>`) +
      `<div class="filters">
        <label class="flt grow"><span>Buscar</span><input type="search" id="fq" value="${esc(f.q)}" placeholder="Código, nombre, residente…" oninput="A.fProy('q',this.value)"></label>
        ${sel('fcat', 'Categoría', A.categorias(), f.cat, "A.fProy('cat',this.value)")}
        ${sel('fest', 'Estado', A.ESTADOS, f.est, "A.fProy('est',this.value)")}
        ${sel('fges', 'Gestor', A.gestores(), f.ges, "A.fProy('ges',this.value)")}
        <label class="check"><input type="checkbox" ${f.soloActivos ? 'checked' : ''} onchange="A.fProy('soloActivos',this.checked)"> Ocultar finalizados</label>
      </div>` + A.table('proy', cols, ps, { onRow: p => `A.go('#/proyecto/${p.id}')`, sort: { k: 'fechaFin', d: 1 } });
  };
  A.fProy = (k, v) => {
    F.proy[k] = v;
    if (k === 'q') { A._after = () => { const el = $('#fq'); el.focus(); el.setSelectionRange(v.length, v.length); }; }
    rerender();
  };

  A.formProyecto = (id) => {
    const p = id ? S.get('proyectos', id) : null;
    A.openForm({
      title: p ? 'Editar proyecto' : 'Nuevo proyecto', wide: true, values: p || { estado: 'PRELIMINARES', anio: new Date().getFullYear() },
      fields: [
        { k: 'codigo', l: 'Código', help: 'Ej. 023-ENEL-128' }, { k: 'nombre', l: 'Nombre del proyecto', req: true },
        { k: 'descripcion', l: 'Descripción / objeto', full: true },
        { k: 'categoria', l: 'Categoría', list: A.categorias }, { k: 'estado', l: 'Estado', t: 'select', options: A.ESTADOS, blank: false },
        { k: 'gestor', l: 'Gestor ENEL', list: A.gestores }, { k: 'residente', l: 'Residente / responsable Mínima', list: A.personas },
        { k: 'contrato', l: 'Contrato marco', list: () => ['JA10177023', 'JA10177024'] }, { k: 'centroCostos', l: 'Centro de costos' },
        { k: 'fechaInicio', l: 'Inicio programado', t: 'date' }, { k: 'fechaFin', l: 'Entrega programada', t: 'date' },
        { k: 'valorContrato', l: 'Valor proyectado (pedido)', t: 'money' }, { k: 'costosOtros', l: 'Otros costos directos (materiales, personal)', t: 'money', help: 'Se suma a los contratos de contratistas para el margen.' },
        { k: 'linkCarpeta', l: 'Enlace carpeta del proyecto', t: 'url', full: true }
      ],
      onSave: (v) => {
        if (p && v.fechaInicio !== p.fechaInicio) v.inicioEstimado = false;
        if (!v.avances) v.avances = { informe: 0, presupuesto: 0, preop: 0, obra: 0, liquidacion: 0, dossier: 0 };
        const r = S.upsert('proyectos', v);
        if (!p) { S.log(r.id, 'Proyecto creado'); A.go('#/proyecto/' + r.id); }
      },
      onDelete: p ? (v) => { S.remove('proyectos', v.id); A.go('#/proyectos'); } : null
    });
  };

  /* ═══════════ FICHA DE PROYECTO ═══════════ */
  const TABS = [['resumen', 'Resumen'], ['cronograma', 'Cronograma'], ['pendientes', 'Pendientes'], ['hse', 'HSE'], ['financiero', 'Financiero'], ['contratistas', 'Contratistas'], ['bitacora', 'Bitácora']];
  VIEWS.proyecto = (id, tab = 'resumen') => {
    const p = S.get('proyectos', id);
    if (!p) return head('Proyecto no encontrado') + empty('El proyecto no existe o fue eliminado.', '<a class="btn" href="#/proyectos">Volver a proyectos</a>');
    const al = A.alertas(p);
    const cnt = {
      pendientes: S.list('pendientes', x => x.proyectoId === id && x.estado !== 'Cerrado').length,
      hse: A.hseResumen(p).pend, contratistas: S.list('contratos', c => c.proyectoId === id).length
    };
    return `<div class="crumbs"><a href="#/proyectos">Proyectos</a> / ${esc(p.categoria || 'Sin categoría')}</div>
      <div class="p-head">
        <div class="p-title"><span class="code">${esc(p.codigo || 'sin código')}</span><h1>${esc(p.nombre)}</h1>
          <div class="p-meta">${A.badge(p.estado)}<span>Gestor ENEL: <b>${esc(p.gestor || '—')}</b></span><span>Residente: <b>${esc(p.residente || '—')}</b></span>${p.linkCarpeta ? `<a href="${esc(p.linkCarpeta)}" target="_blank" rel="noopener">Carpeta del proyecto ↗</a>` : ''}</div></div>
        <div class="p-avance"><div class="big">${U.pct(A.pond(p))}</div><div class="muted small">avance ponderado</div></div>
        <div class="actions"><button class="btn" onclick="A.formProyecto('${id}')">Editar</button><button class="btn ghost" onclick="A.fusionar('${id}')" title="Unir con otro proyecto duplicado">Fusionar</button></div>
      </div>
      ${al.length ? `<div class="p-alerts">${al.map(a => `<a href="#/proyecto/${id}/${a.tab}">${A.dot(a.nivel)}${esc(a.txt)}</a>`).join('')}</div>` : ''}
      <nav class="tabs">${TABS.map(([k, l]) => `<a href="#/proyecto/${id}/${k}" class="${tab === k ? 'on' : ''}">${l}${cnt[k] ? `<sup>${cnt[k]}</sup>` : ''}</a>`).join('')}</nav>
      <div class="tab-body">${(PTAB[tab] || PTAB.resumen)(p)}</div>`;
  };
  const PTAB = {};

  PTAB.resumen = (p) => {
    const w = S.db.meta.pesos, prog = A.programado(p);
    const datos = [
      ['Categoría', p.categoria], ['Año', p.anio], ['Contrato marco', p.contrato], ['Centro de costos', p.centroCostos],
      ['Inicio programado', U.fmtDate(p.fechaInicio)], ['Entrega programada', U.fmtDate(p.fechaFin)],
      ['Valor proyectado', U.money(p.valorContrato)], ['Valor liquidación', U.money(p.valorLiquidacion)],
      ['Estado liquidación', p.estadoLiquidacion], ['Nombres en otros archivos', (p.alias || []).join(' · ')]
    ];
    return `<div class="grid-2">
      <section class="panel">
        <div class="panel-h"><h2>Avance por componente</h2><span class="muted small">Los pesos se ajustan en Datos y conexión</span></div>
        <div class="comp">${A.COMP.map(c => `<div class="comp-row"><span class="comp-l">${c.l}<small>${U.pct(w[c.k])}</small></span>
          <input type="range" min="0" max="100" step="5" value="${Math.round((p.avances[c.k] || 0) * 100)}" oninput="this.nextElementSibling.value=this.value" onchange="A.setAvance('${p.id}','${c.k}',this.value)" aria-label="Avance ${c.l}">
          <input type="number" min="0" max="100" value="${Math.round((p.avances[c.k] || 0) * 100)}" onchange="A.setAvance('${p.id}','${c.k}',this.value)" aria-label="Avance ${c.l} en porcentaje"><span>%</span></div>`).join('')}</div>
        <div class="comp-tot"><div><b>${U.pct(A.pond(p), 1)}</b><span>avance ponderado</span></div><div><b>${U.pct(p.avances.obra)}</b><span>obra ejecutada</span></div><div><b>${prog === null ? '—' : U.pct(prog)}</b><span>obra programada al corte</span></div></div>
      </section>
      <section class="panel">
        <div class="panel-h"><h2>Datos del proyecto</h2></div>
        <dl class="dl">${datos.map(([k, v]) => `<dt>${k}</dt><dd>${esc(v || '—')}</dd>`).join('')}</dl>
      </section>
    </div>
    <div class="grid-2">
      <section class="panel"><div class="panel-h"><h2>Observaciones para ENEL</h2></div>
        <textarea class="note" rows="5" onchange="A.setField('${p.id}','obsCliente',this.value)" placeholder="Lo que se reporta al gestor o a la interventoría">${esc(p.obsCliente || '')}</textarea></section>
      <section class="panel"><div class="panel-h"><h2>Observaciones internas</h2></div>
        <textarea class="note" rows="5" onchange="A.setField('${p.id}','obsInterna',this.value)" placeholder="Notas del equipo de Mínima">${esc(p.obsInterna || '')}</textarea></section>
    </div>`;
  };
  A.setAvance = (id, k, val) => {
    const p = S.get('proyectos', id); const v = U.clamp01((+val || 0) / 100);
    const antes = p.avances[k];
    p.avances = Object.assign({}, p.avances, { [k]: v });
    S.upsert('proyectos', p);
    if (Math.abs((antes || 0) - v) >= 0.05) S.log(id, `Avance ${A.COMP.find(c => c.k === k).l}: ${U.pct(antes)} → ${U.pct(v)}`);
    rerender();
  };
  A.setField = (id, k, v) => { const p = S.get('proyectos', id); p[k] = v; S.upsert('proyectos', p); A.toast('Guardado'); };

  A.fusionar = (id) => {
    const p = S.get('proyectos', id);
    A.openForm({
      title: 'Fusionar proyectos duplicados', saveLabel: 'Fusionar',
      note: `Todo lo registrado en el proyecto seleccionado (pendientes, HSE, actas, contratos, actividades) pasa a <b>${esc(p.nombre)}</b> y el duplicado se elimina. Sus nombres quedan como alias para futuras importaciones.`,
      fields: [{ k: 'otro', l: 'Proyecto duplicado', t: 'select', req: true, full: true, options: () => A.optProyectos().filter(o => o.v !== id) }],
      onSave: (v) => {
        const o = S.get('proyectos', v.otro); if (!o) return false;
        ['actividades', 'pendientes', 'hse', 'contratos', 'actas', 'bitacora'].forEach(c => S.list(c, r => r.proyectoId === o.id).forEach(r => { r.proyectoId = id; S.upsert(c, r); }));
        p.alias = [...new Set([...(p.alias || []), o.nombre, ...(o.alias || [])])];
        ['codigo', 'descripcion', 'categoria', 'gestor', 'residente', 'contrato', 'centroCostos', 'fechaInicio', 'fechaFin', 'valorContrato', 'valorLiquidacion', 'estadoLiquidacion', 'linkCarpeta'].forEach(k => { if (!p[k] && o[k]) p[k] = o[k]; });
        p.obsInterna = [p.obsInterna, o.obsInterna].filter(Boolean).join('\n');
        S.upsert('proyectos', p); S.remove('proyectos', o.id);
        S.log(id, `Fusionado con “${o.nombre}”`);
        A.toast('Proyectos fusionados');
      }
    });
  };

  /* Cronograma del proyecto */
  PTAB.cronograma = (p) => {
    const acts = S.list('actividades', a => a.proyectoId === p.id).sort((a, b) => (a.inicio || '9').localeCompare(b.inicio || '9') || (a.orden || 0) - (b.orden || 0));
    const rows = [{ label: 'Proyecto completo', sub: `${U.pct(p.avances.obra)} obra`, start: p.fechaInicio, end: p.fechaFin, prog: p.avances.obra, cls: 'bar-proj', est: p.inicioEstimado },
      ...acts.map(a => ({ label: a.nombre, sub: (a.responsable ? esc(a.responsable) + ' · ' : '') + U.pct(a.avance), start: a.inicio, end: a.fin, prog: a.avance, cls: a.fin && a.fin < A.corte() && (a.avance || 0) < 1 ? 'bar-late' : '' }))];
    const cols = [
      { k: 'nombre', l: 'Actividad', h: a => `<strong>${esc(a.nombre)}</strong>` }, { k: 'responsable', l: 'Responsable' },
      { k: 'inicio', l: 'Inicio', h: a => U.fmtDate(a.inicio) }, { k: 'fin', l: 'Fin', h: a => U.fmtDate(a.fin) },
      { k: 'avance', l: 'Avance', h: a => `<div class="av">${A.bar(a.avance)}<span>${U.pct(a.avance)}</span></div>` },
      { k: 'pesoObra', l: 'Peso en obra', cls: 'num', h: a => a.peso ? U.pct(a.peso) : '—' }
    ];
    const pesos = U.sum(acts, a => a.peso);
    return `<div class="tab-actions"><button class="btn primary" onclick="A.formActividad('${p.id}')">Nueva actividad</button>${!acts.length ? `<button class="btn" onclick="A.plantillaActividades('${p.id}')">Cargar actividades típicas</button>` : ''}
      ${acts.length && pesos > 0 ? `<button class="btn ghost" onclick="A.calcObraDesdeActividades('${p.id}')" title="Promedio ponderado del avance de las actividades">Actualizar % obra desde actividades</button>` : ''}</div>` +
      A.gantt(rows, { weekW: 34 }) +
      `<h3 class="sec">Actividades</h3>` + A.table('act', cols, acts, { onRow: a => `A.formActividad('${p.id}','${a.id}')`, empty: 'Sin actividades. El avance de obra se puede seguir solo con el porcentaje del resumen, o detallarlo aquí.' });
  };
  A.formActividad = (pid, id) => {
    const a = id ? S.get('actividades', id) : null;
    A.openForm({
      title: a ? 'Editar actividad' : 'Nueva actividad', values: a || { proyectoId: pid, avance: 0 },
      fields: [{ k: 'nombre', l: 'Actividad', req: true, full: true }, { k: 'responsable', l: 'Responsable', list: A.personas },
        { k: 'peso', l: 'Peso en el avance de obra', t: 'pct', help: 'Opcional. Si todas las actividades tienen peso, el % de obra se puede calcular.' },
        { k: 'inicio', l: 'Inicio', t: 'date' }, { k: 'fin', l: 'Fin', t: 'date' }, { k: 'avance', l: 'Avance', t: 'pct' }, { k: 'nota', l: 'Nota', full: true }],
      onSave: v => { if (v.inicio && v.fin && v.fin < v.inicio) { A.toast('La fecha de fin es anterior al inicio.', 'err'); return false; } S.upsert('actividades', v); },
      onDelete: a ? v => S.remove('actividades', v.id) : null
    });
  };
  A.plantillaActividades = (pid) => {
    const p = S.get('proyectos', pid);
    const ini = p.fechaInicio || A.corte(), fin = p.fechaFin || U.addDays(ini, 42);
    const tot = Math.max(7, U.daysBetween(ini, fin));
    const fases = [['Preliminares y localización', 0, 0.12, 0.05], ['Desmontes y demoliciones', 0.08, 0.3, 0.15], ['Obra civil / estructura', 0.25, 0.7, 0.4], ['Instalaciones técnicas', 0.45, 0.85, 0.2], ['Acabados', 0.7, 0.95, 0.15], ['Aseo, entrega y recibo ENEL', 0.92, 1, 0.05]];
    fases.forEach(([n, a, b, w], i) => S.upsert('actividades', { proyectoId: pid, nombre: n, orden: i, inicio: U.addDays(ini, Math.round(tot * a)), fin: U.addDays(ini, Math.round(tot * b)), avance: 0, peso: w, responsable: p.residente || '' }));
    A.toast('Actividades creadas: ajusta nombres, fechas y pesos al alcance real.'); rerender();
  };
  A.calcObraDesdeActividades = (pid) => {
    const acts = S.list('actividades', a => a.proyectoId === pid);
    const w = U.sum(acts, a => a.peso); if (!w) return;
    A.setAvance(pid, 'obra', Math.round(U.sum(acts, a => (a.peso || 0) * (a.avance || 0)) / w * 100));
  };

  /* Pendientes del proyecto */
  const pendCols = (conProyecto) => [
    { k: 'estado', l: '', nosort: true, w: '34px', h: x => `<input type="checkbox" title="Cerrar pendiente" ${x.estado === 'Cerrado' ? 'checked' : ''} onclick="event.stopPropagation();A.togglePend('${x.id}',this.checked)">` },
    { k: 'descripcion', l: 'Pendiente', h: x => `<span class="${x.estado === 'Cerrado' ? 'done' : ''}">${esc(x.descripcion)}</span>${(x.adjuntos || []).length ? ' <span class="muted small">📎</span>' : ''}` },
    ...(conProyecto ? [{ k: 'proyectoId', l: 'Proyecto', v: x => A.pname(x.proyectoId), h: x => `<span class="small">${esc(A.pname(x.proyectoId))}</span>` }] : []),
    { k: 'tipo', l: 'Tipo' }, { k: 'responsable', l: 'Responsable' },
    { k: 'prioridad', l: 'Prioridad', v: x => A.PRIORIDADES.indexOf(x.prioridad), h: x => `<span class="prio p-${U.norm(x.prioridad)}">${esc(x.prioridad || '')}</span>` },
    { k: 'fechaCompromiso', l: 'Compromiso', h: x => { const venc = x.estado !== 'Cerrado' && x.fechaCompromiso && x.fechaCompromiso < U.todayISO(); return `<span class="${venc ? 'neg' : ''}">${U.fmtDate(x.fechaCompromiso)}</span>`; } },
    { k: 'est', l: 'Estado', v: x => x.estado, h: x => `<span class="small">${esc(x.estado)}</span>` }
  ];
  PTAB.pendientes = (p) => {
    const ver = F.pendP = F.pendP || 'abiertos';
    const l = S.list('pendientes', x => x.proyectoId === p.id && (ver === 'todos' || x.estado !== 'Cerrado'));
    return `<div class="tab-actions"><button class="btn primary" onclick="A.formPendiente('${p.id}')">Nuevo pendiente</button>
      <div class="seg"><button class="${ver === 'abiertos' ? 'on' : ''}" onclick="A.setF('pendP','abiertos')">Abiertos</button><button class="${ver === 'todos' ? 'on' : ''}" onclick="A.setF('pendP','todos')">Todos</button></div></div>` +
      A.table('pendp', pendCols(false), l, { onRow: x => `A.formPendiente('${p.id}','${x.id}')`, sort: { k: 'fechaCompromiso', d: 1 }, empty: 'Sin pendientes abiertos.' });
  };
  A.setF = (k, v) => { F[k] = v; rerender(); };
  A.togglePend = (id, closed) => {
    const x = S.get('pendientes', id); x.estado = closed ? 'Cerrado' : 'Abierto'; x.fechaCierre = closed ? U.todayISO() : null;
    S.upsert('pendientes', x); if (closed) S.log(x.proyectoId, 'Pendiente cerrado: ' + x.descripcion); rerender();
  };
  A.formPendiente = (pid, id) => {
    const x = id ? S.get('pendientes', id) : null;
    const p = pid ? S.get('proyectos', pid) : null;
    A.openForm({
      title: x ? 'Editar pendiente' : 'Nuevo pendiente', values: x || { proyectoId: pid || '', estado: 'Abierto', prioridad: 'Media', tipo: 'Obra', responsable: p ? p.residente : '' },
      fields: [
        ...(pid ? [] : [{ k: 'proyectoId', l: 'Proyecto', t: 'select', options: A.optProyectos, req: true, full: true }]),
        { k: 'descripcion', l: 'Descripción', t: 'textarea', req: true, full: true },
        { k: 'tipo', l: 'Tipo', t: 'select', options: A.PEND_TIPOS, blank: false }, { k: 'prioridad', l: 'Prioridad', t: 'select', options: A.PRIORIDADES, blank: false },
        { k: 'responsable', l: 'Responsable', list: A.personas }, { k: 'fechaCompromiso', l: 'Fecha compromiso', t: 'date' },
        { k: 'estado', l: 'Estado', t: 'select', options: A.PEND_ESTADOS, blank: false }, { k: 'adj', l: 'Soporte / foto', t: 'file', folder: v => S.get('proyectos', v.proyectoId)?.codigo || 'sin-codigo' }
      ],
      onSave: v => { if (v.estado === 'Cerrado' && !v.fechaCierre) v.fechaCierre = U.todayISO(); const isNew = !v.id; S.upsert('pendientes', v); if (isNew) S.log(v.proyectoId, 'Pendiente registrado: ' + v.descripcion); },
      onDelete: x ? v => S.remove('pendientes', v.id) : null
    });
  };

  /* HSE del proyecto */
  PTAB.hse = (p) => {
    const h = A.hseResumen(p), ck = p.hseChecklist || {};
    const cols = hseCols(false);
    return `${kpis([
kpi(U.pct(h.cumplimiento), 'requisitos documentales al día', h.pend ? 'k-warn' : ''),
kpi(h.diasSin === null ? '—' : h.diasSin, 'días sin incidentes'),
kpi(h.hallazgos, 'hallazgos abiertos', h.hallazgos ? 'k-warn' : ''),
kpi(h.permisosVencidos, 'permisos vencidos sin cerrar', h.permisosVencidos ? 'k-bad' : '')
], 'small-k')}
      <div class="grid-2 g-hse">
        <section class="panel"><div class="panel-h"><h2>Requisitos HSE del proyecto</h2><span class="muted small">Base para aceptación de ENEL e interventoría</span></div>
          <ul class="req">${h.reqs.map((r, i) => { const e = (ck[r] && ck[r].estado) || 'Pendiente'; return `<li class="rq-${U.norm(e).replace('/', '')}"><span>${esc(r)}</span>
            <select onchange="A.setReq('${p.id}',${i},this.value)" aria-label="Estado de ${esc(r)}">${A.REQ_ESTADOS.map(s => `<option ${s === e ? 'selected' : ''}>${s}</option>`).join('')}</select>
            <small class="muted">${ck[r] && ck[r].fecha ? U.fmtDate(ck[r].fecha) : ''}</small></li>`; }).join('')}</ul>
        </section>
        <section class="panel"><div class="panel-h"><h2>Registros</h2><button class="btn primary sm" onclick="A.formHSE('${p.id}')">Nuevo registro</button></div>
          ${A.table('hsep', cols, h.regs, { onRow: x => `A.formHSE('${p.id}','${x.id}')`, sort: { k: 'fecha', d: -1 }, empty: 'Registra inspecciones, ATS, permisos de trabajo, charlas e incidentes.' })}
        </section>
      </div>`;
  };
  A.setReq = (pid, i, estado) => {
    const p = S.get('proyectos', pid); const r = A.reqHSE()[i];
    p.hseChecklist = Object.assign({}, p.hseChecklist, { [r]: { estado, fecha: U.todayISO(), por: S.user.name } });
    S.upsert('proyectos', p); rerender();
  };
  const hseCols = (conProyecto) => [
    { k: 'fecha', l: 'Fecha', h: x => U.fmtDate(x.fecha) },
    ...(conProyecto ? [{ k: 'proyectoId', l: 'Proyecto', v: x => A.pname(x.proyectoId), h: x => `<span class="small">${esc(A.pname(x.proyectoId))}</span>` }] : []),
    { k: 'tipo', l: 'Tipo', h: x => `${esc(x.tipo)}${x.subtipo ? `<div class="muted small">${esc(x.subtipo)}</div>` : ''}` },
    { k: 'descripcion', l: 'Descripción', h: x => esc((x.descripcion || '').slice(0, 90)) + ((x.adjuntos || []).length ? ' <span class="muted small">📎</span>' : '') },
    { k: 'responsable', l: 'Responsable' },
    { k: 'estado', l: 'Estado', h: x => { const venc = x.tipo === 'Permiso de trabajo' && x.estado !== 'Cerrado' && x.vence && x.vence < U.todayISO(); return `<span class="small ${venc ? 'neg' : ''}">${esc(x.estado || '')}${venc ? ' · vencido' : ''}</span>`; } }
  ];
  A.formHSE = (pid, id) => {
    const x = id ? S.get('hse', id) : null;
    A.openForm({
      title: x ? 'Editar registro HSE' : 'Nuevo registro HSE', values: x || { proyectoId: pid || '', fecha: U.todayISO(), tipo: 'Inspección', estado: 'Abierto' },
      fields: [
        ...(pid ? [] : [{ k: 'proyectoId', l: 'Proyecto', t: 'select', options: A.optProyectos, req: true, full: true }]),
        { k: 'fecha', l: 'Fecha', t: 'date', req: true }, { k: 'tipo', l: 'Tipo', t: 'select', options: A.HSE_TIPOS, blank: false },
        { k: 'subtipo', l: 'Tipo de permiso / tema', t: 'select', options: A.PERMISOS, help: 'Para permisos de trabajo, charlas o ATS.' },
        { k: 'vence', l: 'Vigencia hasta', t: 'date', help: 'Permisos de trabajo: se alerta si vence sin cerrarse.' },
        { k: 'descripcion', l: 'Descripción / hallazgo', t: 'textarea', full: true, req: true },
        { k: 'accion', l: 'Acción correctiva', t: 'textarea', full: true, rows: 2 },
        { k: 'responsable', l: 'Responsable', list: A.personas }, { k: 'personas', l: 'Personas involucradas', t: 'number' },
        { k: 'estado', l: 'Estado', t: 'select', options: ['Abierto', 'Cerrado'], blank: false },
        { k: 'adj', l: 'Evidencia (foto, formato firmado)', t: 'file', folder: v => (S.get('proyectos', v.proyectoId)?.codigo || 'sin-codigo') + '/hse' }
      ],
      onSave: v => { const isNew = !v.id; S.upsert('hse', v); if (isNew && /Incidente|Accidente/.test(v.tipo)) S.log(v.proyectoId, `${v.tipo} HSE registrado: ${v.descripcion}`); },
      onDelete: x ? v => S.remove('hse', v.id) : null
    });
  };

  /* Financiero del proyecto */
  PTAB.financiero = (p) => {
    const f = A.fin(p);
    const cols = actaCols(false);
    return `${kpis([
kpi(U.moneyShort(p.valorContrato), 'valor proyectado'),
kpi(U.moneyShort(f.conciliado), `conciliado en actas (${U.pct(f.pctFact)})`),
kpi(U.moneyShort(f.facturado), 'facturado'),
kpi(U.moneyShort(p.valorLiquidacion), 'valor liquidación', f.difLiq < 0 ? 'k-warn' : '', f.difLiq !== null ? 'Diferencia vs proyectado: ' + U.money(f.difLiq) : ''),
kpi(U.moneyShort(f.costo || null), 'costos registrados'),
kpi(f.margenPct === null ? '—' : U.pct(f.margenPct), 'margen sobre costos registrados', f.margenPct !== null && f.margenPct < 0.1 ? 'k-bad' : '')
], 'small-k')}
      <div class="grid-2">
        <section class="panel"><div class="panel-h"><h2>Liquidación</h2><button class="btn sm" onclick="A.formLiquidacion('${p.id}')">Actualizar</button></div>
          <dl class="dl"><dt>Estado</dt><dd>${esc(p.estadoLiquidacion || '—')}</dd><dt>Valor proyectado</dt><dd>${U.money(p.valorContrato)}</dd><dt>Valor liquidación</dt><dd>${U.money(p.valorLiquidacion)}</dd>
          <dt>Diferencia</dt><dd class="${f.difLiq < 0 ? 'neg' : ''}">${U.money(f.difLiq)}${f.difLiq !== null && p.valorContrato ? ` (${U.pct(f.difLiq / p.valorContrato, 1)})` : ''}</dd><dt>Avance liquidación</dt><dd>${U.pct(p.avances.liquidacion)}</dd></dl>
        </section>
        <section class="panel"><div class="panel-h"><h2>Costos</h2></div>
          <dl class="dl"><dt>Contratos de contratistas</dt><dd>${U.money(f.comprometido)}</dd><dt>Pagado a contratistas</dt><dd>${U.money(f.pagado)}</dd><dt>Saldo por pagar</dt><dd>${U.money(f.saldo)}</dd><dt>Otros costos directos</dt><dd>${U.money(p.costosOtros)}</dd></dl>
          <p class="muted small">El margen solo es confiable si todos los costos del proyecto están registrados (contratos + otros costos directos).</p>
        </section>
      </div>
      <h3 class="sec">Actas de este proyecto <button class="btn sm" onclick="A.formActa(null,'${p.id}')">Nueva acta</button></h3>` +
      A.table('actp', cols, f.actas, { onRow: a => `A.formActa('${a.id}')`, sort: { k: 'fechaConciliacion', d: -1 }, empty: 'Sin actas asociadas. Asigna actas desde la vista Financiero.' });
  };
  A.formLiquidacion = (pid) => {
    const p = S.get('proyectos', pid);
    A.openForm({
      title: 'Liquidación del proyecto', values: p,
      fields: [{ k: 'estadoLiquidacion', l: 'Estado', t: 'select', options: A.ESTADOS_LIQ }, { k: 'valorLiquidacion', l: 'Valor liquidación', t: 'money' }, { k: 'valorContrato', l: 'Valor proyectado', t: 'money' }, { k: 'linkLiquidacion', l: 'Enlace al archivo de liquidación', t: 'url', full: true }],
      onSave: v => { const antes = p.estadoLiquidacion; S.upsert('proyectos', v); if (antes !== v.estadoLiquidacion) S.log(pid, `Liquidación: ${antes || '—'} → ${v.estadoLiquidacion}`); }
    });
  };

  /* Contratistas del proyecto */
  PTAB.contratistas = (p) => {
    const cs = S.list('contratos', c => c.proyectoId === p.id);
    return `<div class="tab-actions"><button class="btn primary" onclick="A.formContrato(null,'${p.id}')">Nuevo contrato</button>${!S.list('contratistas').length ? '<a class="btn ghost" href="#/contratistas">Crear contratistas</a>' : ''}</div>` +
      A.table('ctp', contratoCols(false), cs, { onRow: c => `A.formContrato('${c.id}')`, empty: 'Sin contratos registrados para este proyecto.' });
  };
  const contratoCols = (conProyecto) => [
    { k: 'contratistaId', l: 'Contratista', v: c => S.get('contratistas', c.contratistaId)?.nombre, h: c => { const ct = S.get('contratistas', c.contratistaId); const docs = ct ? A.contratistaDocs(ct) : []; return `<strong>${esc(ct?.nombre || '—')}</strong>${docs.map(d => `<div class="small ${d.nivel === 3 ? 'neg' : 'warn'}">${esc(d.txt)}</div>`).join('')}`; } },
    ...(conProyecto ? [{ k: 'proyectoId', l: 'Proyecto', v: c => A.pname(c.proyectoId), h: c => `<span class="small">${esc(A.pname(c.proyectoId))}</span>` }] : []),
    { k: 'objeto', l: 'Objeto', h: c => esc((c.objeto || '').slice(0, 70)) },
    { k: 'valor', l: 'Valor', cls: 'num', h: c => U.money(c.valor) },
    { k: 'pagado', l: 'Pagado', cls: 'num', v: c => U.sum(c.pagos || [], x => x.valor), h: c => { const pg = U.sum(c.pagos || [], x => x.valor); return `${U.money(pg)}<div class="small muted">${U.pct(c.valor ? pg / c.valor : 0)}</div>`; } },
    { k: 'saldo', l: 'Saldo', cls: 'num', v: c => (c.valor || 0) - U.sum(c.pagos || [], x => x.valor), h: c => U.money((c.valor || 0) - U.sum(c.pagos || [], x => x.valor)) },
    { k: 'estado', l: 'Estado', h: c => `<span class="small">${esc(c.estado || '')}</span>${c.pazYSalvo ? '<div class="small ok">Paz y salvo</div>' : ''}` }
  ];
  A.formContrato = (id, pid) => {
    const c = id ? S.get('contratos', id) : null;
    if (!S.list('contratistas').length) { A.toast('Primero crea el contratista en la vista Contratistas.', 'err'); return; }
    A.openForm({
      title: c ? 'Contrato de obra' : 'Nuevo contrato', wide: true, values: c || { proyectoId: pid || '', estado: 'Vigente', pagos: [] },
      fields: [
        { k: 'proyectoId', l: 'Proyecto', t: 'select', options: A.optProyectos, req: true }, { k: 'contratistaId', l: 'Contratista', t: 'select', options: A.optContratistas, req: true },
        { k: 'objeto', l: 'Objeto / alcance', t: 'textarea', full: true, req: true },
        { k: 'numero', l: 'N.º orden o contrato' }, { k: 'valor', l: 'Valor', t: 'money', req: true },
        { k: 'anticipo', l: 'Anticipo', t: 'money' }, { k: 'retenidoPct', l: 'Retención de garantía', t: 'pct' },
        { k: 'fechaInicio', l: 'Inicio', t: 'date' }, { k: 'fechaFin', l: 'Fin', t: 'date' },
        { k: 'estado', l: 'Estado', t: 'select', options: A.CONTRATO_ESTADOS, blank: false }, { k: 'pazYSalvo', l: 'Paz y salvo', t: 'check', cl: 'Emitido' }
      ],
      onSave: v => { S.upsert('contratos', v); },
      onDelete: c ? v => S.remove('contratos', v.id) : null
    });
    if (c) {
      const form = document.getElementById('modal-form');
      const div = document.createElement('div'); div.className = 'pagos';
      div.innerHTML = `<h3>Pagos</h3>${(c.pagos || []).length ? `<table class="mini"><thead><tr><th>Fecha</th><th>Concepto</th><th class="num">Valor</th><th></th></tr></thead><tbody>${c.pagos.map((x, i) => `<tr><td>${U.fmtDate(x.fecha)}</td><td>${esc(x.concepto || '')}</td><td class="num">${U.money(x.valor)}</td><td><button type="button" class="icon-btn" onclick="A.delPago('${c.id}',${i})" aria-label="Quitar pago">✕</button></td></tr>`).join('')}</tbody></table>` : '<p class="muted small">Sin pagos registrados.</p>'}
        <div class="pago-new"><input type="date" id="pg-f" value="${U.todayISO()}" aria-label="Fecha del pago"><input type="text" id="pg-c" placeholder="Concepto (anticipo, corte 1…)" aria-label="Concepto"><input type="text" id="pg-v" inputmode="numeric" placeholder="Valor" aria-label="Valor"><button type="button" class="btn sm" onclick="A.addPago('${c.id}')">Agregar pago</button></div>`;
      form.insertBefore(div, form.querySelector('footer'));
    }
  };
  A.addPago = (id) => {
    const c = S.get('contratos', id); const v = U.num($('#pg-v').value);
    if (!v) { A.toast('Escribe el valor del pago.', 'err'); return; }
    c.pagos = [...(c.pagos || []), { fecha: $('#pg-f').value, concepto: $('#pg-c').value, valor: v }];
    S.upsert('contratos', c); A.closeModal(); A.formContrato(id); A.toast('Pago registrado');
  };
  A.delPago = (id, i) => { const c = S.get('contratos', id); c.pagos = c.pagos.filter((_, j) => j !== i); S.upsert('contratos', c); A.closeModal(); A.formContrato(id); };

  /* Bitácora */
  PTAB.bitacora = (p) => {
    const l = S.list('bitacora', b => b.proyectoId === p.id).sort((a, b) => b.fecha.localeCompare(a.fecha));
    return `<div class="bit-new"><textarea id="bit-t" rows="2" placeholder="Registrar novedad: visita de interventoría, acuerdo con el gestor, cambio de alcance…"></textarea><button class="btn primary" onclick="A.addBitacora('${p.id}')">Registrar</button></div>
      <ul class="bitacora">${l.map(b => `<li><time>${new Date(b.fecha).toLocaleString('es-CO', { dateStyle: 'medium', timeStyle: 'short' })}</time><span class="who">${esc(b.updatedBy || '')}</span><p>${esc(b.texto)}</p></li>`).join('') || '<li class="muted">Sin novedades registradas.</li>'}</ul>`;
  };
  A.addBitacora = (pid) => { const t = $('#bit-t').value.trim(); if (!t) return; S.log(pid, t); rerender(); };

  /* ═══════════ CRONOGRAMA GENERAL ═══════════ */
  VIEWS.cronograma = () => {
    const f = F.cron = F.cron || { sb: false, fin: false, ges: '' };
    const ps = S.list('proyectos', p => (f.sb || p.estado !== 'STAND BY') && (f.fin || p.estado !== 'FINALIZADO') && (!f.ges || p.gestor === f.ges));
    const grupos = U.groupBy(ps, p => p.categoria || 'SIN CATEGORÍA');
    const rows = [];
    Object.keys(grupos).sort().forEach(g => {
      const l = grupos[g].sort((a, b) => (a.fechaInicio || a.fechaFin || '9').localeCompare(b.fechaInicio || b.fechaFin || '9'));
      const conF = l.filter(p => p.fechaFin);
      rows.push({ group: true, label: `${g} · ${l.length}` });
      l.forEach(p => {
        const late = A.diasFin(p) !== null && A.diasFin(p) < 0 && (p.avances.obra || 0) < 1 && !['FINALIZADO', 'EN LIQUIDACIÓN'].includes(p.estado);
        rows.push({ label: p.nombre, href: `#/proyecto/${p.id}/cronograma`, sub: `${A.badge(p.estado)} <span>${U.pct(p.avances.obra)}</span>`, start: p.fechaInicio, end: p.fechaFin, prog: p.avances.obra, est: p.inicioEstimado, cls: late ? 'bar-late' : (A.ESTADO_CLS[p.estado] || '') });
      });
    });
    return head('Cronograma de proyectos', 'Barras: inicio a entrega programada (borde difuso = inicio sin confirmar). Relleno: avance de obra. Línea naranja: fecha de corte.',
      `<label class="flt"><span>Fecha de corte</span><input type="date" value="${A.corte()}" onchange="A.setCorte(this.value)"></label>`) +
      `<div class="filters">${sel('cges', 'Gestor', A.gestores(), f.ges, "A.fCron('ges',this.value)")}
        <label class="check"><input type="checkbox" ${f.sb ? 'checked' : ''} onchange="A.fCron('sb',this.checked)"> Incluir stand by</label>
        <label class="check"><input type="checkbox" ${f.fin ? 'checked' : ''} onchange="A.fCron('fin',this.checked)"> Incluir finalizados</label>
        <span class="legend"><i class="lg st-ejec"></i>En ejecución <i class="lg st-pre"></i>Preliminares <i class="lg st-liq"></i>Liquidación / post venta <i class="lg bar-late"></i>Entrega vencida</span></div>` +
      (rows.length ? A.gantt(rows) : empty('No hay proyectos para mostrar con estos filtros.'));
  };
  A.fCron = (k, v) => { F.cron[k] = v; rerender(); };

  /* ═══════════ PENDIENTES (todos) ═══════════ */
  VIEWS.pendientes = () => {
    const f = F.pend = F.pend || { est: 'abiertos', tipo: '', resp: '', q: '' };
    const hoy = U.todayISO();
    const q = U.norm(f.q);
    const l = S.list('pendientes', x => (f.est === 'todos' || (f.est === 'vencidos' ? x.estado !== 'Cerrado' && x.fechaCompromiso && x.fechaCompromiso < hoy : x.estado !== 'Cerrado'))
      && (!f.tipo || x.tipo === f.tipo) && (!f.resp || x.responsable === f.resp) && (!q || U.norm(x.descripcion + ' ' + A.pname(x.proyectoId)).includes(q)));
    const abiertos = S.list('pendientes', x => x.estado !== 'Cerrado');
    const porTipo = U.groupBy(abiertos, x => x.tipo || 'Sin tipo');
    return head('Pendientes de obra', `${abiertos.length} abiertos · ${abiertos.filter(x => x.fechaCompromiso && x.fechaCompromiso < hoy).length} vencidos · ${abiertos.filter(x => !x.fechaCompromiso).length} sin fecha compromiso`,
      `<button class="btn primary" onclick="A.formPendiente()">Nuevo pendiente</button>`) +
      `<div class="chips">${Object.entries(porTipo).map(([t, l]) => `<button class="chip ${f.tipo === t ? 'on' : ''}" onclick="A.fPend('tipo','${f.tipo === t ? '' : esc(t)}')">${esc(t)} <b>${l.length}</b></button>`).join('')}</div>
      <div class="filters">
        <label class="flt grow"><span>Buscar</span><input type="search" id="pq" value="${esc(f.q)}" oninput="A.fPend('q',this.value)" placeholder="Descripción o proyecto"></label>
        <div class="seg">${[['abiertos', 'Abiertos'], ['vencidos', 'Vencidos'], ['todos', 'Todos']].map(([k, l]) => `<button class="${f.est === k ? 'on' : ''}" onclick="A.fPend('est','${k}')">${l}</button>`).join('')}</div>
        ${sel('presp', 'Responsable', [...new Set(S.list('pendientes').map(x => x.responsable).filter(Boolean))].sort(), f.resp, "A.fPend('resp',this.value)")}
      </div>` + A.table('pend', pendCols(true), l, { onRow: x => `A.formPendiente(null,'${x.id}')`, sort: { k: 'fechaCompromiso', d: 1 }, empty: 'Sin pendientes con estos filtros.' });
  };
  A.fPend = (k, v) => { F.pend[k] = v; if (k === 'q') A._after = () => { const el = $('#pq'); el.focus(); el.setSelectionRange(v.length, v.length); }; rerender(); };

  /* ═══════════ HSE (todos) ═══════════ */
  VIEWS.hse = () => {
    const f = F.hse = F.hse || { tipo: '', est: '' };
    const obra = S.list('proyectos', A.EN_OBRA);
    const res = obra.map(p => ({ p, h: A.hseResumen(p) }));
    const regs = S.list('hse', x => (!f.tipo || x.tipo === f.tipo) && (!f.est || x.estado === f.est));
    const mes = U.todayISO().slice(0, 7);
    const regsMes = S.list('hse', x => (x.fecha || '').startsWith(mes));
    const ev = S.list('hse', x => /Incidente|Accidente/.test(x.tipo)).map(x => x.fecha).filter(Boolean).sort();
    const docs = S.list('contratistas').map(c => ({ c, d: A.contratistaDocs(c) })).filter(x => x.d.length);
    return head('HSE', `${obra.length} proyectos en obra · registros del mes: ${regsMes.length}`, `<button class="btn primary" onclick="A.formHSE()">Nuevo registro</button>`) +
      `${kpis([
kpi(ev.length ? U.daysBetween(ev[ev.length - 1], U.todayISO()) : '—', 'días sin incidentes registrados'),
kpi(regsMes.filter(x => x.tipo === 'Inspección').length, 'inspecciones este mes'),
kpi(U.sum(res, x => x.h.hallazgos), 'hallazgos abiertos', U.sum(res, x => x.h.hallazgos) ? 'k-warn' : ''),
kpi(U.sum(res, x => x.h.permisosVencidos), 'permisos vencidos sin cerrar', U.sum(res, x => x.h.permisosVencidos) ? 'k-bad' : ''),
kpi(docs.length, 'contratistas con documentos vencidos o por vencer', docs.length ? 'k-warn' : '')
], '')}
      <div class="grid-2">
        <section class="panel"><div class="panel-h"><h2>Cumplimiento documental por proyecto en obra</h2></div>
          <table class="mini"><thead><tr><th>Proyecto</th><th>Requisitos</th><th class="num">Pend.</th></tr></thead><tbody>
          ${res.sort((a, b) => a.h.cumplimiento - b.h.cumplimiento).map(x => `<tr class="clickable" onclick="A.go('#/proyecto/${x.p.id}/hse')"><td>${esc(x.p.nombre)}</td><td><div class="av">${A.bar(x.h.cumplimiento, x.h.cumplimiento < 0.6 ? 'b-bad' : '')}<span>${U.pct(x.h.cumplimiento)}</span></div></td><td class="num">${x.h.pend}</td></tr>`).join('') || '<tr><td colspan="3" class="muted">No hay proyectos en obra.</td></tr>'}
          </tbody></table></section>
        <section class="panel"><div class="panel-h"><h2>Documentos de contratistas</h2><a class="small" href="#/contratistas">Ver contratistas</a></div>
          ${docs.length ? `<ul class="alert-list">${docs.map(x => x.d.map(d => `<li>${A.dot(d.nivel)}<a href="#/contratistas"><strong>${esc(x.c.nombre)}</strong><span>${esc(d.txt)}</span></a></li>`).join('')).join('')}</ul>` : '<p class="muted">ARL, planillas y pólizas al día para los contratistas registrados.</p>'}
        </section>
      </div>
      <div class="filters">${sel('htipo', 'Tipo', A.HSE_TIPOS, f.tipo, "A.fHse('tipo',this.value)")}${sel('hest', 'Estado', ['Abierto', 'Cerrado'], f.est, "A.fHse('est',this.value)")}</div>` +
      A.table('hse', hseCols(true), regs, { onRow: x => `A.formHSE(null,'${x.id}')`, sort: { k: 'fecha', d: -1 }, empty: 'Sin registros HSE.' });
  };
  A.fHse = (k, v) => { F.hse[k] = v; rerender(); };

  /* ═══════════ FINANCIERO ═══════════ */
  const actaCols = (conProyecto) => [
    { k: 'fechaConciliacion', l: 'Conciliación', h: a => U.fmtDate(a.fechaConciliacion) },
    { k: 'acta', l: 'Acta', h: a => `${esc(a.acta)}${a.estimado ? ' <span class="tag">estimado</span>' : ''}` },
    ...(conProyecto ? [{ k: 'proyectoId', l: 'Proyecto', v: a => A.pname(a.proyectoId), h: a => `<select class="inline-sel ${a.revisarAsignacion ? 'revisar' : ''} ${!a.proyectoId ? 'sinp' : ''}" onclick="event.stopPropagation()" onchange="A.asignarActa('${a.id}',this.value)" aria-label="Proyecto del acta"><option value="">Sin asignar</option>${A.optProyectos().map(o => `<option value="${o.v}" ${o.v === a.proyectoId ? 'selected' : ''}>${esc(o.l)}</option>`).join('')}</select>${a.revisarAsignacion ? `<button class="link small" onclick="event.stopPropagation();A.confirmarActa('${a.id}')">Confirmar</button>` : ''}` }] : []),
    { k: 'valor', l: 'Valor', cls: 'num', h: a => U.money(a.valor) },
    { k: 'lineaNegocio', l: 'Línea', h: a => `<span class="small">${esc((a.lineaNegocio || '').replace('Generación y Renovables', 'Gen. y Renov.'))}</span>` },
    { k: 'factura', l: 'Factura', h: a => a.factura ? `${esc(a.factura)}<div class="muted small">${U.fmtDate(a.fechaRadicacion)}</div>` : `<span class="warn small">${esc(a.estadoFactura || 'Sin factura')}</span>` }
  ];
  VIEWS.financiero = () => {
    const f = F.fin = F.fin || { vista: 'proyectos', filtro: '' };
    const ps = S.list('proyectos');
    const actas = S.list('actas');
    const sinAsig = actas.filter(a => !a.proyectoId), revisar = actas.filter(a => a.revisarAsignacion);
    const tot = { proy: U.sum(ps.filter(A.ACTIVO), p => p.valorContrato), conc: U.sum(actas, a => a.valor), fact: U.sum(actas.filter(a => a.estadoFactura === 'Facturada'), a => a.valor) };
    const difLiq = U.sum(ps.filter(p => p.valorLiquidacion != null && p.valorContrato), p => p.valorLiquidacion - p.valorContrato);
    // facturación por mes de radicación
    const porMes = U.groupBy(actas.filter(a => a.fechaRadicacion && a.estadoFactura === 'Facturada'), a => a.fechaRadicacion.slice(0, 7));
    const meses = Object.keys(porMes).sort().slice(-12);
    const maxM = Math.max(1, ...meses.map(m => U.sum(porMes[m], a => a.valor)));
    let body;
    if (f.vista === 'proyectos') {
      const cols = [
        { k: 'nombre', l: 'Proyecto', h: p => `<strong>${esc(p.nombre)}</strong><div class="muted small">${esc(p.codigo || '')}</div>` },
        { k: 'valorContrato', l: 'Proyectado', cls: 'num', h: p => U.moneyShort(p.valorContrato) },
        { k: 'conc', l: 'Conciliado', cls: 'num', v: p => A.fin(p).conciliado, h: p => { const x = A.fin(p); return `${U.moneyShort(x.conciliado)}<div class="small muted">${U.pct(x.pctFact)}</div>`; } },
        { k: 'fact', l: 'Facturado', cls: 'num', v: p => A.fin(p).facturado, h: p => U.moneyShort(A.fin(p).facturado) },
        { k: 'valorLiquidacion', l: 'Liquidación', cls: 'num', h: p => U.moneyShort(p.valorLiquidacion) },
        { k: 'dif', l: 'Diferencia', cls: 'num', v: p => A.fin(p).difLiq, h: p => { const d = A.fin(p).difLiq; return `<span class="${d < 0 ? 'neg' : d > 0 ? 'ok' : ''}">${U.moneyShort(d)}</span>`; } },
        { k: 'estadoLiquidacion', l: 'Estado liq.', h: p => `<span class="small">${esc((p.estadoLiquidacion || '—').toLowerCase())}</span>` },
        { k: 'costo', l: 'Costos reg.', cls: 'num', v: p => A.fin(p).costo, h: p => U.moneyShort(A.fin(p).costo || null) },
        { k: 'margen', l: 'Margen', cls: 'num', v: p => A.fin(p).margenPct, h: p => { const m = A.fin(p).margenPct; return m === null ? '<span class="muted">—</span>' : `<span class="${m < 0.1 ? 'neg' : ''}">${U.pct(m)}</span>`; } }
      ];
      const l = ps.filter(p => !f.filtro || (f.filtro === 'liq' ? !['APROBADA'].includes(p.estadoLiquidacion) && p.estadoLiquidacion : f.filtro === 'neg' ? A.fin(p).difLiq < 0 : true));
      body = `<div class="filters"><div class="seg">${[['', 'Todos'], ['liq', 'Liquidación sin aprobar'], ['neg', 'Liquidación bajo lo proyectado']].map(([k, t]) => `<button class="${f.filtro === k ? 'on' : ''}" onclick="A.fFin('filtro','${k}')">${t}</button>`).join('')}</div></div>` +
        A.table('finp', cols, l, { onRow: p => `A.go('#/proyecto/${p.id}/financiero')`, sort: { k: 'dif', d: 1 } });
    } else {
      const l = actas.filter(a => f.filtro === 'sin' ? !a.proyectoId : f.filtro === 'rev' ? a.revisarAsignacion : f.filtro === 'nofact' ? a.estadoFactura !== 'Facturada' : f.filtro === 'est' ? a.estimado : true);
      body = `<div class="filters"><div class="seg">${[['', 'Todas'], ['sin', `Sin proyecto (${sinAsig.length})`], ['rev', `Por revisar (${revisar.length})`], ['nofact', 'Sin factura'], ['est', 'Estimados']].map(([k, t]) => `<button class="${f.filtro === k ? 'on' : ''}" onclick="A.fFin('filtro','${k}')">${t}</button>`).join('')}</div>
        <span class="muted small">${l.length} actas · ${U.money(U.sum(l, a => a.valor))}</span><span class="spacer"></span><button class="btn sm" onclick="A.formActa()">Nueva acta</button></div>` +
        A.table('actas', actaCols(true), l, { onRow: a => `A.formActa('${a.id}')`, sort: { k: 'fechaConciliacion', d: -1 } });
    }
    return head('Financiero', 'Valores proyectados, actas conciliadas con ENEL, facturación y liquidaciones') +
      `${kpis([
kpi(U.moneyShort(tot.proy), 'valor proyectado cartera activa'),
kpi(U.moneyShort(tot.conc), 'conciliado en actas'),
kpi(U.moneyShort(tot.fact), 'facturado'),
kpi(U.moneyShort(tot.conc - tot.fact), 'conciliado pendiente de factura', tot.conc - tot.fact > 0 ? 'k-warn' : ''),
kpi(U.moneyShort(difLiq), 'diferencia neta liquidación vs proyectado', difLiq < 0 ? 'k-bad' : ''),
kpi(sinAsig.length + revisar.length, 'actas por asignar o revisar', sinAsig.length + revisar.length ? 'k-warn' : '')
], '')}
      ${meses.length ? `<section class="panel"><div class="panel-h"><h2>Facturación radicada por mes</h2></div><div class="months">${meses.map(m => { const v = U.sum(porMes[m], a => a.valor); return `<div class="mcol" title="${U.money(v)}"><span class="mv">${U.moneyShort(v)}</span><i style="height:${Math.max(3, v / maxM * 120)}px"></i><span class="ml">${U.MESES[+m.slice(5) - 1]} ${m.slice(2, 4)}</span></div>`; }).join('')}</div></section>` : ''}
      <nav class="tabs"><a href="javascript:A.fFin('vista','proyectos')" class="${f.vista === 'proyectos' ? 'on' : ''}">Por proyecto</a><a href="javascript:A.fFin('vista','actas')" class="${f.vista === 'actas' ? 'on' : ''}">Actas y facturas${sinAsig.length + revisar.length ? `<sup>${sinAsig.length + revisar.length}</sup>` : ''}</a></nav>` + body;
  };
  A.fFin = (k, v) => { F.fin[k] = v; if (k === 'vista') F.fin.filtro = ''; rerender(); };
  A.asignarActa = (id, pid) => { const a = S.get('actas', id); a.proyectoId = pid || null; a.revisarAsignacion = false; S.upsert('actas', a); A.toast(pid ? 'Acta asignada' : 'Acta sin proyecto'); rerender(); };
  A.confirmarActa = (id) => { const a = S.get('actas', id); a.revisarAsignacion = false; S.upsert('actas', a); rerender(); };
  A.formActa = (id, pid) => {
    const a = id ? S.get('actas', id) : null;
    A.openForm({
      title: a ? 'Acta' : 'Nueva acta', wide: true, values: a || { proyectoId: pid || '', estadoFactura: 'Sin factura', fechaConciliacion: U.todayISO() },
      fields: [
        { k: 'acta', l: 'Acta', req: true, full: true, help: 'Ej. Acta No 89 - Cambio de Cubierta SE Salitre Corte 2' },
        { k: 'proyectoId', l: 'Proyecto', t: 'select', options: A.optProyectos }, { k: 'contrato', l: 'Contrato', list: () => ['JA10177023', 'JA10177024'] },
        { k: 'fechaConciliacion', l: 'Fecha de conciliación', t: 'date' }, { k: 'valor', l: 'Valor', t: 'money', req: true },
        { k: 'lineaNegocio', l: 'Línea de negocio', t: 'select', options: ['Distribución', 'Generación y Renovables'] }, { k: 'pedido', l: 'Pedido' },
        { k: 'conformidad', l: 'Conformidad (HES)' }, { k: 'estimado', l: 'Estimado', t: 'check', cl: 'Valor estimado (no definitivo)' },
        { k: 'estadoFactura', l: 'Estado factura', t: 'select', options: ['Sin factura', 'En emisión de factura', 'En actualización de valores', 'Facturada'], blank: false },
        { k: 'factura', l: 'N.º factura' }, { k: 'fechaRadicacion', l: 'Fecha de radicación', t: 'date' }, { k: 'idDrape', l: 'Id Drape' }, { k: 'gestor', l: 'Gestor', list: A.gestores }
      ],
      onSave: v => { if (v.factura && v.estadoFactura !== 'Facturada') v.estadoFactura = 'Facturada'; v.revisarAsignacion = false; S.upsert('actas', v); },
      onDelete: a ? v => S.remove('actas', v.id) : null
    });
  };

  /* ═══════════ CONTRATISTAS ═══════════ */
  VIEWS.contratistas = () => {
    const cs = S.list('contratistas');
    const contratos = S.list('contratos');
    const cols = [
      { k: 'nombre', l: 'Contratista', h: c => `<strong>${esc(c.nombre)}</strong><div class="muted small">${esc(c.nit || '')}${c.especialidad ? ' · ' + esc(c.especialidad) : ''}</div>` },
      { k: 'contacto', l: 'Contacto', h: c => `${esc(c.contacto || '')}<div class="muted small">${esc(c.telefono || '')}</div>` },
      { k: 'proy', l: 'Proyectos', cls: 'num', v: c => new Set(contratos.filter(x => x.contratistaId === c.id).map(x => x.proyectoId)).size, h: c => new Set(contratos.filter(x => x.contratistaId === c.id).map(x => x.proyectoId)).size },
      { k: 'valor', l: 'Contratado', cls: 'num', v: c => U.sum(contratos.filter(x => x.contratistaId === c.id), x => x.valor), h: c => U.moneyShort(U.sum(contratos.filter(x => x.contratistaId === c.id), x => x.valor)) },
      { k: 'saldo', l: 'Saldo por pagar', cls: 'num', v: c => U.sum(contratos.filter(x => x.contratistaId === c.id), x => (x.valor || 0) - U.sum(x.pagos || [], y => y.valor)), h: c => U.moneyShort(U.sum(contratos.filter(x => x.contratistaId === c.id), x => (x.valor || 0) - U.sum(x.pagos || [], y => y.valor))) },
      { k: 'docs', l: 'Documentos', v: c => -A.contratistaDocs(c).length, h: c => { const d = A.contratistaDocs(c); return d.length ? d.map(x => `<div class="small ${x.nivel === 3 ? 'neg' : 'warn'}">${esc(x.txt)}</div>`).join('') : (c.arlVence || c.planillaVence ? '<span class="small ok">Al día</span>' : '<span class="small muted">Sin fechas</span>'); } },
      { k: 'calificacion', l: 'Calificación', h: c => c.calificacion ? '★'.repeat(+c.calificacion) + '<span class="muted">' + '★'.repeat(5 - c.calificacion) + '</span>' : '<span class="muted">—</span>' }
    ];
    return head('Contratistas', `${cs.length} contratistas · ${contratos.length} contratos`, `<button class="btn" onclick="A.formContrato()">Nuevo contrato</button><button class="btn primary" onclick="A.formContratista()">Nuevo contratista</button>`) +
      A.table('ctr', cols, cs, { onRow: c => `A.formContratista('${c.id}')`, empty: 'Registra los contratistas con su NIT y las fechas de vigencia de ARL, planilla y pólizas.' }) +
      `<h3 class="sec">Contratos</h3>` + A.table('ctrs', contratoCols(true), contratos, { onRow: c => `A.formContrato('${c.id}')`, empty: 'Sin contratos.' });
  };
  A.formContratista = (id) => {
    const c = id ? S.get('contratistas', id) : null;
    A.openForm({
      title: c ? 'Contratista' : 'Nuevo contratista', wide: true, values: c || {},
      fields: [
        { k: 'nombre', l: 'Razón social / nombre', req: true }, { k: 'nit', l: 'NIT / cédula' },
        { k: 'especialidad', l: 'Especialidad', list: () => ['Obra civil', 'Cubiertas', 'Eléctrico', 'Hidrosanitario', 'Red contra incendio', 'Carpintería y mobiliario', 'Drywall y acabados', 'Estructura metálica', 'Señalización y branding', 'Aire acondicionado', 'Vidrios y aluminio'] },
        { k: 'contacto', l: 'Contacto' }, { k: 'telefono', l: 'Teléfono' }, { k: 'correo', l: 'Correo' },
        { k: 'arlVence', l: 'ARL vigente hasta', t: 'date' }, { k: 'planillaVence', l: 'Planilla seg. social vigente hasta', t: 'date', help: 'Normalmente fin del mes pagado.' },
        { k: 'polizaVence', l: 'Póliza vigente hasta', t: 'date' }, { k: 'calificacion', l: 'Calificación de desempeño', t: 'select', options: [{ v: 5, l: '5 · Excelente' }, { v: 4, l: '4 · Bueno' }, { v: 3, l: '3 · Aceptable' }, { v: 2, l: '2 · Deficiente' }, { v: 1, l: '1 · No volver a contratar' }] },
        { k: 'notas', l: 'Notas', t: 'textarea', full: true }
      ],
      onSave: v => { S.upsert('contratistas', v); },
      onDelete: c ? v => { if (S.list('contratos', x => x.contratistaId === v.id).length) { A.toast('Tiene contratos asociados; elimínalos primero.', 'err'); return; } S.remove('contratistas', v.id); } : null
    });
  };

  /* ═══════════ DATOS Y CONEXIÓN ═══════════ */
  VIEWS.datos = () => {
    const m = S.db.meta, w = m.pesos;
    const conn = S.mode === 'm365'
      ? `<p><b>Conectado a Microsoft 365</b> como ${esc(S.user.name)} (${esc(S.user.email)}).</p>
         <p class="muted small">Base: ${esc((root.APP_CONFIG.storage.hostname || 'OneDrive') + (root.APP_CONFIG.storage.sitePath || '') + ' / ' + root.APP_CONFIG.storage.folder + ' / ' + root.APP_CONFIG.storage.file)}${S.webUrl ? ` · <a href="${esc(S.webUrl)}" target="_blank" rel="noopener">abrir en SharePoint</a>` : ''}</p>
         <div class="row-gap"><button class="btn" onclick="Store.pull().then(()=>A.render())">Recargar desde SharePoint</button><button class="btn ghost" onclick="Store.logout()">Cerrar sesión</button></div>`
      : `<p><b>Modo local.</b> Los datos están solo en este navegador. Para compartirlos con el equipo y guardarlos en SharePoint, configura <code>js/config.js</code> (ver README).</p>
         <label class="flt"><span>Tu nombre (queda en los registros)</span><input type="text" value="${esc(S.user.name)}" onchange="A.setLocalUser(this.value)"></label>`;
    return head('Datos y conexión') +
      `<div class="grid-2">
        <section class="panel"><div class="panel-h"><h2>Almacenamiento</h2><span class="sync-inline">${esc(S.status)}</span></div>${conn}</section>
        <section class="panel"><div class="panel-h"><h2>Importar los Excel actuales</h2></div>
          <p class="small">Arrastra uno o varios archivos: seguimiento de liquidaciones, seguimiento semanal de ejecución, cronograma de proyectos y estatus de facturación. La herramienta reconoce cada formato y une la información por código o por nombre del proyecto. Se puede repetir: actualiza lo existente sin duplicar.</p>
          <label class="drop" id="drop"><input type="file" id="xls" multiple accept=".xlsx,.xlsm,.xls" onchange="A.importar(this.files)"><span>Soltar archivos .xlsx aquí o <u>seleccionarlos</u></span></label>
          <div id="imp-rep"></div>
        </section>
      </div>
      <div class="grid-2">
        <section class="panel"><div class="panel-h"><h2>Exportar y respaldar</h2></div>
          <div class="row-gap"><button class="btn" onclick="A.exportExcel()">Exportar a Excel</button><button class="btn" onclick="A.backup()">Descargar copia (JSON)</button>
          <label class="btn ghost">Restaurar copia<input type="file" accept=".json" hidden onchange="A.restore(this.files[0])"></label></div>
          <p class="muted small">El Excel exportado sirve para informes, Power BI o para enviar a ENEL. En modo Microsoft 365 además se guarda una copia diaria en la carpeta <i>copias</i>.</p>
        </section>
        <section class="panel"><div class="panel-h"><h2>Pesos del avance ponderado</h2><span class="muted small">Suman ${U.pct(U.sum(Object.values(w)))}</span></div>
          <div class="pesos">${A.COMP.map(c => `<label class="flt"><span>${c.l}</span><div class="pct-in"><input type="number" min="0" max="100" value="${Math.round((w[c.k] || 0) * 100)}" onchange="A.setPeso('${c.k}',this.value)"><span>%</span></div></label>`).join('')}</div>
        </section>
      </div>
      <section class="panel"><div class="panel-h"><h2>Requisitos HSE por proyecto</h2><span class="muted small">Uno por línea. Se aplican a todos los proyectos en obra.</span></div>
        <textarea class="note" rows="8" onchange="A.setReqs(this.value)">${esc(A.reqHSE().join('\n'))}</textarea>
      </section>`;
  };
  A.setLocalUser = (v) => { localStorage.setItem('enelctl.user', v); S.user.name = v; statusChip(); };
  A.setPeso = (k, v) => { S.setMeta({ pesos: Object.assign({}, S.db.meta.pesos, { [k]: U.clamp01((+v || 0) / 100) }) }); rerender(); };
  A.setReqs = (txt) => { S.setMeta({ hseRequisitos: txt.split('\n').map(s => s.trim()).filter(Boolean) }); A.toast('Requisitos actualizados'); };

  A.importar = async (files) => {
    if (!files || !files.length) return;
    const rep = $('#imp-rep'); rep.innerHTML = '<p class="muted">Leyendo archivos…</p>';
    try {
      const data = await Promise.all([...files].map(f => f.arrayBuffer().then(b => ({ name: f.name, data: new Uint8Array(b) }))));
      const r = root.Importer.run(data, S.db, S.stamp);
      S.setMeta({});    // marca cambios y guarda
      rep.innerHTML = `<div class="imp-ok">
        <p><b>Importación terminada.</b></p>
        <ul>${r.archivos.map(a => `<li>${esc(a.archivo)} → ${esc(a.tipo)} (hoja “${esc(a.hoja)}”, ${a.filas} filas)</li>`).join('')}</ul>
        <p>${r.creados.length} proyectos nuevos · ${r.actualizados.length} actualizados · ${r.pendientes} pendientes · ${r.actasNuevas} actas nuevas${r.actasSinProyecto ? ` (<a href="javascript:A.irActas('sin')">${r.actasSinProyecto} sin proyecto</a>)` : ''}</p>
        ${r.vinculados.length ? `<details><summary>${r.vinculados.length} nombres unidos a un proyecto existente — revisar</summary><ul class="small">${r.vinculados.map(v => `<li>${esc(v)}</li>`).join('')}</ul><p class="small muted">Si alguno no corresponde, edita el proyecto; si quedaron duplicados, usa “Fusionar” en la ficha.</p></details>` : ''}
        ${r.avisos.map(a => `<p class="warn small">${esc(a)}</p>`).join('')}
      </div>`;
      statusChip();
    } catch (e) { console.error(e); rep.innerHTML = `<p class="neg">No se pudo importar: ${esc(e.message)}</p>`; }
  };
  A.irActas = (f) => { F.fin = { vista: 'actas', filtro: f }; A.go('#/financiero'); };

  A.exportExcel = () => {
    const X = root.XLSX, wb = X.utils.book_new();
    const pn = (id) => { const p = S.get('proyectos', id); return p ? p.nombre : ''; };
    const pc = (id) => { const p = S.get('proyectos', id); return p ? p.codigo : ''; };
    const add = (name, rows) => X.utils.book_append_sheet(wb, X.utils.json_to_sheet(rows.length ? rows : [{ '': 'Sin registros' }]), name);
    add('Proyectos', S.list('proyectos').map(p => { const f = A.fin(p); return {
      'Código': p.codigo, 'Proyecto': p.nombre, 'Categoría': p.categoria, 'Gestor ENEL': p.gestor, 'Residente': p.residente, 'Estado': p.estado,
      'Inicio': p.fechaInicio, 'Entrega': p.fechaFin, 'Días para entrega': A.diasFin(p),
      'Avance ponderado': +A.pond(p).toFixed(4), 'Informe': p.avances.informe, 'Presupuesto': p.avances.presupuesto, 'Preoperacional': p.avances.preop, 'Obra': p.avances.obra, 'Liquidación': p.avances.liquidacion, 'Dossier': p.avances.dossier,
      'Obra programada': A.programado(p), 'Valor proyectado': p.valorContrato, 'Conciliado en actas': f.conciliado, 'Facturado': f.facturado,
      'Valor liquidación': p.valorLiquidacion, 'Diferencia liquidación': f.difLiq, 'Estado liquidación': p.estadoLiquidacion,
      'Costos registrados': f.costo, 'Margen %': f.margenPct, 'Alertas': A.alertas(p).map(a => a.txt).join(' | '), 'Obs. ENEL': p.obsCliente, 'Obs. internas': p.obsInterna }; }));
    add('Actividades', S.list('actividades').map(a => ({ 'Código': pc(a.proyectoId), 'Proyecto': pn(a.proyectoId), 'Actividad': a.nombre, 'Responsable': a.responsable, 'Inicio': a.inicio, 'Fin': a.fin, 'Avance': a.avance, 'Peso': a.peso })));
    add('Pendientes', S.list('pendientes').map(x => ({ 'Código': pc(x.proyectoId), 'Proyecto': pn(x.proyectoId), 'Pendiente': x.descripcion, 'Tipo': x.tipo, 'Prioridad': x.prioridad, 'Responsable': x.responsable, 'Compromiso': x.fechaCompromiso, 'Estado': x.estado, 'Cierre': x.fechaCierre })));
    add('HSE', S.list('hse').map(x => ({ 'Fecha': x.fecha, 'Código': pc(x.proyectoId), 'Proyecto': pn(x.proyectoId), 'Tipo': x.tipo, 'Subtipo': x.subtipo, 'Descripción': x.descripcion, 'Acción': x.accion, 'Responsable': x.responsable, 'Vence': x.vence, 'Estado': x.estado, 'Evidencias': (x.adjuntos || []).map(a => a.url).join(' ') })));
    add('Requisitos HSE', S.list('proyectos', A.EN_OBRA).flatMap(p => A.reqHSE().map(r => ({ 'Código': p.codigo, 'Proyecto': p.nombre, 'Requisito': r, 'Estado': (p.hseChecklist?.[r]?.estado) || 'Pendiente', 'Fecha': p.hseChecklist?.[r]?.fecha || '' }))));
    add('Contratistas', S.list('contratistas').map(c => ({ 'Contratista': c.nombre, 'NIT': c.nit, 'Especialidad': c.especialidad, 'Contacto': c.contacto, 'Teléfono': c.telefono, 'Correo': c.correo, 'ARL hasta': c.arlVence, 'Planilla hasta': c.planillaVence, 'Póliza hasta': c.polizaVence, 'Calificación': c.calificacion })));
    add('Contratos', S.list('contratos').map(c => ({ 'Código': pc(c.proyectoId), 'Proyecto': pn(c.proyectoId), 'Contratista': S.get('contratistas', c.contratistaId)?.nombre, 'N.º': c.numero, 'Objeto': c.objeto, 'Valor': c.valor, 'Pagado': U.sum(c.pagos || [], x => x.valor), 'Saldo': (c.valor || 0) - U.sum(c.pagos || [], x => x.valor), 'Estado': c.estado, 'Paz y salvo': c.pazYSalvo ? 'Sí' : 'No' })));
    add('Actas', S.list('actas').map(a => ({ 'Conciliación': a.fechaConciliacion, 'Acta': a.acta, 'Código': pc(a.proyectoId), 'Proyecto': pn(a.proyectoId), 'Contrato': a.contrato, 'Valor': a.valor, 'Línea': a.lineaNegocio, 'Pedido': a.pedido, 'Conformidad': a.conformidad, 'Radicación': a.fechaRadicacion, 'Factura': a.factura, 'Estado factura': a.estadoFactura, 'Id Drape': a.idDrape, 'Gestor': a.gestor })));
    X.writeFile(wb, `Control_ENEL_${A.corte()}.xlsx`);
  };
  A.backup = () => {
    const b = new Blob([S.exportJSON()], { type: 'application/json' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = `control-enel-copia-${U.todayISO()}.json`; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  };
  A.restore = async (file) => {
    if (!file) return;
    if (!confirm('Restaurar reemplaza TODOS los datos actuales por los de la copia. ¿Continuar?')) return;
    try { S.replaceAll(JSON.parse(await file.text())); A.toast('Copia restaurada'); A.render(); }
    catch (e) { A.toast('El archivo no es una copia válida.', 'err'); }
  };

  /* ═══════════ Búsqueda rápida ═══════════ */
  A.quick = (q) => {
    const box = $('#qres'); const n = U.norm(q);
    if (n.length < 2) { box.hidden = true; return; }
    const res = S.list('proyectos', p => U.norm([p.codigo, p.nombre, ...(p.alias || [])].join(' ')).includes(n)).slice(0, 8);
    box.innerHTML = res.length ? res.map(p => `<a href="#/proyecto/${p.id}" onclick="A.quickClose()"><span class="code">${esc(p.codigo || '')}</span>${esc(p.nombre)}</a>`).join('') : '<span class="muted small">Sin coincidencias</span>';
    box.hidden = false;
  };
  A.quickClose = () => { $('#qres').hidden = true; $('#qs').value = ''; };

  /* ═══════════ Arranque ═══════════ */
  A.start = async () => {
    buildNav();
    S.on(() => { statusChip(); if (S.rev !== A._rev) { A._rev = S.rev; A.render(); } });
    window.addEventListener('hashchange', () => { A.render(); $('#view').scrollTop = 0; document.body.classList.remove('nav-open'); });
    document.addEventListener('click', e => { if (!e.target.closest('.qsearch')) $('#qres').hidden = true; });
    document.addEventListener('dragover', e => { if (e.target.closest && e.target.closest('#drop')) { e.preventDefault(); e.target.closest('#drop').classList.add('over'); } });
    document.addEventListener('drop', e => { const d = e.target.closest && e.target.closest('#drop'); if (d) { e.preventDefault(); d.classList.remove('over'); A.importar(e.dataTransfer.files); } });
    try {
      const r = await S.init();
      A._rev = S.rev;
      if (r.needLogin) { $('#login').hidden = false; $('#shell').hidden = true; return; }
      $('#login').hidden = true; $('#shell').hidden = false;
      statusChip(); A.render();
    } catch (e) {
      console.error(e);
      $('#login').hidden = false; $('#shell').hidden = true;
      $('#login-err').textContent = e.message;
    }
  };
})(window);
