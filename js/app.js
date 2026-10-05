/* Núcleo de la aplicación: constantes, cálculos y componentes de interfaz */
(function (root) {
  const U = root.U, S = root.Store;
  const A = root.A = {};

  /* ── Catálogos ── */
  A.ESTADOS = ['PREPARACIÓN', 'PRELIMINARES', 'EN EJECUCIÓN', 'STAND BY', 'POST VENTAS', 'EN LIQUIDACIÓN', 'FINALIZADO'];
  A.ESTADO_CLS = { 'PREPARACIÓN': 'st-prep', 'PRELIMINARES': 'st-pre', 'EN EJECUCIÓN': 'st-ejec', 'STAND BY': 'st-sb', 'POST VENTAS': 'st-pv', 'EN LIQUIDACIÓN': 'st-liq', 'FINALIZADO': 'st-fin' };
  A.ACTIVO = (p) => !['FINALIZADO'].includes(p.estado);
  A.EN_OBRA = (p) => ['PRELIMINARES', 'EN EJECUCIÓN', 'POST VENTAS'].includes(p.estado);
  A.ESTADOS_LIQ = ['POR PRESENTAR', 'PENDIENTE', 'PRESENTADA', 'EN CORRECCIÓN', 'PRESENTADA CON CORRECCIONES', 'APROBADA'];
  A.COMP = [
    { k: 'informe', l: 'Informe' }, { k: 'presupuesto', l: 'Presupuesto' }, { k: 'preop', l: 'Preoperacional' },
    { k: 'obra', l: 'Obra / proyecto' }, { k: 'liquidacion', l: 'Liquidación' }, { k: 'dossier', l: 'Dossier' }
  ];
  A.PEND_TIPOS = ['Obra', 'Compra', 'Cliente / ENEL', 'Documental', 'HSE', 'Liquidación'];
  A.PRIORIDADES = ['Alta', 'Media', 'Baja'];
  A.PEND_ESTADOS = ['Abierto', 'En curso', 'Cerrado'];
  A.HSE_TIPOS = ['Inspección', 'ATS', 'Permiso de trabajo', 'Charla / capacitación', 'Observación / acto inseguro', 'Incidente', 'Accidente', 'Simulacro', 'Auditoría ENEL / interventoría'];
  A.PERMISOS = ['Trabajo en alturas', 'Trabajo en caliente', 'Riesgo eléctrico', 'Espacios confinados', 'Izaje de cargas', 'Excavación', 'Otro'];
  A.HSE_REQ_BASE = [
    'Programa de obra aprobado por ENEL',
    'Matriz de peligros y riesgos (IPEVR) del proyecto',
    'Procedimientos de trabajo seguro por actividad',
    'Afiliación ARL y planilla de seguridad social del personal',
    'Certificados de trabajo en alturas vigentes',
    'Exámenes médicos ocupacionales',
    'Coordinador de alturas / SST designado en obra',
    'Plan de emergencias y MEDEVAC de la sede',
    'Inspección de equipos, herramientas y EPP',
    'Hojas de seguridad (SDS) de productos químicos',
    'Plan de manejo de residuos (RCD) y certificados de disposición',
    'Inducción HSE de ENEL / sede al personal'
  ];
  A.REQ_ESTADOS = ['Pendiente', 'Entregado', 'Aprobado', 'N/A'];
  A.CONTRATO_ESTADOS = ['En negociación', 'Vigente', 'Terminado', 'Liquidado'];

  /* ── Cálculos ── */
  A.corte = () => S.db.meta.corte || U.todayISO();
  A.pond = (p) => {
    const w = S.db.meta.pesos;
    return A.COMP.reduce((s, c) => s + (w[c.k] || 0) * (p.avances?.[c.k] || 0), 0);
  };
  A.programado = (p) => {
    if (!p.fechaInicio || !p.fechaFin || p.inicioEstimado) return null;
    const tot = U.daysBetween(p.fechaInicio, p.fechaFin);
    if (!tot || tot <= 0) return null;
    return U.clamp01(U.daysBetween(p.fechaInicio, A.corte()) / tot);
  };
  A.diasFin = (p) => p.fechaFin ? U.daysBetween(A.corte(), p.fechaFin) : null;
  A.reqHSE = () => (S.db.meta.hseRequisitos && S.db.meta.hseRequisitos.length) ? S.db.meta.hseRequisitos : A.HSE_REQ_BASE;

  A.fin = (p) => {
    const actas = S.list('actas', a => a.proyectoId === p.id);
    const contratos = S.list('contratos', c => c.proyectoId === p.id);
    const facturado = U.sum(actas.filter(a => a.estadoFactura === 'Facturada'), a => a.valor);
    const conciliado = U.sum(actas, a => a.valor);
    const comprometido = U.sum(contratos, c => c.valor);
    const pagado = U.sum(contratos, c => U.sum(c.pagos || [], x => x.valor));
    const ingreso = p.valorLiquidacion || p.valorContrato || 0;
    const costo = comprometido + (+p.costosOtros || 0);
    const difLiq = (p.valorLiquidacion != null && p.valorContrato) ? p.valorLiquidacion - p.valorContrato : null;
    return {
      actas, contratos, facturado, conciliado, comprometido, pagado, saldo: comprometido - pagado, ingreso, costo,
      margen: costo ? ingreso - costo : null, margenPct: costo && ingreso ? (ingreso - costo) / ingreso : null,
      pctFact: p.valorContrato ? conciliado / p.valorContrato : null, difLiq
    };
  };

  A.hseResumen = (p) => {
    const reqs = A.reqHSE();
    const ck = p.hseChecklist || {};
    const pend = reqs.filter(r => !ck[r] || ck[r].estado === 'Pendiente').length;
    const regs = S.list('hse', h => h.proyectoId === p.id);
    const eventos = regs.filter(h => h.tipo === 'Incidente' || h.tipo === 'Accidente').map(h => h.fecha).filter(Boolean).sort();
    const ult = eventos[eventos.length - 1];
    const permisosVencidos = regs.filter(h => h.tipo === 'Permiso de trabajo' && h.estado !== 'Cerrado' && h.vence && h.vence < U.todayISO()).length;
    const hallazgos = regs.filter(h => h.estado === 'Abierto' && h.tipo !== 'Permiso de trabajo').length;
    return { reqs, pend, cumplimiento: reqs.length ? (reqs.length - pend) / reqs.length : 1, regs, ultEvento: ult, diasSin: ult ? U.daysBetween(ult, U.todayISO()) : null, permisosVencidos, hallazgos };
  };

  A.contratistaDocs = (c) => {
    const hoy = U.todayISO(), pronto = U.addDays(hoy, 15);
    const out = [];
    [['arlVence', 'ARL'], ['planillaVence', 'Planilla seg. social'], ['polizaVence', 'Póliza']].forEach(([k, l]) => {
      if (!c[k]) return;
      if (c[k] < hoy) out.push({ nivel: 3, txt: `${l} vencida (${U.fmtDate(c[k])})` });
      else if (c[k] <= pronto) out.push({ nivel: 2, txt: `${l} vence ${U.fmtDate(c[k])}` });
    });
    return out;
  };

  /* Alertas por proyecto: nivel 3 = alta, 2 = media, 1 = baja */
  A.alertas = (p) => {
    const out = [], add = (nivel, txt, tab) => out.push({ nivel, txt, p, tab });
    const obra = p.avances?.obra || 0, prog = A.programado(p), dias = A.diasFin(p);
    const cerrado = ['FINALIZADO', 'EN LIQUIDACIÓN', 'POST VENTAS'].includes(p.estado);
    if (!cerrado && p.estado !== 'STAND BY') {
      if (dias !== null && dias < 0 && obra < 1) add(3, `Entrega vencida hace ${-dias} días con ${U.pct(obra)} de obra`, 'cronograma');
      else if (dias !== null && dias <= 14 && obra < 0.9) add(2, `Entrega en ${dias} días con ${U.pct(obra)} de obra`, 'cronograma');
      if (prog !== null && p.estado === 'EN EJECUCIÓN') {
        const gap = obra - prog;
        if (gap < -0.15) add(3, `Obra ${U.pct(obra)} frente a ${U.pct(prog)} programado`, 'cronograma');
        else if (gap < -0.05) add(2, `Obra ${U.pct(obra)} frente a ${U.pct(prog)} programado`, 'cronograma');
      }
      if (A.EN_OBRA(p) && !p.fechaFin) add(1, 'Sin fecha de entrega en el cronograma', 'resumen');
      if (A.EN_OBRA(p) && p.inicioEstimado) add(1, 'Fecha de inicio sin confirmar: no se puede medir atraso', 'resumen');
      if (p.estado === 'PRELIMINARES' && obra >= 0.5) add(1, `Estado preliminares con ${U.pct(obra)} de avance de obra: verificar dato`, 'resumen');
    }
    const hoy = U.todayISO();
    const venc = S.list('pendientes', x => x.proyectoId === p.id && x.estado !== 'Cerrado' && x.fechaCompromiso && x.fechaCompromiso < hoy).length;
    if (venc) add(2, `${venc} pendiente${venc > 1 ? 's' : ''} con fecha vencida`, 'pendientes');
    if (['FINALIZADO', 'POST VENTAS'].includes(p.estado) && ['PENDIENTE', 'POR PRESENTAR', ''].includes(p.estadoLiquidacion || ''))
      add(2, 'Obra terminada y liquidación sin presentar', 'financiero');
    if (/CORRECCI/.test(p.estadoLiquidacion || '')) add(1, 'Liquidación en corrección', 'financiero');
    const f = A.fin(p);
    if (f.difLiq !== null && p.valorContrato && f.difLiq / p.valorContrato < -0.1)
      add(2, `Liquidación ${U.pct(-f.difLiq / p.valorContrato)} por debajo del valor proyectado (${U.moneyShort(f.difLiq)})`, 'financiero');
    // Solo cuando la liquidación ya tiene un valor real (presentada, en corrección o aprobada)
    const liqReal = p.estadoLiquidacion && !['PENDIENTE', 'POR PRESENTAR'].includes(p.estadoLiquidacion);
    if (liqReal && p.valorLiquidacion && f.conciliado > p.valorLiquidacion * 1.02)
      add(p.estadoLiquidacion === 'APROBADA' ? 3 : 2, `Actas conciliadas (${U.moneyShort(f.conciliado)}) superan la liquidación (${U.moneyShort(p.valorLiquidacion)}): posible nota crédito o actas mal asignadas`, 'financiero');
    if (f.margenPct !== null && f.margenPct < 0.1) add(3, `Margen sobre costos registrados de ${U.pct(f.margenPct)}`, 'financiero');
    if (A.EN_OBRA(p)) {
      const h = A.hseResumen(p);
      if (h.pend) add(1, `${h.pend} requisito${h.pend > 1 ? 's' : ''} HSE pendiente${h.pend > 1 ? 's' : ''}`, 'hse');
      if (h.permisosVencidos) add(3, `${h.permisosVencidos} permiso${h.permisosVencidos > 1 ? 's' : ''} de trabajo vencido${h.permisosVencidos > 1 ? 's' : ''} sin cerrar`, 'hse');
      f.contratos.filter(c => c.estado === 'Vigente').forEach(c => {
        const ct = S.get('contratistas', c.contratistaId);
        if (ct) A.contratistaDocs(ct).filter(d => d.nivel === 3).forEach(d => add(3, `${ct.nombre}: ${d.txt}`, 'contratistas'));
      });
    }
    return out.sort((a, b) => b.nivel - a.nivel);
  };

  /* ── Componentes ── */
  A.badge = (estado) => `<span class="badge ${A.ESTADO_CLS[estado] || ''}">${U.esc(estado || '—')}</span>`;
  A.bar = (v, cls = '') => `<span class="bar ${cls}"><i style="width:${Math.round(U.clamp01(v) * 100)}%"></i></span>`;
  A.dot = (nivel) => `<span class="dot n${nivel}" aria-label="${['', 'baja', 'media', 'alta'][nivel]}"></span>`;
  A.pname = (id) => { const p = S.get('proyectos', id); return p ? p.nombre : '—'; };
  A.plink = (id) => { const p = S.get('proyectos', id); return p ? `<a href="#/proyecto/${p.id}">${U.esc(p.nombre)}</a>` : '<span class="muted">Sin proyecto</span>'; };

  /* Tabla con orden por columna. cols: [{k,l,v:(r)=>valor de orden, h:(r)=>html, cls, w}] */
  const sortState = {};
  A.table = (id, cols, rows, opts = {}) => {
    const st = sortState[id] || opts.sort || null;
    let data = rows.slice();
    if (st) {
      const c = cols.find(x => x.k === st.k);
      if (c) {
        const val = c.v || (r => r[c.k]);
        data.sort((a, b) => {
          const x = val(a), y = val(b);
          if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1;
          return (typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'es')) * st.d;
        });
      }
    }
    const head = cols.map(c => `<th class="${c.cls || ''}" ${c.w ? `style="width:${c.w}"` : ''}>${c.nosort ? c.l : `<button class="th-sort" onclick="A.sortBy('${id}','${c.k}')">${c.l}${st && st.k === c.k ? (st.d > 0 ? ' ▲' : ' ▼') : ''}</button>`}</th>`).join('');
    const body = data.length ? data.map(r => `<tr ${opts.onRow ? `class="clickable" onclick="${opts.onRow(r)}"` : ''}>${cols.map(c => `<td class="${c.cls || ''}">${c.h ? c.h(r) : U.esc(r[c.k] ?? '')}</td>`).join('')}</tr>`).join('')
      : `<tr><td colspan="${cols.length}" class="empty">${opts.empty || 'Sin registros.'}</td></tr>`;
    return `<div class="tbl-wrap"><table class="tbl" id="t-${id}"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  };
  A.sortBy = (id, k) => {
    const st = sortState[id];
    sortState[id] = { k, d: st && st.k === k ? -st.d : 1 };
    A.render();
  };

  /* Diagrama de Gantt semanal. rows: {label, sub, start, end, prog, cls, href} */
  A.gantt = (rows, opts = {}) => {
    const corte = A.corte();
    const dates = rows.flatMap(r => [r.start, r.end]).filter(Boolean).sort();
    let start = U.mondayOf(opts.start || (dates[0] && dates[0] < U.addDays(corte, -21) ? U.addDays(corte, -21) : (dates[0] || corte)));
    if (start > U.mondayOf(corte)) start = U.mondayOf(U.addDays(corte, -7));
    let end = opts.end || dates[dates.length - 1] || U.addDays(corte, 56);
    if (U.daysBetween(start, end) > 7 * 40) end = U.addDays(start, 7 * 40);
    if (U.daysBetween(start, end) < 7 * 8) end = U.addDays(start, 7 * 8);
    const weeks = Math.ceil((U.daysBetween(start, end) + 1) / 7);
    // Ancho de semana: llena el espacio disponible sin bajar de 30 px
    const vw = window.innerWidth, lab = vw > 1100 ? 300 : vw > 760 ? 220 : 150;
    const avail = vw - (vw > 760 ? 232 : 0) - 60 - lab;
    const W = Math.max(opts.weekW || 30, Math.min(64, Math.floor(avail / weeks))), total = weeks * W;
    const x = (iso) => Math.max(0, Math.min(total, U.daysBetween(start, iso) / 7 * W));
    // encabezado de meses
    let months = [], cur = null;
    for (let i = 0; i < weeks; i++) {
      const d = U.fromISO(U.addDays(start, i * 7 + 3));
      const key = d.getFullYear() + '-' + d.getMonth();
      if (!cur || cur.key !== key) { cur = { key, l: U.MESES[d.getMonth()] + ' ' + String(d.getFullYear()).slice(2), n: 0 }; months.push(cur); }
      cur.n++;
    }
    const mh = months.map(m => `<div class="g-month" style="width:${m.n * W}px">${m.l}</div>`).join('');
    const wh = Array.from({ length: weeks }, (_, i) => {
      const d = U.fromISO(U.addDays(start, i * 7));
      return `<div class="g-week ${U.addDays(start, i * 7) === U.mondayOf(corte) ? 'is-corte' : ''}" style="width:${W}px">${d.getDate()}</div>`;
    }).join('');
    const cx = x(corte);
    const body = rows.map(r => {
      if (r.group) return `<div class="g-row g-group"><div class="g-label">${U.esc(r.label)}</div><div class="g-track" style="width:${total}px"></div></div>`;
      let barra = '';
      if (r.start && r.end) {
        const l = x(r.start), w = Math.max(6, x(U.addDays(r.end, 1)) - l);
        barra = `<div class="g-bar ${r.cls || ''} ${r.est ? 'est' : ''}" style="left:${l}px;width:${w}px" title="${U.esc(r.label)}: ${U.fmtDate(r.start)} → ${U.fmtDate(r.end)}${r.est ? ' (inicio sin confirmar)' : ''}"><i style="width:${Math.round(U.clamp01(r.prog) * 100)}%"></i></div>`;
      } else if (r.end) {
        barra = `<div class="g-mile ${r.cls || ''}" style="left:${x(r.end) - 5}px" title="Entrega ${U.fmtDate(r.end)} (sin fecha de inicio)"></div>`;
      } else barra = `<div class="g-nodate">Sin fechas</div>`;
      const lab = r.href ? `<a href="${r.href}">${U.esc(r.label)}</a>` : U.esc(r.label);
      return `<div class="g-row"><div class="g-label"><span class="g-name">${lab}</span>${r.sub ? `<span class="g-sub">${r.sub}</span>` : ''}</div><div class="g-track" style="width:${total}px">${barra}</div></div>`;
    }).join('');
    return `<div class="gantt" style="--gw:${W}px"><div class="g-scroll"><div class="g-inner" style="width:calc(var(--glab) + ${total}px)">
      <div class="g-head"><div class="g-label g-corner">Corte ${U.fmtDate(corte)}</div><div><div class="g-months">${mh}</div><div class="g-weeks">${wh}</div></div></div>
      <div class="g-body">${body}<div class="g-corte" style="left:calc(var(--glab) + ${cx}px)"></div></div>
    </div></div></div>`;
  };

  /* ── Formularios modales ── */
  A.openForm = ({ title, fields, values = {}, onSave, onDelete, saveLabel = 'Guardar', wide = false, note = '' }) => {
    const v = Object.assign({}, values);
    const opts = (f) => typeof f.options === 'function' ? f.options(v) : f.options;
    const input = (f) => {
      const val = v[f.k];
      const id = 'f-' + f.k;
      const req = f.req ? 'required' : '';
      switch (f.t) {
        case 'textarea': return `<textarea id="${id}" rows="${f.rows || 3}" ${req}>${U.esc(val || '')}</textarea>`;
        case 'select': return `<select id="${id}" ${req}>${f.blank !== false ? '<option value="">—</option>' : ''}${opts(f).map(o => { const ov = typeof o === 'object' ? o.v : o, ol = typeof o === 'object' ? o.l : o; return `<option value="${U.esc(ov)}" ${String(val ?? '') === String(ov) ? 'selected' : ''}>${U.esc(ol)}</option>`; }).join('')}</select>`;
        case 'date': return `<input type="date" id="${id}" value="${U.esc(val || '')}" ${req}>`;
        case 'money': return `<input type="text" inputmode="numeric" id="${id}" data-money value="${val == null || val === '' ? '' : Math.round(val).toLocaleString('es-CO')}" ${req}>`;
        case 'number': return `<input type="number" step="any" id="${id}" value="${val ?? ''}" ${req}>`;
        case 'pct': return `<div class="pct-in"><input type="number" min="0" max="100" step="1" id="${id}" value="${val == null ? '' : Math.round(val * 100)}" ${req}><span>%</span></div>`;
        case 'check': return `<label class="check"><input type="checkbox" id="${id}" ${val ? 'checked' : ''}> ${U.esc(f.cl || '')}</label>`;
        case 'file': return `<input type="file" id="${id}" ${S.mode !== 'm365' ? 'disabled' : ''}><small class="muted">${S.mode !== 'm365' ? 'Disponible al conectar Microsoft 365.' : 'Se guarda en SharePoint (máx. 4 MB).'}</small>${(v.adjuntos || []).map(a => `<div class="adj"><a href="${U.esc(a.url)}" target="_blank" rel="noopener">${U.esc(a.nombre)}</a></div>`).join('')}`;
        default: return `<input type="${f.t === 'url' ? 'url' : 'text'}" id="${id}" value="${U.esc(val ?? '')}" ${req} ${f.list ? `list="dl-${f.k}"` : ''}>${f.list ? `<datalist id="dl-${f.k}">${f.list().map(o => `<option value="${U.esc(o)}">`).join('')}</datalist>` : ''}`;
      }
    };
    const html = `<form class="modal-card ${wide ? 'wide' : ''}" id="modal-form" novalidate>
      <header><h2>${U.esc(title)}</h2><button type="button" class="icon-btn" onclick="A.closeModal()" aria-label="Cerrar">✕</button></header>
      ${note ? `<p class="form-note">${note}</p>` : ''}
      <div class="form-grid">${fields.map(f => `<div class="field ${f.full ? 'full' : ''}"><label for="f-${f.k}">${U.esc(f.l)}${f.req ? ' *' : ''}</label>${input(f)}${f.help ? `<small class="muted">${f.help}</small>` : ''}</div>`).join('')}</div>
      <footer>${onDelete ? `<button type="button" class="btn danger ghost" id="f-del">Eliminar</button>` : ''}<span class="spacer"></span><button type="button" class="btn ghost" onclick="A.closeModal()">Cancelar</button><button type="submit" class="btn primary">${saveLabel}</button></footer>
    </form>`;
    A.showModal(html);
    const form = document.getElementById('modal-form');
    form.querySelectorAll('[data-money]').forEach(el => el.addEventListener('blur', () => { const n = U.num(el.value); el.value = n == null ? '' : Math.round(n).toLocaleString('es-CO'); }));
    if (onDelete) document.getElementById('f-del').onclick = () => { if (confirm('¿Eliminar este registro? Queda en el historial de versiones de SharePoint.')) { onDelete(v); A.closeModal(); A.render(); } };
    form.onsubmit = async (e) => {
      e.preventDefault();
      const out = Object.assign({}, values);
      for (const f of fields) {
        const el = document.getElementById('f-' + f.k);
        if (!el) continue;
        let val;
        if (f.t === 'money' || f.t === 'number') val = U.num(el.value);
        else if (f.t === 'pct') val = el.value === '' ? null : U.clamp01(U.num(el.value) / 100);
        else if (f.t === 'check') val = el.checked;
        else if (f.t === 'file') {
          if (el.files && el.files[0]) {
            try { A.toast('Subiendo adjunto…'); const a = await S.upload(el.files[0], f.folder ? f.folder(out) : 'general'); out.adjuntos = [...(out.adjuntos || []), a]; }
            catch (err) { A.toast(err.message, 'err'); return; }
          }
          continue;
        } else val = el.value.trim();
        if (f.req && (val === '' || val === null)) { el.focus(); A.toast(`Completa “${f.l}”.`, 'err'); return; }
        out[f.k] = val;
      }
      const r = onSave(out);
      if (r === false) return;
      A.closeModal(); A.render();
    };
    setTimeout(() => { const first = form.querySelector('input:not([type=checkbox]):not([disabled]),select,textarea'); if (first) first.focus(); }, 30);
  };
  A.showModal = (html) => {
    const m = document.getElementById('modal');
    m.innerHTML = html; m.hidden = false; document.body.classList.add('modal-open');
  };
  A.closeModal = () => {
    const m = document.getElementById('modal');
    m.hidden = true; m.innerHTML = ''; document.body.classList.remove('modal-open');
    if (A._pendingRender) { A._pendingRender = false; A.render(); }
  };
  A.modalOpen = () => !document.getElementById('modal').hidden;
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && A.modalOpen()) A.closeModal(); });

  let toastT;
  A.toast = (msg, kind = '') => {
    const t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast show ' + kind;
    clearTimeout(toastT); toastT = setTimeout(() => t.className = 'toast', kind === 'err' ? 6000 : 2600);
  };

  /* Opciones reutilizables para formularios */
  A.optProyectos = () => S.list('proyectos').sort((a, b) => a.nombre.localeCompare(b.nombre, 'es')).map(p => ({ v: p.id, l: (p.codigo ? p.codigo + ' · ' : '') + p.nombre }));
  A.optContratistas = () => S.list('contratistas').sort((a, b) => a.nombre.localeCompare(b.nombre, 'es')).map(c => ({ v: c.id, l: c.nombre }));
  A.personas = () => [...new Set([...S.list('proyectos').map(p => p.residente), ...S.list('pendientes').map(p => p.responsable), ...S.list('hse').map(h => h.responsable)].filter(Boolean))].sort();
  A.gestores = () => [...new Set(S.list('proyectos').map(p => p.gestor).filter(Boolean))].sort();
  A.categorias = () => [...new Set(S.list('proyectos').map(p => p.categoria).filter(Boolean))].sort();
})(window);
