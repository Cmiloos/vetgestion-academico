/* VetGestión — lógica de la interfaz.
   Los datos vienen del Excel a través de window.pywebview.api (ver app/api.py). */

/* ---------- utilidades ---------- */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const js = v => esc(JSON.stringify(v));            // para pasar valores dentro de onclick="..."
const money = n => '$' + Math.round(Number(n) || 0).toLocaleString('es-AR');
const plural = (n, s, p) => `${n} ${n === 1 ? s : (p || s + 's')}`;
const vpill = {vencido:['p-red','Vencido'], pronto:['p-amber','Pronto'], ok:['p-ok','Vigente'], sin_dato:['p-blue','Sin fecha']};
const api = () => window.pywebview.api;
const FOTO_GENERICA = 'data:image/svg+xml,' + encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" fill="none" stroke="#1d9e75" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
     <rect x="14" y="8" width="36" height="48" rx="6"/><path d="M14 24h36M26 16h12"/></svg>`);

let S = {productos: [], movimientos: [], ventasHoy: {importe: 0, unidades: 0}, stockBajo: 5, diasPronto: 90};
let role = null, usuario = '';
const imgCache = {};
const imgSrc = p => imgCache[p.imagen] || FOTO_GENERICA;
const prod = c => S.productos.find(p => p.codigo === c);
const labs = () => [...new Set(S.productos.map(p => p.laboratorio).filter(Boolean))].sort((a, b) => a.localeCompare(b));
const cats = (lab) => [...new Set(S.productos.filter(p => !lab || p.laboratorio === lab).map(p => p.categoria).filter(Boolean))].sort((a, b) => a.localeCompare(b));

function cuando(iso) {
  const f = new Date(iso); if (isNaN(f)) return '';
  const min = Math.round((Date.now() - f) / 60000);
  const hora = f.toLocaleTimeString('es-AR', {hour: '2-digit', minute: '2-digit'});
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  if (min < 1) return 'recién';
  if (min < 60) return `hace ${min} min`;
  if (f >= hoy) return `hoy, ${hora}`;
  if (f >= hoy - 864e5) return `ayer, ${hora}`;
  return f.toLocaleDateString('es-AR', {day: 'numeric', month: 'long'});
}

/* ---------- carga de datos ---------- */
const enNube = () => S.almacen?.modo === 'sheets';
const donde = () => enNube() ? 'Google Sheets' : 'el Excel';
const fmtPct = n => n.toLocaleString('es-AR');

/* Aviso "Guardando…" para operaciones lentas (Google Sheets tarda 1-3 s). Solo aparece si demora. */
let ocupadoN = 0;
async function conAviso(txt, fn) {
  ocupadoN++;
  const t = setTimeout(() => { $('ocupadoTxt').textContent = txt; $('ocupado').hidden = false; }, 250);
  try { return await fn(); }
  finally { clearTimeout(t); if (--ocupadoN === 0) $('ocupado').hidden = true; }
}

async function cargar(silencioso = false) {
  const r = await (silencioso ? api().estado() : conAviso(enNube() ? 'Leyendo Google Sheets…' : 'Cargando…', () => api().estado()));
  if (!r.ok) { if (!silencioso) mostrarError(r.error); return false; }
  S = r;
  const faltan = [...new Set(S.productos.map(p => p.imagen).filter(n => n && !(n in imgCache)))];
  await Promise.all(faltan.map(async n => { imgCache[n] = await api().imagen(n); }));
  renderTodo();
  return true;
}

function renderTodo() {
  $('gateFoot').textContent = `Uso interno de NeoZoo · ${plural(labs().length, 'laboratorio')} · ${plural(S.productos.length, 'producto')}`;
  renderAlmacen();
  renderPanel();
  renderGrid($('q').value);
  if (window._sel) { prod(window._sel) ? renderFicha(window._sel) : ($('ficha').innerHTML = '', window._sel = null); }
  renderTable($('tq').value);
  llenarSelectsPrecios();
  preview();
  renderHist();
  llenarSelectsForm();
  renderAdminLista();
  if ($('s-movs').classList.contains('show')) cargarMovs();
  if ($('s-lista').classList.contains('show')) renderLista();
}

/* Llama a una función de la API; si falla muestra el error y devuelve null. */
async function llamar(fn, ...args) {
  const r = await conAviso(enNube() ? 'Guardando en Google Sheets…' : 'Guardando…', () => api()[fn](...args));
  if (!r.ok) { mostrarError(r.error); return null; }
  return r;
}

async function iniciar() {
  // con Google Sheets se muestra al instante la copia de la última vez y se actualiza de fondo
  const copia = await api().estado_guardado();
  if (copia.ok) {
    S = copia;
    const faltan = [...new Set(S.productos.map(p => p.imagen).filter(n => n && !(n in imgCache)))];
    await Promise.all(faltan.map(async n => { imgCache[n] = await api().imagen(n); }));
    renderTodo();
    $('cargando').hidden = true;
    avisoSync('Actualizando desde Google Sheets…');
    const ok = await cargar(true);
    if (!pendientes) avisoSync(ok ? 'Actualizado' : 'Sin conexión: se muestra la última copia', ok);
  } else {
    await cargar();
    $('cargando').hidden = true;
  }
  // al volver a la ventana se releen los datos (cambios de otra computadora o del Excel editado a mano)
  window.addEventListener('focus', () => { if (role && !pendientes) cargar(true); });
  // con Google Sheets, además se refresca solo cada 30 s si no estás escribiendo ni con un cartel abierto
  setInterval(() => {
    const escribiendo = ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName);
    if (role && enNube() && $('modal').hidden && !escribiendo && !ocupadoN && !pendientes) cargar(true);
  }, 30000);
}
if (window.pywebview && window.pywebview.api) iniciar();
else window.addEventListener('pywebviewready', iniciar);

/* ---------- acceso ---------- */
function chooseEmpleado() { role = 'empleado'; usuario = 'Empleado'; enterApp('buscar', 'Empleado', 'Mostrador', 'E'); }
function askPw() { $('choices').style.display = 'none'; $('gate-q').style.display = 'none'; $('pwbox').style.display = 'block'; $('pw').focus(); }
async function tryPw() {
  if (await api().login($('pw').value)) { role = 'vet'; usuario = 'Fátima'; enterApp('panel', 'Fátima', 'Administradora', 'F'); }
  else { $('pwerr').style.display = 'block'; $('pw').select(); }
}
function backGate() {
  role = null; usuario = '';
  $('app').style.display = 'none'; $('gate').style.display = 'flex';
  $('choices').style.display = 'grid'; $('gate-q').style.display = 'block'; $('pwbox').style.display = 'none';
  $('pw').value = ''; $('pwerr').style.display = 'none';
  cerrarModal(); limpiarForm();
}
function enterApp(start, name, roleLabel, initial) {
  $('gate').style.display = 'none'; $('app').style.display = 'grid';
  $('un').textContent = name; $('ur').textContent = roleLabel; $('av').textContent = initial;
  document.querySelectorAll('#nav button').forEach(b => { b.style.display = (role === 'vet' || !b.dataset.role) ? 'flex' : 'none'; });
  $('btnExcel').style.display = role === 'vet' ? 'flex' : 'none';
  renderTodo();
  go(start);
}

/* ---------- navegación ---------- */
const titles = {
  panel: ['Panel', 'Estado general de un vistazo'],
  buscar: ['Buscar y vender', 'Encontrá el producto y registrá la venta'],
  tabla: ['Tabla de productos', 'Todos los productos, como una planilla'],
  precios: ['Actualizar precios', 'Aumentos por laboratorio, categoría o archivo'],
  admin: ['Administración', 'Alta, edición, stock y baja de productos'],
  movs: ['Movimientos', 'Historial de ventas, stock y precios de la base de datos'],
};
function go(s) {
  document.querySelectorAll('#nav button').forEach(x => x.classList.toggle('active', x.dataset.s === (s === 'lista' ? 'panel' : s)));
  document.querySelectorAll('.screen').forEach(x => x.classList.remove('show'));
  $('s-' + s).classList.add('show');
  $('ttl').textContent = titles[s][0]; $('sub2').textContent = titles[s][1];
  $('modes').style.display = s === 'buscar' ? 'flex' : 'none';
  if (s === 'panel') animarPanel();
  if (s === 'movs') cargarMovs();
}
document.querySelectorAll('#nav button').forEach(b => b.onclick = () => go(b.dataset.s));
document.querySelectorAll('#modes button').forEach(b => b.onclick = () => {
  document.querySelectorAll('#modes button').forEach(x => x.classList.remove('on'));
  b.classList.add('on'); if (window._sel) renderFicha(window._sel);
});
async function abrirBase() { await llamar('abrir_base'); }

/* ---------- modales ---------- */
function abrirModal(html) { $('modal').innerHTML = html; $('modal').hidden = false; }
function cerrarModal() { $('modal').hidden = true; $('modal').innerHTML = ''; }
$('modal').addEventListener('click', e => { if (e.target.id === 'modal') cerrarModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('modal').hidden) cerrarModal(); });

const ICO_OK = '<path d="M20 6 9 17l-5-5"/>';
const ICO_ERR = '<path d="M12 8v5m0 4h.01"/><circle cx="12" cy="12" r="9"/>';
function confirmar({titulo, texto, detalle, valor, boton = 'Continuar', tipo = ''}) {
  abrirModal(`
    <div class="modal-box modal-ok" role="dialog" aria-modal="true">
      <span class="ok-ic ${tipo}"><svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round">${tipo ? ICO_ERR : ICO_OK}</svg></span>
      <h3>${esc(titulo)}</h3>
      <p>${esc(texto)}</p>
      ${detalle ? `<div class="ok-box"><div class="ok-k">${esc(detalle)}</div>${valor ? `<div class="ok-v">${esc(valor)}</div>` : ''}</div>` : ''}
      <button class="btn" style="margin-top:${detalle ? 0 : 20}px" onclick="cerrarModal()">${esc(boton)}</button>
    </div>`);
}
function mostrarError(msg) { confirmar({titulo: 'No se pudo completar', texto: msg, boton: 'Entendido', tipo: 'err'}); }

/* Diálogo con confirmación y, opcionalmente, un número a ingresar. */
function dialogo({titulo, texto, numero = null, boton = 'Aceptar', peligro = false, alAceptar}) {
  window._dlgOk = async () => {
    const v = numero ? $('dlgNum').value : null;
    cerrarModal();
    await alAceptar(v);
  };
  abrirModal(`
    <div class="modal-box modal-ok" role="dialog" aria-modal="true">
      <span class="ok-ic ${peligro ? 'err' : 'warn'}"><svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">${ICO_ERR}</svg></span>
      <h3>${esc(titulo)}</h3>
      <p>${texto}</p>
      ${numero ? `<div class="dlg-in"><input type="number" id="dlgNum" value="${numero.valor}" min="${numero.min ?? 1}" ${numero.max ? `max="${numero.max}"` : ''} onkeydown="if(event.key==='Enter')_dlgOk()"></div>` : ''}
      <div class="dlg-btns">
        <button class="btn ghost" onclick="cerrarModal()">Cancelar</button>
        <button class="btn ${peligro ? 'danger' : ''}" onclick="_dlgOk()">${esc(boton)}</button>
      </div>
    </div>`);
  if (numero) setTimeout(() => $('dlgNum').select(), 30);
}

/* ---------- panel ---------- */
function renderPanel() {
  const P = S.productos;
  const venc = P.filter(p => p.estado === 'vencido');
  const pronto = P.filter(p => p.estado === 'pronto').sort((a, b) => a.dias - b.dias);
  const bajo = P.filter(p => p.estado !== 'vencido' && p.stock <= S.stockBajo);
  const unidades = P.reduce((a, p) => a + p.stock, 0);
  const ahora = new Date();
  const fecha = ahora.toLocaleDateString('es-AR', {weekday: 'long', day: 'numeric', month: 'long'}).replace(',', '');
  $('heroDate').textContent = fecha.charAt(0).toUpperCase() + fecha.slice(1);
  $('heroHi').textContent = (ahora.getHours() < 13 ? 'Buen día' : ahora.getHours() < 20 ? 'Buenas tardes' : 'Buenas noches') + ', Fátima';
  $('heroTxt').innerHTML = venc.length || pronto.length
    ? `Hay <b>${plural(venc.length, 'producto vencido', 'productos vencidos')}</b> y <b>${pronto.length} por vencer</b> en los próximos ${S.diasPronto} días. Conviene revisarlos antes de abrir el mostrador.`
    : `No hay productos vencidos ni por vencer en los próximos ${S.diasPronto} días. <b>Todo en orden.</b>`;
  $('heroVentas').textContent = money(S.ventasHoy.importe);
  $('heroVentasL').innerHTML = `vendido hoy<br>${plural(S.ventasHoy.unidades, 'unidad', 'unidades')}`;

  const C = 238.76, dMin = pronto.length ? pronto[0].dias : null;
  const off = dMin === null ? C : C * (1 - Math.min(dMin / S.diasPronto, 1));
  $('metricas').innerHTML = `
    <div class="metric crit link" onclick="verLista('vencidos')" title="Ver los productos vencidos">
      <div class="lb"><span class="m-ic" style="background:rgba(181,53,42,.12);color:var(--red)"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M12 9v4m0 4h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg></span>Vencidos</div>
      <div class="vl" data-n="${venc.length}">${venc.length}</div>
      <div class="fx">${venc.length ? 'requieren retiro inmediato' : 'nada para retirar'}</div>
      <div class="ver">Ver lista →</div>
    </div>
    <div class="metric warm gauge-card link" onclick="verLista('pronto')" title="Ver los productos próximos a vencer">
      <div class="gauge">
        <svg width="86" height="86" viewBox="0 0 86 86">
          <circle class="trk" cx="43" cy="43" r="38" fill="none" stroke-width="8"/>
          <circle class="val" cx="43" cy="43" r="38" fill="none" stroke-width="8" stroke-linecap="round" stroke-dasharray="${C}" stroke-dashoffset="${off}" transform="rotate(-90 43 43)"/>
        </svg>
        <div class="gauge-n"><b>${dMin ?? '—'}</b><span>días</span></div>
      </div>
      <div class="gauge-txt">
        <div class="lb">Próximos a vencer</div>
        <div class="vl" data-n="${pronto.length}">${pronto.length}</div>
        <div class="fx">${dMin !== null ? `el más próximo, en ${plural(dMin, 'día')}` : `nada en ${S.diasPronto} días`}</div>
        <div class="ver">Ver lista →</div>
      </div>
    </div>
    <div class="metric link" onclick="verLista('bajo')" title="Ver los productos con stock bajo">
      <div class="lb"><span class="m-ic" style="background:var(--blue-bg);color:var(--blue)"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8v13H3V8M1 3h22v5H1zM10 12h4"/></svg></span>Stock bajo</div>
      <div class="vl" data-n="${bajo.length}">${bajo.length}</div>
      <div class="bar"><div class="bar-fill" style="width:${P.length ? Math.round(bajo.length / P.length * 100) : 0}%"></div></div>
      <div class="fx">${bajo.length ? `${S.stockBajo} unidades o menos · reponer` : 'stock suficiente'}</div>
      <div class="ver" style="color:var(--blue)">Ver lista →</div>
    </div>
    <div class="metric link" onclick="verLista('activos')" title="Ver todos los productos activos">
      <div class="lb"><span class="m-ic" style="background:var(--ok-bg);color:var(--primary-dk)"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg></span>Productos activos</div>
      <div class="vl" data-n="${P.length}">${P.length}</div>
      <div class="fx">de ${plural(labs().length, 'laboratorio')} · ${unidades} u. en stock</div>
      <div class="ver" style="color:var(--primary-dk)">Ver lista →</div>
    </div>`;

  const nombres = arr => arr.slice(0, 3).map(p => p.nombre).join(', ') + (arr.length > 3 ? ` y ${arr.length - 3} más` : '');
  const alerta = (bg, color, ico, titulo, txt, pill, pillTxt, lista) => `<div class="alert${lista ? ' link' : ''}"${lista ? ` onclick="verLista('${lista}')"` : ''}><span class="ic" style="background:var(${bg})"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${ico}</svg></span><div class="tx"><b>${esc(titulo)}</b><span>${esc(txt)}</span></div><span class="pill ${pill}">${pillTxt}</span></div>`;
  let al = '';
  if (venc.length) al += alerta('--red-bg', '#b5352a', '<path d="M12 9v4m0 4h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/>',
    plural(venc.length, 'producto vencido', 'productos vencidos'), `${nombres(venc)} — retirar de la venta`, 'p-red', 'Vencido', 'vencidos');
  if (pronto.length) al += alerta('--amber-bg', '#a9690c', '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    `${plural(pronto[0].stock, 'unidad', 'unidades')} de ${pronto[0].nombre} vencen en ${plural(pronto[0].dias, 'día')}`,
    'Priorizar su venta antes que el lote nuevo' + (pronto.length > 1 ? ` · y ${plural(pronto.length - 1, 'producto')} más por vencer` : ''), 'p-amber', 'Pronto', 'pronto');
  if (bajo.length) al += alerta('--blue-bg', '#185fa5', '<path d="M21 8v13H3V8M1 3h22v5H1zM10 12h4"/>',
    `Reposición sugerida: ${plural(bajo.length, 'producto')}`, `${nombres(bajo)} tienen ${S.stockBajo} unidades o menos`, 'p-blue', 'Reponer', 'bajo');
  if (!al) al = alerta('--ok-bg', '#0f6e56', '<path d="M20 6 9 17l-5-5"/>', 'Todo en orden', 'No hay vencimientos cercanos ni faltantes de stock.', 'p-ok', 'OK');
  $('alertas').innerHTML = al;

  const counts = {};
  P.forEach(p => { counts[p.categoria] = (counts[p.categoria] || 0) + 1; });
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  const max = entries.length ? entries[0][1] : 1;
  $('catBars').innerHTML = entries.map(([cat, n]) => `
    <div class="barrow"><div class="barlabel">${esc(cat)}</div>
      <div class="bartrack"><div class="barfill" style="width:${Math.round(n / max * 100)}%"></div></div>
      <div class="barval">${n}</div></div>`).join('') || '<p class="note">Sin productos.</p>';

  $('actividad').innerHTML = S.movimientos.slice(0, 7).map(m => `
    <div class="activity"><span class="dot"></span><div><b>${esc(m.tipo)}</b>
      <div class="t">${esc(textoMov(m))} · ${cuando(m.fecha)}</div></div></div>`).join('') || '<p class="note">Todavía no hay movimientos.</p>';
}

/* ---------- listas del panel ---------- */
let listaTipo = 'vencidos';
const fechaVenc = p => {  // último día del mes de vencimiento, ej. 31/10/2026
  const m = /^(\d{2})\/(\d{4})$/.exec(p.vencimiento || ''); if (!m) return '—';
  return new Date(+m[2], +m[1], 0).toLocaleDateString('es-AR', {day: '2-digit', month: '2-digit', year: 'numeric'});
};
const LISTAS = {
  vencidos: {t: 'Productos vencidos', d: () => 'Retiralos de la venta o dalos de baja del catálogo',
    f: p => p.estado === 'vencido', orden: (a, b) => a.dias - b.dias, vacio: 'No hay productos vencidos'},
  pronto: {t: 'Próximos a vencer', d: () => `Vencen en los próximos ${S.diasPronto} días · conviene priorizar su venta`,
    f: p => p.estado === 'pronto', orden: (a, b) => a.dias - b.dias, vacio: 'Ningún producto vence pronto'},
  bajo: {t: 'Stock bajo', d: () => `Productos con ${S.stockBajo} unidades o menos · conviene reponer`,
    f: p => p.estado !== 'vencido' && p.stock <= S.stockBajo, orden: (a, b) => a.stock - b.stock, vacio: 'Todos los productos tienen stock suficiente'},
  activos: {t: 'Productos activos', d: () => `Todo el catálogo · ${plural(labs().length, 'laboratorio')}`,
    f: () => true, orden: (a, b) => a.nombre.localeCompare(b.nombre), vacio: 'Todavía no hay productos cargados'},
};
function verLista(tipo) { listaTipo = tipo; renderLista(); go('lista'); }
function renderLista() {
  const L = LISTAS[listaTipo];
  titles.lista = [L.t, L.d()];
  if ($('s-lista').classList.contains('show')) { $('ttl').textContent = L.t; $('sub2').textContent = L.d(); }
  const rows = S.productos.filter(L.f).sort(L.orden);
  $('listaCant').textContent = plural(rows.length, 'producto');
  const col4 = {vencidos: 'Situación', pronto: 'Situación', bajo: 'Situación', activos: 'Estado'}[listaTipo];
  $('lhead').innerHTML = `<tr><th></th><th>Producto</th><th>Vencimiento</th><th>${col4}</th><th class="num">Stock</th><th class="num">Precio</th><th style="text-align:right">Acciones</th></tr>`;
  const b = (txt, fn, extra = '') => `<button class="btn ghost sm" ${extra} onclick="event.stopPropagation();${fn}">${txt}</button>`;
  $('lbody').innerHTML = rows.map(p => {
    const c = js(p.codigo), [cls, lbl] = vpill[p.estado];
    let sit, acc;
    if (listaTipo === 'vencidos') {
      sit = `<span style="color:var(--red);font-weight:600">Venció hace ${plural(-p.dias, 'día')}</span>`;
      acc = b('Retirar stock', `retirarDlg(${c})`, p.stock ? '' : 'disabled') + b('Dar de baja', `bajaDlg(${c})`, 'style="color:var(--red)"');
    } else if (listaTipo === 'pronto') {
      sit = `<span style="color:var(--amber);font-weight:600">Vence en ${plural(p.dias, 'día')}</span>`;
      acc = b('Ver detalle', `verDetalle(${c})`) + b('Retirar', `retirarDlg(${c})`, p.stock ? '' : 'disabled');
    } else if (listaTipo === 'bajo') {
      sit = p.stock === 0 ? `<span style="color:var(--red);font-weight:600">Sin stock</span>` : `<span style="color:var(--amber);font-weight:600">Quedan ${p.stock}</span>`;
      acc = b('+ Stock', `reponerDlg(${c})`) + b('Ver detalle', `verDetalle(${c})`);
    } else {
      sit = `<span class="pill ${cls}">${lbl}</span>`;
      acc = b('Ver detalle', `verDetalle(${c})`) + b('Editar', `editar(${c})`);
    }
    return `<tr class="click" onclick="verDetalle(${c})"><td><div class="thumb"><img src="${imgSrc(p)}" alt=""></div></td>
      <td style="font-weight:500">${esc(p.nombre)}<div class="code" style="font-family:inherit">${esc(p.laboratorio)} · ${esc(p.categoria)} · ${esc(p.codigo)}</div></td>
      <td>${fechaVenc(p)}</td><td>${sit}</td><td class="num">${p.stock}</td><td class="num">${money(p.precio)}</td>
      <td><div class="acts">${acc}</div></td></tr>`;
  }).join('') || `<tr><td colspan="7"><div class="lista-vacia"><b>${L.vacio}</b>Todo en orden por acá.</div></td></tr>`;
}

function textoMov(m) {
  if (m.tipo === 'Venta') return `${m.producto} · ${plural(+m.cantidad, 'unidad', 'unidades')} · ${money(m.importe)}`;
  if (m.tipo === 'Aumento de precio') return `${m.detalle} · ${plural(+m.cantidad, 'producto')}`;
  if (m.tipo === 'Reposición de stock') return `${m.producto} · +${m.cantidad} u.`;
  return m.producto || m.detalle;
}

function animarPanel() {
  const panel = $('s-panel');
  panel.classList.add('arranca');
  setTimeout(() => panel.classList.remove('arranca'), 30);
  panel.querySelectorAll('.vl[data-n]').forEach(el => {
    const fin = +el.dataset.n; if (!fin) return;
    const pasos = 26; let i = 0;
    clearInterval(el._t);
    el._t = setInterval(() => {
      i++;
      el.textContent = i >= pasos ? fin : Math.round(fin * (1 - Math.pow(1 - i / pasos, 3)));
      if (i >= pasos) clearInterval(el._t);
    }, 34);
  });
}

/* ---------- buscar y vender ---------- */
function score(q, p) {
  q = q.toLowerCase(); const hay = (p.nombre + ' ' + p.droga + ' ' + p.laboratorio + ' ' + p.categoria).toLowerCase();
  if (hay.includes(q)) return 0;
  let s = 0, i = 0;
  for (const ch of q) { const f = hay.indexOf(ch, i); if (f < 0) s += 2; else { s += f - i > 3 ? 1 : 0; i = f + 1; } }
  return s + (hay.includes(q.slice(0, 3)) ? 0 : 1);
}
function stockInfo(p) {
  if (p.stock === 0) return ['b-red', 'Sin stock'];
  if (p.stock <= S.stockBajo) return ['b-amber', 'Stock bajo'];
  return ['b-ok', 'En stock'];
}
function cardHtml(p) {
  const [dotCls, stockLbl] = stockInfo(p);
  const [pillCls, pillLbl] = vpill[p.estado];
  return `<button class="pcard" onclick="renderFicha(${js(p.codigo)})">
    <div class="ph"><span class="pill ${pillCls}">${pillLbl}</span><img src="${imgSrc(p)}" alt="${esc(p.nombre)}"></div>
    <div class="pbody">
      <div class="pn">${esc(p.nombre)}</div>
      <div class="pl">${esc(p.laboratorio)} · ${esc(p.categoria)}</div>
      <div class="pcard-foot"><span class="pp">${money(p.precio)}</span><span class="stockdot"><span class="dot ${dotCls}"></span>${stockLbl} · ${p.stock}</span></div>
    </div>
  </button>`;
}
function renderGrid(q = '') {
  const query = q.trim();
  const list = query
    ? S.productos.filter(p => score(query, p) <= query.length).sort((a, b) => score(query, a) - score(query, b))
    : S.productos;
  $('grid').innerHTML = list.length ? list.map(cardHtml).join('')
    : `<div class="empty"><b>No encontramos productos con ese filtro</b>Probá con el nombre del producto, la droga o el laboratorio.</div>`;
}
$('q').addEventListener('input', e => renderGrid(e.target.value));

function renderFicha(code) {
  window._sel = code; const p = prod(code); if (!p) return;
  const vender = document.querySelector('#modes button.on')?.dataset.m === 'vender';
  const stockTxt = p.stock === 0 ? `<span style="color:var(--red);font-weight:600">Sin stock</span>` : plural(p.stock, 'unidad', 'unidades');
  const verBtn = `<button class="btn ghost" onclick="verDetalle(${js(p.codigo)})">Ver detalle</button>`;
  let accion = '';
  if (p.estado === 'vencido') accion = `<button class="btn ghost" disabled>Vencido — no se vende</button>`;
  else if (p.stock === 0) accion = `<button class="btn ghost" disabled>Sin stock</button>`;
  else if (vender) accion = `<div class="venta-row"><label for="qtyVenta">Cantidad</label>
      <input type="number" class="qty" id="qtyVenta" value="1" min="1" max="${p.stock}" oninput="totVenta(${p.precio})" onkeydown="if(event.key==='Enter')registrarVenta(${js(p.codigo)})">
      <button class="btn" onclick="registrarVenta(${js(p.codigo)})">Descontar venta</button>
      <span class="venta-tot" id="ventaTot">Total: <b>${money(p.precio)}</b></span></div>`;
  else accion = `<button class="btn" onclick="modoVender()">Vender</button>`;
  $('ficha').innerHTML = `
    <div class="ficha">
      <div class="ph"><img src="${imgSrc(p)}" alt="${esc(p.nombre)}"></div>
      <div>
        <h3>${esc(p.nombre)}</h3><div class="lab">${esc(p.laboratorio)} · código ${esc(p.codigo)}</div>
        <div class="frow">
          <div><div class="k">Precio</div><div class="price">${money(p.precio)}</div></div>
          <div><div class="k">Stock</div><div class="v">${stockTxt}</div></div>
          <div><div class="k">Droga</div><div class="v">${esc(p.droga)}</div></div>
          <div><div class="k">Para</div><div class="v">${esc(p.especie)}</div></div>
          <div><div class="k">Vence</div><div class="v">${esc(p.vencimiento || '—')} <span class="pill ${vpill[p.estado][0]}">${vpill[p.estado][1]}</span></div></div>
        </div>
        <div class="tags"><span class="tag">${esc(p.categoria)}</span><span class="tag">${esc(p.presentacion)}</span></div>
        <div style="margin-top:20px;display:flex;gap:10px;align-items:center;flex-wrap:wrap">${accion}${verBtn}</div>
      </div>
    </div>`;
}
function modoVender() { document.querySelector('#modes button[data-m="vender"]').click(); setTimeout(() => $('qtyVenta')?.select(), 20); }
function totVenta(precio) { $('ventaTot').innerHTML = `Total: <b>${money(precio * (parseInt($('qtyVenta').value) || 0))}</b>`; }

function registrarVenta(code) {
  const p = prod(code), cant = parseInt($('qtyVenta')?.value) || 1;
  if (!p) return;
  if (cant < 1 || cant > p.stock) { mostrarError(`No alcanza el stock: quedan ${plural(p.stock, 'unidad', 'unidades')}.`); return; }
  cambioStockAlInstante(p, -cant, 'Venta', p.precio * cant);
  confirmar({titulo: 'Venta registrada', texto: textoGuardado(),
    detalle: `${p.nombre} · ${plural(cant, 'unidad vendida', 'unidades vendidas')} · ${money(p.precio * cant)}`,
    valor: `Quedan ${plural(p.stock, 'unidad', 'unidades')}`});
  enSegundoPlano('vender', code, cant, usuario);
}

/* ---------- cambios al instante ----------
   Ventas, reposiciones y retiros se ven en pantalla en el momento y se guardan en segundo plano.
   Si la base los rechaza (p. ej. otra computadora vendió la última unidad) se avisa y se recargan los datos. */
let pendientes = 0;
const textoGuardado = () => enNube() ? 'Se está guardando en Google Sheets en segundo plano.' : 'El stock quedó actualizado en el Excel.';

function cambioStockAlInstante(p, delta, tipo, importe = 0) {
  p.stock += delta;
  S.movimientos.unshift({fecha: new Date().toISOString(), tipo, codigo: p.codigo, producto: p.nombre,
    cantidad: Math.abs(delta), importe, detalle: '', usuario});
  if (tipo === 'Venta') { S.ventasHoy.importe += importe; S.ventasHoy.unidades += -delta; }
  renderTodo();
}

/* Indicador abajo a la derecha. txt=null lo oculta; listo=true lo muestra en verde y se va solo. */
function avisoSync(txt, listo = false) {
  if (!txt) { $('sync').hidden = true; return; }
  $('syncTxt').textContent = txt; $('sync').className = listo ? 'sync ok' : 'sync'; $('sync').hidden = false;
  if (listo) setTimeout(() => { if (!pendientes && $('syncTxt').textContent === txt) $('sync').hidden = true; }, 1500);
}

/* Las operaciones se mandan de a una y en orden (una fila de espera), sin frenar la pantalla. */
let cola = Promise.resolve();
function enSegundoPlano(fn, ...args) {
  pendientes++;
  avisoSync(enNube() ? `Guardando en Google Sheets…${pendientes > 1 ? ` (${pendientes})` : ''}` : 'Guardando…');
  cola = cola.then(async () => {
    let r;
    try { r = await api()[fn](...args); } catch (e) { r = {ok: false, error: String(e)}; }
    pendientes--;
    if (!r.ok) {
      mostrarError(r.error + ' Ese cambio no se guardó: se volvieron a cargar los datos.');
      avisoSync(null);
      await cargar(true);
      return;
    }
    if (pendientes === 0) {
      await cargar(true);  // trae lo que devolvió la base (incluye cambios de otras computadoras)
      avisoSync('Guardado', true);
    } else {
      avisoSync(`Guardando en Google Sheets… (${pendientes})`);
    }
  });
}

/* Estado de vencimiento calculado en pantalla (igual que en db.py: vence el último día del mes). */
function estadoVenc(venc) {
  const m = /^(\d{1,2})\/(\d{4})$/.exec(venc || ''); if (!m) return ['sin_dato', null];
  const fin = new Date(+m[2], +m[1], 0), hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const dias = Math.round((fin - hoy) / 864e5);
  return [dias < 0 ? 'vencido' : dias <= S.diasPronto ? 'pronto' : 'ok', dias];
}

/* Código nuevo con el prefijo del laboratorio (o uno nuevo), sin repetir ninguno existente. */
function nuevoCodigo(lab) {
  const usados = new Set(S.productos.map(p => p.codigo));
  const delLab = S.productos.filter(p => p.laboratorio === lab && /^\d{3}-/.test(p.codigo)).map(p => p.codigo.slice(0, 3)).sort();
  const nums = S.productos.filter(p => /^\d{3}-/.test(p.codigo)).map(p => +p.codigo.slice(0, 3));
  const prefijo = delLab[0] || String((nums.length ? Math.max(...nums) : 0) + 1).padStart(3, '0');
  let c;
  do { c = `${prefijo}-${1000 + Math.floor(Math.random() * 9000)}`; } while (usados.has(c));
  return c;
}

function movAlInstante(tipo, codigo, producto, cantidad = 0, detalle = '') {
  S.movimientos.unshift({fecha: new Date().toISOString(), tipo, codigo, producto, cantidad, importe: 0, detalle, usuario});
}

/* sugerencias del buscador */
const SUG_ICONOS = {
  Producto: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/>',
  Droga: '<path d="M10.5 20.5 20.5 10.5a4.95 4.95 0 1 0-7-7l-10 10a4.95 4.95 0 1 0 7 7Z"/><path d="m8.5 8.5 7 7"/>',
  Laboratorio: '<path d="M9 3h6M10 3v6.6a2 2 0 0 1-.3 1L4.6 19a2 2 0 0 0 1.7 3h11.4a2 2 0 0 0 1.7-3l-5.1-8.4a2 2 0 0 1-.3-1V3"/>'
};
let sugActual = [], sugSel = -1;
const resaltar = (txt, q) => {
  const i = txt.toLowerCase().indexOf(q.toLowerCase());
  if (i < 0) return esc(txt);
  return esc(txt.slice(0, i)) + '<b>' + esc(txt.slice(i, i + q.length)) + '</b>' + esc(txt.slice(i + q.length));
};
function armarSugerencias(q) {
  q = q.trim(); if (!q) return [];
  const ql = q.toLowerCase(), out = [], vistos = new Set(), P = S.productos;
  const add = (tipo, txt, meta, accion) => { const k = tipo + '|' + txt; if (vistos.has(k)) return; vistos.add(k); out.push({tipo, txt, meta, accion}); };
  P.forEach(p => { if (p.nombre.toLowerCase().includes(ql)) add('Producto', p.nombre, `${p.laboratorio} · ${money(p.precio)}`, {t: 'prod', c: p.codigo}); });
  [...new Set(P.map(p => p.droga))].forEach(d => {
    if (!d || d === '—' || !d.toLowerCase().includes(ql)) return;
    const n = P.filter(p => p.droga === d).length;
    add('Droga', d, `${plural(n, 'producto')} con esta droga`, {t: 'filtro', v: d});
  });
  labs().forEach(l => {
    if (!l.toLowerCase().includes(ql)) return;
    add('Laboratorio', l, `${plural(P.filter(p => p.laboratorio === l).length, 'producto')} del laboratorio`, {t: 'filtro', v: l});
  });
  P.forEach(p => {
    if (p.nombre.toLowerCase().includes(ql)) return;
    if ((p.droga + p.laboratorio + p.categoria).toLowerCase().includes(ql)) add('Producto', p.nombre, `${p.droga} · ${p.laboratorio}`, {t: 'prod', c: p.codigo});
  });
  return out.slice(0, 6);
}
function pintarSugerencias(q) {
  const caja = $('sug');
  sugActual = armarSugerencias(q); sugSel = -1;
  if (!sugActual.length) { caja.hidden = true; caja.innerHTML = ''; return; }
  caja.innerHTML = sugActual.map((s, i) => `
    <button class="sug-item" onclick="elegirSugerencia(${i})">
      <span class="sug-ic"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${SUG_ICONOS[s.tipo]}</svg></span>
      <span class="sug-tx"><span class="sug-n">${resaltar(s.txt, q.trim())}</span><span class="sug-meta">${esc(s.meta)}</span></span>
      <span class="sug-tag">${s.tipo}</span>
    </button>`).join('');
  caja.hidden = false;
}
function marcarSugerencia() { document.querySelectorAll('#sug .sug-item').forEach((el, i) => el.classList.toggle('on', i === sugSel)); }
function elegirSugerencia(i) {
  const s = sugActual[i]; if (!s) return;
  cerrarSugerencias();
  if (s.accion.t === 'prod') {
    const p = prod(s.accion.c);
    $('q').value = p.nombre; renderGrid(p.nombre); renderFicha(p.codigo);
    $('ficha').scrollIntoView({behavior: 'smooth', block: 'nearest'});
  } else { $('q').value = s.accion.v; renderGrid(s.accion.v); }
}
function cerrarSugerencias() { $('sug').hidden = true; sugSel = -1; }
$('q').addEventListener('input', e => pintarSugerencias(e.target.value));
$('q').addEventListener('focus', e => { if (e.target.value.trim()) pintarSugerencias(e.target.value); });
$('q').addEventListener('keydown', e => {
  if ($('sug').hidden || !sugActual.length) { if (e.key === 'Escape') $('q').blur(); return; }
  if (e.key === 'ArrowDown') { e.preventDefault(); sugSel = (sugSel + 1) % sugActual.length; marcarSugerencia(); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); sugSel = (sugSel - 1 + sugActual.length) % sugActual.length; marcarSugerencia(); }
  else if (e.key === 'Enter') { if (sugSel >= 0) { e.preventDefault(); elegirSugerencia(sugSel); } else cerrarSugerencias(); }
  else if (e.key === 'Escape') cerrarSugerencias();
});
document.addEventListener('click', e => { if (!e.target.closest('.search')) cerrarSugerencias(); });

/* ---------- ficha ampliada ---------- */
const fila = (k, v) => `<div class="spec"><div class="spec-k">${k}</div><div class="spec-v">${v}</div></div>`;
function verDetalle(code) {
  const p = prod(code); if (!p) return;
  const [pillCls, pillLbl] = vpill[p.estado];
  const t = v => esc(v || '—');
  const stockTxt = p.stock === 0 ? `<span style="color:var(--red);font-weight:650">Sin stock</span>` : plural(p.stock, 'unidad', 'unidades');
  abrirModal(`
    <div class="modal-box" role="dialog" aria-modal="true">
      <button class="modal-x" onclick="cerrarModal()" aria-label="Cerrar"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg></button>
      <div class="modal-head">
        <div class="ph"><img src="${imgSrc(p)}" alt="${esc(p.nombre)}"></div>
        <div class="modal-ttl">
          <span class="pill ${pillCls}">${pillLbl}</span>
          <h3>${esc(p.nombre)}</h3>
          <div class="lab">${esc(p.laboratorio)} · código ${esc(p.codigo)}</div>
          <div class="price">${money(p.precio)}</div>
        </div>
      </div>
      <div class="modal-body">
        <div class="spec-grid">
          ${fila('Droga principal', t(p.droga))}${fila('Concentración', t(p.concentracion))}
          ${fila('Presentación', t(p.presentacion))}${fila('Vía de administración', t(p.via))}
          ${fila('Dosis sugerida', t(p.dosis))}${fila('Especie', t(p.especie))}
          ${fila('Categoría', t(p.categoria))}${fila('Stock actual', stockTxt)}
          ${fila('Lote', t(p.lote))}${fila('Vencimiento', t(p.vencimiento))}
        </div>
        <div class="spec-full"><div class="spec-k">Composición</div><div class="spec-v">${t(p.composicion)}</div></div>
        <div class="spec-full"><div class="spec-k">Indicaciones</div><div class="spec-v">${t(p.indicaciones)}</div></div>
        <div class="spec-full"><div class="spec-k">Conservación</div><div class="spec-v">${t(p.conservacion)}</div></div>
      </div>
      <div class="modal-foot">
        ${role === 'vet' ? `<button class="btn ghost" onclick="cerrarModal();editar(${js(p.codigo)})">Editar producto</button>` : ''}
        <button class="btn" onclick="cerrarModal()">Cerrar</button>
      </div>
    </div>`);
}

/* ---------- tabla ---------- */
const statIc = {
  productos: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/>',
  labs: '<path d="M9 3h6M10 3v6.6a2 2 0 0 1-.3 1L4.6 19a2 2 0 0 0 1.7 3h11.4a2 2 0 0 0 1.7-3l-5.1-8.4a2 2 0 0 1-.3-1V3"/>',
  plata: '<path d="M12 2v20M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
  lista: '<path d="M4 6h16M4 12h16M4 18h10"/>'
};
const statCard = (ic, n, l) => `<div class="stat">
  <span class="stat-ic"><svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${statIc[ic]}</svg></span>
  <div><div class="stat-n">${n}</div><div class="stat-l">${l}</div></div></div>`;

function renderTable(f = '') {
  f = f.toLowerCase();
  const P = S.productos;
  $('statStrip').innerHTML = statCard('productos', P.length, 'Productos en catálogo') +
    statCard('labs', labs().length, 'Laboratorios') +
    statCard('plata', money(P.reduce((a, p) => a + p.precio * p.stock, 0)), `Valor del stock · ${P.reduce((a, p) => a + p.stock, 0)} unidades`);
  const rows = P.filter(p => (p.nombre + p.laboratorio + p.categoria + p.droga + p.codigo).toLowerCase().includes(f));
  $('tbody').innerHTML = rows.map(p => {
    const [cls, lbl] = vpill[p.estado];
    const st = p.stock === 0 ? `<span style="color:var(--red);font-weight:600">0</span>` : p.stock <= S.stockBajo ? `<span style="color:var(--amber);font-weight:600">${p.stock}</span>` : p.stock;
    return `<tr class="click" onclick="verDetalle(${js(p.codigo)})"><td><div class="thumb"><img src="${imgSrc(p)}" alt=""></div></td><td class="code">${esc(p.codigo)}</td><td style="font-weight:500">${esc(p.nombre)}</td><td>${esc(p.laboratorio)}</td><td>${esc(p.categoria)}</td><td class="num">${money(p.precio)}</td><td class="num">${st}</td><td><span class="pill ${cls}">${lbl}</span> <span class="code">${esc(p.vencimiento)}</span></td></tr>`;
  }).join('') || `<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:30px">Sin resultados</td></tr>`;
}
$('tq').addEventListener('input', e => renderTable(e.target.value));

/* ---------- precios ---------- */
const nuevoPrecio = (p, pct) => Math.round(p * (1 + pct / 100));
const pctAumento = () => parseFloat($('pp').value.replace('%', '').replace(',', '.')) || 0;  // acepta "5,5"
function llenarSelect(el, opciones, primera) {
  const actual = el.value;
  el.innerHTML = (primera ? `<option value="">${primera}</option>` : '') + opciones.map(o => `<option>${esc(o)}</option>`).join('');
  if ([...el.options].some(o => o.value === actual)) el.value = actual;
}
function llenarSelectsPrecios() { llenarSelect($('pl'), labs()); llenarSelect($('pc'), cats($('pl').value), 'Todas'); }
function afectados() { const l = $('pl').value, c = $('pc').value; return S.productos.filter(p => p.laboratorio === l && (!c || p.categoria === c)); }
function preview() {
  const l = $('pl').value, c = $('pc').value, pct = pctAumento();
  const af = afectados();
  if (!af.length) { $('prev').innerHTML = `<div class="big">Sin productos para ese filtro</div><div class="sm">Probá con otra categoría o elegí “Todas”.</div>`; return; }
  const total = af.reduce((a, p) => a + p.precio, 0), totalNuevo = af.reduce((a, p) => a + nuevoPrecio(p.precio, pct), 0);
  $('prev').innerHTML = `
    <div class="big">${plural(af.length, 'producto')} de ${esc(l)}${c ? ` · ${esc(c)}` : ''}</div>
    <div class="prev-rows">${af.map(p => `<div class="prev-row"><span class="prev-n">${esc(p.nombre)}</span><span class="prev-a">${money(p.precio)}</span>
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
      <span class="prev-b">${money(nuevoPrecio(p.precio, pct))}</span></div>`).join('')}</div>
    <div class="prev-tot"><div><span>Precio actual</span><b>${money(total)}</b></div><div><span>Aumento</span><b>${fmtPct(pct)} %</b></div><div><span>Nuevo precio</span><b class="prev-hi">${money(totalNuevo)}</b></div></div>`;
}
$('pl').addEventListener('change', () => { llenarSelect($('pc'), cats($('pl').value), 'Todas'); preview(); });
['pc', 'pp'].forEach(id => $(id).addEventListener('input', preview));

async function aplicarAumento() {
  const l = $('pl').value, c = $('pc').value, pct = pctAumento();
  const af = afectados();
  if (!af.length || !pct) { confirmar({titulo: 'No hay cambios para aplicar', texto: 'Elegí un laboratorio con productos y un porcentaje distinto de cero.', boton: 'Volver', tipo: 'warn'}); return; }
  af.forEach(p => { p.precio = nuevoPrecio(p.precio, pct); });
  movAlInstante('Aumento de precio', '', l, af.length, `${l} ${pct > 0 ? '+' : ''}${fmtPct(pct)} %${c ? ` · ${c}` : ''}`);
  renderTodo();
  confirmar({titulo: 'Precio actualizado', texto: enNube() ? 'Se está guardando en Google Sheets en segundo plano.' : 'El aumento quedó guardado en el Excel.',
    detalle: `${l}${c ? ` — ${c}` : ''} · ${plural(af.length, 'producto')} · ${pct > 0 ? '+' : ''}${fmtPct(pct)} %`,
    valor: money(af.reduce((a, p) => a + p.precio, 0))});
  enSegundoPlano('aumentar', l, c, pct, usuario);
}
function resetPrecios() { $('pl').selectedIndex = 0; llenarSelect($('pc'), cats($('pl').value), 'Todas'); $('pc').value = ''; $('pp').value = 5; preview(); }
function renderHist() {
  const h = S.movimientos.filter(m => m.tipo === 'Aumento de precio').slice(0, 6);
  $('histAumentos').innerHTML = h.map(m => `<div class="activity"><span class="dot"></span><div><b>${esc(m.detalle)}</b><div class="t">${plural(+m.cantidad, 'producto')} · ${cuando(m.fecha)}</div></div></div>`).join('')
    || '<p class="note">Todavía no se aplicaron aumentos.</p>';
}

/* importación CSV */
let csvTexto = '';
function leerCsv(file) {
  if (!file) return;
  const fr = new FileReader();
  fr.onload = async e => {
    csvTexto = e.target.result; $('csvInput').value = '';
    const r = await llamar('importar_csv', csvTexto, false, usuario); if (!r) return;
    abrirModal(`
      <div class="modal-box modal-ok" style="max-width:520px" role="dialog" aria-modal="true">
        <h3>Vista previa del archivo</h3>
        <p>Revisá los aumentos antes de aplicarlos.</p>
        <table class="csv-tbl"><thead><tr><th>Laboratorio</th><th>Categoría</th><th class="num">%</th><th class="num">Productos</th></tr></thead>
          <tbody>${r.reglas.map(x => `<tr><td>${esc(x.laboratorio)}</td><td>${esc(x.categoria || 'Todas')}</td><td class="num">${x.pct}</td><td class="num">${x.productos}</td></tr>`).join('')}</tbody></table>
        ${r.errores.length ? `<p class="csv-err">Se van a ignorar: ${r.errores.map(esc).join(' · ')}</p>` : ''}
        <div class="dlg-btns"><button class="btn ghost" onclick="cerrarModal()">Cancelar</button><button class="btn" onclick="aplicarCsv()">Aplicar aumentos</button></div>
      </div>`);
  };
  fr.readAsText(file, 'utf-8');
}
async function aplicarCsv() {
  cerrarModal();
  const r = await llamar('importar_csv', csvTexto, true, usuario); if (!r) return;
  await cargar();
  const n = r.resumen.reduce((a, x) => a + x.productos, 0);
  confirmar({titulo: 'Archivo aplicado', texto: `Los nuevos precios quedaron guardados en ${donde()}.`,
    detalle: r.resumen.map(x => `${x.laboratorio} ${x.pct > 0 ? '+' : ''}${x.pct} %`).join(' · '), valor: plural(n, 'producto actualizado', 'productos actualizados')});
}
function zonaArchivo(zona, input, alSoltar) {
  zona.addEventListener('click', () => input.click());
  input.addEventListener('change', e => alSoltar(e.target.files[0]));
  ['dragenter', 'dragover'].forEach(ev => zona.addEventListener(ev, e => { e.preventDefault(); zona.classList.add('drop-on'); }));
  ['dragleave', 'drop'].forEach(ev => zona.addEventListener(ev, e => { e.preventDefault(); zona.classList.remove('drop-on'); }));
  zona.addEventListener('drop', e => alSoltar(e.dataTransfer.files[0]));
}
zonaArchivo($('csvZona'), $('csvInput'), leerCsv);

/* ---------- administración ---------- */
let editando = null;     // código del producto que se está editando (null = alta)
let fotoNueva = null;    // null = sin cambios · '' = sin foto · 'data:...' = foto nueva
const OTRO = '__otro';

function llenarSelectsForm() {
  [['fLab', labs(), 'fLabOtro', 'Otro laboratorio…'], ['fCat', cats(), 'fCatOtro', 'Otra categoría…']].forEach(([id, ops, otro, txt]) => {
    const el = $(id), actual = el.value;
    el.innerHTML = ops.map(o => `<option>${esc(o)}</option>`).join('') + `<option value="${OTRO}">${txt}</option>`;
    if ([...el.options].some(o => o.value === actual)) el.value = actual;
    $(otro).hidden = el.value !== OTRO;
  });
}
['fLab', 'fCat'].forEach(id => $(id).addEventListener('change', () => {
  $(id + 'Otro').hidden = $(id).value !== OTRO;
  if ($(id).value === OTRO) $(id + 'Otro').focus();
}));
function elegir(sel, valor) {
  if (![...sel.options].some(o => o.value === valor)) sel.insertAdjacentHTML('afterbegin', `<option>${esc(valor)}</option>`);
  sel.value = valor;
}
const CAMPOS = {fNombre: 'nombre', fDroga: 'droga', fPrecio: 'precio', fStock: 'stock', fConc: 'concentracion', fVenc: 'vencimiento',
  fCod: 'codigo', fVia: 'via', fDosis: 'dosis', fLote: 'lote', fComp: 'composicion', fInd: 'indicaciones', fCons: 'conservacion'};

function editar(code) {
  const p = prod(code); if (!p) return;
  go('admin');
  editando = code;
  Object.entries(CAMPOS).forEach(([id, k]) => { $(id).value = (k === 'droga' && p[k] === '—') ? '' : p[k]; });
  elegir($('fLab'), p.laboratorio); elegir($('fCat'), p.categoria);
  $('fLabOtro').hidden = $('fCatOtro').hidden = true;
  if (p.presentacion) elegir($('fPres'), p.presentacion);
  if (p.especie) elegir($('fEsp'), p.especie);
  fotoNueva = null;
  if (p.imagen && imgCache[p.imagen]) mostrarFoto(imgCache[p.imagen], false); else quitarFoto(false);
  $('formT').textContent = `Editar producto · ${p.nombre}`;
  $('btnGuardar').textContent = 'Guardar cambios';
  $('btnLimpiar').textContent = 'Cancelar edición';
  $('fAviso').hidden = true;
  document.querySelector('.main').scrollTo?.({top: 0, behavior: 'smooth'});
  window.scrollTo({top: 0, behavior: 'smooth'});
}

async function guardarProducto() {
  const d = {};
  Object.entries(CAMPOS).forEach(([id, k]) => { d[k] = $(id).value.trim(); });
  d.laboratorio = $('fLab').value === OTRO ? $('fLabOtro').value.trim() : $('fLab').value;
  d.categoria = $('fCat').value === OTRO ? $('fCatOtro').value.trim() : $('fCat').value;
  d.presentacion = $('fPres').value; d.especie = $('fEsp').value;
  // las mismas validaciones que hace api.py, para avisar en el momento y no después
  const aviso = t => { $('fAviso').textContent = t; $('fAviso').hidden = false; };
  const precio = Math.round(parseFloat(d.precio.replace(',', '.'))), stock = d.stock === '' ? 0 : Math.round(parseFloat(d.stock.replace(',', '.')));
  if (!d.nombre || !d.precio) return aviso('Completá al menos el nombre y el precio de venta.');
  if (!d.laboratorio || !d.categoria) return aviso('Elegí o escribí el laboratorio y la categoría.');
  if (!(precio >= 1)) return aviso('Precio de venta: tiene que ser un número mayor a cero.');
  if (!(stock >= 0)) return aviso('Stock: tiene que ser 0 o más.');
  const mv = /^\s*(\d{1,2})\s*\/\s*(\d{4})\s*$/.exec(d.vencimiento);
  if (d.vencimiento && (!mv || +mv[1] < 1 || +mv[1] > 12)) return aviso('El vencimiento tiene que tener el formato MM/AAAA (por ejemplo 05/2027).');
  if (mv) d.vencimiento = `${mv[1].padStart(2, '0')}/${mv[2]}`;
  if (d.codigo && !/^[\w-]{3,20}$/.test(d.codigo)) return aviso('El código solo puede tener letras, números y guiones.');
  if (!d.codigo) d.codigo = nuevoCodigo(d.laboratorio);
  if (S.productos.some(p => p.codigo === d.codigo && p.codigo !== editando)) return aviso(`Ya existe un producto con el código ${d.codigo}.`);
  $('fAviso').hidden = true;

  // se ve al instante; se guarda en segundo plano
  const eraEdicion = !!editando, original = editando;
  const foto = eraEdicion ? fotoNueva : (fotoNueva || '');
  const anterior = eraEdicion ? prod(original) : null;
  const nuevo = {...(anterior || {}), ...d, precio, stock, droga: d.droga || '—'};
  if (foto) { nuevo.imagen = 'local:' + d.codigo; imgCache[nuevo.imagen] = foto; }
  else if (foto === '') nuevo.imagen = '';
  [nuevo.estado, nuevo.dias] = estadoVenc(nuevo.vencimiento);
  if (eraEdicion) S.productos = S.productos.map(p => p.codigo === original ? nuevo : p);
  else S.productos.push(nuevo);
  movAlInstante(eraEdicion ? 'Edición de producto' : 'Alta de producto', nuevo.codigo, nuevo.nombre, eraEdicion ? 0 : stock,
    eraEdicion ? 'Datos de la ficha' : `${nuevo.laboratorio} · ${nuevo.categoria}`);
  limpiarForm();
  renderTodo();
  confirmar({titulo: eraEdicion ? 'Cambios guardados' : 'Producto guardado',
    texto: enNube() ? 'Se está guardando en Google Sheets en segundo plano.' : (eraEdicion ? 'La ficha quedó actualizada en el Excel.' : 'El producto quedó guardado en el Excel.'),
    detalle: `${nuevo.nombre} · código ${nuevo.codigo}`, valor: money(precio)});
  enSegundoPlano('guardar_producto', d, foto, original, usuario);
}

function limpiarForm() {
  editando = null;
  Object.keys(CAMPOS).forEach(id => { $(id).value = ''; });
  $('fLabOtro').value = $('fCatOtro').value = '';
  $('fLab').selectedIndex = 0; $('fCat').selectedIndex = 0; $('fPres').selectedIndex = 0; $('fEsp').selectedIndex = 0;
  $('fLabOtro').hidden = $('fCatOtro').hidden = true;
  $('formT').textContent = 'Alta de producto';
  $('btnGuardar').textContent = 'Guardar producto';
  $('btnLimpiar').textContent = 'Limpiar formulario';
  $('fAviso').hidden = true;
  quitarFoto(false); fotoNueva = null;
}

/* foto del producto */
function mostrarFoto(dataUrl, esNueva = true) {
  if (esNueva) fotoNueva = dataUrl;
  $('fotoZona').hidden = true; $('fotoPrev').hidden = false; $('fotoImg').src = dataUrl;
}
function quitarFoto(marcar = true) {
  if (marcar) fotoNueva = '';
  $('fotoPrev').hidden = true; $('fotoZona').hidden = false; $('fotoInput').value = '';
}
function leerFoto(file) {
  if (!file || !file.type.startsWith('image/')) return;
  if (file.size > 5 * 1024 * 1024) { mostrarError('La imagen pesa más de 5 MB. Elegí una más liviana.'); return; }
  const fr = new FileReader();
  fr.onload = e => mostrarFoto(e.target.result);
  fr.readAsDataURL(file);
}
zonaArchivo($('fotoZona'), $('fotoInput'), leerFoto);

/* lista de productos con acciones */
function renderAdminLista() {
  const f = $('aq').value.toLowerCase();
  const rows = S.productos.filter(p => (p.nombre + p.codigo + p.laboratorio + p.droga).toLowerCase().includes(f));
  $('abody').innerHTML = rows.map(p => {
    const [cls, lbl] = vpill[p.estado];
    return `<tr><td><div class="thumb"><img src="${imgSrc(p)}" alt=""></div></td><td class="code">${esc(p.codigo)}</td>
      <td style="font-weight:500">${esc(p.nombre)}<div class="code" style="font-family:inherit">${esc(p.laboratorio)} · ${esc(p.categoria)}</div></td>
      <td class="num">${money(p.precio)}</td><td class="num">${p.stock}</td><td><span class="pill ${cls}">${lbl}</span></td>
      <td><div class="acts">
        <button class="btn ghost sm" onclick="editar(${js(p.codigo)})">Editar</button>
        <button class="btn ghost sm" onclick="reponerDlg(${js(p.codigo)})">+ Stock</button>
        <button class="btn ghost sm" onclick="retirarDlg(${js(p.codigo)})" ${p.stock ? '' : 'disabled'}>Retirar</button>
        <button class="btn ghost sm" style="color:var(--red)" onclick="bajaDlg(${js(p.codigo)})">Baja</button>
      </div></td></tr>`;
  }).join('') || `<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:30px">Sin resultados</td></tr>`;
}
$('aq').addEventListener('input', renderAdminLista);

function reponerDlg(code) {
  const p = prod(code);
  dialogo({titulo: 'Reponer stock', texto: `¿Cuántas unidades de <b>${esc(p.nombre)}</b> ingresaron? Hoy hay ${p.stock}.`,
    numero: {valor: 10}, boton: 'Sumar al stock', alAceptar: async n => {
      n = parseInt(n);
      if (!(n >= 1)) { mostrarError('La cantidad tiene que ser 1 o más.'); return; }
      cambioStockAlInstante(p, n, 'Reposición de stock');
      confirmar({titulo: 'Stock repuesto', texto: textoGuardado(), detalle: p.nombre, valor: `Stock actual: ${p.stock}`});
      enSegundoPlano('reponer', code, n, usuario);
    }});
}
function retirarDlg(code) {
  const p = prod(code);
  dialogo({titulo: 'Retirar de la venta', texto: `Unidades de <b>${esc(p.nombre)}</b> a retirar${p.estado === 'vencido' ? ' por vencimiento' : ''}. Hay ${p.stock}.`,
    numero: {valor: p.stock, max: p.stock}, boton: 'Retirar', peligro: true, alAceptar: async n => {
      n = parseInt(n);
      if (!(n >= 1) || n > p.stock) { mostrarError(`Podés retirar entre 1 y ${p.stock} unidades.`); return; }
      cambioStockAlInstante(p, -n, p.estado === 'vencido' ? 'Baja por vencimiento' : 'Retiro de stock');
      confirmar({titulo: 'Unidades retiradas', texto: textoGuardado(), detalle: p.nombre, valor: `Stock actual: ${p.stock}`});
      enSegundoPlano('retirar', code, n, p.estado, usuario);
    }});
}
function bajaDlg(code) {
  const p = prod(code);
  dialogo({titulo: 'Dar de baja el producto', texto: `<b>${esc(p.nombre)}</b> se va a eliminar del catálogo. El historial de movimientos se conserva.`,
    boton: 'Dar de baja', peligro: true, alAceptar: async () => {
      if (editando === code) limpiarForm();
      S.productos = S.productos.filter(x => x.codigo !== code);
      if (window._sel === code) window._sel = null;
      movAlInstante('Baja de producto', code, p.nombre, p.stock, 'Eliminado del catálogo');
      renderTodo();
      confirmar({titulo: 'Producto dado de baja', texto: 'Ya no aparece en el catálogo.', detalle: p.nombre});
      enSegundoPlano('baja', code, usuario);
    }});
}

/* ---------- movimientos ---------- */
let todosMovs = [];
const tipoPill = {'Venta': 'p-ok', 'Aumento de precio': 'p-blue', 'Reposición de stock': 'p-ok', 'Alta de producto': 'p-blue',
  'Edición de producto': 'p-amber', 'Baja de producto': 'p-red', 'Baja por vencimiento': 'p-red', 'Retiro de stock': 'p-amber'};
async function cargarMovs() {
  // primero lo que ya está en pantalla; el historial completo llega de fondo, sin cartel de espera
  if (!todosMovs.length || pendientes) { todosMovs = S.movimientos; renderMovs(); }
  if (pendientes) return;
  const r = await api().movimientos(); if (!r.ok) return;
  todosMovs = r.movimientos;
  llenarSelect($('mTipo'), [...new Set(todosMovs.map(m => m.tipo))].sort(), 'Todos los movimientos');
  renderMovs();
}
function renderMovs() {
  const t = $('mTipo').value, q = $('mq').value.toLowerCase();
  const rows = todosMovs.filter(m => (!t || m.tipo === t) && (m.producto + ' ' + m.detalle + ' ' + m.codigo).toLowerCase().includes(q));
  const ventas = todosMovs.filter(m => m.tipo === 'Venta');
  $('movStrip').innerHTML = statCard('lista', todosMovs.length, 'Movimientos registrados') +
    statCard('plata', money(S.ventasHoy.importe), `Ventas de hoy · ${plural(S.ventasHoy.unidades, 'unidad', 'unidades')}`) +
    statCard('plata', money(ventas.reduce((a, m) => a + (+m.importe || 0), 0)), `Total vendido · ${plural(ventas.length, 'venta')}`);
  $('mbody').innerHTML = rows.map(m => {
    const f = new Date(m.fecha);
    const fecha = isNaN(f) ? esc(m.fecha) : f.toLocaleDateString('es-AR') + ' ' + f.toLocaleTimeString('es-AR', {hour: '2-digit', minute: '2-digit'});
    return `<tr><td class="code" style="white-space:nowrap">${fecha}</td><td><span class="pill ${tipoPill[m.tipo] || 'p-blue'}">${esc(m.tipo)}</span></td>
      <td style="font-weight:500">${esc(m.producto)}${m.codigo ? `<div class="code">${esc(m.codigo)}</div>` : ''}</td>
      <td class="num">${m.cantidad || ''}</td><td class="num">${+m.importe ? money(m.importe) : ''}</td>
      <td style="color:var(--muted)">${esc(m.detalle)}</td><td>${esc(m.usuario)}</td></tr>`;
  }).join('') || `<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:30px">Sin movimientos</td></tr>`;
}
$('mTipo').addEventListener('change', renderMovs);
$('mq').addEventListener('input', renderMovs);

/* ---------- base de datos: Excel local o Google Sheets ---------- */
function renderAlmacen() {
  const a = S.almacen; if (!a) return;
  $('btnExcelTxt').textContent = a.modo === 'sheets' ? 'Abrir base de datos (Sheets)' : 'Abrir base de datos (Excel)';
  $('bdEstado').innerHTML = a.modo === 'sheets'
    ? `<span class="pill p-ok">Google Sheets · en la nube</span>`
    : `<span class="pill p-blue">Excel · solo en esta computadora</span>`;
  $('bdTexto').innerHTML = a.modo === 'sheets'
    ? `Todas las computadoras que usen VetGestión con esta planilla ven los mismos datos al instante.${a.planilla ? ` Planilla: <b>${esc(a.planilla)}</b>.` : ''}`
    : 'Los datos se guardan en un Excel de esta computadora (se abre con «Abrir base de datos»). Conectá una planilla de Google Sheets para compartirlos entre varias computadoras.';
  $('bdConectar').hidden = a.modo === 'sheets';
  $('bdConectado').hidden = a.modo !== 'sheets';
}

async function conectarSheets() {
  const url = $('bdUrl').value.trim();
  if (!url) { mostrarError('Pegá la URL de la aplicación web (termina en /exec). Tocá «Cómo conectarla» para ver los pasos.'); return; }
  const r = await conAviso('Conectando con Google Sheets…', () => api().conectar_sheets(url));
  if (!r.ok) { mostrarError(r.error); return; }
  $('bdUrl').value = '';
  await cargar();
  confirmar({titulo: 'Conectado a Google Sheets',
    texto: r.subidos ? 'Se subieron los productos y movimientos de esta computadora a la planilla.' : 'La planilla ya tenía datos: desde ahora se usan esos.',
    detalle: r.planilla || 'Planilla conectada', valor: 'Datos en la nube'});
}

function desconectarSheets() {
  dialogo({titulo: 'Volver a Excel local', texto: 'Se copian los datos de la planilla al Excel de esta computadora y se deja de usar Google Sheets en esta PC. La planilla en la nube no se borra.',
    boton: 'Volver a Excel', peligro: true, alAceptar: async () => {
      const r = await conAviso('Copiando datos al Excel…', () => api().desconectar_sheets());
      if (!r.ok) { mostrarError(r.error); return; }
      await cargar();
      confirmar({titulo: 'Usando Excel local', texto: r.copiados ? 'Los datos de la nube quedaron copiados en el Excel.' : 'No había conexión: el Excel quedó con los datos que tenía antes.'});
    }});
}

async function verInstrucciones() {
  const codigo = await api().codigo_script();
  abrirModal(`
    <div class="modal-box" role="dialog" aria-modal="true" style="max-width:720px">
      <button class="modal-x" onclick="cerrarModal()" aria-label="Cerrar"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg></button>
      <div class="modal-body" style="padding-top:28px">
        <h3 style="margin:0 0 14px;font-size:22px;font-weight:800">Conectar con Google Sheets</h3>
        <ul class="help">
          <li><span>1</span><div><b>Creá una planilla nueva</b> en <b>sheets.google.com</b> (en blanco). Ponele un nombre, por ejemplo “VetGestión NeoZoo”.</div></li>
          <li><span>2</span><div>En la planilla: <b>Extensiones › Apps Script</b>. Borrá lo que aparece, pegá el código de abajo y tocá <b>Guardar</b>.</div></li>
          <li><span>3</span><div><b>Implementar › Nueva implementación</b>, elegí el tipo <b>Aplicación web</b>. Ejecutar como: <b>Yo</b>. Quién tiene acceso: <b>Cualquier persona</b> (la <b>última</b> opción, la que <b>no</b> dice “con una cuenta de Google”). Tocá Implementar.</div></li>
          <li><span>4</span><div><b>Autorizá el acceso</b> con tu cuenta. Si aparece “Google no verificó esta app”: <b>Configuración avanzada › Ir a …</b></div></li>
          <li><span>5</span><div>Copiá la <b>URL de la aplicación web</b> (termina en <b>/exec</b>), pegala en VetGestión y tocá <b>Conectar</b>. En las otras computadoras pegás la misma URL.</div></li>
        </ul>
        <div class="field mt"><label>Código para pegar en Apps Script</label><textarea id="codigoGs" readonly style="min-height:190px;font-family:ui-monospace,monospace;font-size:12px">${esc(codigo)}</textarea></div>
      </div>
      <div class="modal-foot">
        <button class="btn ghost" onclick="cerrarModal()">Cerrar</button>
        <button class="btn" id="btnCopiar" onclick="copiarCodigo()">Copiar código</button>
      </div>
    </div>`);
}
function copiarCodigo() {
  const t = $('codigoGs'); t.focus(); t.select();
  document.execCommand('copy');
  $('btnCopiar').textContent = '¡Copiado!';
}
