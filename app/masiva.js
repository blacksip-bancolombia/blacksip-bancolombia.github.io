/* masiva.js · Carga masiva por Excel (usa api(), esc() y fmt de index.html, y Excel de excel.js)
 * 1) Descargar: lee cada producto EN VIVO desde VTEX (acción "producto") y arma un Excel con hojas Productos y SKUs.
 *    Las hojas ocultas _original y _campos guardan los valores del momento y el tipo de cada columna.
 * 2) Subir: compara cada celda con _original, valida, muestra la revisión y aplica con las acciones
 *    "guardar" y "guardar_sku" (origen "Carga masiva"). n8n vuelve a leer VTEX: si algo cambió desde la descarga → CONFLICTO.
 */
const M = { cambios: [], archivo: '', aplicando: false };
const SEP = '; ';                                      // separador de varios valores en una celda
const P_FIJAS = [['Nombre', 'nombre'], ['Título SEO', 'titulo'], ['Descripción', 'descripcion'], ['Descripción corta', 'descripcion_corta'], ['Meta descripción', 'meta'], ['Palabras clave', 'palabras']];
const S_FIJAS = ['Activo', 'Nombre SKU', 'RefId', 'Precio base', 'Precio lista'];
const LISTA_UNA = [5, 6], LISTA_VARIAS = [7], MULTILINEA = [2, 9];
const txt = v => String(v ?? '').replace(/\r\n/g, '\n').trim();
const normM = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const partes = v => txt(v).split(/\s*;\s*/).map(x => x.trim()).filter(Boolean);
const hoyM = () => new Date().toLocaleDateString('en-CA', { timeZone: 'America/Bogota' });
const slug = s => normM(s).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40) || 'productos';

function vistaModo(m) {
  document.querySelectorAll('.modos button').forEach(b => b.classList.toggle('activo', b.dataset.m === m));
  document.querySelector('#vistaApp .layout').style.display = m === 'uno' ? '' : 'none';
  $('vistaMasiva').style.display = m === 'masiva' ? '' : 'none';
}

/* ---------------- 1. Descargar ---------------- */
async function idsParaDescargar() {
  const lista = $('mIds').value.split(/[^0-9]+/).map(Number).filter(n => Number.isInteger(n) && n > 0);
  if (lista.length) return { ids: [...new Set(lista)].slice(0, 300), nombre: 'ids' };
  const grupo = $('mGrupo').value.trim();
  if (!grupo) throw new Error('Elija una constructora o escriba ProductIds.');
  const j = await api({ accion: 'buscar', grupo, todos: true });
  const ids = (j.resultados || []).map(r => r.id).filter(id => id !== 1);
  if (!ids.length) throw new Error('No se encontraron proyectos de esa constructora.');
  return { ids, nombre: grupo };
}
async function descargarExcel() {
  const btn = $('mBtnDescargar'), est = $('mEstadoDescarga');
  btn.disabled = true; est.innerHTML = '<div class="cargando">Buscando proyectos…</div>';
  try {
    const { ids, nombre } = await idsParaDescargar();
    const fichas = [], fallos = []; let hechos = 0;
    const pintar = () => { est.innerHTML = '<div class="progreso"><div style="width:' + Math.round(hechos / ids.length * 100) + '%"></div></div>' +
      '<div class="m">Leyendo de VTEX: ' + hechos + ' de ' + ids.length + ' proyectos…</div>'; };
    pintar();
    let i = 0;
    const trabajador = async () => { while (i < ids.length) { const id = ids[i++];
      try { const f = await api({ accion: 'producto', product_id: id }); if (f.error) throw new Error(f.error); fichas.push(f); }
      catch (e) { fallos.push(id + ' (' + e.message + ')'); }
      hechos++; pintar(); } };
    await Promise.all([trabajador(), trabajador(), trabajador()]);       // 3 a la vez para no saturar n8n
    if (!fichas.length) throw new Error('No se pudo leer ningún proyecto. ' + fallos.slice(0, 3).join('; '));
    fichas.sort((a, b) => ids.indexOf(a.producto.id) - ids.indexOf(b.producto.id));
    const blob = await armarExcel(fichas);
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'carga_' + slug(nombre) + '_' + hoyM() + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    est.innerHTML = '<div class="ok-msg">Listo: ' + fichas.length + ' proyectos y ' + fichas.reduce((n, f) => n + (f.skus || []).filter(s => !s.error).length, 0) + ' SKUs.</div>' +
      (fallos.length ? '<div class="error">No se pudieron leer: ' + esc(fallos.join(', ')) + '</div>' : '');
  } catch (e) { est.innerHTML = '<div class="error">' + esc(e.message) + '</div>'; }
  btn.disabled = false;
}

async function armarExcel(fichas) {
  // Columnas de especificaciones: unión de todos los proyectos, en el orden en que aparecen
  const specP = new Map(), specS = new Map();
  for (const f of fichas) {
    for (const s of f.especificaciones || []) if (!specP.has(s.nombre)) specP.set(s.nombre, { tipo_id: s.tipo_id ?? null, opciones: s.opciones || null });
    for (const s of f.sku_campos || []) if (!specS.has(s.nombre)) specS.set(s.nombre, { tipo_id: s.tipo_id ?? null, opciones: s.opciones || null });
  }
  const colsP = ['ProductId', 'Proyecto', 'Constructora', ...P_FIJAS.map(x => x[0]), ...specP.keys()];
  const colsS = ['SkuId', 'ProductId', 'Proyecto', ...S_FIJAS, ...specS.keys()];
  const original = [['hoja', 'id', 'columna', 'valor']];
  const filasP = [colsP], filasS = [colsS];
  const celda = (tipo, vals) => MULTILINEA.includes(tipo) ? vals.join('\n') : vals.join(SEP);
  for (const f of fichas) {
    const p = f.producto, fila = [p.id, p.nombre || '', p.marca || '']; original.push(['Productos', p.id, '__fila', '1']);
    for (const [k, c] of P_FIJAS) { const v = txt(p[c]); fila.push(v); if (v) original.push(['Productos', p.id, k, v]); }
    for (const [nombre, t] of specP) {
      const s = (f.especificaciones || []).find(x => x.nombre === nombre); const v = s ? celda(t.tipo_id, s.valores || []) : '';
      fila.push(v); if (v) original.push(['Productos', p.id, nombre, v]);
    }
    filasP.push(fila);
    for (const s of f.skus || []) {
      if (s.error) continue;
      const pr = s.precio || {};
      const fs = [s.id, p.id, p.nombre || '', s.activo ? 'Sí' : 'No', txt(s.nombre), txt(s.ref_id), pr.base ?? '', pr.lista ?? ''];
      original.push(['SKUs', s.id, '__fila', '1']); original.push(['SKUs', s.id, 'Activo', s.activo ? 'Sí' : 'No']);
      if (txt(s.nombre)) original.push(['SKUs', s.id, 'Nombre SKU', txt(s.nombre)]);
      if (txt(s.ref_id)) original.push(['SKUs', s.id, 'RefId', txt(s.ref_id)]);
      if (pr.base != null) original.push(['SKUs', s.id, 'Precio base', String(pr.base)]);
      if (pr.lista != null) original.push(['SKUs', s.id, 'Precio lista', String(pr.lista)]);
      for (const [nombre, t] of specS) {
        const x = (s.specs || []).find(y => y.nombre === nombre); const v = x ? celda(t.tipo_id, x.valores || []) : '';
        fs.push(v); if (v) original.push(['SKUs', s.id, nombre, v]);
      }
      filasS.push(fs);
    }
  }
  // Hoja de opciones (para las listas desplegables de Excel) y descripción de columnas
  const opciones = [], campos = [['hoja', 'indice', 'columna', 'tipo_id', 'opciones']];
  const valP = [], valS = [];
  const conLista = (hoja, cols, mapa, desde, vals, nfilas) => cols.forEach((nombre, i) => {
    if (i < desde) return; const t = mapa.get(nombre); if (!t) return;
    campos.push([hoja, i, nombre, t.tipo_id ?? '', t.opciones ? JSON.stringify(t.opciones) : '']);
    if (LISTA_UNA.includes(t.tipo_id) && Array.isArray(t.opciones) && t.opciones.length) {
      const c = opciones.length; opciones.push([hoja + ': ' + nombre, ...t.opciones]);
      vals.push({ rango: Excel.col(i) + '2:' + Excel.col(i) + Math.max(2, nfilas + 200), lista: 'Opciones!$' + Excel.col(c) + '$2:$' + Excel.col(c) + '$' + (t.opciones.length + 1) });
    }
  });
  conLista('Productos', colsP, specP, 3 + P_FIJAS.length, valP, filasP.length);
  conLista('SKUs', colsS, specS, 3 + S_FIJAS.length, valS, filasS.length);
  valS.push({ rango: 'D2:D' + Math.max(2, filasS.length + 200), lista: '"Sí,No"' });
  colsP.forEach((n, i) => { if (i >= 3 && i < 3 + P_FIJAS.length) campos.push(['Productos', i, n, '', '']); });
  colsS.forEach((n, i) => { if (i >= 3 && i < 3 + S_FIJAS.length) campos.push(['SKUs', i, n, '', '']); });
  const maxOp = Math.max(1, ...opciones.map(o => o.length));
  const hojaOp = [opciones.map(o => o[0])]; for (let r = 1; r < maxOp; r++) hojaOp.push(opciones.map(o => o[r] ?? ''));
  const instrucciones = [['Cómo usar este archivo'],
    ['1. Edite solo las columnas blancas. Las grises (ProductId, SkuId, Proyecto, Constructora) son de referencia: el nombre se cambia en la columna Nombre.'],
    ['2. No cambie el orden ni el nombre de las columnas, y no borre las hojas ocultas.'],
    ['3. Campos de lista: elija de la flecha desplegable. Varias opciones en una celda: sepárelas con punto y coma (;).'],
    ['4. Activo: Sí o No. Precios: solo números, sin puntos de miles (ej. 964064622).'],
    ['5. Celdas vacías que tenían valor no se borran en VTEX (vaciar no está soportado).'],
    ['6. Suba el archivo en la app (Carga masiva → Subir). Verá cada cambio antes de aplicarlo.'],
    ['7. Si alguien cambió algo en VTEX después de esta descarga, ese campo sale como CONFLICTO y no se sobrescribe.'],
    [''], ['Descargado el ' + new Date().toLocaleString('es-CO', { timeZone: 'America/Bogota' }) + ' (valores leídos de VTEX en ese momento).']];
  const anchos = cols => cols.map((n, i) => i === 0 ? 11 : (n === 'Descripción' ? 60 : Math.min(40, Math.max(12, n.length + 4))));
  return Excel.crear([
    { nombre: 'Productos', filas: filasP, anchos: anchos(colsP), soloLectura: [0, 1, 2], validaciones: valP },
    { nombre: 'SKUs', filas: filasS, anchos: anchos(colsS), soloLectura: [0, 1, 2], validaciones: valS },
    { nombre: 'Opciones', filas: hojaOp, anchos: hojaOp[0].map(() => 28) },
    { nombre: 'Instrucciones', filas: instrucciones, anchos: [110] },
    { nombre: '_original', filas: original, oculta: true },
    { nombre: '_campos', filas: campos, oculta: true }
  ]);
}

/* ---------------- 2. Subir y revisar ---------------- */
async function leerArchivo(input) {
  const f = input.files && input.files[0]; if (!f) return;
  const est = $('mEstadoSubida'); est.innerHTML = '<div class="cargando">Leyendo el archivo…</div>'; $('mRevision').innerHTML = '';
  try {
    const libro = await Excel.leer(f);
    for (const h of ['Productos', 'SKUs', '_original', '_campos'])
      if (!libro[h]) throw new Error('Falta la hoja "' + h + '". Use un archivo descargado desde esta app (sin borrar hojas).');
    M.cambios = detectar(libro); M.archivo = f.name;
    est.innerHTML = '';
    pintarRevision();
  } catch (e) { est.innerHTML = '<div class="error">' + esc(e.message) + '</div>'; }
  input.value = '';
}

function detectar(libro) {
  const orig = {}; for (const [h, id, c, v] of libro._original.slice(1)) if (h) orig[h + '|' + id + '|' + c] = txt(v);
  const campos = {}; for (const [h, i, n, t, o] of libro._campos.slice(1)) if (h) {
    let op = null; try { op = o ? JSON.parse(o) : null; } catch (e) { op = null; }
    (campos[h] ||= {})[Number(i)] = { nombre: n, tipo_id: t === '' || t == null ? null : Number(t), opciones: op };
  }
  const cambios = [], avisos = [];
  for (const hoja of ['Productos', 'SKUs']) {
    const filas = libro[hoja], enc = (filas[0] || []).map(txt);
    for (const [i, c] of Object.entries(campos[hoja] || {})) if (enc[i] !== c.nombre)
      throw new Error('En la hoja ' + hoja + ' la columna ' + Excel.col(Number(i)) + ' debería llamarse "' + c.nombre + '". No cambie el orden ni el nombre de las columnas.');
    const conocidos = new Set(Object.keys(orig).filter(k => k.startsWith(hoja + '|') && k.endsWith('|__fila')).map(k => Number(k.split('|')[1])));
    for (const fila of filas.slice(1)) {
      const id = Number(fila[0]); if (!Number.isInteger(id) || id <= 0) continue;
      const pid = hoja === 'Productos' ? id : Number(fila[1]);
      if (!conocidos.has(id)) { avisos.push((hoja === 'SKUs' ? 'SKU ' : 'Producto ') + id); continue; }
      for (const [i, c] of Object.entries(campos[hoja] || {})) {
        const crudo = fila[Number(i)]; const antes = orig[hoja + '|' + id + '|' + c.nombre] ?? '';
        let nuevo = typeof crudo === 'number' ? String(crudo) : txt(crudo);
        const fija = hoja === 'Productos' ? P_FIJAS.some(x => x[0] === c.nombre) : S_FIJAS.includes(c.nombre);
        const lista = !fija && (LISTA_UNA.includes(c.tipo_id) || LISTA_VARIAS.includes(c.tipo_id));
        const igual = lista ? partes(antes).map(normM).sort().join('|') === partes(nuevo).map(normM).sort().join('|')
          : (c.nombre.startsWith('Precio') && nuevo !== '' && antes !== '' ? Number(nuevo) === Number(antes) : nuevo === antes);
        if (igual) continue;
        const ch = { hoja, product_id: pid, sku_id: hoja === 'SKUs' ? id : null, columna: c.nombre, antes, nuevo, error: '', alerta: '', aplicar: true, resultado: '', detalle: '' };
        ch.proyecto = txt(hoja === 'Productos' ? fila[1] : fila[2]);
        // Validaciones (n8n vuelve a validar todo antes de escribir)
        if (!nuevo && antes) ch.error = 'Vaciar un valor no está soportado.';
        else if (lista) {
          const op = Array.isArray(c.opciones) ? c.opciones : []; const vs = partes(nuevo);
          const fuera = vs.filter(x => !op.some(o => normM(o) === normM(x)));
          if (fuera.length) ch.error = '"' + fuera.join('", "') + '" no es una opción de la lista.';
          else if (LISTA_UNA.includes(c.tipo_id) && vs.length > 1) ch.error = 'Este campo admite una sola opción.';
          else if (hoja === 'SKUs' && LISTA_VARIAS.includes(c.tipo_id)) ch.error = 'En SKUs, los campos de varias opciones se editan en VTEX.';
          ch.envio = vs.map(x => op.find(o => normM(o) === normM(x)) || x).join('\n'); ch.envioAntes = partes(antes).join('\n');
        } else if (c.nombre === 'Activo') {
          const n = normM(nuevo); const b = ['si', 'activo', '1', 'verdadero', 'true'].includes(n) ? 'Sí' : ['no', 'inactivo', '0', 'falso', 'false'].includes(n) ? 'No' : null;
          if (!b) ch.error = 'Escriba Sí o No.'; else { ch.envio = b; if (b === antes) continue; }
        } else if (c.nombre.startsWith('Precio')) {
          const n = Number(String(nuevo).replace(/\s/g, ''));
          if (!Number.isFinite(n) || n <= 0) ch.error = 'Debe ser un número mayor que cero, sin puntos de miles.';
          else { ch.envio = String(n); const a = Number(antes); if (a) { const pct = Math.round((n - a) / a * 100); if (Math.abs(pct) > 30) ch.alerta = (pct > 0 ? '+' : '') + pct + ' %'; } }
        } else if (c.tipo_id === 4 && !/^-?\d+([.,]\d+)?$/.test(nuevo)) ch.error = 'Este campo solo admite números.';
        else if (!fija && !MULTILINEA.includes(c.tipo_id) && nuevo.includes('\n')) ch.error = 'Este campo admite una sola línea.';
        if (ch.envio === undefined) ch.envio = nuevo;
        if (ch.envioAntes === undefined) ch.envioAntes = antes;
        ch.campo = hoja === 'Productos' ? (fija ? c.nombre : 'Spec: ' + c.nombre) : (fija ? c.nombre : 'Spec SKU: ' + c.nombre);
        if (ch.error) ch.aplicar = false;
        cambios.push(ch);
      }
    }
  }
  M.avisos = avisos;
  return cambios;
}

function pintarRevision() {
  const c = M.cambios, ok = c.filter(x => !x.error).length, mal = c.length - ok;
  const proyectos = new Set(c.map(x => x.product_id)).size;
  if (!c.length) { $('mRevision').innerHTML = '<div class="vacio">El archivo no tiene cambios frente a la descarga.</div>'; return; }
  const corto = s => { s = String(s ?? ''); return s ? esc(s.length > 90 ? s.slice(0, 90) + '…' : s) : '<span class="sin">vacío</span>'; };
  const clase = r => r === 'OK' ? 'res-ok' : r === 'CONFLICTO' ? 'res-conflicto' : r ? 'res-error' : '';
  $('mRevision').innerHTML =
    '<div class="resumen-m"><b>' + c.length + '</b> cambios en <b>' + proyectos + '</b> proyectos · <span class="res-ok">' + ok + ' listos</span>' +
    (mal ? ' · <span class="res-error">' + mal + ' con error de validación (no se aplicarán)</span>' : '') +
    (M.avisos && M.avisos.length ? ' · filas nuevas ignoradas: ' + esc(M.avisos.slice(0, 8).join(', ')) : '') + '</div>' +
    '<div class="tabla-m"><table class="hist"><thead><tr><th><input type="checkbox" checked onchange="marcarTodos(this.checked)"' + (M.aplicando ? ' disabled' : '') + '></th>' +
    '<th>Proyecto</th><th>SKU</th><th>Campo</th><th>Antes</th><th>Nuevo</th><th>Estado</th></tr></thead><tbody>' +
    c.map((x, i) => '<tr><td><input type="checkbox" ' + (x.aplicar ? 'checked' : '') + (x.error || M.aplicando || x.resultado ? ' disabled' : '') +
      ' onchange="M.cambios[' + i + '].aplicar=this.checked;contarM()"></td><td>' + x.product_id + '<div class="m">' + esc(x.proyecto) + '</div></td><td>' + (x.sku_id || '') + '</td><td>' +
      esc(x.columna) + '</td><td>' + corto(x.antes) + '</td><td>' + corto(x.nuevo) + (x.alerta ? '<span class="alerta">⚠ ' + esc(x.alerta) + '</span>' : '') + '</td><td>' +
      (x.error ? '<span class="res-error">No válido</span><div class="m">' + esc(x.error) + '</div>'
        : x.resultado ? '<span class="' + clase(x.resultado) + '">' + esc(x.resultado) + '</span><div class="m">' + esc(x.detalle) + '</div>' : '<span class="m">Pendiente</span>') +
      '</td></tr>').join('') + '</tbody></table></div>' +
    '<div class="barra-m"><span id="mInfo"></span><button class="btn-sec" onclick="descargarResultado()">Descargar este reporte</button>' +
    '<button class="btn-morado" id="mBtnAplicar" onclick="confirmarAplicar()">Aplicar</button></div>';
  contarM();
}
function marcarTodos(si) { M.cambios.forEach(x => { if (!x.error && !x.resultado) x.aplicar = si; }); pintarRevision(); }
function contarM() {
  const n = M.cambios.filter(x => x.aplicar && !x.error && !x.resultado).length, b = $('mBtnAplicar');
  if (b) { b.disabled = !n || M.aplicando; b.textContent = M.aplicando ? 'Aplicando…' : 'Aplicar ' + n + (n === 1 ? ' cambio' : ' cambios') + ' en VTEX'; }
}
function confirmarAplicar() {
  const sel = M.cambios.filter(x => x.aplicar && !x.error && !x.resultado); if (!sel.length) return;
  const proy = new Set(sel.map(x => x.product_id)).size, alertas = sel.filter(x => x.alerta).length;
  const m = document.createElement('div'); m.className = 'modal-fondo'; m.id = 'modal';
  m.innerHTML = '<div class="modal"><h3>Aplicar carga masiva</h3><div class="adv">Se aplicarán <b>' + sel.length + ' cambios en ' + proy + ' proyectos</b> directamente en <b>VTEX (producción)</b>.' +
    (alertas ? '<br><b>' + alertas + ' precios cambian más del 30 %.</b>' : '') + '<br>Antes de escribir se vuelve a leer VTEX: si algo cambió desde la descarga, ese campo sale como CONFLICTO y no se sobrescribe.</div>' +
    '<div class="acc"><button class="btn-sec" onclick="cerrarModal()">Volver</button><button class="btn-morado" onclick="cerrarModal();aplicarMasiva()">Sí, aplicar</button></div></div>';
  document.body.appendChild(m);
}
async function aplicarMasiva() {
  const sel = M.cambios.filter(x => x.aplicar && !x.error && !x.resultado); if (!sel.length) return;
  M.aplicando = true; pintarRevision();
  const porProducto = new Map(); for (const x of sel) { if (!porProducto.has(x.product_id)) porProducto.set(x.product_id, []); porProducto.get(x.product_id).push(x); }
  let hechos = 0; const total = porProducto.size;
  const progreso = () => { $('mInfo').innerHTML = '<div class="progreso"><div style="width:' + Math.round(hechos / total * 100) + '%"></div></div>Proyecto ' + hechos + ' de ' + total; };
  progreso();
  for (const [pid, lista] of porProducto) {
    const prod = lista.filter(x => !x.sku_id), skus = lista.filter(x => x.sku_id);
    const marcar = (grupo, r) => { for (const x of grupo) {
      const etiqueta = x.sku_id ? 'SKU ' + x.sku_id + ' · ' + x.campo : x.campo;
      const res = (r || []).find(y => y.campo === etiqueta);
      x.resultado = res ? res.resultado : 'ERROR'; x.detalle = res ? res.detalle : 'n8n no devolvió resultado para este campo.'; } };
    for (const [grupo, accion] of [[prod, 'guardar'], [skus, 'guardar_sku']]) {
      if (!grupo.length) continue;
      try {
        const r = await api({ accion, product_id: pid, origen: 'Carga masiva',
          cambios: grupo.map(x => ({ campo: x.campo, antes: x.envioAntes, nuevo: x.envio, ...(x.sku_id ? { sku_id: x.sku_id } : {}) })) });
        marcar(grupo, r.resultados);
      } catch (e) { grupo.forEach(x => { x.resultado = 'ERROR'; x.detalle = e.message; }); }
    }
    hechos++; pintarRevision(); progreso();
  }
  M.aplicando = false; pintarRevision();
  const ok = sel.filter(x => x.resultado === 'OK').length, conf = sel.filter(x => x.resultado === 'CONFLICTO').length, err = sel.length - ok - conf;
  $('mInfo').innerHTML = '<b>Terminado:</b> <span class="res-ok">' + ok + ' OK</span> · <span class="res-conflicto">' + conf + ' conflictos</span> · <span class="res-error">' + err + ' errores</span>';
}
async function descargarResultado() {
  const filas = [['ProductId', 'Proyecto', 'SkuId', 'Campo', 'Antes', 'Nuevo', 'Estado', 'Detalle']];
  for (const x of M.cambios) filas.push([x.product_id, x.proyecto, x.sku_id || '', x.columna, x.antes, x.nuevo,
    x.error ? 'No válido' : (x.resultado || (x.aplicar ? 'Pendiente' : 'No seleccionado')), x.error || x.detalle]);
  const blob = await Excel.crear([{ nombre: 'Resultado', filas, anchos: [11, 34, 10, 26, 30, 30, 14, 60] }]);
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'resultado_' + (M.archivo || 'carga').replace(/\.xlsx$/i, '') + '.xlsx';
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
