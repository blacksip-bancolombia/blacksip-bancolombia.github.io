/**
 * VTEX Trazabilidad · BlackSip — datos del dashboard.
 *
 * El navegador pide los datos directo al webhook de n8n de producción (https://bic.blacksip.com), que solo acepta llamadas
 * desde esta página (CORS) y solo devuelve los datos del dashboard (lectura). Aquí se convierte esa respuesta al formato que usan las páginas
 * (el mismo que usaba la versión de Google Apps Script).
 */

const MKTS = [
  {
    clave: 'inmobiliario', nombre: 'Inmobiliario', descripcion: 'Proyectos de vivienda publicados por las constructoras',
    etq: { marketplace: 'Inmobiliario', grupo: 'Constructora', grupos: 'Constructoras', producto: 'Proyecto', productos: 'Proyectos', marcaAparte: false }
  },
  {
    clave: 'movilidad', nombre: 'Movilidad', descripcion: 'Vehículos publicados por los concesionarios',
    etq: { marketplace: 'Movilidad', grupo: 'Concesionario', grupos: 'Concesionarios', producto: 'Producto', productos: 'Productos', marcaAparte: false }
  }
];
const CATEGORIA_FUERA = 'Otra categoría (revisar)';

/** Marketplace pedido en la dirección: /inmobiliario, /movilidad o dashboard.html?m=… */
function mktDeLaUrl() {
  const q = new URLSearchParams(location.search).get('m') || 'inmobiliario';
  return MKTS.find(m => m.clave === String(q || '').toLowerCase()) || MKTS[0];
}

async function cargarDatosMkt(mkt) {
  const r = await fetch('https://bic.blacksip.com/webhook/trz-dashboard-datos?marketplace=' + encodeURIComponent(mkt.nombre));
  let j = null;
  try { j = await r.json(); } catch (e) { /* respuesta vacía */ }
  if (!r.ok || !j) throw new Error((j && j.error) || 'Error ' + r.status);
  return j;
}

/** "2026-11-18 21:49:15" (UTC, como sale de MySQL) → ISO con zona. */
function utcAIso(s) { return s ? String(s).replace(' ', 'T') + 'Z' : null; }
/** Día en hora Colombia (UTC−5, sin horario de verano). */
function diaColombia(s) {
  if (!s) return '';
  return new Date(new Date(utcAIso(s)).getTime() - 5 * 3600e3).toISOString().slice(0, 10);
}

/** Respuesta del webhook → { dash, historico } en el formato de las páginas. */
function adaptarDatos(mkt, j) {
  const etq = mkt.etq;
  const productos = j.productos || [];
  if (!productos.length) return { etq, dash: null, historico: null };

  const campos = (j.campos || []).map(c => ({ columna: c.columna, tipo: String(c.tipo || '').toLowerCase(), grafico: !!c.grafico }));
  const conRevision = c => c.tipo === 'youtube' || c.tipo === 'telefono';

  const columnas = ['ProductId', etq.producto, etq.grupo, 'Categoría', 'Fecha de lanzamiento'];
  campos.forEach(c => { columnas.push(c.columna); if (conRevision(c)) columnas.push(c.columna + ' · revisión'); });
  columnas.push('Estado', 'SKUs', 'SKUs activos', 'Precio desde', 'Última actualización', 'Origen última actualización', 'Días sin actualizar');
  const tipos = columnas.map(c => {
    if (c === 'Fecha de lanzamiento' || c === 'Última actualización') return 'fecha';
    if (['SKUs', 'SKUs activos', 'Precio desde', 'Días sin actualizar'].includes(c)) return 'numero';
    if (c === 'ProductId') return 'id';
    const k = campos.find(x => x.columna === c);
    return k && k.tipo !== 'texto' ? k.tipo : 'texto';
  });
  const idx = n => columnas.indexOf(n);

  const ahora = Date.now();
  const sinRegion = {};
  const filas = productos.map(p => {
    const specs = p.specs || {}, rev = p.revisiones || {};
    const valores = [];
    campos.forEach(c => {
      valores.push(specs[c.columna] ?? '');
      if (conRevision(c)) valores.push(rev[c.columna] ?? '');
    });
    const r = rev['Región'];
    if (r && r.indexOf('Departamento sin región') === 0) {
      const dep = r.split(': ').slice(1).join(': ');
      sinRegion[dep] = (sinRegion[dep] || 0) + 1;
    }
    const dias = p.ultima_actualizacion
      ? Math.max(0, Math.floor((ahora - new Date(utcAIso(p.ultima_actualizacion)).getTime()) / 86400000)) : '';
    // Igual que la versión anterior: activo = producto activo con al menos un SKU activo
    const activo = !!p.activo && Number(p.skus_activos) > 0;
    return [p.id, p.nombre || '', p.grupo || '', p.categoria || '', p.lanzamiento || '']
      .concat(valores)
      .concat([activo ? 'Activo' : 'Inactivo', p.skus, p.skus_activos, p.precio_desde == null ? '' : p.precio_desde,
        diaColombia(p.ultima_actualizacion), p.ultima_origen || '', dias]);
  });

  const presentes = [...new Set(filas.map(f => f[3]).filter(c => c && c !== CATEGORIA_FUERA))].sort((a, b) => a.localeCompare(b));
  const avisos = [];
  const faltan = Object.keys(sinRegion);
  if (faltan.length) avisos.push(faltan.length + (faltan.length === 1 ? ' departamento no tiene' : ' departamentos no tienen') +
    ' región asignada (' + faltan.slice(0, 5).join(', ') + (faltan.length > 5 ? '…' : '') + '). Agréguelos en la tabla de regiones.');

  const dash = {
    fecha: utcAIso(j.portada && j.portada.datos_utc) || new Date().toISOString(),
    marketplace: mkt.nombre, columnas, tipos,
    graficos: campos.filter(c => c.grafico).map(c => idx(c.columna)).filter(i => i >= 0),
    categorias: presentes.concat([CATEGORIA_FUERA]),
    roles: {
      pid: 0, producto: 1, grupo: 2, marca: -1, categoria: 3, lanzamiento: 4, estado: idx('Estado'),
      ultima: idx('Última actualización'), dias: idx('Días sin actualizar'), primera: -1,
      precio: idx('Precio desde'), skusAct: idx('SKUs activos')
    },
    hayBase: true, avisos, filas, etq
  };

  // Histórico: total diario y, desde el 19/11/2026, el detalle por categoría y grupo ([categoría, grupo, activos])
  const historico = { cats: [], grupos: [], dias: {} };
  const pos = (lista, v) => { let k = lista.indexOf(v); if (k < 0) { lista.push(v); k = lista.length - 1; } return k; };
  (j.historico || []).forEach(h => {
    const det = Array.isArray(h.detalle) ? h.detalle : [];
    historico.dias[h.fecha] = {
      t: h.activos, det: det.length > 0,
      d: det.map(x => [pos(historico.cats, String(x[0] || '')), pos(historico.grupos, String(x[1] || '')), Number(x[2]) || 0])
    };
  });

  return { etq, dash, historico };
}

/** Números de la tarjeta de la portada, con la misma regla de "activo" del dashboard. */
function resumenPortada(datos) {
  const D = datos.dash;
  if (!D) return null;
  const R = D.roles;
  const activos = D.filas.filter(f => f[R.estado] === 'Activo');
  const grupos = new Set();
  activos.forEach(f => String(f[R.grupo] || '').split(' | ').filter(Boolean).forEach(g => grupos.add(g)));
  return { fecha: D.fecha, activos: activos.length, grupos: grupos.size, total: D.filas.length };
}
