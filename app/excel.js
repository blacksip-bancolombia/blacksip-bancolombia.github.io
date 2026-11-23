/* excel.js · lectura y escritura mínima de .xlsx (con JSZip). Sin dependencias externas.
 * Excel.crear(hojas) → Blob. hojas: [{ nombre, filas: [[valor,...]], anchos: [n], oculta, validaciones: [{ rango, lista }], soloLectura: [col] }]
 *   Opcionales (planillas con diseño): estilo(r, c) → id de estilo o null · combinar: ['A1:G1'] · altos: { fila0: alto } · congelar: filas fijas (1)
 *   · ocultas: [col] · filtro: 'A3:G90' (en vez del automático)
 * Excel.leer(archivo) → { NombreHoja: [[valor,...]] }  (valores: texto o número)
 */
(function () {
  const escXml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const col = n => { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };

  function hojaXml(h) {
    const filas = h.filas || [], lect = new Set(h.soloLectura || []);
    const altos = h.altos || {};
    const sd = filas.map((fila, r) => '<row r="' + (r + 1) + '"' + (altos[r] ? ' ht="' + altos[r] + '" customHeight="1"' : '') + '>' + fila.map((v, c) => {
      if (v === null || v === undefined || v === '') { const e = h.estilo ? h.estilo(r, c) : null; return e != null ? '<c r="' + col(c) + (r + 1) + '" s="' + e + '"/>' : ''; }
      const e = h.estilo ? h.estilo(r, c) : null;
      const ref = col(c) + (r + 1), s = e != null ? e : (r === 0 ? 1 : (lect.has(c) ? 2 : 3));
      if (typeof v === 'number' && Number.isFinite(v)) return '<c r="' + ref + '" s="' + s + '"><v>' + v + '</v></c>';
      return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + escXml(v) + '</t></is></c>';
    }).join('') + '</row>').join('');
    const ocultas = new Set(h.ocultas || []);
    const anchos = (h.anchos || []).map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"' + (ocultas.has(i) ? ' hidden="1"' : '') + '/>').join('');
    const fijas = h.congelar || 1;
    const comb = (h.combinar || []).length ? '<mergeCells count="' + h.combinar.length + '">' + h.combinar.map(r => '<mergeCell ref="' + r + '"/>').join('') + '</mergeCells>' : '';
    const val = (h.validaciones || []).map(v => '<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Valor no permitido" error="Elija una opción de la lista." sqref="' +
      v.rango + '"><formula1>' + escXml(v.lista) + '</formula1></dataValidation>').join('');
    const ncols = Math.max(1, ...filas.map(f => f.length));
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      (h.horizontal ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : '') + '<sheetViews><sheetView workbookViewId="0"' + (h.oculta ? '' : '') + '><pane ySplit="' + fijas + '" topLeftCell="A' + (fijas + 1) + '" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' + (anchos ? '<cols>' + anchos + '</cols>' : '') + '<sheetData>' + sd + '</sheetData>' +
      (h.filtro ? '<autoFilter ref="' + h.filtro + '"/>' : (filas.length > 1 && !h.oculta ? '<autoFilter ref="A1:' + col(ncols - 1) + filas.length + '"/>' : '')) +
      comb + (val ? '<dataValidations count="' + h.validaciones.length + '">' + val + '</dataValidations>' : '') +
      (h.horizontal ? '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' : '') + '</worksheet>';
  }

  async function crear(hojas) {
    const z = new JSZip();
    z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      hojas.map((h, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>');
    z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    const absRef = r => r.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
    const filtros = hojas.map((h, i) => h.filtro ? '<definedName name="_xlnm._FilterDatabase" localSheetId="' + i + '" hidden="1">\'' + escXml(h.nombre) + '\'!' + absRef(h.filtro) + '</definedName>'
      : (h.filas || []).length > 1 && !h.oculta
      ? '<definedName name="_xlnm._FilterDatabase" localSheetId="' + i + '" hidden="1">\'' + escXml(h.nombre) + '\'!$A$1:$' + col(Math.max(1, ...h.filas.map(f => f.length)) - 1) + '$' + h.filas.length + '</definedName>' : '').join('');
    z.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
      hojas.map((h, i) => '<sheet name="' + escXml(h.nombre) + '" sheetId="' + (i + 1) + '"' + (h.oculta ? ' state="hidden"' : '') + ' r:id="rId' + (i + 1) + '"/>').join('') +
      '</sheets>' + (filtros ? '<definedNames>' + filtros + '</definedNames>' : '') + '</workbook>');
    z.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      hojas.map((h, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
      '<Relationship Id="rId' + (hojas.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    // Estilos: 0 normal · 1 encabezado (negrita, fondo morado claro) · 2 solo lectura (fondo gris) · 3 editable (ajuste de texto)
    z.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="10"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +                        // 2 título
      '<font><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +                              // 3 descripción
      '<font><b/><sz val="11"/><color rgb="FF2C2A29"/><name val="Calibri"/></font>' +                          // 4 encabezado oscuro
      '<font><b/><sz val="14"/><color rgb="FFFDDA24"/><name val="Calibri"/></font>' +                    // 5 marca (amarillo)
      '<font><b/><sz val="11"/><color rgb="FF00654A"/><name val="Calibri"/></font>' +                    // 6 OK
      '<font><b/><sz val="11"/><color rgb="FFB23A0E"/><name val="Calibri"/></font>' +                    // 7 ERROR
      '<font><b/><sz val="11"/><color rgb="FF8A6A00"/><name val="Calibri"/></font>' +                    // 8 CONFLICTO
      '<font><i/><sz val="10"/><color rgb="FF66625F"/><name val="Calibri"/></font></fonts>' +            // 9 subtítulo
      '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF9063CD"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEFEDEC"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF2C2A29"/><bgColor indexed="64"/></patternFill></fill>' +   // 4 negro
      '<fill><patternFill patternType="solid"><fgColor rgb="FF454240"/><bgColor indexed="64"/></patternFill></fill>' +   // 5 gris oscuro
      '<fill><patternFill patternType="solid"><fgColor rgb="FFE6E4E3"/><bgColor indexed="64"/></patternFill></fill>' +   // 6 gris claro
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFDDA24"/><bgColor indexed="64"/></patternFill></fill>' +   // 7 amarillo
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF3A3"/><bgColor indexed="64"/></patternFill></fill></fills>' +  // 8 amarillo claro
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color rgb="FFC9C5C2"/></left><right style="thin"><color rgb="FFC9C5C2"/></right><top style="thin"><color rgb="FFC9C5C2"/></top><bottom style="thin"><color rgb="FFC9C5C2"/></bottom><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="19"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      // Planilla con diseño: 4 título · 5 descripción · 6 encabezado informativo · 7 encabezado editable · 8 dato informativo · 9 dato editable · 10 marca
      '<xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="4" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="4" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="8" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="5" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
      // 11 encabezado de tabla (negro) · 12 OK · 13 ERROR · 14 CONFLICTO · 15 número · 16 número editable · 17 subtítulo · 18 texto con ajuste
      '<xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="6" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="7" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="8" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>' +
      '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
      '<xf numFmtId="3" fontId="0" fillId="8" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
    hojas.forEach((h, i) => z.file('xl/worksheets/sheet' + (i + 1) + '.xml', hojaXml(h)));
    return z.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
  }

  async function leer(archivo) {
    const z = await JSZip.loadAsync(archivo);
    const dp = new DOMParser();
    const xml = async p => { const f = z.file(p); return f ? dp.parseFromString(await f.async('string'), 'application/xml') : null; };
    const hijos = (n, tag) => [...n.getElementsByTagName(tag)];
    const textoDe = n => hijos(n, 't').map(t => t.textContent).join('');
    const wb = await xml('xl/workbook.xml'); if (!wb) throw new Error('El archivo no es un Excel (.xlsx) válido.');
    const rels = await xml('xl/_rels/workbook.xml.rels'); const destino = {};
    for (const r of hijos(rels, 'Relationship')) destino[r.getAttribute('Id')] = r.getAttribute('Target');
    const ss = await xml('xl/sharedStrings.xml'); const compartidos = ss ? hijos(ss, 'si').map(textoDe) : [];
    const out = {};
    for (const s of hijos(wb, 'sheet')) {
      const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      let t = destino[rid] || ''; t = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
      const h = await xml(t); if (!h) continue;
      const filas = [];
      for (const row of hijos(h, 'row')) {
        const r = Number(row.getAttribute('r')) - 1; const fila = filas[r] = [];
        for (const c of hijos(row, 'c')) {
          const ref = c.getAttribute('r') || ''; const ci = colNum(ref.replace(/\d+/g, '')); const tipo = c.getAttribute('t');
          const v = c.getElementsByTagName('v')[0];
          let val = '';
          if (tipo === 's') val = compartidos[Number(v ? v.textContent : -1)] ?? '';
          else if (tipo === 'inlineStr') val = textoDe(c);
          else if (tipo === 'str' || tipo === 'e') val = v ? v.textContent : '';
          else if (tipo === 'b') val = v && v.textContent === '1' ? 'VERDADERO' : 'FALSO';
          else val = v ? Number(v.textContent) : '';
          fila[ci] = val;
        }
      }
      for (let i = 0; i < filas.length; i++) if (!filas[i]) filas[i] = [];
      out[s.getAttribute('name')] = filas;
    }
    return out;
  }
  /* Hoja con diseño BlackSip: fila 1 título, fila 2 contexto (filtros, fecha), fila 3 encabezados, datos desde la fila 4.
   * op: { nombre, titulo, subtitulo, encabezados, filas, anchos?, numericas?: [col], editables?: [col], resultado?: col, ocultas?, validaciones?, horizontal? } */
  function tabla(op) {
    const n = op.encabezados.length, num = new Set(op.numericas || []), edit = new Set(op.editables || []);
    const filas = [[op.titulo].concat(Array(n - 1).fill('')), [op.subtitulo || ''].concat(Array(n - 1).fill('')), op.encabezados].concat(op.filas);
    const anchos = op.anchos || op.encabezados.map((h, c) => Math.min(48, Math.max(10, String(h).length + 3,
      ...op.filas.slice(0, 300).map(f => Math.min(48, String(f[c] ?? '').length + 2)))));
    const res = r => { const v = String(filas[r][op.resultado] || ''); return v === 'OK' ? 12 : v === 'CONFLICTO' ? 14 : (v === 'ERROR' || v === 'No válido') ? 13 : 8; };
    const ultima = col(n - 1);
    return { nombre: op.nombre, filas, anchos, congelar: 3, combinar: ['A1:' + ultima + '1', 'A2:' + ultima + '2'], altos: { 0: 30, 1: 20, 2: 32 },
      filtro: 'A3:' + ultima + Math.max(3, filas.length), ocultas: op.ocultas, validaciones: op.validaciones, horizontal: op.horizontal !== false,
      estilo: (r, c) => r === 0 ? 4 : r === 1 ? 17 : r === 2 ? (edit.has(c) ? 7 : 11)
        : c === op.resultado ? res(r) : num.has(c) ? (edit.has(c) ? 16 : 15) : edit.has(c) ? 9 : 8 };
  }
  // Nombre de archivo seguro: sin tildes ni símbolos
  const archivo = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
  async function bajar(hojas, nombre) {
    const blob = await crear(hojas);
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = archivo(nombre) + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  window.Excel = { crear, leer, col, tabla, bajar, archivo };
})();
