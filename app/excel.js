/* excel.js · lectura y escritura mínima de .xlsx (con JSZip). Sin dependencias externas.
 * Excel.crear(hojas) → Blob. hojas: [{ nombre, filas: [[valor,...]], anchos: [n], oculta, validaciones: [{ rango, lista }], soloLectura: [col] }]
 * Excel.leer(archivo) → { NombreHoja: [[valor,...]] }  (valores: texto o número)
 */
(function () {
  const escXml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const col = n => { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };

  function hojaXml(h) {
    const filas = h.filas || [], lect = new Set(h.soloLectura || []);
    const sd = filas.map((fila, r) => '<row r="' + (r + 1) + '">' + fila.map((v, c) => {
      if (v === null || v === undefined || v === '') return '';
      const ref = col(c) + (r + 1), s = r === 0 ? 1 : (lect.has(c) ? 2 : 3);
      if (typeof v === 'number' && Number.isFinite(v)) return '<c r="' + ref + '" s="' + s + '"><v>' + v + '</v></c>';
      return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + escXml(v) + '</t></is></c>';
    }).join('') + '</row>').join('');
    const anchos = (h.anchos || []).map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('');
    const val = (h.validaciones || []).map(v => '<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Valor no permitido" error="Elija una opción de la lista." sqref="' +
      v.rango + '"><formula1>' + escXml(v.lista) + '</formula1></dataValidation>').join('');
    const ncols = Math.max(1, ...filas.map(f => f.length));
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"' + (h.oculta ? '' : '') + '><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' + (anchos ? '<cols>' + anchos + '</cols>' : '') + '<sheetData>' + sd + '</sheetData>' +
      (filas.length > 1 && !h.oculta ? '<autoFilter ref="A1:' + col(ncols - 1) + filas.length + '"/>' : '') +
      (val ? '<dataValidations count="' + h.validaciones.length + '">' + val + '</dataValidations>' : '') + '</worksheet>';
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
    const filtros = hojas.map((h, i) => (h.filas || []).length > 1 && !h.oculta
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
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
      '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF9063CD"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEFEDEC"/><bgColor indexed="64"/></patternFill></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
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
  window.Excel = { crear, leer, col };
})();
