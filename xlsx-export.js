// Generador de .xlsx con el formato de la plantilla Morph.
// Sin dependencias: escribe el ZIP (método store) a mano.

const crcTable = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = crcTable[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function zip(files) {
  const enc = new TextEncoder();
  const parts = [];
  const central = [];
  let offset = 0;
  files.forEach(f => {
    const name = enc.encode(f.name);
    const data = enc.encode(f.data);
    const crc = crc32(data);
    const local = new Uint8Array(30 + name.length);
    const dv = new DataView(local.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 0, true);
    dv.setUint16(8, 0, true);
    dv.setUint16(10, 0, true);
    dv.setUint16(12, 0, true);
    dv.setUint32(14, crc, true);
    dv.setUint32(18, data.length, true);
    dv.setUint32(22, data.length, true);
    dv.setUint16(26, name.length, true);
    dv.setUint16(28, 0, true);
    local.set(name, 30);
    parts.push(local, data);

    const cd = new Uint8Array(46 + name.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, name.length, true);
    cv.setUint32(42, offset, true);
    cd.set(name, 46);
    central.push(cd);
    offset += local.length + data.length;
  });
  let cdSize = 0;
  central.forEach(c => { cdSize += c.length; });
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, central.length, true);
  ev.setUint16(10, central.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, offset, true);
  return new Blob(parts.concat(central, [end]), {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

const esc = s => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// Estilos heredados de la plantilla:
// 1 título barra gris · 4/5 cabeceras · 6/7/8 banda de sección
// 9/10 fila de datos · 11 fila vacía · 12 separador · 13..16 bloque de totales
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/></font><font><color rgb="FF000000"/><name val="Arial"/></font><font/></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="lightGray"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCCCCCC"/><bgColor rgb="FFCCCCCC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFEFEF"/><bgColor rgb="FFEFEFEF"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF3F3F3"/><bgColor rgb="FFF3F3F3"/></patternFill></fill></fills><borders count="5"><border/><border><left style="thin"><color rgb="FF000000"/></left><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border><border><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border><border><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border></borders><cellStyleXfs count="1"><xf borderId="0" fillId="0" fontId="0" numFmtId="0" applyAlignment="1" applyFont="1"/></cellStyleXfs><cellXfs count="17"><xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="bottom"/></xf><xf borderId="1" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" wrapText="1"/></xf><xf borderId="2" fillId="0" fontId="2" numFmtId="0" xfId="0" applyBorder="1" applyFont="1"/><xf borderId="3" fillId="0" fontId="2" numFmtId="0" xfId="0" applyBorder="1" applyFont="1"/><xf borderId="4" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="1" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="4" fillId="3" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="4" fillId="4" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="1" fillId="4" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="4" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="1" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="4" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment wrapText="1"/></xf><xf borderId="4" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" wrapText="1"/></xf><xf borderId="4" fillId="3" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" wrapText="1"/></xf><xf borderId="4" fillId="4" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" wrapText="1"/></xf><xf borderId="1" fillId="4" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle xfId="0" name="Normal" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>`;

/**
 * data = {
 *   title:   'P20808-01-Punta Cana',
 *   renders: [{ name, desc, qty, hours }],
 *   extras:  [{ name, desc, qty, hours }],
 *   totals:  { renders, arquitectura, assets, adicionales }
 * }
 */
export function buildRenderXlsx(data) {
  const rows = [];
  const merges = [];
  let r = 14;

  const txt = (ref, s, v) => v == null || v === ''
    ? `<c r="${ref}" s="${s}"/>`
    : `<c r="${ref}" s="${s}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
  const num = (ref, s, v) => v == null || v === ''
    ? `<c r="${ref}" s="${s}"/>`
    : `<c r="${ref}" s="${s}"><v>${v}</v></c>`;

  // Barra de título
  rows.push(`<row r="${r}">${txt('E' + r, 1, data.title)}<c r="F${r}" s="2"/><c r="G${r}" s="2"/><c r="H${r}" s="2"/><c r="I${r}" s="2"/><c r="J${r}" s="3"/></row>`);
  merges.push(`E${r}:J${r}`);
  r++;

  // Cabecera de columnas
  rows.push(`<row r="${r}">${txt('E' + r, 4, 'Items')}${txt('F' + r, 4, 'Descripcion')}${txt('G' + r, 5, 'Cantidad')}<c r="H${r}" s="3"/>${txt('I' + r, 5, 'Horas')}<c r="J${r}" s="3"/></row>`);
  merges.push(`G${r}:H${r}`, `I${r}:J${r}`);
  r++;

  const band = label => {
    rows.push(`<row r="${r}">${txt('E' + r, 6, label)}<c r="F${r}" s="7"/><c r="G${r}" s="8"/><c r="H${r}" s="3"/><c r="I${r}" s="8"/><c r="J${r}" s="3"/></row>`);
    merges.push(`G${r}:H${r}`, `I${r}:J${r}`);
    r++;
  };
  const head = (a, b) => {
    rows.push(`<row r="${r}">${txt('E' + r, 9, a)}${txt('F' + r, 9, b)}<c r="G${r}" s="10"/><c r="H${r}" s="3"/><c r="I${r}" s="10"/><c r="J${r}" s="3"/></row>`);
    merges.push(`G${r}:H${r}`, `I${r}:J${r}`);
    r++;
  };
  const item = it => {
    rows.push(`<row r="${r}">${txt('E' + r, 11, it.name)}${txt('F' + r, 11, it.desc)}${num('G' + r, 10, it.qty)}<c r="H${r}" s="3"/>${num('I' + r, 10, it.hours)}<c r="J${r}" s="3"/></row>`);
    merges.push(`G${r}:H${r}`, `I${r}:J${r}`);
    r++;
  };
  const blank = () => {
    rows.push(`<row r="${r}"><c r="E${r}" s="11"/><c r="F${r}" s="11"/><c r="G${r}" s="10"/><c r="H${r}" s="3"/><c r="I${r}" s="10"/><c r="J${r}" s="3"/></row>`);
    merges.push(`G${r}:H${r}`, `I${r}:J${r}`);
    r++;
  };
  const spacer = () => {
    rows.push(`<row r="${r}"><c r="E${r}" s="12"/><c r="F${r}" s="12"/><c r="G${r}" s="12"/><c r="H${r}" s="12"/><c r="I${r}" s="12"/><c r="J${r}" s="12"/></row>`);
    r++;
  };

  band('Renders');
  head('Render Name', 'Short Description');
  (data.renders || []).forEach(item);
  blank();

  band('Arquitectura');
  head('Plano Name', 'Short Description');
  blank();

  band('Addicionales');
  head('Addicionales Names', 'Short Description');
  (data.extras || []).forEach(item);

  spacer();
  spacer();

  // Bloque de horas totales
  rows.push(`<row r="${r}">${txt('E' + r, 1, 'Horas Totales')}<c r="F${r}" s="2"/><c r="G${r}" s="2"/><c r="H${r}" s="2"/><c r="I${r}" s="2"/><c r="J${r}" s="3"/></row>`);
  merges.push(`E${r}:J${r}`);
  r++;
  rows.push(`<row r="${r}">${txt('E' + r, 13, 'Renders')}${txt('F' + r, 13, 'Arquitectura')}${txt('G' + r, 1, 'Assets')}<c r="H${r}" s="3"/>${txt('I' + r, 1, 'Addicionales')}<c r="J${r}" s="3"/></row>`);
  merges.push(`G${r}:H${r}`, `I${r}:J${r}`);
  r++;
  const t = data.totals || {};
  rows.push(`<row r="${r}">${num('E' + r, 14, t.renders || 0)}${num('F' + r, 15, t.arquitectura || 0)}${num('G' + r, 16, t.assets || 0)}<c r="H${r}" s="3"/>${num('I' + r, 16, t.adicionales || 0)}<c r="J${r}" s="3"/></row>`);
  merges.push(`G${r}:H${r}`, `I${r}:J${r}`);
  r++;
  spacer();

  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.75"/><cols><col customWidth="1" min="5" max="5" width="31.5"/><col customWidth="1" min="6" max="6" width="33.5"/></cols><sheetData>${rows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells></worksheet>`;

  return zip([
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Presupuesto" sheetId="1" r:id="rId1"/></sheets></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    { name: 'xl/styles.xml', data: STYLES },
    { name: 'xl/worksheets/sheet1.xml', data: sheet }
  ]);
}
