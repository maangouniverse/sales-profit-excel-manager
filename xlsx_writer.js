const fs = require('fs');
const zlib = require('zlib');

// CRC32 implementation for ZIP format
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) {
    c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
  }
  crcTable[n] = c >>> 0;
}

function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) {
    crc = crcTable[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function createZipBuffer(files) {
  const localHeaders = [];
  const centralHeaders = [];
  let offset = 0;

  for (const file of files) {
    const nameBuf = Buffer.from(file.name, 'utf8');
    const rawData = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, 'utf8');
    const crc = crc32(rawData);
    const compressed = zlib.deflateRawSync(rawData);

    const localHeader = Buffer.alloc(30 + nameBuf.length);
    localHeader.writeUInt32LE(0x04034b50, 0); // signature
    localHeader.writeUInt16LE(20, 4);         // version needed
    localHeader.writeUInt16LE(0x0800, 6);     // UTF-8 flag
    localHeader.writeUInt16LE(8, 8);          // compression: deflate
    localHeader.writeUInt16LE(0, 10);         // mod time
    localHeader.writeUInt16LE(0x5521, 12);    // mod date
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(compressed.length, 18);
    localHeader.writeUInt32LE(rawData.length, 22);
    localHeader.writeUInt16LE(nameBuf.length, 26);
    localHeader.writeUInt16LE(0, 28);
    nameBuf.copy(localHeader, 30);

    localHeaders.push(localHeader, compressed);

    const centralHeader = Buffer.alloc(46 + nameBuf.length);
    centralHeader.writeUInt32LE(0x02014b50, 0); // central dir signature
    centralHeader.writeUInt16LE(20, 4);         // version made by
    centralHeader.writeUInt16LE(20, 6);         // version needed
    centralHeader.writeUInt16LE(0x0800, 8);     // UTF-8 flag
    centralHeader.writeUInt16LE(8, 10);         // compression
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0x5521, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(compressed.length, 20);
    centralHeader.writeUInt32LE(rawData.length, 24);
    centralHeader.writeUInt16LE(nameBuf.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    nameBuf.copy(centralHeader, 46);

    centralHeaders.push(centralHeader);
    offset += localHeader.length + compressed.length;
  }

  const centralSize = centralHeaders.reduce((acc, b) => acc + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localHeaders, ...centralHeaders, eocd]);
}

function escXml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function colToLetter(col) {
  let temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function dateToSerial(yyyyMmDd) {
  const [y, m, d] = yyyyMmDd.split('-').map(Number);
  const utc = Date.UTC(y, m - 1, d);
  const epoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - epoch) / 86400000);
}

class SheetBuilder {
  constructor(name, options = {}) {
    this.name = name;
    this.rows = new Map();
    this.cols = options.cols || [];
    this.merges = [];
    this.validations = [];
    this.condFormats = [];
    this.freezeRow = options.freezeRow || 0;
    this.freezeCol = options.freezeCol || 0;
    this.drawingRelId = null;
    this.rowHeights = new Map();
  }

  setRowHeight(r, h) {
    this.rowHeights.set(r, h);
  }

  setCell(r, c, cellObj) {
    if (!this.rows.has(r)) this.rows.set(r, new Map());
    this.rows.get(r).set(c, cellObj);
  }

  merge(ref) {
    this.merges.push(ref);
  }

  addValidation(v) {
    this.validations.push(v);
  }

  addCondFormat(cf) {
    this.condFormats.push(cf);
  }

  toXml() {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
    xml += `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;

    if (this.freezeRow > 0 || this.freezeCol > 0) {
      const topLeft = `${colToLetter(this.freezeCol + 1)}${this.freezeRow + 1}`;
      xml += `<sheetViews><sheetView workbookViewId="0" showGridLines="1">`;
      xml += `<pane ${this.freezeCol > 0 ? `xSplit="${this.freezeCol}" ` : ''}${this.freezeRow > 0 ? `ySplit="${this.freezeRow}" ` : ''}topLeftCell="${topLeft}" activePane="bottomLeft" state="frozen"/>`;
      xml += `</sheetView></sheetViews>`;
    } else {
      xml += `<sheetViews><sheetView workbookViewId="0" showGridLines="1"/></sheetViews>`;
    }

    xml += `<sheetFormatPr defaultRowHeight="20"/>`;

    if (this.cols.length > 0) {
      xml += `<cols>`;
      for (const col of this.cols) {
        xml += `<col min="${col.min}" max="${col.max}" width="${col.width}" customWidth="1"/>`;
      }
      xml += `</cols>`;
    }

    xml += `<sheetData>`;
    const sortedRows = Array.from(this.rows.keys()).sort((a, b) => a - b);
    for (const r of sortedRows) {
      const h = this.rowHeights.get(r);
      xml += `<row r="${r}"${h ? ` ht="${h}" customHeight="1"` : ''}>`;
      const rowMap = this.rows.get(r);
      const sortedCols = Array.from(rowMap.keys()).sort((a, b) => a - b);
      for (const c of sortedCols) {
        const cell = rowMap.get(c);
        const ref = `${colToLetter(c)}${r}`;
        const sAttr = cell.s !== undefined ? ` s="${cell.s}"` : '';

        if (cell.f !== undefined) {
          const isStrResult = typeof cell.v === 'string';
          const tAttr = isStrResult ? ` t="str"` : '';
          xml += `<c r="${ref}"${sAttr}${tAttr}><f>${escXml(cell.f)}</f>`;
          if (cell.v !== undefined && cell.v !== null && cell.v !== '') {
            xml += `<v>${escXml(cell.v)}</v>`;
          }
          xml += `</c>`;
        } else if (cell.v === undefined || cell.v === null || cell.v === '') {
          if (cell.s !== undefined) {
            xml += `<c r="${ref}"${sAttr}/>`;
          }
        } else if (typeof cell.v === 'number') {
          xml += `<c r="${ref}"${sAttr}><v>${cell.v}</v></c>`;
        } else {
          xml += `<c r="${ref}"${sAttr} t="inlineStr"><is><t>${escXml(cell.v)}</t></is></c>`;
        }
      }
      xml += `</row>`;
    }
    xml += `</sheetData>`;

    if (this.merges.length > 0) {
      xml += `<mergeCells count="${this.merges.length}">`;
      for (const m of this.merges) {
        xml += `<mergeCell ref="${m}"/>`;
      }
      xml += `</mergeCells>`;
    }

    let priority = 1;
    for (const cf of this.condFormats) {
      xml += `<conditionalFormatting sqref="${cf.sqref}">`;
      for (const rule of cf.rules) {
        if (rule.type === 'cellIs') {
          xml += `<cfRule type="cellIs" dxfId="${rule.dxfId}" priority="${priority++}" operator="${rule.operator}"><formula>${escXml(rule.formula)}</formula></cfRule>`;
        } else if (rule.type === 'containsText') {
          xml += `<cfRule type="containsText" dxfId="${rule.dxfId}" priority="${priority++}" operator="containsText" text="${escXml(rule.text)}"><formula>NOT(ISERROR(SEARCH("${escXml(rule.text)}",${rule.topLeftCell})))</formula></cfRule>`;
        }
      }
      xml += `</conditionalFormatting>`;
    }

    if (this.validations.length > 0) {
      xml += `<dataValidations count="${this.validations.length}">`;
      for (const v of this.validations) {
        const allowBlank = v.allowBlank !== false ? '1' : '0';
        const showInput = v.promptTitle ? ` showInputMessage="1" promptTitle="${escXml(v.promptTitle)}" prompt="${escXml(v.prompt)}"` : ' showInputMessage="1"';
        const showError = v.errorTitle ? ` showErrorMessage="1" errorStyle="stop" errorTitle="${escXml(v.errorTitle)}" error="${escXml(v.error)}"` : ' showErrorMessage="1"';
        const opAttr = v.operator ? ` operator="${v.operator}"` : '';
        xml += `<dataValidation type="${v.type}"${opAttr} allowBlank="${allowBlank}"${showInput}${showError} sqref="${v.sqref}">`;
        xml += `<formula1>${escXml(v.formula1)}</formula1>`;
        if (v.formula2 !== undefined) {
          xml += `<formula2>${escXml(v.formula2)}</formula2>`;
        }
        xml += `</dataValidation>`;
      }
      xml += `</dataValidations>`;
    }

    if (this.drawingRelId) {
      xml += `<drawing r:id="${this.drawingRelId}"/>`;
    }

    xml += `</worksheet>`;
    return xml;
  }
}

module.exports = {
  createZipBuffer,
  escXml,
  colToLetter,
  dateToSerial,
  SheetBuilder
};
