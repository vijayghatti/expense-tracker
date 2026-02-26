/**
 * MiniXLSX - Lightweight Excel read/write library
 * Export: true XLSX (ZIP of XML) + CSV fallback
 * Import: CSV, TSV parsing
 * Zero dependencies
 */
const MiniXLSX = (() => {

  /* =========================================
   *  XLSX BINARY EXPORT (true .xlsx format)
   * ========================================= */

  function exportToExcel(rows, sheetName = 'Sheet1', filename = 'export.xlsx') {
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);

    // Build shared strings table
    const sharedStrings = [];
    const ssMap = {};
    function getSSIndex(str) {
      str = String(str);
      if (ssMap[str] !== undefined) return ssMap[str];
      ssMap[str] = sharedStrings.length;
      sharedStrings.push(str);
      return ssMap[str];
    }

    // Pre-populate shared strings with all string values
    headers.forEach(h => getSSIndex(h));
    rows.forEach(row => {
      headers.forEach(h => {
        const val = row[h];
        if (val !== null && val !== undefined && typeof val !== 'number') {
          getSSIndex(String(val));
        }
      });
    });

    // Build sheet XML
    let sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    sheetXml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    sheetXml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n';
    sheetXml += '<sheetData>\n';

    // Header row (row 1)
    sheetXml += '<row r="1">\n';
    headers.forEach((h, ci) => {
      const col = colLetter(ci);
      sheetXml += `<c r="${col}1" t="s" s="1"><v>${getSSIndex(h)}</v></c>\n`;
    });
    sheetXml += '</row>\n';

    // Data rows
    rows.forEach((row, ri) => {
      const rowNum = ri + 2;
      sheetXml += `<row r="${rowNum}">\n`;
      headers.forEach((h, ci) => {
        const col = colLetter(ci);
        const ref = `${col}${rowNum}`;
        const val = row[h];
        if (val === null || val === undefined || val === '') {
          // skip empty
        } else if (typeof val === 'number') {
          sheetXml += `<c r="${ref}" s="2"><v>${val}</v></c>\n`;
        } else {
          sheetXml += `<c r="${ref}" t="s"><v>${getSSIndex(String(val))}</v></c>\n`;
        }
      });
      sheetXml += '</row>\n';
    });

    sheetXml += '</sheetData>\n</worksheet>';

    // Shared strings XML
    let ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    ssXml += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">\n`;
    sharedStrings.forEach(s => {
      ssXml += `<si><t>${escapeXml(s)}</t></si>\n`;
    });
    ssXml += '</sst>';

    // Styles XML (minimal)
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF4CAF50"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>`;

    // Workbook XML
    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

    // Relationships
    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;

    // Create ZIP
    const files = [
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'xl/workbook.xml', data: workbookXml },
      { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
      { name: 'xl/worksheets/sheet1.xml', data: sheetXml },
      { name: 'xl/styles.xml', data: stylesXml },
      { name: 'xl/sharedStrings.xml', data: ssXml }
    ];

    const zipBlob = createZip(files);
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename.endsWith('.xlsx') ? filename : filename.replace(/\.[^.]+$/, '.xlsx');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /* =========================================
   *  MINIMAL ZIP CREATOR (store-only, no compression)
   * ========================================= */

  function createZip(files) {
    const encoder = new TextEncoder();
    const parts = [];
    const centralDir = [];
    let offset = 0;

    files.forEach(file => {
      const nameBytes = encoder.encode(file.name);
      const dataBytes = encoder.encode(file.data);
      const crc = crc32(dataBytes);

      // Local file header (30 + name + data)
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(localHeader.buffer);
      lv.setUint32(0, 0x04034b50, true);  // signature
      lv.setUint16(4, 20, true);           // version needed
      lv.setUint16(6, 0, true);            // flags
      lv.setUint16(8, 0, true);            // compression (store)
      lv.setUint16(10, 0, true);           // mod time
      lv.setUint16(12, 0, true);           // mod date
      lv.setUint32(14, crc, true);         // crc32
      lv.setUint32(18, dataBytes.length, true); // compressed size
      lv.setUint32(22, dataBytes.length, true); // uncompressed size
      lv.setUint16(26, nameBytes.length, true); // filename length
      lv.setUint16(28, 0, true);           // extra field length
      localHeader.set(nameBytes, 30);

      // Central directory entry (46 + name)
      const cdEntry = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(cdEntry.buffer);
      cv.setUint32(0, 0x02014b50, true);  // signature
      cv.setUint16(4, 20, true);           // version made by
      cv.setUint16(6, 20, true);           // version needed
      cv.setUint16(8, 0, true);            // flags
      cv.setUint16(10, 0, true);           // compression
      cv.setUint16(12, 0, true);           // mod time
      cv.setUint16(14, 0, true);           // mod date
      cv.setUint32(16, crc, true);         // crc32
      cv.setUint32(20, dataBytes.length, true); // compressed size
      cv.setUint32(24, dataBytes.length, true); // uncompressed size
      cv.setUint16(28, nameBytes.length, true); // filename length
      cv.setUint16(30, 0, true);           // extra field length
      cv.setUint16(32, 0, true);           // comment length
      cv.setUint16(34, 0, true);           // disk start
      cv.setUint16(36, 0, true);           // internal attrs
      cv.setUint32(38, 0, true);           // external attrs
      cv.setUint32(42, offset, true);      // local header offset
      cdEntry.set(nameBytes, 46);

      parts.push(localHeader);
      parts.push(dataBytes);
      centralDir.push(cdEntry);

      offset += localHeader.length + dataBytes.length;
    });

    // End of central directory
    const cdOffset = offset;
    let cdSize = 0;
    centralDir.forEach(cd => { parts.push(cd); cdSize += cd.length; });

    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);      // signature
    ev.setUint16(4, 0, true);                // disk number
    ev.setUint16(6, 0, true);                // disk with CD
    ev.setUint16(8, files.length, true);     // entries on disk
    ev.setUint16(10, files.length, true);    // total entries
    ev.setUint32(12, cdSize, true);          // CD size
    ev.setUint32(16, cdOffset, true);        // CD offset
    ev.setUint16(20, 0, true);              // comment length
    parts.push(eocd);

    return new Blob(parts, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // CRC32 implementation
  const crc32Table = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[i] = c;
    }
    return table;
  })();

  function crc32(data) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++) {
      crc = crc32Table[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function colLetter(index) {
    let s = '';
    index++;
    while (index > 0) {
      index--;
      s = String.fromCharCode(65 + (index % 26)) + s;
      index = Math.floor(index / 26);
    }
    return s;
  }

  /* =========================================
   *  CSV EXPORT
   * ========================================= */

  function exportToCSV(rows, filename = 'export.csv') {
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);
    // BOM for Excel UTF-8 recognition
    let csv = '\uFEFF';
    csv += headers.map(h => csvEscape(h)).join(',') + '\n';
    rows.forEach(row => {
      csv += headers.map(h => csvEscape(row[h] ?? '')).join(',') + '\n';
    });
    downloadBlob(csv, filename, 'text/csv;charset=utf-8');
  }

  /* =========================================
   *  IMPORT (CSV, TSV, XLSX)
   * ========================================= */

  function importFile(file) {
    return new Promise((resolve, reject) => {
      if (file.name.match(/\.xlsx$/i)) {
        // Read as binary for XLSX
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            resolve(parseXLSX(new Uint8Array(e.target.result)));
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
      } else {
        // Read as text for CSV/TSV/XLS
        const reader = new FileReader();
        reader.onload = (e) => {
          const text = e.target.result;
          try {
            if (file.name.endsWith('.tsv')) {
              resolve(parseCSV(text, '\t'));
            } else {
              resolve(parseCSV(text, ','));
            }
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsText(file);
      }
    });
  }

  /* =========================================
   *  MINIMAL XLSX READER (unzip + parse XML)
   * ========================================= */

  function parseXLSX(data) {
    const files = unzip(data);

    // Find shared strings
    let sharedStrings = [];
    const ssFile = files['xl/sharedStrings.xml'];
    if (ssFile) {
      const ssText = new TextDecoder().decode(ssFile);
      const parser = new DOMParser();
      const doc = parser.parseFromString(ssText, 'text/xml');
      const sis = doc.getElementsByTagName('si');
      for (let i = 0; i < sis.length; i++) {
        // Handle both <t> and <r><t> structures
        const tEls = sis[i].getElementsByTagName('t');
        let str = '';
        for (let j = 0; j < tEls.length; j++) str += tEls[j].textContent;
        sharedStrings.push(str);
      }
    }

    // Find sheet1
    const sheetFile = files['xl/worksheets/sheet1.xml'];
    if (!sheetFile) return [];
    const sheetText = new TextDecoder().decode(sheetFile);
    const parser = new DOMParser();
    const doc = parser.parseFromString(sheetText, 'text/xml');
    const rowEls = doc.getElementsByTagName('row');

    const allRows = [];
    for (let i = 0; i < rowEls.length; i++) {
      const cells = rowEls[i].getElementsByTagName('c');
      const row = [];
      let maxCol = 0;
      for (let j = 0; j < cells.length; j++) {
        const ref = cells[j].getAttribute('r') || '';
        const colIdx = refToCol(ref);
        const type = cells[j].getAttribute('t');
        const vEl = cells[j].getElementsByTagName('v')[0];
        let val = vEl ? vEl.textContent : '';

        if (type === 's' && sharedStrings[parseInt(val)] !== undefined) {
          val = sharedStrings[parseInt(val)];
        } else if (type !== 's' && val !== '' && !isNaN(Number(val))) {
          val = Number(val);
        }

        // Pad row to fill any gaps
        while (row.length < colIdx) row.push('');
        row[colIdx] = val;
        maxCol = Math.max(maxCol, colIdx);
      }
      allRows.push(row);
    }

    if (allRows.length < 2) return [];
    const headers = allRows[0].map((h, i) => String(h || `Column${i + 1}`));
    return allRows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] !== undefined ? row[i] : '';
      });
      return obj;
    });
  }

  function refToCol(ref) {
    const match = ref.match(/^([A-Z]+)/);
    if (!match) return 0;
    const letters = match[1];
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
      col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return col - 1;
  }

  /* =========================================
   *  MINIMAL ZIP READER (store + deflate)
   * ========================================= */

  function unzip(data) {
    const view = new DataView(data.buffer);
    const files = {};
    let offset = 0;

    while (offset < data.length - 4) {
      const sig = view.getUint32(offset, true);
      if (sig !== 0x04034b50) break; // Not a local file header

      const compression = view.getUint16(offset + 8, true);
      const compSize = view.getUint32(offset + 18, true);
      const uncompSize = view.getUint32(offset + 22, true);
      const nameLen = view.getUint16(offset + 26, true);
      const extraLen = view.getUint16(offset + 28, true);
      const nameBytes = data.slice(offset + 30, offset + 30 + nameLen);
      const name = new TextDecoder().decode(nameBytes);
      const dataStart = offset + 30 + nameLen + extraLen;
      const fileData = data.slice(dataStart, dataStart + compSize);

      if (compression === 0) {
        // Stored (no compression)
        files[name] = fileData;
      } else if (compression === 8) {
        // Deflated - use DecompressionStream
        try {
          files[name] = inflateSync(fileData);
        } catch (e) {
          // Skip files we can't decompress
          files[name] = fileData;
        }
      }

      offset = dataStart + compSize;
    }

    return files;
  }

  // Simple inflate using DecompressionStream (async workaround with sync fallback)
  function inflateSync(data) {
    // Try raw inflate using a minimal implementation
    // For browser compat, we'll use a synchronous tiny inflate
    return tinyInflate(data);
  }

  // Minimal DEFLATE decompressor
  function tinyInflate(data) {
    let pos = 0;
    let outBuf = [];

    function readBit() {
      if (pos >= data.length * 8) return 0;
      const byte = data[pos >> 3];
      const bit = (byte >> (pos & 7)) & 1;
      pos++;
      return bit;
    }

    function readBits(n) {
      let val = 0;
      for (let i = 0; i < n; i++) val |= readBit() << i;
      return val;
    }

    function buildHuffman(lengths) {
      const maxLen = Math.max(...lengths.filter(l => l > 0), 1);
      const blCount = new Array(maxLen + 1).fill(0);
      lengths.forEach(l => { if (l > 0) blCount[l]++; });
      const nextCode = new Array(maxLen + 1).fill(0);
      let code = 0;
      for (let bits = 1; bits <= maxLen; bits++) {
        code = (code + blCount[bits - 1]) << 1;
        nextCode[bits] = code;
      }
      const table = {};
      lengths.forEach((len, i) => {
        if (len > 0) {
          const c = nextCode[len]++;
          const key = c.toString(2).padStart(len, '0');
          table[key] = i;
        }
      });
      return table;
    }

    function decodeSymbol(table) {
      let code = '';
      for (let i = 0; i < 20; i++) {
        code += readBit();
        if (table[code] !== undefined) return table[code];
      }
      throw new Error('Invalid huffman code');
    }

    // Fixed Huffman tables
    function getFixedLitTable() {
      const lengths = new Array(288);
      for (let i = 0; i <= 143; i++) lengths[i] = 8;
      for (let i = 144; i <= 255; i++) lengths[i] = 9;
      for (let i = 256; i <= 279; i++) lengths[i] = 7;
      for (let i = 280; i <= 287; i++) lengths[i] = 8;
      return buildHuffman(lengths);
    }

    function getFixedDistTable() {
      const lengths = new Array(32).fill(5);
      return buildHuffman(lengths);
    }

    const lenBase = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
    const lenExtra = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
    const distBase = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
    const distExtra = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];

    function inflateBlock(litTable, distTable) {
      while (true) {
        const sym = decodeSymbol(litTable);
        if (sym < 256) {
          outBuf.push(sym);
        } else if (sym === 256) {
          break;
        } else {
          const lenIdx = sym - 257;
          const length = lenBase[lenIdx] + readBits(lenExtra[lenIdx]);
          const distSym = decodeSymbol(distTable);
          const distance = distBase[distSym] + readBits(distExtra[distSym]);
          const start = outBuf.length - distance;
          for (let i = 0; i < length; i++) {
            outBuf.push(outBuf[start + i]);
          }
        }
      }
    }

    let bfinal;
    do {
      bfinal = readBit();
      const btype = readBits(2);

      if (btype === 0) {
        // Stored
        pos = (pos + 7) & ~7; // align to byte
        const len = readBits(16);
        readBits(16); // nlen
        for (let i = 0; i < len; i++) outBuf.push(readBits(8));
      } else if (btype === 1) {
        // Fixed Huffman
        inflateBlock(getFixedLitTable(), getFixedDistTable());
      } else if (btype === 2) {
        // Dynamic Huffman
        const hlit = readBits(5) + 257;
        const hdist = readBits(5) + 1;
        const hclen = readBits(4) + 4;
        const clOrder = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
        const clLengths = new Array(19).fill(0);
        for (let i = 0; i < hclen; i++) clLengths[clOrder[i]] = readBits(3);
        const clTable = buildHuffman(clLengths);

        const allLengths = [];
        while (allLengths.length < hlit + hdist) {
          const sym = decodeSymbol(clTable);
          if (sym < 16) {
            allLengths.push(sym);
          } else if (sym === 16) {
            const rep = readBits(2) + 3;
            const prev = allLengths[allLengths.length - 1] || 0;
            for (let i = 0; i < rep; i++) allLengths.push(prev);
          } else if (sym === 17) {
            const rep = readBits(3) + 3;
            for (let i = 0; i < rep; i++) allLengths.push(0);
          } else if (sym === 18) {
            const rep = readBits(7) + 11;
            for (let i = 0; i < rep; i++) allLengths.push(0);
          }
        }

        const litLengths = allLengths.slice(0, hlit);
        const distLengths = allLengths.slice(hlit, hlit + hdist);
        inflateBlock(buildHuffman(litLengths), buildHuffman(distLengths));
      }
    } while (!bfinal);

    return new Uint8Array(outBuf);
  }

  /* =========================================
   *  CSV PARSING
   * ========================================= */

  function parseCSV(text, separator = ',') {
    // Remove BOM
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) return [];
    const headers = parseCSVLine(lines[0], separator);
    return lines.slice(1).map(line => {
      const vals = parseCSVLine(line, separator);
      const obj = {};
      headers.forEach((h, i) => {
        let v = vals[i] || '';
        const num = Number(v);
        if (v !== '' && !isNaN(num) && isFinite(num)) v = num;
        obj[h.trim()] = v;
      });
      return obj;
    });
  }

  function parseCSVLine(line, sep) {
    const result = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (inQuotes) {
        if (c === '"' && line[i + 1] === '"') { current += '"'; i++; }
        else if (c === '"') inQuotes = false;
        else current += c;
      } else {
        if (c === '"') inQuotes = true;
        else if (c === sep) { result.push(current); current = ''; }
        else current += c;
      }
    }
    result.push(current);
    return result;
  }

  /* =========================================
   *  HELPERS
   * ========================================= */

  function csvEscape(val) {
    const s = String(val);
    if (s.includes(',') || s.includes('"') || s.includes('\n')) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function downloadBlob(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /* =========================================
   *  XLSX EXPORT WITH CHARTS
   * ========================================= */

  function exportToExcelWithCharts(rows, catData, monthData, filename) {
    if (!rows.length) return;

    const headers = Object.keys(rows[0]);
    const sharedStrings = [];
    const ssMap = {};

    function getSSIndex(str) {
      str = String(str);
      if (ssMap[str] !== undefined) return ssMap[str];
      ssMap[str] = sharedStrings.length;
      sharedStrings.push(str);
      return ssMap[str];
    }

    // Pre-populate shared strings
    headers.forEach(h => getSSIndex(h));
    rows.forEach(row => {
      headers.forEach(h => {
        const val = row[h];
        if (val !== null && val !== undefined && typeof val !== 'number') getSSIndex(String(val));
      });
    });
    ['Category', 'Amount', 'Month'].forEach(s => getSSIndex(s));
    catData.forEach(c => getSSIndex(c.name));
    monthData.forEach(m => getSSIndex(m.month));

    // ---- Sheet 1: Expenses ----
    let sheet1Xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    sheet1Xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    sheet1Xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n';
    sheet1Xml += '<sheetData>\n';
    sheet1Xml += '<row r="1">\n';
    headers.forEach((h, ci) => {
      sheet1Xml += `<c r="${colLetter(ci)}1" t="s" s="1"><v>${getSSIndex(h)}</v></c>\n`;
    });
    sheet1Xml += '</row>\n';
    rows.forEach((row, ri) => {
      const rowNum = ri + 2;
      sheet1Xml += `<row r="${rowNum}">\n`;
      headers.forEach((h, ci) => {
        const ref = `${colLetter(ci)}${rowNum}`;
        const val = row[h];
        if (val === null || val === undefined || val === '') { /* skip */ }
        else if (typeof val === 'number') sheet1Xml += `<c r="${ref}" s="2"><v>${val}</v></c>\n`;
        else sheet1Xml += `<c r="${ref}" t="s"><v>${getSSIndex(String(val))}</v></c>\n`;
      });
      sheet1Xml += '</row>\n';
    });
    sheet1Xml += '</sheetData>\n</worksheet>';

    // ---- Sheet 2: Summary + Charts ----
    const nCats = catData.length;
    const nMonths = monthData.length;
    const maxRows = Math.max(nCats, nMonths, 1);
    const chartStartRow0 = maxRows + 3; // 0-indexed row for drawing anchor

    let sheet2Xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    sheet2Xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    sheet2Xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n';
    sheet2Xml += '<sheetData>\n';
    sheet2Xml += '<row r="1">\n';
    sheet2Xml += `<c r="A1" t="s" s="1"><v>${getSSIndex('Category')}</v></c>\n`;
    sheet2Xml += `<c r="B1" t="s" s="1"><v>${getSSIndex('Amount')}</v></c>\n`;
    sheet2Xml += `<c r="D1" t="s" s="1"><v>${getSSIndex('Month')}</v></c>\n`;
    sheet2Xml += `<c r="E1" t="s" s="1"><v>${getSSIndex('Amount')}</v></c>\n`;
    sheet2Xml += '</row>\n';
    for (let i = 0; i < maxRows; i++) {
      const r = i + 2;
      sheet2Xml += `<row r="${r}">\n`;
      if (i < nCats) {
        sheet2Xml += `<c r="A${r}" t="s"><v>${getSSIndex(catData[i].name)}</v></c>\n`;
        sheet2Xml += `<c r="B${r}" s="2"><v>${catData[i].amount}</v></c>\n`;
      }
      if (i < nMonths) {
        sheet2Xml += `<c r="D${r}" t="s"><v>${getSSIndex(monthData[i].month)}</v></c>\n`;
        sheet2Xml += `<c r="E${r}" s="2"><v>${monthData[i].amount}</v></c>\n`;
      }
      sheet2Xml += '</row>\n';
    }
    sheet2Xml += '</sheetData>\n';
    sheet2Xml += '<drawing r:id="rId1"/>\n';
    sheet2Xml += '</worksheet>';

    // ---- Shared Strings ----
    let ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    ssXml += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">\n`;
    sharedStrings.forEach(s => { ssXml += `<si><t>${escapeXml(s)}</t></si>\n`; });
    ssXml += '</sst>';

    // ---- Styles ----
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF4CAF50"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>`;

    // ---- Workbook ----
    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Expenses" sheetId="1" r:id="rId1"/>
<sheet name="Summary" sheetId="2" r:id="rId2"/>
</sheets>
</workbook>`;

    // ---- Relationships ----
    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    const sheet2Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

    const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`;

    // ---- Drawing: two charts side by side below the data ----
    const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor moveWithCells="1" sizeWithCells="1">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${chartStartRow0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${chartStartRow0 + 20}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Category Chart"/>
        <xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor moveWithCells="1" sizeWithCells="1">
    <xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${chartStartRow0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${chartStartRow0 + 20}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="3" name="Monthly Chart"/>
        <xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

    // ---- Chart 1: Bar chart (Spending by Category) ----
    const catLblPts = catData.map((c, i) => `<c:pt idx="${i}"><c:v>${escapeXml(c.name)}</c:v></c:pt>`).join('');
    const catValPts = catData.map((c, i) => `<c:pt idx="${i}"><c:v>${c.amount.toFixed(2)}</c:v></c:pt>`).join('');
    const chart1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="en-US"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Spending by Category</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat>
            <c:strRef>
              <c:f>Summary!$A$2:$A$${nCats + 1}</c:f>
              <c:strCache><c:ptCount val="${nCats}"/>${catLblPts}</c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Summary!$B$2:$B$${nCats + 1}</c:f>
              <c:numCache><c:formatCode>#,##0.00</c:formatCode><c:ptCount val="${nCats}"/>${catValPts}</c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1001"/><c:axId val="1002"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:numFmt formatCode="#,##0.00" sourceLinked="0"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;

    // ---- Chart 2: Line chart (Monthly Spending Trend) ----
    const monLblPts = monthData.map((m, i) => `<c:pt idx="${i}"><c:v>${escapeXml(m.month)}</c:v></c:pt>`).join('');
    const monValPts = monthData.map((m, i) => `<c:pt idx="${i}"><c:v>${m.amount.toFixed(2)}</c:v></c:pt>`).join('');
    const chart2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="en-US"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Monthly Spending Trend</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>
          <c:cat>
            <c:strRef>
              <c:f>Summary!$D$2:$D$${nMonths + 1}</c:f>
              <c:strCache><c:ptCount val="${nMonths}"/>${monLblPts}</c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Summary!$E$2:$E$${nMonths + 1}</c:f>
              <c:numCache><c:formatCode>#,##0.00</c:formatCode><c:ptCount val="${nMonths}"/>${monValPts}</c:numCache>
            </c:numRef>
          </c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:axId val="2001"/><c:axId val="2002"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="2001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="2002"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:numFmt formatCode="#,##0.00" sourceLinked="0"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="2001"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;

    // ---- Content Types ----
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`;

    // ---- Build ZIP ----
    const files = [
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'xl/workbook.xml', data: workbookXml },
      { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
      { name: 'xl/worksheets/sheet1.xml', data: sheet1Xml },
      { name: 'xl/worksheets/sheet2.xml', data: sheet2Xml },
      { name: 'xl/worksheets/_rels/sheet2.xml.rels', data: sheet2Rels },
      { name: 'xl/styles.xml', data: stylesXml },
      { name: 'xl/sharedStrings.xml', data: ssXml },
      { name: 'xl/drawings/drawing1.xml', data: drawingXml },
      { name: 'xl/drawings/_rels/drawing1.xml.rels', data: drawingRels },
      { name: 'xl/charts/chart1.xml', data: chart1Xml },
      { name: 'xl/charts/chart2.xml', data: chart2Xml }
    ];

    const zipBlob = createZip(files);
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename.endsWith('.xlsx') ? filename : filename.replace(/\.[^.]+$/, '') + '.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  return { exportToExcel, exportToExcelWithCharts, exportToCSV, importFile, parseCSV };
})();
