/*
 * Ekspor rekap penjualan per outlet ke Excel (.xlsx) dan PDF.
 * Dimuat hanya saat tombol Unduh ditekan. Excel dibuat sendiri (tanpa library); PDF memakai jsPDF + autoTable
 * dari folder vendor/ (MIT License).
 *
 * model = { monthLabel, owner, printedAt, outlets: [ { name, transactions, qty, sub, disc, tax, total,
 *           rows: [ { date: 'yyyy-MM-dd', dateLabel, product, qty, price, disc, tax, sub, discAmt, taxAmt, total, note } ] } ],
 *           grand: { transactions, qty, sub, disc, tax, total } }
 */
(function (global) {
  'use strict';

  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  function num(n) { return Number(n || 0); }
  function fmt(n) { return num(n).toLocaleString('id-ID'); }
  function pctText(n) { return n ? String(n).replace('.', ',') : '0'; }

  /* ------------------------------------------------------------------ ZIP tanpa kompresi */
  var crcTable = (function () {
    var t = [], c, n, k;
    for (n = 0; n < 256; n++) {
      c = n;
      for (k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = crcTable[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function zip(files, mime) {
    var enc = new TextEncoder();
    var parts = [], central = [], offset = 0;
    var d = new Date();
    var dosTime = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    var dosDate = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data);
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true);
      lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), name, data);
      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true);
      ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var csize = 0;
    central.forEach(function (p) { csize += p.length; });
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, csize, true); end.setUint32(16, offset, true);
    return new Blob(parts.concat(central, [new Uint8Array(end.buffer)]), { type: mime });
  }

  /* ------------------------------------------------------------------ XLSX */
  function colName(i) {
    var s = '';
    i++;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  function xe(s) {
    return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  // gaya sel (indeks di cellXfs pada styles.xml)
  var S = { title: 1, muted: 2, head: 3, text: 4, int: 5, date: 6, gen: 7, totLabel: 8, totInt: 9, totBlank: 10 };

  function cs(ref, text, s) { return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + xe(text) + '</t></is></c>'; }
  function cn(ref, n, s) { return '<c r="' + ref + '" s="' + s + '"><v>' + num(n) + '</v></c>'; }
  function cf(ref, formula, v, s) { return '<c r="' + ref + '" s="' + s + '"><f>' + xe(formula) + '</f><v>' + num(v) + '</v></c>'; }
  function cb(ref, s) { return '<c r="' + ref + '" s="' + s + '"/>'; }
  function serial(iso) {
    var p = iso.split('-');
    return Math.round(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])) / 86400000) + 25569;
  }

  function sheetXml(rows, widths, freezeAfterRow) {
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
      '<sheetViews><sheetView workbookViewId="0">' +
      (freezeAfterRow ? '<pane ySplit="' + freezeAfterRow + '" topLeftCell="A' + (freezeAfterRow + 1) + '" activePane="bottomLeft" state="frozen"/>' : '') +
      '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>' +
      widths.map(function (w, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('') +
      '</cols><sheetData>' +
      rows.map(function (r, i) { return r ? '<row r="' + (i + 1) + '">' + r.join('') + '</row>' : ''; }).join('') +
      '</sheetData><pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>' +
      '<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/></worksheet>';
    return xml;
  }

  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="165" formatCode="dd\\ mmm\\ yyyy"/></numFmts>' +
    '<fonts count="5">' +
    '<font><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="14"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
    '<font><i/><sz val="10"/><color rgb="FF667085"/><name val="Calibri"/></font></fonts>' +
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1A73E8"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE8F0FE"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FFBFC7D5"/></left><right style="thin"><color rgb="FFBFC7D5"/></right>' +
    '<top style="thin"><color rgb="FFBFC7D5"/></top><bottom style="thin"><color rgb="FFBFC7D5"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="11">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="3" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function sheetName(raw, used) {
    var base = String(raw || 'Outlet').replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '') || 'Outlet';
    base = base.slice(0, 31).trim();
    var name = base, i = 2;
    while (used[name.toLowerCase()]) {
      var suffix = ' (' + i++ + ')';
      name = base.slice(0, 31 - suffix.length) + suffix;
    }
    used[name.toLowerCase()] = true;
    return name;
  }

  var OUTLET_HEAD = ['Tanggal', 'Produk', 'Qty', 'Harga', 'Disc %', 'Pajak %', 'Subtotal', 'Diskon (Rp)', 'Pajak (Rp)', 'Total', 'Catatan'];

  function outletSheet(model, o) {
    var rows = [];
    rows[0] = [cs('A1', 'Rekap Penjualan - ' + o.name, S.title)];
    rows[1] = [cs('A2', 'Bulan: ' + model.monthLabel + (model.owner ? '   |   Sales: ' + model.owner : ''), S.muted)];
    rows[2] = [cs('A3', 'Dibuat: ' + model.printedAt + '   |   Nominal dalam Rupiah', S.muted)];
    rows[4] = OUTLET_HEAD.map(function (h, i) { return cs(colName(i) + '5', h, S.head); });
    var r = 6;
    o.rows.forEach(function (x) {
      var n = String(r);
      rows[r - 1] = [
        cn('A' + n, serial(x.date), S.date), cs('B' + n, x.product, S.text), cn('C' + n, x.qty, S.int),
        cn('D' + n, x.price, S.int), cn('E' + n, x.disc, S.gen), cn('F' + n, x.tax, S.gen),
        cn('G' + n, x.sub, S.int), cn('H' + n, x.discAmt, S.int), cn('I' + n, x.taxAmt, S.int),
        cn('J' + n, x.total, S.int), cs('K' + n, x.note || '', S.text)
      ];
      r++;
    });
    var first = 6, last = r - 1, t = String(r);
    function sum(col, v) { return cf(col + t, 'SUM(' + col + first + ':' + col + last + ')', v, S.totInt); }
    rows[r - 1] = [
      cs('A' + t, 'Total', S.totLabel), cb('B' + t, S.totBlank), sum('C', o.qty), cb('D' + t, S.totBlank),
      cb('E' + t, S.totBlank), cb('F' + t, S.totBlank), sum('G', o.sub), sum('H', o.disc), sum('I', o.tax),
      sum('J', o.total), cb('K' + t, S.totBlank)
    ];
    return sheetXml(rows, [13, 34, 8, 13, 9, 9, 15, 14, 14, 15, 28], 5);
  }

  function summarySheet(model) {
    var rows = [];
    rows[0] = [cs('A1', 'Rekap Penjualan per Outlet', S.title)];
    rows[1] = [cs('A2', 'Bulan: ' + model.monthLabel + (model.owner ? '   |   Sales: ' + model.owner : ''), S.muted)];
    rows[2] = [cs('A3', 'Dibuat: ' + model.printedAt + '   |   Nominal dalam Rupiah. Tiap outlet punya sheet sendiri.', S.muted)];
    var head = ['No', 'Outlet', 'Transaksi', 'Item', 'Subtotal', 'Diskon (Rp)', 'Pajak (Rp)', 'Total'];
    rows[4] = head.map(function (h, i) { return cs(colName(i) + '5', h, S.head); });
    var r = 6;
    model.outlets.forEach(function (o, i) {
      var n = String(r);
      rows[r - 1] = [cn('A' + n, i + 1, S.int), cs('B' + n, o.name, S.text), cn('C' + n, o.transactions, S.int), cn('D' + n, o.qty, S.int),
        cn('E' + n, o.sub, S.int), cn('F' + n, o.disc, S.int), cn('G' + n, o.tax, S.int), cn('H' + n, o.total, S.int)];
      r++;
    });
    var first = 6, last = r - 1, t = String(r), g = model.grand;
    function sum(col, v) { return cf(col + t, 'SUM(' + col + first + ':' + col + last + ')', v, S.totInt); }
    rows[r - 1] = [cs('A' + t, 'Total', S.totLabel), cb('B' + t, S.totBlank), sum('C', g.transactions), sum('D', g.qty),
      sum('E', g.sub), sum('F', g.disc), sum('G', g.tax), sum('H', g.total)];
    return sheetXml(rows, [6, 36, 11, 9, 15, 14, 14, 16], 5);
  }

  function book(sheets) {
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join('') + '</Types>';
    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>' +
      sheets.map(function (s, i) { return '<sheet name="' + xe(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join('') +
      '</sheets></workbook>';
    var wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function (s, i) {
        return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      }).join('') +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

    var files = [
      { name: '[Content_Types].xml', data: ct },
      { name: '_rels/.rels', data: rels },
      { name: 'xl/workbook.xml', data: wb },
      { name: 'xl/_rels/workbook.xml.rels', data: wbRels },
      { name: 'xl/styles.xml', data: STYLES }
    ];
    sheets.forEach(function (s, i) { files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: s.xml }); });
    return zip(files, XLSX_MIME);
  }

  function xlsx(model) {
    var used = {};
    var sheets = [{ name: sheetName('Ringkasan', used), xml: summarySheet(model) }];
    model.outlets.forEach(function (o) { sheets.push({ name: sheetName(o.name, used), xml: outletSheet(model, o) }); });
    return book(sheets);
  }

  /* ------------------------------------------------------------------ ESTIMASI */

  function estMeta(m) {
    return 'Rentang: ' + m.horizon + ' hari ke depan' + (m.owner ? '   |   Sales: ' + m.owner : '') + '   |   Dibuat: ' + m.printedAt;
  }
  var EST_NOTE = 'Perkiraan dari kebiasaan order lalu (' + 'jeda order tengah, rata-rata qty 3 order terakhir), bukan kepastian. Nilai hanya menghitung produk yang sudah waktunya atau jatuh tempo dalam rentang.';
  function estSummarySheet(m) {
    var rows = [];
    rows[0] = [cs('A1', 'Estimasi Kebutuhan Order per Outlet', S.title)];
    rows[1] = [cs('A2', estMeta(m), S.muted)];
    rows[2] = [cs('A3', EST_NOTE, S.muted)];
    var head = ['No', 'Outlet', 'Status', 'Order berikutnya', 'Produk', 'Perkiraan nilai (Rp)'];
    rows[4] = head.map(function (h, i) { return cs(colName(i) + '5', h, S.head); });
    var r = 6;
    m.outlets.forEach(function (o, i) {
      var n = String(r);
      rows[r - 1] = [cn('A' + n, i + 1, S.int), cs('B' + n, o.name, S.text), cs('C' + n, o.statusText, S.text),
        o.nextDate ? cn('D' + n, serial(o.nextDate), S.date) : cs('D' + n, '-', S.text), cn('E' + n, o.active, S.int), cn('F' + n, o.value, S.int)];
      r++;
    });
    var t = String(r);
    rows[r - 1] = [cs('A' + t, 'Total', S.totLabel), cb('B' + t, S.totBlank), cb('C' + t, S.totBlank), cb('D' + t, S.totBlank),
      cb('E' + t, S.totBlank), cf('F' + t, 'SUM(F6:F' + (r - 1) + ')', m.summary.value, S.totInt)];
    return sheetXml(rows, [6, 36, 20, 16, 10, 20], 5);
  }
  function estNeedsSheet(m) {
    var rows = [];
    rows[0] = [cs('A1', 'Produk yang Perlu Disiapkan', S.title)];
    rows[1] = [cs('A2', estMeta(m), S.muted)];
    rows[2] = [cs('A3', EST_NOTE, S.muted)];
    var head = ['No', 'Produk', 'Perkiraan Qty', 'Jumlah outlet', 'Perkiraan nilai (Rp)'];
    rows[4] = head.map(function (h, i) { return cs(colName(i) + '5', h, S.head); });
    var r = 6;
    m.needs.forEach(function (x, i) {
      var n = String(r);
      rows[r - 1] = [cn('A' + n, i + 1, S.int), cs('B' + n, x.product, S.text), cn('C' + n, x.qty, S.int), cn('D' + n, x.outlets, S.int), cn('E' + n, x.value, S.int)];
      r++;
    });
    var t = String(r);
    rows[r - 1] = [cs('A' + t, 'Total', S.totLabel), cb('B' + t, S.totBlank), cb('C' + t, S.totBlank), cb('D' + t, S.totBlank),
      cf('E' + t, 'SUM(E6:E' + (r - 1) + ')', m.summary.value, S.totInt)];
    return sheetXml(rows, [6, 38, 14, 14, 20], 5);
  }
  var EST_HEAD = ['Produk', 'Status', 'Order terakhir', 'Qty terakhir', 'Jumlah order', 'Jeda biasa (hari)', 'Order berikutnya', 'Perkiraan qty', 'Perkiraan nilai (Rp)'];
  function estOutletSheet(m, o) {
    var rows = [];
    rows[0] = [cs('A1', 'Estimasi Order - ' + o.name, S.title)];
    rows[1] = [cs('A2', estMeta(m), S.muted)];
    rows[2] = [cs('A3', o.statusText + (o.nextDate ? '   |   Order berikutnya sekitar ' + o.nextLabel : ''), S.muted)];
    rows[4] = EST_HEAD.map(function (h, i) { return cs(colName(i) + '5', h, S.head); });
    var r = 6;
    o.products.forEach(function (x) {
      var n = String(r);
      rows[r - 1] = [cs('A' + n, x.product, S.text), cs('B' + n, x.statusText, S.text), cn('C' + n, serial(x.lastDate), S.date),
        cn('D' + n, x.lastQty, S.int), cn('E' + n, x.orders, S.int),
        x.interval ? cn('F' + n, x.interval, S.int) : cs('F' + n, '-', S.text),
        x.nextDate ? cn('G' + n, serial(x.nextDate), S.date) : cs('G' + n, '-', S.text),
        x.counted ? cn('H' + n, x.qty, S.int) : cs('H' + n, '-', S.text), cn('I' + n, x.counted ? x.value : 0, S.int)];
      r++;
    });
    var t = String(r);
    rows[r - 1] = [cs('A' + t, 'Total', S.totLabel), cb('B' + t, S.totBlank), cb('C' + t, S.totBlank), cb('D' + t, S.totBlank), cb('E' + t, S.totBlank),
      cb('F' + t, S.totBlank), cb('G' + t, S.totBlank), cb('H' + t, S.totBlank), cf('I' + t, 'SUM(I6:I' + (r - 1) + ')', o.value, S.totInt)];
    return sheetXml(rows, [34, 20, 14, 12, 12, 16, 16, 14, 20], 5);
  }
  function xlsxEst(m) {
    var used = {};
    var sheets = [{ name: sheetName('Ringkasan', used), xml: estSummarySheet(m) }, { name: sheetName('Kebutuhan Produk', used), xml: estNeedsSheet(m) }];
    m.outlets.forEach(function (o) { sheets.push({ name: sheetName(o.name, used), xml: estOutletSheet(m, o) }); });
    return book(sheets);
  }


  /* ------------------------------------------------------------------ PDF */
  function pdf(model) {
    var jsPDF = global.jspdf && global.jspdf.jsPDF;
    if (!jsPDF) throw new Error('Pustaka PDF belum termuat.');
    var doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    if (typeof doc.autoTable !== 'function') throw new Error('Pustaka tabel PDF belum termuat.');
    var W = doc.internal.pageSize.getWidth();
    var M = 14;
    var BLUE = [26, 115, 232], LIGHT = [232, 240, 254];
    var meta = 'Bulan: ' + model.monthLabel + (model.owner ? '   |   Sales: ' + model.owner : '');

    function heading(title, sub1, sub2) {
      doc.setTextColor(20);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(15);
      var lines = doc.splitTextToSize(title, W - 2 * M);
      doc.text(lines, M, 16);
      var y = 16 + (lines.length - 1) * 6.5;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(100);
      doc.text(sub1, M, y + 6);
      doc.text(sub2, M, y + 11);
      doc.setTextColor(20);
      return y + 16;
    }
    var common = {
      theme: 'grid', margin: { left: M, right: M, bottom: 14 },
      styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 1.8, lineColor: [191, 199, 213], lineWidth: 0.2, textColor: 20 },
      headStyles: { fillColor: BLUE, textColor: 255, fontStyle: 'bold', halign: 'center', valign: 'middle' },
      footStyles: { fillColor: LIGHT, textColor: 20, fontStyle: 'bold' },
      showFoot: 'lastPage'
    };
    function R(t) { return { content: t, styles: { halign: 'right' } }; }
    function opts(extra) { var o = {}, k; for (k in common) o[k] = common[k]; for (k in extra) o[k] = extra[k]; return o; }

    // halaman 1: ringkasan semua outlet
    var y = heading('Rekap Penjualan per Outlet', meta, 'Dibuat: ' + model.printedAt + '   |   Nominal dalam Rupiah');
    var g = model.grand;
    doc.autoTable(opts({
      startY: y,
      head: [['No', 'Outlet', 'Transaksi', 'Item', 'Subtotal', 'Diskon', 'Pajak', 'Total']],
      body: model.outlets.map(function (o, i) {
        return [String(i + 1), o.name, fmt(o.transactions), fmt(o.qty), fmt(o.sub), fmt(o.disc), fmt(o.tax), fmt(o.total)];
      }),
      foot: [['', 'Total semua outlet', R(fmt(g.transactions)), R(fmt(g.qty)), R(fmt(g.sub)), R(fmt(g.disc)), R(fmt(g.tax)), R(fmt(g.total))]],
      columnStyles: { 0: { halign: 'center', cellWidth: 12 }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' },
        5: { halign: 'right' }, 6: { halign: 'right' }, 7: { halign: 'right', fontStyle: 'bold' } }
    }));

    // satu tabel per outlet, tiap outlet mulai di halaman baru
    model.outlets.forEach(function (o) {
      doc.addPage();
      var sy = heading(o.name, meta, o.transactions + ' transaksi   |   ' + fmt(o.qty) + ' item   |   Total ' + fmt(o.total));
      doc.autoTable(opts({
        startY: sy,
        head: [['Tanggal', 'Produk', 'Qty', 'Harga', 'Disc %', 'Pajak %', 'Subtotal', 'Diskon', 'Pajak', 'Total']],
        body: o.rows.map(function (x) {
          return [x.dateLabel, x.product + (x.note ? '\n(' + x.note + ')' : ''), fmt(x.qty), fmt(x.price), pctText(x.disc), pctText(x.tax),
            fmt(x.sub), fmt(x.discAmt), fmt(x.taxAmt), fmt(x.total)];
        }),
        foot: [['Total', '', R(fmt(o.qty)), '', '', '', R(fmt(o.sub)), R(fmt(o.disc)), R(fmt(o.tax)), R(fmt(o.total))]],
        columnStyles: { 0: { cellWidth: 24 }, 2: { halign: 'right', cellWidth: 14 }, 3: { halign: 'right', cellWidth: 24 },
          4: { halign: 'right', cellWidth: 16 }, 5: { halign: 'right', cellWidth: 16 }, 6: { halign: 'right', cellWidth: 28 },
          7: { halign: 'right', cellWidth: 26 }, 8: { halign: 'right', cellWidth: 26 }, 9: { halign: 'right', cellWidth: 30, fontStyle: 'bold' } }
      }));
    });

    var pages = doc.getNumberOfPages();
    var H = doc.internal.pageSize.getHeight();
    for (var i = 1; i <= pages; i++) {
      doc.setPage(i);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(130);
      doc.text('Rekap Sales', M, H - 7);
      doc.text('Halaman ' + i + ' dari ' + pages, W - M, H - 7, { align: 'right' });
    }
    return doc.output('blob');
  }

  function pdfEst(m) {
    var jsPDF = global.jspdf && global.jspdf.jsPDF;
    if (!jsPDF) throw new Error('Pustaka PDF belum termuat.');
    var doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    if (typeof doc.autoTable !== 'function') throw new Error('Pustaka tabel PDF belum termuat.');
    var W = doc.internal.pageSize.getWidth(), M = 14;
    var BLUE = [26, 115, 232], LIGHT = [232, 240, 254];
    function heading(title, sub1, sub2) {
      doc.setTextColor(20); doc.setFont('helvetica', 'bold'); doc.setFontSize(15);
      var lines = doc.splitTextToSize(title, W - 2 * M);
      doc.text(lines, M, 16);
      var y = 16 + (lines.length - 1) * 6.5;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(100);
      var l1 = doc.splitTextToSize(sub1, W - 2 * M); doc.text(l1, M, y + 6);
      y += 6 + (l1.length - 1) * 4;
      var l2 = doc.splitTextToSize(sub2, W - 2 * M); doc.text(l2, M, y + 5);
      y += 5 + (l2.length - 1) * 4;
      doc.setTextColor(20);
      return y + 6;
    }
    var common = {
      theme: 'grid', margin: { left: M, right: M, bottom: 14 },
      styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 1.8, lineColor: [191, 199, 213], lineWidth: 0.2, textColor: 20 },
      headStyles: { fillColor: BLUE, textColor: 255, fontStyle: 'bold', halign: 'center', valign: 'middle' },
      footStyles: { fillColor: LIGHT, textColor: 20, fontStyle: 'bold' }, showFoot: 'lastPage'
    };
    function R(t) { return { content: t, styles: { halign: 'right' } }; }
    function opts(extra) { var o = {}, k; for (k in common) o[k] = common[k]; for (k in extra) o[k] = extra[k]; return o; }
    var y = heading('Estimasi Kebutuhan Order per Outlet', estMeta(m), EST_NOTE);
    doc.autoTable(opts({
      startY: y,
      head: [['No', 'Outlet', 'Status', 'Order berikutnya', 'Produk', 'Perkiraan nilai']],
      body: m.outlets.map(function (o, i) { return [String(i + 1), o.name, o.statusText, o.nextDate ? o.nextLabel : '-', fmt(o.active), fmt(o.value)]; }),
      foot: [['', 'Total', '', '', '', R(fmt(m.summary.value))]],
      columnStyles: { 0: { halign: 'center', cellWidth: 12 }, 4: { halign: 'right' }, 5: { halign: 'right', fontStyle: 'bold' } }
    }));
    if (m.needs.length) {
      doc.addPage();
      var ny = heading('Produk yang Perlu Disiapkan', estMeta(m), 'Gabungan perkiraan qty dari semua outlet.');
      doc.autoTable(opts({
        startY: ny,
        head: [['No', 'Produk', 'Perkiraan Qty', 'Jumlah outlet', 'Perkiraan nilai']],
        body: m.needs.map(function (x, i) { return [String(i + 1), x.product, fmt(x.qty), fmt(x.outlets), fmt(x.value)]; }),
        foot: [['', 'Total', '', '', R(fmt(m.summary.value))]],
        columnStyles: { 0: { halign: 'center', cellWidth: 12 }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right', fontStyle: 'bold' } }
      }));
    }
    m.outlets.forEach(function (o) {
      doc.addPage();
      var sy = heading(o.name, estMeta(m), o.statusText + (o.nextDate ? '   |   Order berikutnya sekitar ' + o.nextLabel : '') + '   |   Perkiraan ' + fmt(o.value));
      doc.autoTable(opts({
        startY: sy,
        head: [['Produk', 'Status', 'Order terakhir', 'Qty terakhir', 'Jumlah order', 'Jeda (hari)', 'Order berikutnya', 'Perkiraan qty', 'Perkiraan nilai']],
        body: o.products.map(function (x) {
          return [x.product, x.statusText, x.lastLabel, fmt(x.lastQty), fmt(x.orders), x.interval ? fmt(x.interval) : '-', x.nextDate ? x.nextLabel : '-',
            x.counted ? fmt(x.qty) : '-', x.counted ? fmt(x.value) : '0'];
        }),
        foot: [['Total', '', '', '', '', '', '', '', R(fmt(o.value))]],
        columnStyles: { 3: { halign: 'right' }, 4: { halign: 'right' }, 5: { halign: 'right' }, 7: { halign: 'right' }, 8: { halign: 'right', fontStyle: 'bold' } }
      }));
    });
    var pages = doc.getNumberOfPages(), H = doc.internal.pageSize.getHeight();
    for (var i = 1; i <= pages; i++) {
      doc.setPage(i); doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(130);
      doc.text('Rekap Sales - Estimasi', M, H - 7);
      doc.text('Halaman ' + i + ' dari ' + pages, W - M, H - 7, { align: 'right' });
    }
    return doc.output('blob');
  }

  global.RSExport = { xlsx: xlsx, pdf: pdf, xlsxEst: xlsxEst, pdfEst: pdfEst };
})(window);
