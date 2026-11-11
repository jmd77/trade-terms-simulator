/*
 * Trade terms simulator — export.js (SPEC5 9: the formatted Excel export, Part 3)
 *
 * Loaded after app.js (and vendor/exceljs.min.js). Adds two functions to TTS:
 *   TTS.buildStyledWorkbook(ExcelJS, state, scope) -> an ExcelJS workbook (pure: no DOM, runs in Node)
 *   TTS.exportStyled(scope)                        builds it and downloads "Trade terms - <Customer>.xlsx" or
 *                                                  "Trade terms - all customers.xlsx" (Blob + temporary link), toast
 *                                                  "Excel file exported."; without ExcelJS it falls back to the old
 *                                                  two-sheet SheetJS export (TTS.excel.exportWorkbook)
 * scope = a customer id, or 'all' (or nothing). One sheet per screen: Priorities, Contract analysis, Agreement today,
 * Track and trace, Negotiation plan; with all customers every sheet has a Customer column first.
 * Everything comes from TTS.engine, TTS.fmt, TTS.plan.texts and the state (never from the DOM), so the labels on screen
 * can change without breaking the export.
 */
(function () {
  'use strict';
  var NODE = typeof module !== 'undefined' && module.exports;
  var TTS = NODE ? require('./app.js') : window.TTS;
  var E = TTS.engine, fmt = TTS.fmt;

  // ---------------------------------------------------------------- look
  var COL = { navy: 'FF232341', white: 'FFFFFFFF', grid: 'FFD5D9E1', band: 'FFE8EBF2', mint: 'FFD9F2E7', yellow: 'FFFFF1C2',
    grey: 'FFECEEF1', muted: 'FF7D8396' };
  var FONT = 'Calibri';
  var THIN = { style: 'thin', color: { argb: COL.grid } };
  var NAVY_RULE = { style: 'thin', color: { argb: COL.navy } };
  var FILL_BY_LABEL = {
    'Conditional': COL.mint, 'Fully executed': COL.mint,
    'Partially conditional': COL.yellow, 'Partially executed': COL.yellow,
    'Unconditional': COL.grey, 'Not executed': COL.grey
  };
  // column types: number format and alignment
  var TYPES = {
    text: { h: 'left' },
    num1: { fmt: '0.0', h: 'right' },          // SEK m
    pct1: { fmt: '0.0%', h: 'right' },         // a real fraction
    rate: { fmt: '0.00%', h: 'right' },
    date: { fmt: 'd mmm yyyy', h: 'right' },
    int: { fmt: '0', h: 'right' },
    num: { fmt: null, h: 'right' },            // whole number as 0, otherwise 0.0
    yn: { h: 'center' },
    cls: { h: 'left' },                        // classification / status: lightly filled by value
    any: { fmt: null, h: 'right' }             // sheet 5: format given per cell
  };
  function fill(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }
  function finite(n) { return typeof n === 'number' && isFinite(n); }
  function r1(n) { return Math.round((+n || 0) * 10 + 1e-9) / 10; }
  function xDate(iso) { if (!iso) return null; var p = String(iso).split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])); }
  function xNum(n) { return finite(n) ? n : null; }
  function yn(b) { return b ? 'Yes' : 'No'; }
  function neg(n) { return finite(n) && n !== 0 ? -n : 0; }
  function pName(id) { var p = E.priority(id); return p ? p.name : ''; }

  // How many lines a wrapped text takes in a column of that width, so the row height can be set (Excel does not grow rows
  // by itself when it opens a file). Approximate Calibri 11 glyph widths in pixels at 100% zoom; greedy word wrap.
  var GLYPH = {}, G_DEFAULT = 6.1;
  [['fijlrtI.,:;\'!|() ', 3.6], ['sczJ-', 5.1], ['kvxyEFLSTZ?', 5.9], ['0123456789abdeghnopqu$', 6.6], ['ABCKPRVXY#+=', 7.2],
    ['DGHNU–', 7.9], ['OQw&', 8.6], ['m%', 10.3], ['MW', 11.6]].forEach(function (g) {
    for (var i = 0; i < g[0].length; i++) GLYPH[g[0].charAt(i)] = g[1];
  });
  function textPx(s, bold) {
    var px = 0;
    for (var i = 0; i < s.length; i++) px += GLYPH[s.charAt(i)] || G_DEFAULT;
    return px * (bold ? 1.29 : 1.23);
  }
  function wrapLines(s, width, bold) {
    var avail = Math.max(20, width * 7 - 4), space = textPx(' ', bold);
    return String(s).split('\n').reduce(function (n, para) {
      var count = 1, len = 0;
      para.split(' ').forEach(function (w) {
        var L = textPx(w, bold);
        if (len === 0) len = L;
        else if (len + space + L <= avail) len += space + L;
        else { count++; len = L; }
        while (len > avail) { count++; len -= avail; }
      });
      return n + count;
    }, 0);
  }

  // ---------------------------------------------------------------- a sheet = a table with a styled header
  // cols: [{ h: header, w: width, t: type }]; a cell value may be { v, fmt, bold, italic, muted, fill, indent, h }
  function Table(wb, name, cols, opt) {
    opt = opt || {};
    var ws = wb.addWorksheet(name, {
      views: [{ state: 'frozen', xSplit: opt.xSplit || 0, ySplit: opt.headRows || 1, activeCell: 'A' + ((opt.headRows || 1) + 1) }],
      pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.3, footer: 0.3 },
        printTitlesRow: '1:' + (opt.headRows || 1) },
      headerFooter: { oddFooter: '&L&8Trade terms simulator · ' + name + '&R&8Page &P of &N' }
    });
    ws.columns = cols.map(function (c) { return { width: c.w }; });
    this.ws = ws; this.cols = cols; this.n = cols.length;
  }
  // a styled header row (texts: one per column; null = part of a merge)
  Table.prototype.header = function (texts) {
    var self = this, row = this.ws.addRow(texts), lines = 1;
    for (var i = 1; i <= this.n; i++) {
      var cell = row.getCell(i), t = texts[i - 1];
      cell.font = { name: FONT, size: 11, bold: true, color: { argb: COL.white } };
      cell.fill = fill(COL.navy);
      cell.alignment = { wrapText: true, vertical: 'middle', horizontal: TYPES[self.cols[i - 1].t].h === 'right' ? 'right' : 'left' };
      cell.border = { top: NAVY_RULE, left: { style: 'thin', color: { argb: 'FF3A3A5C' } }, bottom: NAVY_RULE, right: { style: 'thin', color: { argb: 'FF3A3A5C' } } };
      if (t) lines = Math.max(lines, wrapLines(t, self.cols[i - 1].w, true));
    }
    row.height = Math.max(30, lines * 15 + 4);
    return row;
  };
  // a data row; style: { bold, fill, rule (navy rule above), italic }
  Table.prototype.row = function (values, style) {
    style = style || {};
    var self = this, row = this.ws.addRow(values.map(function (v) { return v && typeof v === 'object' && !(v instanceof Date) ? v.v : v; }));
    var lines = 1;
    for (var i = 1; i <= this.n; i++) {
      var col = self.cols[i - 1], ty = TYPES[col.t], raw = values[i - 1];
      var o = raw && typeof raw === 'object' && !(raw instanceof Date) ? raw : {};
      var v = o.v !== undefined ? o.v : raw;
      var cell = row.getCell(i);
      var bold = !!(style.bold || o.bold), italic = !!(style.italic || o.italic);
      cell.font = { name: FONT, size: 11, bold: bold, italic: italic, color: o.muted ? { argb: COL.muted } : undefined };
      var nf = o.fmt || ty.fmt;
      if (!nf && finite(v)) nf = (col.t === 'num' || col.t === 'any') ? (Math.abs(v - Math.round(v)) < 1e-9 ? '0' : '0.0') : null;
      if (nf && (finite(v) || v instanceof Date)) cell.numFmt = nf;
      var h = o.h || (typeof v === 'string' && (col.t === 'any' || col.t === 'num') ? 'right' : ty.h);
      cell.alignment = { vertical: 'top', horizontal: h, wrapText: typeof v === 'string', indent: o.indent || undefined };
      var bg = o.fill || style.fill || (col.t === 'cls' && FILL_BY_LABEL[v]);
      if (bg) cell.fill = fill(bg);
      cell.border = { top: style.rule ? NAVY_RULE : THIN, left: THIN, bottom: THIN, right: THIN };
      if (typeof v === 'string' && v) lines = Math.max(lines, wrapLines(v, col.w - (o.indent || 0) * 1.5, bold));
    }
    if (lines > 1) row.height = lines * 15 + 2;
    return row;
  };
  Table.prototype.blank = function () { return this.ws.addRow([]); };
  Table.prototype.filter = function (headRow) {
    this.ws.autoFilter = { from: { row: headRow || 1, column: 1 }, to: { row: headRow || 1, column: this.n } };
  };

  function withCustomer(all, cols) { return all ? [{ h: 'Customer', w: 16, t: 'text' }].concat(cols) : cols; }
  function pre(all, c, values) { return all ? [c ? c.name : 'All customers'].concat(values) : values; }

  // ---------------------------------------------------------------- 1. Priorities
  function sheetPriorities(wb, st, cs, all) {
    var cols = withCustomer(all, [
      { h: '#', w: 5, t: 'int' }, { h: 'Priority', w: 28, t: 'text' }, { h: 'What winning looks like', w: 64, t: 'text' },
      { h: 'Spend today (SEK m)', w: 12, t: 'num1' }, { h: 'Share today', w: 12, t: 'pct1' },
      { h: 'Spend proposed (SEK m)', w: 12, t: 'num1' }, { h: 'Share proposed', w: 12, t: 'pct1' }]);
    var T = new Table(wb, 'Priorities', cols, { xSplit: all ? 1 : 0 });
    T.header(cols.map(function (c) { return c.h; }));
    function block(c, lines, first) {
      var today = E.spendByPriority(lines, 'today'), prop = E.spendByPriority(lines, 'proposed');
      today.forEach(function (e, i) {
        var p = e.priorityId ? E.priority(e.priorityId) : null, none = !e.priorityId;
        T.row(pre(all, c, [none ? null : i + 1, none ? { v: 'No priority', italic: true } : e.name,
          none ? { v: 'Spend that serves none of the priorities', muted: true } : (p && p.winning) || '',
          e.amount, e.share, prop[i].amount, prop[i].share]), { rule: all && i === 0 && !first });
      });
    }
    cs.forEach(function (c, k) { block(c, c.lines, k === 0); });
    if (all) {
      var lines = []; cs.forEach(function (c) { lines = lines.concat(c.lines); });
      block(null, lines, false);
    }
    return T;
  }

  // ---------------------------------------------------------------- 2. Contract analysis
  function sheetContract(wb, st, cs, all) {
    var cols = withCustomer(all, [
      { h: 'Contract point', w: 32, t: 'text' }, { h: 'Contract wording', w: 60, t: 'text' }, { h: 'Justification found', w: 40, t: 'text' },
      { h: 'SEK m', w: 12, t: 'num1' }, { h: '% of gross revenue', w: 12, t: 'pct1' }, { h: 'Fixed or variable', w: 12, t: 'text' },
      { h: 'Rate', w: 12, t: 'rate' }, { h: 'Rate is on', w: 14, t: 'text' }, { h: 'Paid how', w: 15, t: 'text' },
      { h: 'Classification', w: 16, t: 'cls' }, { h: 'Bucket', w: 18, t: 'text' }, { h: 'Priority', w: 28, t: 'text' },
      { h: 'Internal counterpart', w: 36, t: 'text' }, { h: 'Target', w: 12, t: 'num' }, { h: 'Unit', w: 14, t: 'text' },
      { h: 'From', w: 12, t: 'date' }, { h: 'To', w: 12, t: 'date' }, { h: 'Owner', w: 20, t: 'text' }, { h: 'Data source', w: 24, t: 'text' },
      { h: 'Specific', w: 11, t: 'yn' }, { h: 'Measurable', w: 12, t: 'yn' }, { h: 'With an owner', w: 11, t: 'yn' }, { h: 'Time-bound', w: 11, t: 'yn' },
      { h: 'SMART (of 4)', w: 11, t: 'int' }, { h: 'Aligned with a priority', w: 12, t: 'yn' },
      { h: 'Proposed action', w: 24, t: 'text' }, { h: 'Why', w: 44, t: 'text' }, { h: 'Proposed SEK m', w: 12, t: 'num1' },
      { h: 'Proposed classification', w: 16, t: 'cls' }, { h: 'Proposed priority', w: 28, t: 'text' }]);
    var n = cols.length;
    var T = new Table(wb, 'Contract analysis', cols, { xSplit: all ? 2 : 1 });
    T.header(cols.map(function (c) { return c.h; }));
    cs.forEach(function (c, k) {
      var prop = {};
      E.waterfall(c, 'proposed').forEach(function (r) { prop[r.lineId || r.key] = r; });
      E.waterfall(c, 'today').forEach(function (r, i) {
        var pr = prop[r.lineId || r.key];
        if (r.kind !== 'line') {
          var v = pre(all, c, [r.label, null, null, r.amount, r.pct]);
          while (v.length < n - 3) v.push(null);
          v.push(pr ? pr.amount : null, null, null);
          T.row(v, { bold: true, fill: COL.band, rule: r.kind === 'anchor' || (all && i === 0 && k > 0) });
          return;
        }
        var l = null;
        c.lines.forEach(function (x) { if (x.id === r.lineId) l = x; });
        var cond = E.type(l, 'today') === 'conditional', sm = E.smart(l, 'today');
        var sug = E.suggestPriority(l), rate = E.rateOf(l, c);
        var dash = { v: '–', muted: true, h: 'center' };
        T.row(pre(all, c, [
          l.name, l.clause || '', l.justification || '', r.amount, r.pct, r.basis, finite(rate) ? rate / 100 : null,
          E.rateBase(l) === 'gross' ? 'Gross revenue' : 'First net', l.paidHow || '', E.CLASS_LABEL[r.cls],
          cond ? (l.bucket || 'No bucket') : dash,
          E.priorityId(l, 'today') ? pName(l.priorityId) : sug ? { v: 'Suggested: ' + pName(sug), fill: COL.yellow } : { v: 'No priority', muted: true },
          l.counterpart || '', xNum(l.target), l.unit || '', xDate(l.start), xDate(l.end), l.owner || '', l.source || '',
          cond ? yn(sm.specific) : dash, cond ? yn(sm.measurable) : dash, cond ? yn(sm.owner) : dash, cond ? yn(sm.timebound) : dash,
          cond ? sm.score : dash, yn(E.aligned(l, 'today')),
          l.action || '', l.why || '', pr ? pr.amount : E.amount(l, 'proposed'), E.CLASS_LABEL[E.classOf(l, 'proposed')],
          E.priorityId(l, 'proposed') ? pName(l.propPriorityId) : { v: 'No priority', muted: true }
        ]));
      });
    });
    T.filter(1);
    return T;
  }

  // ---------------------------------------------------------------- 3. Agreement today (+ By pack)
  function sheetWaterfall(wb, st, cs, all) {
    var cols = withCustomer(all, [
      { h: 'Row', w: 36, t: 'text' }, { h: 'Bucket', w: 20, t: 'text' }, { h: 'Fixed or variable', w: 18, t: 'text' },
      { h: 'SEK m today', w: 12, t: 'num1' }, { h: '% of gross revenue today', w: 12, t: 'pct1' },
      { h: 'SEK m proposed', w: 12, t: 'num1' }, { h: '% of gross revenue proposed', w: 12, t: 'pct1' }]);
    var T = new Table(wb, 'Agreement today', cols, { xSplit: all ? 2 : 1 });
    T.header(cols.map(function (c) { return c.h; }));
    cs.forEach(function (c, k) {
      var prop = {};
      E.waterfall(c, 'proposed').forEach(function (r) { prop[r.lineId || r.key] = r; });
      E.waterfall(c, 'today').forEach(function (r, i) {
        var pr = prop[r.lineId || r.key] || { amount: 0, pct: 0 };
        if (r.kind === 'anchor') {
          T.row(pre(all, c, [r.label, null, null, r.amount, r.pct, pr.amount, pr.pct]), { bold: true, fill: COL.band, rule: true });
        } else if (r.kind === 'promo') {
          T.row(pre(all, c, [{ v: r.label, indent: 2, italic: true }, null, { v: 'Estimate', italic: true, muted: true },
            neg(r.amount), r.pct, neg(pr.amount), pr.pct]));
        } else {
          var fv = r.basis === 'Variable' && finite(r.rate) ? 'Variable (' + fmt.rate(r.rate) + ')' : r.basis;
          T.row(pre(all, c, [{ v: r.label, indent: 2 }, r.bucket, fv, neg(r.amount), r.pct, neg(pr.amount), pr.pct]));
        }
      });
    });
    // "By pack": one small table under the waterfall, today's view, for the customers that have pack groups
    var withPacks = cs.filter(function (c) { return E.packWaterfall(c, 'today'); });
    if (withPacks.length) {
      T.blank();
      var off = all ? 1 : 0, maxCols = 0;
      withPacks.forEach(function (c) { maxCols = Math.max(maxCols, E.packWaterfall(c, 'today').length); });
      // widen the columns the pack table uses beyond the waterfall's own
      for (var j = T.n + 1; j <= off + 1 + maxCols * 2; j++) T.ws.getColumn(j).width = 12;
      var head1 = all ? ['Customer'] : [], head2 = all ? [null] : [];
      head1.push('By pack (today)'); head2.push(null);
      var first = E.packWaterfall(withPacks[0], 'today');
      first.forEach(function (p) { head1.push(p.name, null); head2.push('SEK m', '% of gross revenue'); });
      var width = head1.length;
      var r1row = T.ws.addRow(head1), r2row = T.ws.addRow(head2);
      [r1row, r2row].forEach(function (row) {
        for (var i = 1; i <= width; i++) {
          var cell = row.getCell(i);
          cell.font = { name: FONT, size: 11, bold: true, color: { argb: COL.white } };
          cell.fill = fill(COL.navy);
          cell.alignment = { wrapText: true, vertical: 'middle', horizontal: i > off + 1 ? (row === r1row ? 'center' : 'right') : 'left' };
          cell.border = { top: NAVY_RULE, left: NAVY_RULE, bottom: NAVY_RULE, right: NAVY_RULE };
        }
      });
      r1row.height = 20; r2row.height = 30;
      for (var m = 0; m < off + 1; m++) T.ws.mergeCells(r1row.number, m + 1, r2row.number, m + 1);
      first.forEach(function (p, i) { var col = off + 2 + i * 2; T.ws.mergeCells(r1row.number, col, r1row.number, col + 1); });
      var ROWS = [['gross', 'Gross revenue at list price', 'net'], ['onInvoice', 'On invoice', 'ded'], ['firstNet', 'First net (net invoice sales)', 'net'],
        ['offInvoice', 'Off invoice', 'ded'], ['secondNet', 'Second net', 'net'], ['promo', 'Promotional spend (estimate)', 'ded'], ['thirdNet', 'Third net (net-net)', 'net']];
      withPacks.forEach(function (c, k) {
        var pw = E.packWaterfall(c, 'today');
        ROWS.forEach(function (rw, i) {
          var vals = all ? [c.name] : [];
          vals.push(rw[1]);
          pw.forEach(function (p) { vals.push(rw[2] === 'ded' ? neg(p[rw[0]]) : p[rw[0]], p.pct[rw[0]]); });
          var row = T.ws.addRow(vals), net = rw[2] === 'net';
          for (var q = 1; q <= vals.length; q++) {
            var cell = row.getCell(q), isNum = q > off + 1, isPct = isNum && (q - off - 1) % 2 === 0;
            cell.font = { name: FONT, size: 11, bold: net };
            if (isNum) cell.numFmt = isPct ? '0.0%' : '0.0';
            cell.alignment = { vertical: 'top', horizontal: isNum ? 'right' : 'left', indent: !isNum && q === off + 1 && !net ? 2 : undefined, wrapText: !isNum };
            if (net) cell.fill = fill(COL.band);
            cell.border = { top: net || (all && i === 0 && k > 0) ? NAVY_RULE : THIN, left: THIN, bottom: THIN, right: THIN };
          }
        });
      });
    }
    return T;
  }

  // ---------------------------------------------------------------- 4. Track and trace
  function sheetTrack(wb, st, cs, all) {
    var MON = E.MONTHS;
    var lead = withCustomer(all, [
      { h: 'Investment line', w: 28, t: 'text' }, { h: 'KPI', w: 48, t: 'text' }, { h: 'Target', w: 9.5, t: 'num' }, { h: 'Unit', w: 14, t: 'text' },
      { h: 'From', w: 12, t: 'date' }, { h: 'To', w: 12, t: 'date' }, { h: 'Data source', w: 24, t: 'text' }, { h: 'Owner', w: 22, t: 'text' },
      { h: 'Proof of execution', w: 18, t: 'text' }, { h: 'Paid how', w: 15, t: 'text' }]);
    var months = [];
    MON.forEach(function () { months.push({ h: 'Planned', w: 8.5, t: 'num' }, { h: 'Executed', w: 9.5, t: 'num' }); });
    var tail = [{ h: 'Actual to date', w: 10, t: 'num' }, { h: 'Expected by now', w: 10, t: 'num' }, { h: 'Status', w: 14, t: 'cls' },
      { h: 'SEK m at stake', w: 10, t: 'num1' }, { h: 'Paid to date (SEK m)', w: 10, t: 'num1' }, { h: 'Paid ahead of delivery', w: 11, t: 'yn' },
      { h: 'KPI tracked', w: 9.5, t: 'yn' }, { h: 'Checked', w: 9.5, t: 'yn' }, { h: 'Payment already due', w: 10, t: 'yn' }, { h: 'Paid in full', w: 9.5, t: 'yn' }];
    var cols = lead.concat(months, tail);
    var T = new Table(wb, 'Track and trace', cols, { xSplit: all ? 3 : 2, headRows: 2 });
    // two header rows: the month names merged over (Planned, Executed); every other header merged over both rows
    var h1 = [], h2 = [];
    lead.forEach(function (c) { h1.push(c.h); h2.push(null); });
    MON.forEach(function (m) { h1.push(m, null); h2.push('Planned', 'Executed'); });
    tail.forEach(function (c) { h1.push(c.h); h2.push(null); });
    var hr1 = T.header(h1), hr2 = T.header(h2);
    var fm = lead.length + 1;
    for (var i = 1; i <= cols.length; i++) {
      if (i >= fm && i < fm + 24) {
        hr1.getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
        hr2.getCell(i).alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };
        if ((i - fm) % 2 === 0) T.ws.mergeCells(1, i, 1, i + 1);
      } else {
        T.ws.mergeCells(1, i, 2, i);
      }
    }
    var lines1 = 1;
    lead.concat(tail).forEach(function (c) { lines1 = Math.max(lines1, wrapLines(c.h, c.w, true)); });
    hr1.height = 20; hr2.height = Math.max(20, lines1 * 15 + 4 - 20);
    var monthIdx = E.currentMonth();
    cs.forEach(function (c, k) {
      var lineById = {};
      c.lines.forEach(function (l) { lineById[l.id] = l; });
      E.kpis(c).forEach(function (kp, i) {
        var l = lineById[kp.lineId] || {}, inPeriod = E.periodMonths(kp.start, kp.end);
        var vals = pre(all, c, [kp.lineName, kp.text || '', xNum(kp.target), kp.unit || '', xDate(kp.start), xDate(kp.end), kp.source || '',
          kp.owner || '', kp.proof || 'None', l.paidHow || '']);
        for (var m = 0; m < 12; m++) {
          var out = !inPeriod[m], later = monthIdx >= 0 && m > monthIdx;
          vals.push({ v: xNum(kp.plan[m]), muted: out }, { v: xNum(kp.done[m]), muted: out || later });
        }
        vals.push(xNum(kp.actual), xNum(kp.expected), E.statusLabel(kp.status), kp.money, xNum(kp.paid),
          kp.ahead ? { v: 'Yes', fill: COL.yellow, bold: true } : 'No', yn(kp.tracked), yn(kp.checked), yn(kp.due), yn(kp.paidFull));
        T.row(vals, { rule: all && i === 0 && k > 0 });
      });
    });
    T.filter(2);
    return T;
  }

  // ---------------------------------------------------------------- 5. Negotiation plan
  function sheetPlan(wb, st, cs, all) {
    var cols = withCustomer(all, [
      { h: 'Section', w: 32, t: 'text' }, { h: '#', w: 5, t: 'int' }, { h: 'Item', w: 40, t: 'text' },
      { h: 'Today', w: 12, t: 'any' }, { h: 'Proposed', w: 12, t: 'any' }, { h: 'Detail', w: 80, t: 'text' }]);
    var T = new Table(wb, 'Negotiation plan', cols, { xSplit: all ? 1 : 0 });
    T.header(cols.map(function (c) { return c.h; }));
    var SEK = '0.0', PCT = '0.0%';
    function num(v, f) { return finite(v) ? { v: v, fmt: f } : null; }
    function section(c, title, rows) {
      rows.forEach(function (r, i) {
        T.row(pre(all, c, [i === 0 ? { v: title, bold: true } : null, r[0], r[1], r[2], r[3], r[4] || null]), { rule: i === 0 });
      });
    }
    function keyFigures(c, kf) {
      var cl = kf.classes, cp = kf.counterparts, kp = kf.kpis;
      var rows = [
        [null, 'Trade terms (SEK m)', num(kf.total, SEK)],
        [null, 'Trade terms, % of gross revenue', num(kf.pctOfGross, PCT)],
        [null, 'Third net (SEK m)', num(kf.sales.thirdNet, SEK)],
        [null, 'Third net, % of gross revenue', num(kf.sales.gross > 0 ? kf.sales.thirdNet / kf.sales.gross : 0, PCT)],
        [null, { v: 'Unconditional', fill: COL.grey }, num(cl.unconditional.share, PCT), null, 'Share of the trade terms: ' + fmt.money(cl.unconditional.amount) + '.'],
        [null, { v: 'Partially conditional', fill: COL.yellow }, num(cl.partial.share, PCT), null, 'Share of the trade terms: ' + fmt.money(cl.partial.amount) + '.'],
        [null, { v: 'Conditional', fill: COL.mint }, num(cl.conditional.share, PCT), null, 'Share of the trade terms: ' + fmt.money(cl.conditional.amount) + '.'],
        [null, 'Counterparts SMART', cp.smart + ' of ' + cp.lines, null, 'Conditional lines whose counterpart is specific, measurable, with an owner and time-bound.'],
        [null, 'Aligned with a priority', cp.aligned + ' of ' + cp.lines, null, 'Conditional lines linked to one of the priorities.']
      ];
      if (r1(cp.smartNotAlignedAmount) > 0) rows.push([null, 'SMART but not aligned (SEK m)', num(cp.smartNotAlignedAmount, SEK)]);
      rows.push([null, { v: 'KPIs fully executed', fill: COL.mint }, { v: kp.executed, fmt: '0' }],
        [null, { v: 'KPIs partially executed', fill: COL.yellow }, { v: kp.partly, fmt: '0' }],
        [null, { v: 'KPIs not executed', fill: COL.grey }, { v: kp.notexecuted, fmt: '0' }],
        [null, 'Paid ahead of delivery (SEK m)', num(kf.paidAhead, SEK), null, kf.paidAhead > 0.05 ? 'Paid before the counterpart was delivered.' : null]);
      section(c, 'Key figures', rows);
    }
    function spend(c, lines) {
      var t = E.spendByPriority(lines, 'today'), p = E.spendByPriority(lines, 'proposed');
      section(c, 'Spend against priorities (SEK m)', t.map(function (e, i) {
        return [e.priorityId ? i + 1 : null, e.priorityId ? e.name : { v: 'No priority', italic: true }, num(e.amount, SEK), num(p[i].amount, SEK),
          'Share of the trade terms: ' + fmt.pct(e.share) + ' today, ' + fmt.pct(p[i].share) + ' proposed.'];
      }));
    }
    function entry(list, id) { for (var i = 0; i < list.length; i++) if (list[i].priorityId === id) return list[i]; return { amount: 0 }; }
    function one(c) {
      keyFigures(c, E.keyFigures(c));
      spend(c, c.lines);
      var mwId = E.mustWin(c), mw = E.priority(mwId), sug = E.suggestedMustWin(c);
      var tAmt = entry(E.spendByPriority(c.lines, 'today'), mwId).amount, pAmt = entry(E.spendByPriority(c.lines, 'proposed'), mwId).amount;
      var note = mwId === sug ? (r1(tAmt) === 0 ? 'Suggested: no spend today.' : 'Suggested: lowest spend today, ' + fmt.money(tAmt) + '.')
        : 'Your choice. Suggested: ' + pName(sug) + '.';
      section(c, 'Must-win battle (SEK m)', [[null, { v: mw ? mw.name : '–', bold: true }, num(tAmt, SEK), num(pAmt, SEK),
        ((mw && mw.winning) ? mw.winning + '\n' : '') + note]]);
      var opps = E.opportunities(c), size = E.opportunitySize(c), total = E.summary(c.lines, 'today').total;
      var orows = opps.map(function (o, i) { return [i + 1, o.title, num(o.money, SEK), null, o.detail]; });
      if (!orows.length) orows.push([null, 'Nothing stands out.', null, null, 'Every line is conditional, has a number and serves a priority.']);
      orows.push([null, { v: 'Not yet conditional with a number and linked to a priority (SEK m)', bold: true }, { v: size, fmt: SEK, bold: true }, null,
        fmt.pct(total > 0 ? size / total : 0) + ' of the trade terms.']);
      section(c, 'Top opportunities (SEK m)', orows);
      if (TTS.plan && typeof TTS.plan.texts === 'function') {
        var t = TTS.plan.texts(c);
        section(c, 'Script', [[null, 'Opening', null, null, t.opening], [null, 'What we propose', null, null, t.propose],
          [null, 'What we ask', null, null, t.ask], [null, 'What we give', null, null, t.give]]);
        section(c, 'Objections', [1, 2, 3, 4].map(function (n) { return [n, t['obj' + n + 'q'], null, null, t['obj' + n + 'a']]; }));
      }
    }
    cs.forEach(one);
    if (all) {
      var lines = []; cs.forEach(function (c) { lines = lines.concat(c.lines); });
      keyFigures(null, E.keyFigures(cs));
      spend(null, lines);
      var ranked = cs.map(function (c) {
        var total = E.summary(c.lines, 'today').total, size = E.opportunitySize(c);
        return { c: c, size: size, share: total > 0 ? size / total : 0 };
      }).sort(function (a, b) { return b.size - a.size; });
      section(null, 'Customers ranked by opportunity (SEK m)', ranked.map(function (x, i) {
        return [i + 1, x.c.name, num(x.size, SEK), null, fmt.pct(x.share) + ' of the spend is not yet conditional with a number and linked to a priority. Must-win battle: ' +
          (pName(E.mustWin(x.c)) || '–') + '.'];
      }));
      section(null, 'Investment growth against revenue growth', cs.map(function (c) {
        var g = E.growth(c);
        return [null, c.name, num(g.investment, PCT), null, 'Investment ' + fmt.pct1s(g.investment) + ', revenue (first net) ' + fmt.pct1s(g.revenue) + '.' +
          (g.faster ? ' Faster than revenue: ' + fmt.money(g.excess) + ' above revenue growth.' : '')];
      }));
    }
    return T;
  }

  // ---------------------------------------------------------------- the workbook
  TTS.buildStyledWorkbook = function (ExcelJS, state, scope) {
    var prev = TTS.state;
    TTS.state = state || prev;
    try {
      var st = TTS.state, all = !scope || scope === 'all';
      var cs = all ? st.customers.slice() : st.customers.filter(function (c) { return c.id === scope; });
      if (!cs.length) { cs = st.customers.slice(); all = true; }
      var wb = new ExcelJS.Workbook();
      wb.creator = 'Trade terms simulator';
      wb.title = all ? 'Trade terms - all customers' : 'Trade terms - ' + cs[0].name;
      wb.created = new Date();
      sheetPriorities(wb, st, cs, all);
      sheetContract(wb, st, cs, all);
      sheetWaterfall(wb, st, cs, all);
      sheetTrack(wb, st, cs, all);
      sheetPlan(wb, st, cs, all);
      return wb;
    } finally {
      TTS.state = prev;
    }
  };

  TTS.exportStyled = function (scope) {
    var g = typeof window !== 'undefined' ? window : {};
    var toast = (TTS.ui && TTS.ui.toast) || function () {};
    var name = TTS.excel.exportName(TTS.state, scope);
    if (typeof g.ExcelJS === 'undefined') {
      // fallback: the old two-sheet SheetJS export
      if (typeof g.XLSX === 'undefined') { toast('The Excel library did not load.'); return; }
      try {
        g.XLSX.writeFile(TTS.excel.exportWorkbook(g.XLSX, TTS.state, scope), name, { compression: true });
        toast('Excel file exported.');
      } catch (e) { toast('The Excel file could not be written.'); }
      return;
    }
    var wb;
    try { wb = TTS.buildStyledWorkbook(g.ExcelJS, TTS.state, scope); } catch (e) { toast('The Excel file could not be written.'); return; }
    return wb.xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var url = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = url; a.download = name; a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1000);
      toast('Excel file exported.');
    }, function () { toast('The Excel file could not be written.'); });
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = require('./app.js');
