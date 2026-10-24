(function (g) {
  const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  const IE_RE = /^(INS|INSERT|INT|EXT|I\/E|E\/I|INT\/EXT|EXT\/INT|INT\.?\/EXT\.?|EXT\.?\/INT\.?)$/i;

  // ---------- scene field parsing (shared by PDF and Excel) ----------

  function parseSceneField(raw) {
    let s = String(raw ?? '').replace(/\s+/g, ' ').trim();
    if (!s) return [];
    const epSc = s.match(/\bEP\.?\s*(\d+)\s*[,\-–]?\s*SC\.?\s*(\d+)\s*([A-Z](?![A-Z]))?/i);
    if (epSc) return [{ ep: +epSc[1], num: +epSc[2], suf: (epSc[3] || '').toUpperCase(), part: null }];

    let part = null;
    // "pt", "pt1", "PT2/2", "pt 1 of 2", "pr2/2" (typo) — part markers, not scenes
    s = s.replace(/p[tr](?![a-z])\.?(?:\s*(\d+)(?:\s*(?:\/|of)\s*(\d+))?|\s*of\s*(\d+))?/gi, (m, p, n, n2) => {
      if (!part) part = { p: p ? +p : null, n: n || n2 ? +(n || n2) : null };
      return ' ';
    });
    const out = [];
    const tokRe = /\(?\b([1-9])\s*(?:\/|of)\s*([1-9])\b\)?|(\d+)(?:\.(\d+))?([A-Za-z])?(?![A-Za-z])/g;
    let m;
    while ((m = tokRe.exec(s))) {
      if (m[1]) {
        // "97 2/2", "38 (2/2)" — a bare small fraction after a scene is a part marker
        if (out.length) { if (!part) part = { p: +m[1], n: +m[2] }; continue; }
        m = [m[0], null, null, m[1], null, null];
      }
      m = [m[0], m[3], m[4], m[5]];
      let ep = null, num = +m[1];
      if (m[2] != null) { ep = +m[1]; num = +m[2]; }
      else if (out.length && out[0].ep != null) {
        // "6.24/625" — typo'd second scene that dropped the dot
        const e = String(out[0].ep);
        ep = out[0].ep;
        if (m[1].length === e.length + 2 && m[1].startsWith(e)) num = +m[1].slice(e.length);
      }
      out.push({ ep, num, suf: (m[3] || '').toUpperCase(), part: null });
    }
    for (const o of out) o.part = part;
    return out;
  }

  const keyOf = (o) => `${o.ep ?? ''}|${o.num}|${o.suf}`;

  // ---------- PDF schedule parsing ----------

  async function pdfLines(pdf) {
    const lines = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      const items = tc.items
        .filter((i) => i.str && i.str.trim())
        .map((i) => ({ s: i.str.replace(/\s+/g, ' ').trim(), x: i.transform[4], y: i.transform[5] }));
      items.sort((a, b) => b.y - a.y || a.x - b.x);
      const pageLines = [];
      for (const it of items) {
        // Chain on the lowest item so far: small labels sit a point or two off the heading baseline
        const L = pageLines.find((l) => Math.abs(l.low - it.y) <= 3.5);
        if (L) { L.items.push(it); L.low = Math.min(L.low, it.y); } else pageLines.push({ y: it.y, low: it.y, items: [it] });
      }
      pageLines.sort((a, b) => b.y - a.y);
      for (const l of pageLines) {
        l.items.sort((a, b) => a.x - b.x);
        lines.push({ page: p, items: l.items, text: l.items.map((i) => i.s).join(' ') });
      }
    }
    return lines;
  }

  function parseDate(text, fallbackYear) {
    const m = text.match(/\b(\d{1,2})(?:st|nd|rd|th)?[\s\-]+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?(?:[\s\-,]+(\d{4}))?/i);
    if (!m) return null;
    const y = m[3] ? +m[3] : fallbackYear;
    return { d: +m[1], m: MONTHS[m[2].toLowerCase()], y: y ?? null };
  }

  // Handles "UNIT CALL 0800 EST WRAP 1800", "CALLTIME: 0800 - 1800", "07:30 Call", "08:00 - 18:00 TBC"
  function parseTimes(text) {
    if (/\bS[RST]\b\s*[-:]/i.test(text) && !/call|wrap/i.test(text)) return null;
    const hhmm = (h, m) => `${h.padStart(2, '0')}:${m}`;
    let call = null, wrap = null, m;
    if ((m = text.match(/\b(?:unit\s+)?call(?:\s*time)?\s*:?\s*(\d{1,2}):?(\d{2})\b(?:\s*-\s*(\d{1,2}):?(\d{2})\b)?/i))) {
      call = hhmm(m[1], m[2]); if (m[3]) wrap = hhmm(m[3], m[4]);
    } else if ((m = text.match(/\b(\d{1,2}):(\d{2})\s*call\b/i))) {
      call = hhmm(m[1], m[2]);
    } else if ((m = text.match(/^(?:night\s+shoot:?\s*)?(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/i))) {
      call = hhmm(m[1], m[2]); wrap = hhmm(m[3], m[4]);
    }
    if ((m = text.match(/\bwrap\s*:?\s*(\d{1,2}):?(\d{2})\b/i))) wrap = hhmm(m[1], m[2]);
    return call || wrap ? { call, wrap } : null;
  }

  const TOD_RE = /^(day|night|morn(?:ing)?|evening|dawn|dusk|sunset|sunrise|afternoon)\b/i;

  function sceneLineOf(line, next) {
    const items = line.items;
    for (let k = 1; k < items.length; k++) {
      const t = items[k].s;
      const first = t.split(' ')[0];
      if (!(IE_RE.test(t) || IE_RE.test(first))) continue;
      const field = items.slice(0, k).map((i) => i.s).join(' ');
      if (!/^\d/.test(field)) return null;
      const ie = IE_RE.test(t) ? t : first;
      const set = IE_RE.test(t) ? (items[k + 1] ? items[k + 1].s : '') : t.slice(first.length).trim();
      // Some layouts put Day/Night on the line below the scene heading
      const tod = items.slice(k + 1).concat(next ? next.items.slice(0, 2) : []).map((i) => i.s.match(TOD_RE)).find(Boolean);
      return { field, ie: ie.toUpperCase(), set, tod: tod ? tod[1] : '' };
    }
    return null;
  }

  async function parseSchedule(pdf) {
    const lines = await pdfLines(pdf);
    const days = [];
    let pending = [];       // scenes since the last "end of day"
    let header = null;      // most recent day heading text
    let call = null, wrap = null;
    let order = 0;
    const warnings = [];

    const endRe = /end\s+of\s+(shooting\s+)?(pre[\s\-]?shoot\s+day|day\s*(\d+)|shoot(?:ing)?\s+day\s*(\d+))/i;
    const headRe = /(pre[\s\-]?shoot\s+day|shoot(?:ing)?\s+day\s*\d+|^\W*day\s+\d+\b)/i;

    for (let li = 0; li < lines.length; li++) {
      const line = lines[li];
      const t = line.text.trim();
      const end = /--|total|pages/i.test(t) ? t.match(endRe) : null;
      if (end) {
        const pre = /pre[\s\-]?shoot/i.test(end[2]);
        const n = end[3] || end[4] || (end[2].match(/\d+/) || [])[0];
        const label = pre ? 'Pre-shoot' : `Day ${n}`;
        let date = parseDate(t.slice(end.index + end[0].length));
        if (!date && header) date = parseDate(header);
        days.push({ label, date, call, wrap, scenes: pending });
        pending = []; call = null; wrap = null; header = null;
        continue;
      }
      if (headRe.test(t) && !/^\d/.test(t) && t.length < 120) {
        header = t;
        const tm = parseTimes(t);
        if (tm) { call = tm.call || call; wrap = tm.wrap || wrap; }
        continue;
      }
      if (!call || !wrap) {
        const tm = parseTimes(t);
        if (tm) { call = call || tm.call; wrap = wrap || tm.wrap; continue; }
      }
      const sc = sceneLineOf(line, lines[li + 1]);
      if (sc) {
        const parsed = parseSceneField(sc.field);
        if (!parsed.length) continue;
        pending.push({ ...sc, parsed, order: order++, page: line.page });
      }
    }
    if (pending.length) {
      // Whatever follows the last "end of day" is the unscheduled / 2nd unit / to-be-scheduled list
      days.push({ label: 'Unscheduled', date: null, call: null, wrap: null, scenes: pending, unscheduled: true });
    }

    // Fill missing years from the nearest dated day
    for (let i = 0; i < days.length; i++) {
      const d = days[i].date;
      if (d && d.y == null) {
        const near = days.slice(i + 1).concat(days.slice(0, i).reverse()).find((x) => x.date && x.date.y != null);
        d.y = near ? near.date.y : new Date().getFullYear();
      }
      days[i].jsDate = d ? new Date(Date.UTC(d.y, d.m, d.d)) : null;
    }

    const occ = new Map();
    days.forEach((day, di) => {
      day.index = di;
      day.scenes.forEach((sc) => {
        for (const p of sc.parsed) {
          const k = keyOf(p);
          if (!occ.has(k)) occ.set(k, []);
          occ.get(k).push({ day, dayIndex: di, order: sc.order, part: p.part, field: sc.field, ie: sc.ie, set: sc.set, tod: sc.tod });
        }
      });
    });
    const episodic = days.some((d) => d.scenes.some((s) => s.parsed.some((p) => p.ep != null)));
    return { days, occ, episodic, warnings, sceneCount: order };
  }

  // ---------- Excel processing ----------

  const NEW_COLS = ['Shoot Day', 'Shoot Date', 'Call Time', 'Wrap Time', 'Schedule Order'];
  const DEFAULT_SCENE_HEADER = /^\s*(sc|scn|scene)\b\.?/i;

  function cellText(v) {
    if (v == null) return '';
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map((r) => r.text).join('');
      if (v.text != null) return String(v.text);
      if (v.result != null) return String(v.result);
      if (v instanceof Date) return v.toISOString();
      return '';
    }
    return String(v);
  }

  const normHeader = (s) => String(s).toLowerCase().replace(/[^a-z0-9#]+/g, ' ').trim();

  // columnName (optional) is the user's own header text, compared ignoring case, spaces and punctuation
  function findHeader(ws, columnName) {
    const want = columnName && columnName.trim() ? normHeader(columnName) : null;
    for (let r = 1; r <= Math.min(20, ws.rowCount); r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= Math.max(row.cellCount, 1); c++) {
        const text = cellText(row.getCell(c).value);
        if (want ? normHeader(text) === want : DEFAULT_SCENE_HEADER.test(text)) return { row: r, col: c };
      }
    }
    return null;
  }

  function sheetEpisode(ws) {
    const m = ws.name.match(/\bep(?:isode)?\.?\s*(\d+)\b/i);
    return m ? +m[1] : null;
  }

  function rowHasContent(row, maxCol) {
    for (let c = 1; c <= maxCol; c++) if (cellText(row.getCell(c).value).trim()) return true;
    return false;
  }

  function matchScene(sched, parsed, rowEp) {
    const hits = [];
    for (const p of parsed) {
      const ep = p.ep ?? rowEp;
      let list = sched.occ.get(keyOf({ ...p, ep: sched.episodic ? ep : null })) || [];
      // A lettered breakdown row (7a) whose scene the schedule only lists plain (7) takes the plain scene's days
      if (!list.length && p.suf) list = sched.occ.get(keyOf({ ...p, suf: '', ep: sched.episodic ? ep : null })) || [];
      if (p.part && p.part.p != null) {
        const same = list.filter((o) => o.part && o.part.p === p.part.p);
        if (same.length) list = same;
      }
      hits.push(...list);
    }
    const seen = new Set();
    return hits.filter((h) => (seen.has(h.order) ? false : seen.add(h.order))).sort((a, b) => a.order - b.order);
  }

  function processWorkbook(wb, sched, opts = {}) {
    const report = { sheets: [], unmatchedSchedule: [], newScenes: [] };
    const usedOrders = new Set();
    const sheets = [];

    wb.eachSheet((ws) => {
      const hdr = findHeader(ws, opts.columnName);
      if (!hdr) { sheets.push({ ws, skipped: true }); return; }
      const maxCol = Math.max(ws.columnCount, ws.getRow(hdr.row).cellCount);
      let last = hdr.row;
      for (let r = ws.rowCount; r > hdr.row; r--) if (rowHasContent(ws.getRow(r), maxCol)) { last = r; break; }
      const sheetEp = sheetEpisode(ws);
      const epCol = (() => {
        const row = ws.getRow(hdr.row);
        for (let c = 1; c <= maxCol; c++) if (/^\s*ep(isode)?\b/i.test(cellText(row.getCell(c).value))) return c;
        return null;
      })();

      // Rows without a scene number belong to the scene above them
      const blocks = [];
      for (let r = hdr.row + 1; r <= last; r++) {
        const raw = cellText(ws.getRow(r).getCell(hdr.col).value).trim();
        const parsed = raw ? parseSceneField(raw) : [];
        if (raw || !blocks.length) {
          let rowEp = sheetEp;
          if (epCol) { const e = parseInt(cellText(ws.getRow(r).getCell(epCol).value), 10); if (!isNaN(e)) rowEp = e; }
          blocks.push({ rows: [r], raw, parsed, rowEp, idx: blocks.length });
        } else blocks[blocks.length - 1].rows.push(r);
      }
      sheets.push({ ws, hdr, maxCol, last, blocks });
    });

    // A lettered schedule scene (128A) with no row of its own is a scene newly added to the script:
    // it gets a new, highlighted row in shooting order so it can be read and broken down
    const used = sheets.filter((sh) => !sh.skipped);
    const keysIn = (sh) => {
      const s = new Set();
      for (const b of sh.blocks) for (const p of b.parsed) s.add(keyOf({ ...p, ep: sched.episodic ? (p.ep ?? b.rowEp) : null }));
      return s;
    };
    used.forEach((sh) => { sh.keys = keysIn(sh); });
    for (const [k, list] of sched.occ) {
      const [ep, num, suf] = k.split('|');
      if (!suf || used.some((sh) => sh.keys.has(k))) continue;
      // Goes on the sheet(s) holding the plain scene, else the episode's sheet(s), else the first sheet
      const epOf = (sh) => sh.blocks.some((b) => b.parsed.some((p) => String(p.ep ?? b.rowEp ?? '') === ep));
      let targets = used.filter((sh) => sh.keys.has(`${ep}|${num}|`));
      if (!targets.length && ep) targets = used.filter(epOf);
      if (!targets.length && !ep) targets = used.slice(0, 1);
      if (!targets.length) continue;
      const scene = (ep ? ep + '.' : '') + num + suf;
      for (const sh of targets) {
        sh.blocks.push({ rows: [], raw: scene, parsed: [], rowEp: ep ? +ep : null, idx: sh.blocks.length, added: list[0], hits: list.slice() });
      }
      report.newScenes.push({ scene, ...list[0], days: [...new Set(list.map((o) => o.day.label))], sheets: targets.map((sh) => sh.ws.name) });
    }
    report.newScenes.sort((a, b) => a.order - b.order);

    for (const sh of sheets) {
      if (sh.skipped) { report.sheets.push({ name: sh.ws.name, skipped: true }); continue; }
      const { ws, hdr, maxCol, last, blocks } = sh;
      let matched = 0;
      const unmatched = [];
      for (const b of blocks) {
        if (b.added) { b.hits.forEach((h) => usedOrders.add(h.order)); continue; }
        b.hits = b.parsed.length ? matchScene(sched, b.parsed, b.rowEp) : [];
        if (b.hits.length) { matched++; b.hits.forEach((h) => usedOrders.add(h.order)); }
        else if (b.raw) unmatched.push(b.raw);
      }
      const sorted = blocks.slice().sort((a, b) => {
        const ao = a.hits.length ? a.hits[0].order : Infinity;
        const bo = b.hits.length ? b.hits[0].order : Infinity;
        return ao - bo || a.idx - b.idx;
      });

      rewriteSheet(ws, hdr, last, maxCol, sorted);
      report.sheets.push({
        name: ws.name, skipped: false, scenes: blocks.filter((b) => b.raw && !b.added).length, matched, unmatched,
        rows: sorted.map((b) => ({ raw: b.raw, ep: b.rowEp, hits: b.hits })),
      });
    }

    sched.days.forEach((d) => d.scenes.forEach((s) => { if (!usedOrders.has(s.order)) report.unmatchedSchedule.push({ day: d, field: s.field }); }));
    report.usedOrders = usedOrders;
    return report;
  }

  function snapshotRow(ws, r, maxCol) {
    const row = ws.getRow(r);
    const cells = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      let v = cell.value;
      // Moving a formula would silently re-point its references; keep the computed value instead
      if (v && typeof v === 'object' && (v.formula || v.sharedFormula)) v = v.result ?? null;
      cells.push({ v, style: JSON.parse(JSON.stringify(cell.style || {})) });
    }
    return { cells, height: row.height, hidden: row.hidden, outline: row.outlineLevel };
  }

  function rewriteSheet(ws, hdr, last, maxCol, sorted) {
    const ins = hdr.col + 1; // new columns go right after the scene column
    const shift = (c) => (c >= ins ? c + NEW_COLS.length : c);
    const N = NEW_COLS.length;

    // Old row -> new row mapping
    const rowMap = new Map();
    for (let r = 1; r <= hdr.row; r++) rowMap.set(r, r);
    let nr = hdr.row + 1;
    const plan = [];
    for (const b of sorted) {
      if (b.added) { plan.push({ from: null, to: nr++, block: b, first: true }); continue; }
      b.rows.forEach((r, i) => { rowMap.set(r, nr); plan.push({ from: r, to: nr, block: b, first: i === 0 }); nr++; });
    }

    const snaps = new Map();
    for (let r = 1; r <= last; r++) snaps.set(r, snapshotRow(ws, r, maxCol));

    // Column widths shift with the cells
    const colProps = [];
    for (let c = 1; c <= maxCol; c++) { const col = ws.getColumn(c); colProps.push({ width: col.width, hidden: col.hidden }); }

    // Merges and validations are keyed by address — remap them
    const merges = (ws.model.merges || []).slice();
    merges.forEach((m) => { try { ws.unMergeCells(m); } catch (e) {} });
    const dv = ws.dataValidations && ws.dataValidations.model ? { ...ws.dataValidations.model } : {};
    if (ws.dataValidations) ws.dataValidations.model = {};

    for (let r = 1; r <= last; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= maxCol + N; c++) { const cell = row.getCell(c); cell.value = null; cell.style = {}; }
    }

    const write = (to, snap) => {
      const row = ws.getRow(to);
      snap.cells.forEach((cl, i) => { const cell = row.getCell(shift(i + 1)); cell.value = cl.v; cell.style = cl.style; });
      row.height = snap.height; row.hidden = snap.hidden;
    };
    for (let r = 1; r <= hdr.row; r++) write(r, snaps.get(r));

    const hdrSnap = snaps.get(hdr.row);
    const hdrStyle = hdrSnap.cells[hdr.col - 1].style;
    NEW_COLS.forEach((name, i) => {
      const cell = ws.getRow(hdr.row).getCell(ins + i);
      cell.value = name;
      cell.style = JSON.parse(JSON.stringify(hdrStyle));
    });

    for (const p of plan) {
      const snap = p.from != null ? snaps.get(p.from) : newSceneSnap(p.block, maxCol, hdr.col);
      write(p.to, snap);
      if (p.block.added) {
        const a = p.block.added;
        ws.getRow(p.to).getCell(hdr.col).note = `NEW SCENE added to the script — read it and break it down.\n${[a.ie, a.set, a.tod].filter(Boolean).join(' ')}`;
      }
      const baseStyle = snap.cells[hdr.col - 1].style;
      const row = ws.getRow(p.to);
      const vals = p.first ? newColValues(p.block) : NEW_COLS.map(() => null);
      vals.forEach((v, i) => {
        const cell = row.getCell(ins + i);
        cell.value = v;
        const st = JSON.parse(JSON.stringify(baseStyle));
        if (i === 1 && v instanceof Date) st.numFmt = 'ddd dd mmm yyyy';
        st.alignment = { ...(st.alignment || {}), wrapText: true, vertical: (st.alignment && st.alignment.vertical) || 'top' };
        cell.style = st;
      });
    }

    for (let c = maxCol; c >= 1; c--) {
      const col = ws.getColumn(shift(c));
      col.width = colProps[c - 1].width; col.hidden = colProps[c - 1].hidden;
    }
    [16, 22, 11, 11, 10].forEach((w, i) => { ws.getColumn(ins + i).width = w; ws.getColumn(ins + i).hidden = false; });

    const decode = (a) => { const m = a.match(/^\$?([A-Z]+)\$?(\d+)$/); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { c, r: +m[2] }; };
    const encode = (c, r) => { let s = ''; while (c > 0) { const k = (c - 1) % 26; s = String.fromCharCode(65 + k) + s; c = Math.floor((c - 1) / 26); } return s + r; };
    const moveAddr = (a) => { const { c, r } = decode(a); return encode(shift(c), rowMap.get(r) ?? r); };

    for (const m of merges) {
      const [a, b] = m.split(':');
      const A = decode(a), B = decode(b || a);
      // Only single-row merges can travel with a re-sorted row
      if (A.r !== B.r && A.r > hdr.row) continue;
      try { ws.mergeCells(`${moveAddr(a)}:${moveAddr(b || a)}`); } catch (e) {}
    }
    if (ws.dataValidations) {
      const out = {};
      for (const [addr, rule] of Object.entries(dv)) {
        if (/^[A-Z]+\d+$/.test(addr)) out[moveAddr(addr)] = rule; else out[addr] = rule;
      }
      ws.dataValidations.model = out;
    }
    if (ws.autoFilter) ws.autoFilter = null;
  }

  const NEW_SCENE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB84' } };

  // A breakdown row for a scene the breakdown doesn't have yet: just the scene number, highlighted
  function newSceneSnap(block, maxCol, sceneCol) {
    const cells = [];
    for (let c = 1; c <= maxCol; c++) {
      cells.push({ v: c === sceneCol ? block.raw : null, style: { fill: NEW_SCENE_FILL, font: c === sceneCol ? { bold: true } : undefined } });
    }
    return { cells, height: undefined, hidden: false };
  }

  function newColValues(block) {
    const hits = block.hits;
    if (!block.raw) return NEW_COLS.map(() => null);
    if (!hits.length) return ['Not in schedule', null, null, null, null];
    const days = [];
    for (const h of hits) if (!days.includes(h.day)) days.push(h.day);
    const partOf = (d) => { const h = hits.find((x) => x.day === d); return h.part ? ` (pt${h.part.p != null ? ' ' + h.part.p : ''}${h.part.n ? '/' + h.part.n : ''})` : ''; };
    const label = days.length === 1 ? days[0].label : days.map((d) => d.label + partOf(d)).join(' + ');
    const first = days[0];
    return [label, first.jsDate || null, first.call || null, first.wrap || null, hits[0].order + 1];
  }

  g.SchedCore = { parseSceneField, parseSchedule, processWorkbook, pdfLines, NEW_COLS };
})(typeof window !== 'undefined' ? window : globalThis);
