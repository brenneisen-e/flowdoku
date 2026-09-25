/**
 * v31.96: Concur-Attendee-Import — schreibt Personen IN die Original-Vorlage
 * (AttendeeImportTemplate.xls), statt eine neue Datei zu bauen.
 *
 * **Warum kein XLSX.read/XLSX.write:** Die Vorlage ist eine BIFF8-Datei mit
 * Anleitungsblatt samt Bildern, grünen Pflicht-Köpfen, einem ausgeblendeten
 * Blatt „Properties" und der Typ-Liste an Position 2. SheetJS (Community)
 * liest davon die Werte und schreibt eine neue Mappe — Bilder, Farben und das
 * versteckte Blatt wären weg, und die Anleitung sagt ausdrücklich: Blattfolge
 * und Aufbau nicht verändern. Deshalb wird hier nur der Workbook-Stream
 * gepatcht: Die Zelltabelle des ERSTEN Blatts wird neu geschrieben, alles
 * andere bleibt Byte für Byte, wie es war.
 *
 * Was sich dabei verschiebt und nachgezogen werden muss:
 *  - INDEX des Blatts (Zeilenbereich, Offset von DEFCOLWIDTH, DBCELL-Offsets)
 *  - DIMENSIONS (letzte Zeile)
 *  - DBCELL je 32er-Zeilenblock (relative Offsets)
 *  - BOUNDSHEET im Globals-Teil: die Startposition jedes FOLGENDEN Blatts
 *  - INDEX-Records der folgenden Blätter (absolute Offsets)
 * Den Container (Compound File) schreibt `XLSX.CFB` neu; die anderen Streams
 * (Summary-Info, MsoDataStore) werden unverändert mitgenommen.
 */

/* eslint-disable no-bitwise */

const BOF = 0x0809;
const EOF = 0x000a;
const BOUNDSHEET = 0x0085;
const INDEX = 0x020b;
const DEFCOLWIDTH = 0x0055;
const DIMENSIONS = 0x0200;
const ROW = 0x0208;
const DBCELL = 0x00d7;
const LABEL = 0x0204;
const BLANK = 0x0201;
const MULBLANK = 0x00be;
const LABELSST = 0x00fd;
const RK = 0x027e;
const NUMBER = 0x0203;
const FORMULA = 0x0006;
const STRING = 0x0207;
const BOOLERR = 0x0205;
const MULRK = 0x00bd;

/** Zelltabellen-Records: ROW, Zellen, DBCELL. Alles andere ist Kopf oder Fuß des Blatts. */
const CELL_TABLE = [ROW, DBCELL, LABEL, BLANK, MULBLANK, LABELSST, RK, NUMBER, FORMULA, STRING, BOOLERR, MULRK];

/** Die Spalten der Vorlage, Zeile 1 — daran wird die Datei als die richtige erkannt. */
export const CONCUR_ATTENDEE_COLUMNS = ['AtnTypeKey', 'FirstName', 'LastName', 'Title', 'Company', 'Custom5'];

/** Concur nimmt höchstens 500 Personen je Beleg (Anleitungsblatt, Punkt 1). */
export const CONCUR_MAX_ATTENDEES = 500;

interface Rec { type: number; start: number; data: Uint8Array; }

function readRecords(buf: Uint8Array, from: number, to: number): Rec[] {
  const out: Rec[] = [];
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let p = from;
  while (p + 4 <= to) {
    const type = dv.getUint16(p, true);
    const len = dv.getUint16(p + 2, true);
    out.push({ type, start: p, data: buf.subarray(p + 4, p + 4 + len) });
    p += 4 + len;
    if (type === EOF && out.length > 1) {
      // Ein Substream endet mit seinem EOF — danach beginnt das nächste Blatt.
      const bofCount = out.filter(r => r.type === BOF).length;
      if (bofCount === 1) break;
    }
  }
  return out;
}

function rec(type: number, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(4 + data.length);
  const dv = new DataView(out.buffer);
  dv.setUint16(0, type, true);
  dv.setUint16(2, data.length, true);
  out.set(data, 4);
  return out;
}

function recBytes(r: Rec): Uint8Array { return rec(r.type, r.data); }

function concat(parts: Uint8Array[]): Uint8Array {
  let n = 0;
  parts.forEach(p => { n += p.length; });
  const out = new Uint8Array(n);
  let o = 0;
  parts.forEach(p => { out.set(p, o); o += p.length; });
  return out;
}

function u16(dv: DataView, o: number): number { return dv.getUint16(o, true); }

/** LABEL (BIFF8): Zeile, Spalte, XF, dann XLUnicodeString — immer UTF-16, dann ist kein Zeichen ein Sonderfall. */
function labelRec(row: number, col: number, xf: number, text: string): Uint8Array {
  const s = text.length > 255 ? text.slice(0, 255) : text;
  const data = new Uint8Array(6 + 3 + s.length * 2);
  const dv = new DataView(data.buffer);
  dv.setUint16(0, row, true);
  dv.setUint16(2, col, true);
  dv.setUint16(4, xf, true);
  dv.setUint16(6, s.length, true);
  data[8] = 1;
  for (let i = 0; i < s.length; i++) dv.setUint16(9 + i * 2, s.charCodeAt(i), true);
  return rec(LABEL, data);
}

function blankRec(row: number, col: number, xf: number): Uint8Array {
  const data = new Uint8Array(6);
  const dv = new DataView(data.buffer);
  dv.setUint16(0, row, true);
  dv.setUint16(2, col, true);
  dv.setUint16(4, xf, true);
  return rec(BLANK, data);
}

export interface ConcurXlsResult { bytes: Uint8Array; rowsWritten: number; }

/**
 * Schreibt `rows` (je sechs Spalten, Reihenfolge wie CONCUR_ATTENDEE_COLUMNS)
 * ab Zeile 3 in das erste Blatt der Vorlage. Die Kopfzeilen 1 und 2 bleiben
 * unverändert. Wirft mit deutscher Meldung, wenn die Datei nicht die
 * erwartete Vorlage ist — lieber kein Download als eine Datei, die Concur
 * erst beim Hochladen ablehnt.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function fillConcurAttendeeTemplate(CFB: any, template: Uint8Array, rows: string[][]): ConcurXlsResult {
  const cfb = CFB.read(template, { type: 'array' });
  const entry = CFB.find(cfb, 'Workbook') || CFB.find(cfb, 'Book');
  if (!entry || !entry.content) throw new Error('Die Vorlage enthält keinen Workbook-Stream — ist es wirklich die .xls-Vorlage?');
  const wb: Uint8Array = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content);
  const wdv = new DataView(wb.buffer, wb.byteOffset, wb.byteLength);

  // 1) Globals: BOUNDSHEETs einsammeln (Offset des lbPlyPos im Stream merken).
  const globals = readRecords(wb, 0, wb.length);
  const sheets: Array<{ recStart: number; pos: number }> = [];
  globals.forEach(r => {
    if (r.type === BOUNDSHEET) sheets.push({ recStart: r.start, pos: wdv.getUint32(r.start + 4, true) });
  });
  if (sheets.length === 0) throw new Error('Die Vorlage hat keine Blätter.');
  const first = sheets.reduce((a, b) => (a.pos < b.pos ? a : b));
  const sheetStart = first.pos;

  // 2) Das erste Blatt zerlegen: Kopf | Zelltabelle | Fuß.
  const sheet = readRecords(wb, sheetStart, wb.length);
  const last = sheet[sheet.length - 1];
  if (sheet[0].type !== BOF || last.type !== EOF) throw new Error('Das erste Blatt der Vorlage ist nicht lesbar.');
  const sheetEnd = last.start + 4 + last.data.length;
  const firstCell = sheet.findIndex(r => CELL_TABLE.indexOf(r.type) >= 0);
  let lastCell = -1;
  sheet.forEach((r, i) => { if (CELL_TABLE.indexOf(r.type) >= 0) lastCell = i; });
  if (firstCell < 0) throw new Error('Das erste Blatt der Vorlage hat keine Zellen.');
  const head = sheet.slice(0, firstCell);
  const tail = sheet.slice(lastCell + 1);

  // 3) Bestand lesen: ROW-Records je Zeile, Zellen je Zeile, XF je Zelle.
  const rowRec: Record<number, Uint8Array> = {};
  const cellsOf: Record<number, Uint8Array[]> = {};
  const xfAt: Record<string, number> = {};
  let maxRow = -1;
  for (let i = firstCell; i <= lastCell; i++) {
    const r = sheet[i];
    const dv = new DataView(r.data.buffer, r.data.byteOffset, r.data.byteLength);
    if (r.type === ROW) {
      const rw = u16(dv, 0);
      rowRec[rw] = new Uint8Array(r.data);
      if (rw > maxRow) maxRow = rw;
      continue;
    }
    if (r.type === DBCELL) continue;
    const rw = u16(dv, 0);
    if (rw > maxRow) maxRow = rw;
    if (r.type === MULBLANK) {
      const c0 = u16(dv, 2);
      const n = (r.data.length - 6) / 2;
      for (let k = 0; k < n; k++) xfAt[`${rw}:${c0 + k}`] = u16(dv, 4 + k * 2);
    } else {
      xfAt[`${rw}:${u16(dv, 2)}`] = u16(dv, 4);
    }
    (cellsOf[rw] = cellsOf[rw] || []).push(recBytes(r));
  }

  // Die Kopfzeile muss die Concur-Spalten tragen — sonst ist es eine andere Datei.
  // (Die Texte stehen im SST; hier genügt, dass Zeile 0 und 1 existieren und
  // die Prüfung auf die Spaltennamen läuft vorher über XLSX.read im Aufrufer.)
  if (!rowRec[0] || !rowRec[1]) throw new Error('Die Kopfzeilen der Vorlage fehlen.');

  const DATA_FROM = 2;
  const templateRowFor = (rw: number): Uint8Array => rowRec[rw] || rowRec[DATA_FROM] || rowRec[1];
  const xfFor = (rw: number, col: number): number => {
    const k = `${rw}:${col}`;
    if (xfAt[k] !== undefined) return xfAt[k];
    // Neue Zeilen unterhalb der Vorlage: dieselben Formate wie die erste Datenzeile.
    const k2 = `${DATA_FROM}:${col}`;
    return xfAt[k2] !== undefined ? xfAt[k2] : 15;
  };

  const lastRow = Math.max(maxRow, DATA_FROM + rows.length - 1);

  // 4) Zeilen erzeugen: Kopf unverändert, Datenzeilen neu, leere Vorlagenzeilen
  // mit ihren Leerzellen (Format bleibt, z.B. Textformat in Spalte A).
  const rowCells: Uint8Array[][] = [];
  for (let rw = 0; rw <= lastRow; rw++) {
    if (rw < DATA_FROM) { rowCells.push(cellsOf[rw] || []); continue; }
    const vals = rows[rw - DATA_FROM];
    const cells: Uint8Array[] = [];
    for (let c = 0; c < CONCUR_ATTENDEE_COLUMNS.length; c++) {
      const v = vals ? String(vals[c] || '').trim() : '';
      if (v) cells.push(labelRec(rw, c, xfFor(rw, c), v));
      else if (xfAt[`${rw}:${c}`] !== undefined) cells.push(blankRec(rw, c, xfAt[`${rw}:${c}`]));
    }
    // DBCELL verlangt je Zeile einen Zellanfang — eine ganz leere Zeile
    // bekommt deshalb eine Leerzelle in Spalte A.
    if (cells.length === 0) cells.push(blankRec(rw, 0, xfFor(rw, 0)));
    rowCells.push(cells);
  }
  const rowRecFor = (rw: number): Uint8Array => {
    const d = new Uint8Array(templateRowFor(rw));
    const dv = new DataView(d.buffer);
    dv.setUint16(0, rw, true);
    // colMic/colMac aus den tatsächlich geschriebenen Zellen.
    let cmin = 0xffff, cmax = 0;
    rowCells[rw].forEach(c => {
      const cdv = new DataView(c.buffer, c.byteOffset, c.byteLength);
      const t = cdv.getUint16(0, true);
      const col = cdv.getUint16(6, true);
      let colEnd = col + 1;
      if (t === MULBLANK) colEnd = col + (c.length - 4 - 6) / 2;
      if (col < cmin) cmin = col;
      if (colEnd > cmax) cmax = colEnd;
    });
    if (cmin === 0xffff) cmin = 0;
    dv.setUint16(2, cmin, true);
    dv.setUint16(4, cmax, true);
    return rec(ROW, d);
  };

  // 5) Kopf neu: INDEX braucht die Anzahl Blöcke, bevor Offsets feststehen.
  const blocks = Math.ceil((lastRow + 1) / 32);
  const indexLen = 16 + 4 * blocks;
  const headParts: Uint8Array[] = [];
  let indexAt = -1;
  let defColAt = -1;
  let dimAt = -1;
  let off = sheetStart;
  head.forEach(r => {
    let bytes: Uint8Array;
    if (r.type === INDEX) { indexAt = off; bytes = new Uint8Array(4 + indexLen); }
    else if (r.type === DIMENSIONS) {
      dimAt = off;
      const d = new Uint8Array(r.data);
      new DataView(d.buffer).setUint32(4, lastRow + 1, true);
      bytes = rec(DIMENSIONS, d);
    } else {
      if (r.type === DEFCOLWIDTH) defColAt = off;
      bytes = recBytes(r);
    }
    headParts.push(bytes);
    off += bytes.length;
  });
  if (dimAt < 0) throw new Error('Die Vorlage hat keinen DIMENSIONS-Eintrag.');

  // 6) Zelltabelle mit DBCELLs.
  const tableParts: Uint8Array[] = [];
  const dbcellAt: number[] = [];
  for (let b = 0; b < blocks; b++) {
    const r0 = b * 32;
    const r1 = Math.min(lastRow, r0 + 31);
    const firstRowAt = off;
    for (let rw = r0; rw <= r1; rw++) { const x = rowRecFor(rw); tableParts.push(x); off += x.length; }
    const secondRowAt = firstRowAt + 20;
    const cellStarts: number[] = [];
    for (let rw = r0; rw <= r1; rw++) {
      cellStarts.push(off);
      rowCells[rw].forEach(c => { tableParts.push(c); off += c.length; });
    }
    const n = r1 - r0 + 1;
    const d = new Uint8Array(4 + 2 * n);
    const dv = new DataView(d.buffer);
    dv.setUint32(0, off - firstRowAt, true);
    for (let k = 0; k < n; k++) {
      const rel = k === 0 ? cellStarts[0] - secondRowAt : cellStarts[k] - cellStarts[k - 1];
      if (rel > 0xffff) throw new Error('Zu viele Daten in einem Zeilenblock.');
      dv.setUint16(4 + k * 2, rel, true);
    }
    dbcellAt.push(off);
    const x = rec(DBCELL, d);
    tableParts.push(x);
    off += x.length;
  }
  const tailParts = tail.map(recBytes);
  tailParts.forEach(x => { off += x.length; });
  const newSheetEnd = off;
  const delta = newSheetEnd - sheetEnd;

  const newSheet = concat([...headParts, ...tableParts, ...tailParts]);
  if (indexAt >= 0) {
    const d = new Uint8Array(indexLen);
    const dv = new DataView(d.buffer);
    dv.setUint32(4, 0, true);
    dv.setUint32(8, lastRow + 1, true);
    dv.setUint32(12, defColAt >= 0 ? defColAt : 0, true);
    dbcellAt.forEach((p, i) => dv.setUint32(16 + 4 * i, p, true));
    newSheet.set(rec(INDEX, d), indexAt - sheetStart);
  }

  // 7) Stream zusammensetzen und die Offsets der folgenden Blätter nachziehen.
  const before = new Uint8Array(wb.subarray(0, sheetStart));
  const after = new Uint8Array(wb.subarray(sheetEnd));
  const adv = new DataView(after.buffer);
  const bdv = new DataView(before.buffer);
  sheets.forEach(s => {
    if (s.pos > sheetStart) bdv.setUint32(s.recStart + 4, s.pos + delta, true);
  });
  // INDEX-Records der folgenden Blätter tragen absolute Offsets.
  let p = 0;
  while (p + 4 <= after.length) {
    const t = adv.getUint16(p, true);
    const len = adv.getUint16(p + 2, true);
    if (t === INDEX && len >= 16) {
      const def = adv.getUint32(p + 4 + 12, true);
      if (def) adv.setUint32(p + 4 + 12, def + delta, true);
      for (let q = 16; q + 4 <= len; q += 4) {
        const v = adv.getUint32(p + 4 + q, true);
        if (v) adv.setUint32(p + 4 + q, v + delta, true);
      }
    }
    p += 4 + len;
  }
  const out = concat([before, newSheet, after]);

  entry.content = out;
  entry.size = out.length;
  const bytes = CFB.write(cfb, { type: 'array', fileType: 'cfb' });
  return { bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), rowsWritten: rows.length };
}
