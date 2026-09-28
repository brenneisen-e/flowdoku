/**
 * v32.2.2: Auswahl-Felder, deren Antworten Daten sind, werden auf der
 * Anmeldeseite als Kalender gezeigt. Nutzer-Ansage 28.09.2026: „Kann man
 * Auswahl / Dropdown so anpassen, dass es als Kalender gerendert wird, wenn
 * man hier Daten angibt?"
 *
 * Bewusst KEIN eigener Feldtyp und kein Schalter: Erkannt wird, wenn JEDE
 * Antwortmöglichkeit genau ein Datum ist („19.11.2026", „Do, 19.11.",
 * „2026-11-19", „19. November 2026", „Thu 19 Nov"). Eine Antwort mit Text
 * drumherum („19 November only", „Both days") ist kein Datum — dann bleibt
 * das Dropdown. Gespeichert wird weiter der Options-Text selbst; Mails,
 * Excel-Export und Organizer Center sehen also keinen Unterschied.
 *
 * Fehlt die Jahreszahl, gilt das Jahr des Events (bzw. das nächste Vorkommen
 * ab heute, wenn kein Event-Datum da ist).
 */

const MONTHS: Record<string, number> = {
  jan: 1, januar: 1, january: 1, jän: 1, jänner: 1,
  feb: 2, februar: 2, february: 2,
  mär: 3, mrz: 3, märz: 3, maerz: 3, mar: 3, march: 3,
  apr: 4, april: 4,
  mai: 5, may: 5,
  jun: 6, juni: 6, june: 6,
  jul: 7, juli: 7, july: 7,
  aug: 8, august: 8,
  sep: 9, sept: 9, september: 9,
  okt: 10, oktober: 10, oct: 10, october: 10,
  nov: 11, november: 11,
  dez: 12, dezember: 12, dec: 12, december: 12,
};

// Wochentag vorne („Do, ", "Thursday ") wird ignoriert.
const WEEKDAY_PREFIX = /^(mo|di|mi|do|fr|sa|so|mon|tue|tues|wed|thu|thur|thurs|fri|sat|sun|montag|dienstag|mittwoch|donnerstag|freitag|samstag|sonntag|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\.?,?\s+/i;

function valid(y: number, m: number, d: number): Date | null {
  if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
  const dt = new Date(y, m - 1, d);
  return dt.getMonth() === m - 1 && dt.getDate() === d ? dt : null;
}

function fullYear(y: string | undefined, fallbackYear: number): number {
  if (!y) return fallbackYear;
  const n = parseInt(y, 10);
  return y.length <= 2 ? 2000 + n : n;
}

/** Ein Options-Text als Datum — oder null, wenn er mehr ist als ein Datum. */
export function parseOptionDate(raw: string, fallbackYear: number): Date | null {
  let s = (raw || '').trim().replace(/[.,;:]+$/, (m) => (m.indexOf('.') >= 0 ? '.' : '')).trim();
  s = s.replace(WEEKDAY_PREFIX, '').trim();
  if (!s) return null;
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return valid(+m[1], +m[2], +m[3]);
  m = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})?$/.exec(s);
  if (m) return valid(fullYear(m[3], fallbackYear), +m[2], +m[1]);
  m = /^(\d{1,2})\.?\s+([a-zäöü]+)\.?,?(?:\s+(\d{4}))?$/i.exec(s);
  if (m) {
    const mo = MONTHS[m[2].toLowerCase()];
    return mo ? valid(fullYear(m[3], fallbackYear), mo, +m[1]) : null;
  }
  m = /^([a-zäöü]+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?(?:\s+(\d{4}))?$/i.exec(s);
  if (m) {
    const mo = MONTHS[m[1].toLowerCase()];
    return mo ? valid(fullYear(m[3], fallbackYear), mo, +m[2]) : null;
  }
  return null;
}

/** Alle Optionen als Daten (gleiche Reihenfolge) — null, sobald eine keine ist. */
export function optionsAsDates(options: string[] | undefined, eventStart?: string): Date[] | null {
  const opts = (options || []).map(o => (o || '').trim()).filter(Boolean);
  if (opts.length < 2) return null;
  const ev = eventStart ? new Date(eventStart) : null;
  const fallbackYear = ev && isFinite(ev.getTime()) ? ev.getFullYear() : new Date().getFullYear();
  const out: Date[] = [];
  for (const o of opts) {
    const d = parseOptionDate(o, fallbackYear);
    if (!d) return null;
    out.push(d);
  }
  // Mehr als ein halbes Jahr Spanne ist kein Kalender mehr, sondern eine Liste.
  const times = out.map(d => d.getTime());
  if (Math.max(...times) - Math.min(...times) > 183 * 86400000) return null;
  return out;
}

export function dayKey(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * Soll dieses Feld als Kalender erscheinen? Schalter aus (`asCalendar: false`)
 * gewinnt; sonst Kalender, wenn alle Optionen Daten sind. Bei Schalter an mit
 * Nicht-Daten bleibt es das Dropdown — ein Kalender ohne Tage wäre leer.
 */
export function calendarDatesFor(field: { type?: string; options?: string[]; asCalendar?: boolean }, eventStart?: string): Date[] | null {
  if (field.type !== 'select' || field.asCalendar === false) return null;
  const opts = (field.options || []).filter(o => (o || '').trim());
  if (field.asCalendar === true && opts.length === 1) {
    const one = optionsAsDates([opts[0], opts[0]], eventStart);
    return one ? [one[0]] : null;
  }
  return optionsAsDates(field.options, eventStart);
}
