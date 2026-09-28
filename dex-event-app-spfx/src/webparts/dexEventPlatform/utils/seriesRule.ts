/* v31.99: Serien-Termine.
 *
 * Eine Serie ist KEIN neuer Datentyp. Sie erzeugt ganz normale Kalender-
 * Sub-Events (ein Tag je Termin, v28.91) — mit eigener Teilnehmerliste,
 * Kapazität, Frist und Outlook-Termin. Damit bleiben Flows, `ParentEventId`,
 * Warteliste und Anmeldeseite unangetastet (dieselbe Abwägung wie bei den
 * Stunden-Slots in CLAUDE.md: ein eigener Typ unterhalb des Sub-Events wäre
 * die teure Variante).
 *
 * Die Regel selbst ist nur die Vorlage. Gespeichert wird die zuletzt
 * ANGEWENDETE Regel (`_seriesRule`), damit eine spätere Änderung gegen sie
 * abgeglichen werden kann: Was kommt dazu, was fällt weg, wessen Uhrzeit
 * ändert sich. Dieser Abgleich (`planSeries`) ist rein — er rechnet nur und
 * fasst keinen State an; entschieden wird im Dialog.
 *
 * Alle Tage sind Berliner Kalendertage `YYYY-MM-DD`. Gerechnet wird über
 * `Date.UTC`, damit eine Sommerzeit-Umstellung keinen Tag verschluckt. */
import { SeriesRule } from '../types';

/** Obergrenze je Serie. Jeder Termin ist ein Sub-Event mit eigener Subsite —
 *  mehr als hundert davon in einem Speichervorgang anzulegen dauert
 *  Viertelstunden und trifft zuverlässig die Drosselung. */
export const SERIES_MAX = 100;

const DAY_MS = 86400000;

const parseKey = (k: string): number | null => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(k || '');
  if (!m) return null;
  return Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
};
const keyOf = (ms: number): string => {
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
};
const isHm = (v: string): boolean => /^([01]\d|2[0-3]):[0-5]\d$/.test(v || '');

export const defaultSeriesRule = (startKey: string, startTime: string, endTime: string, allDay: boolean): SeriesRule => {
  const startMs = parseKey(startKey);
  const wd = startMs !== null ? new Date(startMs).getUTCDay() : 1;
  const untilMs = startMs !== null ? startMs + 7 * 7 * DAY_MS : null;
  return {
    freq: 'weekly',
    interval: 1,
    weekdays: [wd],
    skipWeekends: false,
    start: startKey,
    endMode: 'count',
    until: untilMs !== null ? keyOf(untilMs) : '',
    count: 8,
    allDay,
    startTime: isHm(startTime) ? startTime : '09:00',
    endTime: isHm(endTime) ? endTime : '17:00',
    exceptions: [],
  };
};

/** Fehlertext für eine unvollständige Regel, sonst ''. */
export function seriesRuleError(r: SeriesRule, isDe: boolean): string {
  if (parseKey(r.start) === null) return isDe ? 'Bitte den ersten Termin wählen.' : 'Please choose the first date.';
  if (!(r.interval >= 1 && r.interval <= 52)) return isDe ? 'Der Abstand muss zwischen 1 und 52 liegen.' : 'The interval must be between 1 and 52.';
  if (r.freq === 'weekly' && !(r.weekdays && r.weekdays.length > 0)) return isDe ? 'Bitte mindestens einen Wochentag wählen.' : 'Please pick at least one weekday.';
  if (r.endMode === 'until') {
    const u = parseKey(r.until || '');
    if (u === null) return isDe ? 'Bitte das Ende der Serie wählen.' : 'Please choose when the series ends.';
    if (u < (parseKey(r.start) as number)) return isDe ? 'Das Ende der Serie liegt vor dem ersten Termin.' : 'The series ends before its first date.';
  } else if (!(r.count && r.count >= 1 && r.count <= SERIES_MAX)) {
    return isDe ? `Die Anzahl muss zwischen 1 und ${SERIES_MAX} liegen.` : `The number of dates must be between 1 and ${SERIES_MAX}.`;
  }
  if (!r.allDay) {
    if (!isHm(r.startTime) || !isHm(r.endTime)) return isDe ? 'Bitte Uhrzeiten im Format HH:MM angeben.' : 'Please enter times as HH:MM.';
    if (r.endTime <= r.startTime) return isDe ? 'Die Endzeit muss nach der Startzeit liegen.' : 'The end time must be after the start time.';
  }
  return '';
}

/**
 * Alle Tage der Regel, aufsteigend. `truncated` heißt: Die Regel hätte mehr
 * als SERIES_MAX Termine ergeben, die Liste ist abgeschnitten. Ausnahmen
 * (`exceptions`) sind hier NICHT abgezogen — sie gehören zur Regel, sind nur
 * einzeln gelöscht; das Abziehen macht `planSeries`.
 */
export function seriesDates(r: SeriesRule): { dates: string[]; truncated: boolean } {
  const out: string[] = [];
  const start = parseKey(r.start);
  if (start === null) return { dates: out, truncated: false };
  const until = r.endMode === 'until' ? parseKey(r.until || '') : null;
  if (r.endMode === 'until' && until === null) return { dates: out, truncated: false };
  const want = r.endMode === 'count' ? Math.max(0, Math.min(SERIES_MAX, r.count || 0)) : SERIES_MAX;
  const step = Math.max(1, Math.floor(r.interval || 1));
  let truncated = false;
  // Rückgabe true = abbrechen.
  const push = (ms: number): boolean => {
    if (until !== null && ms > until) return true;
    if (out.length >= want) { truncated = r.endMode === 'until'; return true; }
    out.push(keyOf(ms));
    return false;
  };
  // Schutz gegen Endlosschleifen bei exotischen Eingaben (z.B. nur der
  // 31. und lauter kurze Monate): höchstens 20 Jahre durchlaufen.
  const hardStop = start + 20 * 366 * DAY_MS;
  if (r.freq === 'daily') {
    for (let ms = start; ms <= hardStop; ms += step * DAY_MS) {
      const wd = new Date(ms).getUTCDay();
      if (r.skipWeekends && (wd === 0 || wd === 6)) {
        if (until !== null && ms > until) break;
        continue;
      }
      if (push(ms)) break;
    }
  } else if (r.freq === 'weekly') {
    // Wochen beginnen am Montag (deutscher Kalender) — „alle 2 Wochen Mo+Do"
    // heißt: in jeder zweiten Kalenderwoche beide Tage.
    const days = Array.from(new Set((r.weekdays || []).filter(d => d >= 0 && d <= 6)))
      .map(d => (d + 6) % 7)
      .sort((a, b) => a - b);
    if (days.length === 0) return { dates: out, truncated: false };
    const monday = start - ((new Date(start).getUTCDay() + 6) % 7) * DAY_MS;
    let stop = false;
    for (let wk = monday; !stop && wk <= hardStop; wk += step * 7 * DAY_MS) {
      for (const off of days) {
        const ms = wk + off * DAY_MS;
        if (ms < start) continue;
        if (push(ms)) { stop = true; break; }
      }
    }
  } else {
    // Monatlich am selben Tag. Gibt es den Tag im Monat nicht (31.), fällt
    // der Termin in diesem Monat aus — kein stilles Verschieben auf den
    // Letzten, das wäre ein Termin, den niemand angelegt hat.
    const d0 = new Date(start);
    const y0 = d0.getUTCFullYear(); const m0 = d0.getUTCMonth(); const day = d0.getUTCDate();
    for (let k = 0; k < 20 * 12; k += step) {
      const y = y0 + Math.floor((m0 + k) / 12);
      const m = (m0 + k) % 12;
      const ms = Date.UTC(y, m, day);
      if (new Date(ms).getUTCMonth() !== m) {
        if (until !== null && Date.UTC(y, m, 1) > until) break;
        continue;
      }
      if (push(ms)) break;
    }
  }
  return { dates: out, truncated };
}

/** Lesbare Kurzform der Regel, z.B. „Jede Woche Mo, Do · 8 Termine". */
export function seriesSummary(r: SeriesRule, isDe: boolean): string {
  const names = isDe ? ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'] : ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const n = Math.max(1, r.interval || 1);
  let what = '';
  if (r.freq === 'daily') {
    what = n === 1 ? (isDe ? 'Täglich' : 'Daily') : (isDe ? `Alle ${n} Tage` : `Every ${n} days`);
    if (r.skipWeekends) what += isDe ? ' (ohne Wochenende)' : ' (weekdays only)';
  } else if (r.freq === 'weekly') {
    const wd = (r.weekdays || []).slice().sort((a, b) => ((a + 6) % 7) - ((b + 6) % 7)).map(d => names[d]).join(', ');
    what = (n === 1 ? (isDe ? 'Jede Woche' : 'Every week') : (isDe ? `Alle ${n} Wochen` : `Every ${n} weeks`)) + (wd ? ` ${wd}` : '');
  } else {
    const day = parseKey(r.start) !== null ? new Date(parseKey(r.start) as number).getUTCDate() : 0;
    what = (n === 1 ? (isDe ? 'Jeden Monat' : 'Every month') : (isDe ? `Alle ${n} Monate` : `Every ${n} months`)) + (day ? (isDe ? ` am ${day}.` : ` on day ${day}`) : '');
  }
  const time = r.allDay ? (isDe ? 'ganztägig' : 'all day') : `${r.startTime}–${r.endTime}`;
  return `${what} · ${time}`;
}

export interface SeriesPlan<T> {
  /** Tage, die neu angelegt würden. */
  add: string[];
  /** Bestehende Termine, die nach der neuen Regel wegfallen (nur künftige). */
  remove: T[];
  /** Bestehende künftige Termine der Regel mit abweichender Uhrzeit. */
  retime: T[];
  /** Einzeln gelöschte Tage, die die neue Regel wieder träfe. */
  restorable: string[];
  /** Ausnahmen, die mit der neuen Regel gespeichert werden (ohne Wiederherstellung). */
  exceptions: string[];
  /** Bestehende Termine außerhalb der Regel, die bleiben (manuell angelegt oder vergangen). */
  outside: number;
  truncated: boolean;
}

/**
 * Abgleich einer (geänderten) Regel mit den vorhandenen Terminen.
 *
 * - Vergangene Termine werden nie angefasst — weder gelöscht noch umgestellt;
 *   an ihnen hängen Anwesenheit und Abrechnung.
 * - Entfernt werden nur Termine, die zur ALTEN Regel gehörten. Ein von Hand
 *   im Kalender dazugeklickter Tag ist keine Serien-Instanz und bleibt.
 * - Einzeln gelöschte Tage der alten Regel werden zu Ausnahmen; eine
 *   Verlängerung legt sie nicht wieder an, der Dialog bietet es nur an.
 */
export function planSeries<T>(
  prev: SeriesRule | null | undefined,
  next: SeriesRule,
  subs: T[],
  dayOf: (t: T) => string,
  timesOf: (t: T) => { start: string; end: string; allDay: boolean },
  today: string,
): SeriesPlan<T> {
  const nextRes = seriesDates(next);
  const nextSet = new Set(nextRes.dates);
  const prevDates = prev ? seriesDates(prev).dates : [];
  const prevSet = new Set(prevDates);
  const present = new Set(subs.map(dayOf).filter(Boolean));
  const exc = new Set<string>([...(prev && prev.exceptions ? prev.exceptions : []), ...(next.exceptions || [])]);
  for (const k of prevDates) if (!present.has(k)) exc.add(k);
  const isFutureOk = (k: string): boolean => !prev || k >= today;
  const restorable = nextRes.dates.filter(k => exc.has(k) && !present.has(k) && isFutureOk(k));
  const add = nextRes.dates.filter(k => !present.has(k) && !exc.has(k) && isFutureOk(k));
  const remove = prev
    ? subs.filter(s => { const k = dayOf(s); return !!k && k >= today && prevSet.has(k) && !nextSet.has(k); })
    : [];
  const wantStart = next.allDay ? '00:00' : next.startTime;
  const wantEnd = next.allDay ? '23:59' : next.endTime;
  const retime = subs.filter(s => {
    const k = dayOf(s);
    if (!k || k < today || !nextSet.has(k)) return false;
    const t = timesOf(s);
    return t.allDay !== !!next.allDay || t.start !== wantStart || t.end !== wantEnd;
  });
  const removeSet = new Set(remove);
  const outside = subs.filter(s => { const k = dayOf(s); return !!k && !nextSet.has(k) && !removeSet.has(s); }).length;
  const exceptions = Array.from(exc).filter(k => nextSet.has(k)).sort();
  return { add, remove, retime, restorable, exceptions, outside, truncated: nextRes.truncated };
}

/** Heute als Berliner Kalendertag. */
export function berlinTodayKey(): string {
  try {
    // en-CA formatiert als YYYY-MM-DD.
    return new Intl.DateTimeFormat('en-CA', { timeZone: 'Europe/Berlin', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  } catch {
    return keyOf(Date.now());
  }
}

/** Regel aus dem Overrides-Blob lesen; kaputte Werte → undefined. */
export function parseSeriesRule(v: unknown): SeriesRule | undefined {
  if (!v || typeof v !== 'object') return undefined;
  const r = v as Record<string, unknown>;
  const freq = r.freq === 'daily' || r.freq === 'weekly' || r.freq === 'monthly' ? r.freq : null;
  if (!freq || typeof r.start !== 'string') return undefined;
  return {
    freq,
    interval: typeof r.interval === 'number' && r.interval >= 1 ? Math.floor(r.interval) : 1,
    weekdays: Array.isArray(r.weekdays) ? (r.weekdays as unknown[]).filter((d): d is number => typeof d === 'number' && d >= 0 && d <= 6) : [],
    skipWeekends: !!r.skipWeekends,
    start: r.start,
    endMode: r.endMode === 'until' ? 'until' : 'count',
    until: typeof r.until === 'string' ? r.until : '',
    count: typeof r.count === 'number' ? r.count : 0,
    allDay: !!r.allDay,
    startTime: typeof r.startTime === 'string' ? r.startTime : '',
    endTime: typeof r.endTime === 'string' ? r.endTime : '',
    exceptions: Array.isArray(r.exceptions) ? (r.exceptions as unknown[]).filter((k): k is string => typeof k === 'string') : [],
  };
}
