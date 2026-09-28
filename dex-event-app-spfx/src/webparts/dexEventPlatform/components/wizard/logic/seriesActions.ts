/* v31.99: Serien-Termine im Assistenten — die Rechenseite.
 *
 * Zwei Vorgänge, beide rein (nehmen Drafts, geben Drafts zurück, fassen
 * keinen State an):
 *
 * 1. Eine Regel anwenden: Termine anlegen, wegfallende entfernen, Uhrzeiten
 *    nachziehen (`applySeriesPlan`). Den Abgleich rechnet `planSeries` in
 *    utils/seriesRule; entschieden wird im Dialog.
 * 2. Eine Änderung an EINEM Termin auf die anderen übertragen
 *    (`seriesSnapshot` / `changedSeriesGroups` / `propagateSeriesChange`).
 *    Verglichen wird gegen den Stand beim Betreten des Reiters — gefragt wird
 *    beim Verlassen und vor dem Speichern, nie je Tastendruck.
 *
 * Kommunikationsfelder (Mailtexte, Outlook-Text, Logos) gehören bewusst NICHT
 * in die Übertragung: Sie liegen nicht laufend im Draft (switchCommTab), und
 * für „alle Termine gleich" gibt es seit v30.71 den Schalter „gemeinsam". */
import { SubEventDraft } from '../wizardTypes';

export interface SeriesGroup {
  key: string;
  de: string;
  en: string;
  /** Vorauswahl im Dialog. Der Titel ist aus — Serien-Termine heißen nach
   *  ihrem Tag, ein umbenannter Termin ist meist die Ausnahme. */
  defaultOn: boolean;
}

export const SERIES_PROPAGATE_GROUPS: SeriesGroup[] = [
  { key: 'title', de: 'Titel', en: 'Title', defaultOn: false },
  { key: 'description', de: 'Beschreibung', en: 'Description', defaultOn: true },
  { key: 'image', de: 'Event-Bild', en: 'Event image', defaultOn: true },
  { key: 'times', de: 'Uhrzeit (Beginn, Ende, ganztägig)', en: 'Time of day (start, end, all-day)', defaultOn: true },
  { key: 'showAsFree', de: 'Kalender blockieren ja/nein', en: 'Blocks the calendar yes/no', defaultOn: true },
  { key: 'place', de: 'Ort & Adresse', en: 'Location & address', defaultOn: true },
  { key: 'capacity', de: 'Plätze & Warteliste', en: 'Seats & waitlist', defaultOn: true },
  { key: 'agenda', de: 'Programm', en: 'Agenda', defaultOn: true },
  { key: 'fields', de: 'Fragen im Anmeldeformular', en: 'Registration form questions', defaultOn: true },
  { key: 'visibility', de: 'Sichtbarkeit (Standorte, Verteiler, Ausschlüsse)', en: 'Visibility (locations, lists, exclusions)', defaultOn: true },
];

const GROUP_FIELDS: Record<string, Array<keyof SubEventDraft>> = {
  title: ['title'],
  description: ['description'],
  showAsFree: ['showAsFree'],
  place: ['location', 'locationAddress'],
  capacity: ['maxParticipants', 'waitlistEnabled'],
  agenda: ['agenda'],
  fields: ['customFields', 'askSalutation'],
  visibility: ['locationFilter', 'audience', 'filterMode', 'excludedUsers'],
};

/** Felder, die ein neu angelegter Termin von einem bestehenden übernimmt. */
const TEMPLATE_GROUPS = ['description', 'image', 'showAsFree', 'place', 'capacity', 'agenda', 'fields', 'visibility'];

export interface SeriesTimeHelpers {
  /** Berliner Lokalzeit „YYYY-MM-DDTHH:MM" eines Sub-Event-ISO. */
  isoToLocal: (iso: string) => string;
  berlinLocalToUtcIso: (local: string) => string;
}

const hmOf = (h: SeriesTimeHelpers, iso: string): string => ((iso && h.isoToLocal(iso)) || '').slice(11, 16);
const dayOf = (h: SeriesTimeHelpers, iso: string): string => ((iso && h.isoToLocal(iso)) || '').slice(0, 10);

/** Tag eines Termins (Berliner Kalendertag), '' ohne Start. */
export const subDayKey = (h: SeriesTimeHelpers, s: SubEventDraft): string => dayOf(h, s.startDate);

export const subTimes = (h: SeriesTimeHelpers, s: SubEventDraft): { start: string; end: string; allDay: boolean } => ({
  start: hmOf(h, s.startDate), end: hmOf(h, s.endDate), allDay: !!s.allDay,
});

const clone = <T>(v: T): T => (v && typeof v === 'object' ? JSON.parse(JSON.stringify(v)) : v);

const shiftKey = (k: string, days: number): string => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(k || '');
  if (!m) return k;
  const d = new Date(Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10) + days));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
};
const dayDiff = (a: string, b: string): number => {
  const pa = /^(\d{4})-(\d{2})-(\d{2})$/.exec(a); const pb = /^(\d{4})-(\d{2})-(\d{2})$/.exec(b);
  if (!pa || !pb) return 0;
  return Math.round((Date.UTC(+pb[1], +pb[2] - 1, +pb[3]) - Date.UTC(+pa[1], +pa[2] - 1, +pa[3])) / 86400000);
};

/** Vergleichswerte je Gruppe. Das Bild zählt als geändert, sobald im
 *  Reiter ein neues gewählt oder das alte entfernt wurde — nur dann gibt es
 *  etwas, das sich hochladen lässt (`persistSubEventImage` braucht die Datei). */
export function seriesSnapshot(h: SeriesTimeHelpers, s: SubEventDraft): Record<string, string> {
  const out: Record<string, string> = {};
  for (const g of Object.keys(GROUP_FIELDS)) {
    out[g] = JSON.stringify(GROUP_FIELDS[g].map(f => s[f] === undefined ? null : s[f]));
  }
  const f = s.imageFile;
  out.image = JSON.stringify([s.imagePreview || '', f ? `${f.name}|${f.size}|${f.lastModified}` : '', !!s.imageRemoved]);
  const t = subTimes(h, s);
  out.times = JSON.stringify([t.allDay, t.start, t.end]);
  return out;
}

export function changedSeriesGroups(before: Record<string, string>, after: Record<string, string>, s: SubEventDraft): string[] {
  return SERIES_PROPAGATE_GROUPS
    .map(g => g.key)
    .filter(k => before[k] !== after[k])
    // Ein „geändertes" Bild ohne Datei und ohne Entfernen ist nur ein
    // zurückgenommener Wechsel — nichts, was sich übertragen ließe.
    .filter(k => k !== 'image' || !!s.imageFile || !!s.imageRemoved);
}

/** Überträgt die gewählten Gruppen von `src` auf `dst`. */
function copyGroups(h: SeriesTimeHelpers, src: SubEventDraft, dst: SubEventDraft, groups: string[]): SubEventDraft {
  const next: SubEventDraft = { ...dst };
  const srcDay = subDayKey(h, src);
  const dstDay = subDayKey(h, dst);
  for (const g of groups) {
    if (g === 'image') {
      // Dieselbe Datei für alle — hochgeladen wird je Termin einzeln.
      next.imageFile = src.imageFile || null;
      next.imagePreview = src.imagePreview || '';
      next.imageRemoved = !!src.imageRemoved;
      continue;
    }
    if (g === 'times') {
      if (!dstDay) continue;
      const t = subTimes(h, src);
      const st = src.allDay ? '00:00' : (t.start || '00:00');
      const en = src.allDay ? '23:59' : (t.end || '23:59');
      next.allDay = !!src.allDay;
      next.startDate = h.berlinLocalToUtcIso(`${dstDay}T${st}`);
      next.endDate = h.berlinLocalToUtcIso(`${dstDay}T${en}`);
      continue;
    }
    const fields = GROUP_FIELDS[g] || [];
    for (const f of fields) {
      (next as unknown as Record<string, unknown>)[f] = clone(src[f]);
    }
    // Programmpunkte tragen ein Datum. Kopiert auf einen anderen Tag,
    // wandern sie um denselben Abstand mit — sonst stünde das Programm
    // des 06.10. im Termin vom 13.10.
    if (g === 'agenda' && Array.isArray(next.agenda) && srcDay && dstDay && srcDay !== dstDay) {
      const diff = dayDiff(srcDay, dstDay);
      next.agenda = next.agenda.map(a => ({ ...a, date: a.date ? shiftKey(a.date, diff) : a.date }));
    }
  }
  return next;
}

export function propagateSeriesChange(
  h: SeriesTimeHelpers,
  subs: SubEventDraft[],
  srcId: string,
  groups: string[],
  target: 'all' | 'following',
): { next: SubEventDraft[]; changed: number } {
  const src = subs.find(s => s.id === srcId);
  if (!src || groups.length === 0) return { next: subs, changed: 0 };
  const srcDay = subDayKey(h, src);
  let changed = 0;
  const next = subs.map(s => {
    if (s.id === srcId) return s;
    if (target === 'following' && !(subDayKey(h, s) > srcDay)) return s;
    changed++;
    return copyGroups(h, src, s, groups);
  });
  return { next, changed };
}

export interface ApplySeriesChoice {
  add: boolean;
  remove: boolean;
  retime: boolean;
  restore: boolean;
  copyTemplate: boolean;
}

/**
 * Wendet einen bestätigten Abgleich an. Liefert die neuen Drafts, die zum
 * Löschen geparkten gespeicherten Termine (wie das X an der Karte, v29.22)
 * und die Ausnahmen für die Regel.
 */
export function applySeriesPlan(
  h: SeriesTimeHelpers,
  subs: SubEventDraft[],
  plan: { add: string[]; remove: SubEventDraft[]; retime: SubEventDraft[]; restorable: string[]; exceptions: string[] },
  rule: { allDay?: boolean; startTime: string; endTime: string },
  choice: ApplySeriesChoice,
  makeDraft: (dayKey: string, startIso: string, endIso: string, allDay: boolean) => SubEventDraft,
  template: SubEventDraft | null,
): { next: SubEventDraft[]; parked: SubEventDraft[]; exceptions: string[] } {
  const st = rule.allDay ? '00:00' : rule.startTime;
  const en = rule.allDay ? '23:59' : rule.endTime;
  const removeIds = new Set(choice.remove ? plan.remove.map(s => s.id) : []);
  const retimeIds = new Set(choice.retime ? plan.retime.map(s => s.id) : []);
  let next = subs
    .filter(s => !removeIds.has(s.id))
    .map(s => {
      if (!retimeIds.has(s.id)) return s;
      const k = subDayKey(h, s);
      return { ...s, allDay: !!rule.allDay, startDate: h.berlinLocalToUtcIso(`${k}T${st}`), endDate: h.berlinLocalToUtcIso(`${k}T${en}`) };
    });
  const days = [...(choice.add ? plan.add : []), ...(choice.restore ? plan.restorable : [])];
  const created = days.map(k => {
    const d = makeDraft(k, h.berlinLocalToUtcIso(`${k}T${st}`), h.berlinLocalToUtcIso(`${k}T${en}`), !!rule.allDay);
    return (choice.copyTemplate && template) ? copyGroups(h, template, d, TEMPLATE_GROUPS) : d;
  });
  next = next.concat(created).sort((a, b) => (a.startDate || '').localeCompare(b.startDate || ''));
  const parked = subs.filter(s => removeIds.has(s.id) && !!s.dbId);
  // Nicht angelegte Tage der Regel bleiben Ausnahmen; wiederhergestellte nicht.
  const restored = new Set(choice.restore ? plan.restorable : []);
  const skippedAdds = choice.add ? [] : plan.add;
  const exceptions = Array.from(new Set([...plan.exceptions.filter(k => !restored.has(k)), ...skippedAdds])).sort();
  return { next, parked, exceptions };
}
