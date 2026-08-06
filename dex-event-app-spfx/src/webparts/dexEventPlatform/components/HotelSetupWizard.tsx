/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import Modal from './Modal';
import { SPRegistration } from '../services/EventService';
import { DeloitteEvent, DexHotel, DexHotelStay, DexHotelRules } from '../types';
import DatePicker, { registerLocale } from 'react-datepicker';
import { de } from 'date-fns/locale';
import 'react-datepicker/dist/react-datepicker.css';

registerLocale('de', de);

/**
 * HotelSetupWizard (v28.58)
 * -------------------------
 * Geführte Erst-Einrichtung der Hotel-Planung — bewusst entlang der Fragen,
 * die ein Organizer sich ohnehin stellt, nicht entlang der Datenstruktur:
 *
 *   1. **Wann** braucht ihr Zimmer?      → Zeiträume
 *   2. **Wo** habt ihr Zimmer?           → Hotels + Kontingent
 *   3. **Wer kommt wohin?**              → Regeln
 *   4. **Was heißt das für die Buchung?** → Verteilung + Extranächte
 *
 * Warum es den Assistenten überhaupt gibt: Hotels geben ihr Kontingent fast
 * immer nur für bestimmte Nächte heraus („40 Zimmer für 24./25."). Wer früher
 * anreist, ist NICHT abgedeckt — diese Nächte muss jemand zusätzlich buchen
 * (bei vielen Events zahlt sie die Person selbst vor Ort). Die Zahl dafür war
 * bisher Handarbeit in Excel; hier fällt sie in Schritt 4 heraus.
 *
 * Der Clou: Die meisten Anmeldeformulare fragen genau das schon ab —
 * „Hotel (24-25 Sept): Do you require accommodation?" plus „Hotel (additional
 * nights): Do you require additional nights beforehand?" mit Antworten wie
 * „Yes, I need ONE additional night from 23 - 24 Sept.". Der Assistent erkennt
 * beide Fragen, zählt die Antworten und schlägt je Antwort den passenden
 * Zeitraum vor. Gibt es solche Felder nicht, fällt er auf die manuelle Pflege
 * zurück — das Modal funktioniert für jedes Event.
 *
 * Rechenmodell:
 *  - Jede Person bekommt einen **Zeitraum**: aus ihrer Formular-Antwort und/oder
 *    einer Sub-Event-Regel (beides zusammen als Hülle — früheste Anreise,
 *    späteste Abreise), sonst aus einer bestehenden Zuordnung, sonst der
 *    Haupt-Zeitraum.
 *  - Jedes Hotel hat ein **Kontingent** und einen **Kontingent-Zeitraum**.
 *    Eine Person belegt einen Platz, egal wie lange sie bleibt.
 *  - **Extranächte** = Nächte außerhalb des Kontingent-Zeitraums, getrennt nach
 *    „vorab" (frühere Anreise) und „danach" (spätere Abreise).
 */

export interface IHotelSetupWizardProps {
  open: boolean;
  event: DeloitteEvent;
  /** Aktive Anmeldungen (bereits gefiltert). */
  people: SPRegistration[];
  /** Sub-Events der Klammer (leer bei einfachen Events). */
  childEvents: DeloitteEvent[];
  /** Sub-Event-Id → E-Mail-Adressen der aktiven Anmeldungen. */
  subEmails: Record<string, string[]>;
  hotels: DexHotel[];
  stays: DexHotelStay[];
  rules: DexHotelRules;
  isDe: boolean;
  busy: boolean;
  /** Hat die Person im Formular eine Unterkunft gewünscht? null = keine Frage. */
  wishOf: (p: SPRegistration) => boolean | null;
  onClose: () => void;
  onApply: (payload: {
    stays: DexHotelStay[];
    hotels: DexHotel[];
    rules: DexHotelRules;
    assignments: Array<{ reg: SPRegistration; hotel: string; from: string; to: string }>;
  }) => Promise<void>;
}

const uid = (prefix: string): string =>
  `${prefix}${Date.now().toString(36)}${Math.floor(Math.random() * 1e6).toString(36)}`;

const toDay = (iso?: string): string => {
  if (!iso) return '';
  const s = String(iso);
  return s.length >= 10 ? s.substring(0, 10) : '';
};

const addDays = (day: string, n: number): string => {
  const t = Date.parse(`${day}T00:00:00Z`);
  if (isNaN(t)) return day;
  return new Date(t + n * 86400000).toISOString().substring(0, 10);
};

const nightsBetween = (from: string, to: string): number => {
  if (!from || !to) return 0;
  const a = Date.parse(`${from}T00:00:00Z`);
  const b = Date.parse(`${to}T00:00:00Z`);
  if (isNaN(a) || isNaN(b) || b <= a) return 0;
  return Math.round((b - a) / 86400000);
};

const toLocalDay = (d: Date): string =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

const fmtDay = (day: string, isDe: boolean): string => {
  if (!day) return '—';
  const t = Date.parse(`${day}T00:00:00Z`);
  if (isNaN(t)) return day;
  return new Date(t).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' });
};

const nightLabel = (n: number, isDe: boolean): string =>
  isDe ? `${n} ${n === 1 ? 'Nacht' : 'Nächte'}` : `${n} ${n === 1 ? 'night' : 'nights'}`;

const shortTitle = (t: string): string => {
  const parts = (t || '').split('|');
  return (parts.length > 1 ? parts[parts.length - 1] : t).trim();
};

/** Kern-Aufenthalt aus dem Event: Anreise am Starttag, Abreise am Endtag
 *  (mindestens eine Nacht, damit ein Tagesevent nicht 0 Nächte ergibt). */
const coreStay = (event: DeloitteEvent): { from: string; to: string } => {
  const start = toDay(event.startDate) || toLocalDay(new Date());
  const endRaw = toDay(event.endDate) || start;
  const to = nightsBetween(start, endRaw) >= 1 ? endRaw : addDays(start, 1);
  return { from: start, to };
};

/* ---------------------------------------------------------------------- *
 * Formular-Erkennung
 * ---------------------------------------------------------------------- */

const HOTEL_RE = /hotel|unterkunft|übernacht|uebernacht|accommodation|lodging|zimmer/i;
const EXTRA_RE = /additional|extra|zusätzlich|zusaetzlich|weitere|vorab|beforehand|früher|frueher|verlänger|verlaenger|longer/i;
const NO_RE = /^\s*(nein|no|kein|none|nicht)\b|^\s*-\s*$/i;
const AFTER_RE = /after|danach|länger|laenger|abreis|departure|extend|following/i;
const BEFORE_RE = /before|beforehand|vorab|vorher|früher|frueher|anreis|prior/i;
const WORD_NUMS: Record<string, number> = {
  one: 1, two: 2, three: 3, four: 4, five: 5,
  ein: 1, eine: 1, zwei: 2, drei: 3, vier: 4, fünf: 5, fuenf: 5,
};

/**
 * Antwort auf die Zusatznächte-Frage deuten.
 * `null` = nicht deutbar (der Organizer wählt dann selbst).
 */
const parseExtraAnswer = (ansIn: string): { nights: number; after: boolean } | null => {
  const ans = (ansIn || '').trim();
  if (!ans) return { nights: 0, after: false };
  if (NO_RE.test(ans)) return { nights: 0, after: false };
  const s = ans.toLowerCase();
  let n = 0;
  // „2 additional nights" / „2 Nächte" — nur Zahlen direkt vor dem Nacht-Wort,
  // damit Datumsangaben wie „from 23 - 24 Sept." nicht mitgezählt werden.
  const digit = s.match(/(\d+)\s*(additional\s+|extra\s+|weitere\s+|zusätzliche\s+|zusaetzliche\s+)?(night|nacht|nächte|naechte)/);
  if (digit) n = parseInt(digit[1], 10);
  else {
    const keys = Object.keys(WORD_NUMS);
    for (const w of keys) {
      if (new RegExp(`\\b${w}\\b`, 'i').test(s)) { n = WORD_NUMS[w]; break; }
    }
  }
  if (!n) return null;
  const after = AFTER_RE.test(s) && !BEFORE_RE.test(s);
  return { nights: n, after };
};

export const HotelSetupWizard: React.FC<IHotelSetupWizardProps> = (props: IHotelSetupWizardProps) => {
  const { open, event, people, childEvents, subEmails, hotels, stays, rules, isDe, busy, wishOf, onClose, onApply } = props;

  /* ------------------------------------------------------------------ *
   * Was das Anmeldeformular schon weiß
   * ------------------------------------------------------------------ */

  const answersOf = React.useCallback((p: SPRegistration): Record<string, string> => {
    try { return JSON.parse(p.CustomData || '{}') as Record<string, string>; } catch { return {}; }
  }, []);

  /** Die beiden Hotel-Fragen im Formular: „braucht ihr ein Zimmer?" und
   *  „braucht ihr Zusatznächte?". Erkannt über die Beschriftung — das Formular
   *  ist pro Event frei definiert, eine feste Feld-Id gibt es nicht. */
  const fields = React.useMemo(() => {
    const all = (event.eventSpecificFields || []) as Array<{ id: string; label?: string; helpText?: string; options?: string[] }>;
    const hotelish = all.filter(f => HOTEL_RE.test(`${f.label || ''} ${f.helpText || ''}`));
    const extra = hotelish.filter(f => EXTRA_RE.test(`${f.label || ''} ${f.helpText || ''}`))[0] || null;
    const main = hotelish.filter(f => !extra || f.id !== extra.id)[0] || null;
    return { main, extra };
  }, [event.eventSpecificFields]);

  /** Braucht diese Person überhaupt ein Zimmer? Erkanntes Hauptfeld schlägt
   *  die allgemeine Heuristik — sonst würde ein „No" bei den Zusatznächten
   *  fälschlich als „kein Hotel" gelesen. */
  const needsRoom = React.useCallback((p: SPRegistration): boolean | null => {
    if (!fields.main) return wishOf(p);
    const v = (answersOf(p)[fields.main.id] || '').trim();
    if (!v) return null;
    if (NO_RE.test(v)) return false;
    return /\bja\b|\byes\b/i.test(v) ? true : null;
  }, [fields.main, answersOf, wishOf]);

  /** Die verschiedenen Antworten auf die Zusatznächte-Frage, mit Häufigkeit. */
  const extraAnswers = React.useMemo(() => {
    if (!fields.extra) return [] as Array<{ value: string; count: number }>;
    const counts: Record<string, number> = { '': 0 };
    // Auch die noch ungenutzten Auswahlmöglichkeiten aufnehmen — dann steht die
    // Zuordnung schon, bevor die erste Anmeldung mit Zusatznacht eintrudelt.
    for (const o of (fields.extra.options || [])) {
      const v = (o || '').trim();
      if (v) counts[v] = 0;
    }
    for (const p of people) {
      const v = (answersOf(p)[fields.extra.id] || '').trim();
      counts[v] = (counts[v] || 0) + 1;
    }
    return Object.keys(counts)
      .map(v => ({ value: v, count: counts[v] }))
      .sort((a, b) => (a.value === '' ? -1 : b.value === '' ? 1 : b.count - a.count));
  }, [fields.extra, people, answersOf]);

  /* ------------------------------------------------------------------ *
   * Zustand
   * ------------------------------------------------------------------ */

  /** Manuell gepflegte Zeiträume: der Haupt-Zeitraum plus bewusst angelegte
   *  Ausnahmen. Die aus Formular-Antworten abgeleiteten kommen separat dazu
   *  (siehe `answerStays`) — die müssen mitwandern, wenn sich der
   *  Haupt-Zeitraum ändert. */
  const [wStays, setWStays] = React.useState<DexHotelStay[]>([]);
  const [wHotels, setWHotels] = React.useState<DexHotel[]>([]);
  const [wRules, setWRules] = React.useState<DexHotelRules>({});
  /** Antwort-Text → wie viele Nächte, vorab oder danach. */
  const [answerMap, setAnswerMap] = React.useState<Record<string, { nights: number; after: boolean }>>({});
  const [step, setStep] = React.useState(1);
  const [showIntro, setShowIntro] = React.useState(true);
  const [showCustom, setShowCustom] = React.useState(false);
  const [doAssign, setDoAssign] = React.useState(true);
  const [overwrite, setOverwrite] = React.useState(false);
  const [newStay, setNewStay] = React.useState<{ label: string; from: string; to: string }>({ label: '', from: '', to: '' });

  React.useEffect(() => {
    if (!open) return;
    setStep(1);
    setShowCustom(false);
    // Der Überblick ist für den ersten Kontakt — wer schon eingerichtet hat,
    // will direkt in die Felder.
    setShowIntro(hotels.length === 0 && stays.length === 0);
    const core = coreStay(event);
    setWStays(stays.length > 0
      ? stays.map(s => ({ ...s }))
      : [{ id: uid('s'), label: `${nightLabel(nightsBetween(core.from, core.to), isDe)} · ${isDe ? 'Standard' : 'standard'}`, from: core.from, to: core.to, isDefault: true }]);
    setWHotels(hotels.length > 0
      ? hotels.map((h, i) => ({ ...h, priority: typeof h.priority === 'number' ? h.priority : i }))
      : []);
    setWRules({
      bySub: (rules && rules.bySub) ? { ...rules.bySub } : {},
      byAnswer: rules ? rules.byAnswer : undefined,
      keepGroups: rules && typeof rules.keepGroups === 'boolean' ? rules.keepGroups : true,
      skipNoWish: rules && typeof rules.skipNoWish === 'boolean' ? rules.skipNoWish : true,
    });
    setDoAssign(true);
    setOverwrite(false);
    setNewStay({ label: '', from: core.from, to: core.to });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  /** Antworten deuten, sobald das Modal offen ist. Nicht deutbare Antworten
   *  landen auf 0 Nächte — sichtbar in der Tabelle und dort korrigierbar. */
  React.useEffect(() => {
    if (!open) return;
    const next: Record<string, { nights: number; after: boolean }> = {};
    for (const a of extraAnswers) {
      const parsed = parseExtraAnswer(a.value);
      next[a.value] = parsed || { nights: 0, after: false };
    }
    setAnswerMap(next);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, extraAnswers]);

  const mainStay = React.useMemo<DexHotelStay | null>(
    () => wStays.filter(s => s.isDefault)[0] || wStays[0] || null,
    [wStays],
  );

  /** Aus den Antworten abgeleitete Zeiträume — relativ zum Haupt-Zeitraum,
   *  damit sie mitwandern, wenn der sich ändert. */
  const answerStays = React.useMemo<DexHotelStay[]>(() => {
    if (!mainStay) return [];
    const seen: Record<string, DexHotelStay> = {};
    for (const key of Object.keys(answerMap)) {
      const m = answerMap[key];
      if (!m || m.nights <= 0) continue;
      const id = `auto_${m.after ? 'a' : 'b'}_${m.nights}`;
      if (seen[id]) continue;
      seen[id] = {
        id,
        label: isDe
          ? `${nightLabel(m.nights, isDe)} ${m.after ? 'länger' : 'früher'}`
          : `${nightLabel(m.nights, isDe)} ${m.after ? 'longer' : 'earlier'}`,
        from: m.after ? mainStay.from : addDays(mainStay.from, -m.nights),
        to: m.after ? addDays(mainStay.to, m.nights) : mainStay.to,
      };
    }
    return Object.keys(seen).map(k => seen[k]);
  }, [answerMap, mainStay, isDe]);

  /** Alles, was am Ende gespeichert wird und überall zur Auswahl steht. */
  const allStays = React.useMemo<DexHotelStay[]>(() => {
    const out = wStays.slice();
    for (const a of answerStays) {
      if (!out.some(s => s.from === a.from && s.to === a.to)) out.push(a);
    }
    return out;
  }, [wStays, answerStays]);

  const setMainRange = (from: string, to: string): void => {
    const f = from || (mainStay ? mainStay.from : '');
    let t = to || (mainStay ? mainStay.to : '');
    if (!f) return;
    if (!t || nightsBetween(f, t) <= 0) t = addDays(f, 1);
    const label = `${nightLabel(nightsBetween(f, t), isDe)} · ${isDe ? 'Standard' : 'standard'}`;
    if (!mainStay) {
      setWStays([{ id: uid('s'), label, from: f, to: t, isDefault: true }]);
      return;
    }
    setWStays(wStays.map(x => x.id === mainStay.id ? { ...x, from: f, to: t, label, isDefault: true } : x));
  };

  /* ------------------------------------------------------------------ *
   * Verteilung + Extranächte
   * ------------------------------------------------------------------ */

  const subsOf = React.useCallback((p: SPRegistration): DeloitteEvent[] => {
    const em = (p.ParticipantEmail || '').trim().toLowerCase();
    if (!em) return [];
    return childEvents.filter(c => (subEmails[c.id] || []).indexOf(em) >= 0);
  }, [childEvents, subEmails]);

  /** Der Zeitraum dieser Person: Formular-Antwort und Sub-Event-Regeln werden
   *  als Hülle vereinigt (früheste Anreise, späteste Abreise) — wer zum
   *  Vorabend-Dinner kommt UND eine Zusatznacht gebucht hat, braucht beides. */
  const stayFor = React.useCallback((p: SPRegistration): { from: string; to: string } => {
    let from = ''; let to = '';
    const take = (s?: { from: string; to: string } | null): void => {
      if (!s || !s.from || !s.to) return;
      if (!from || s.from < from) from = s.from;
      if (!to || s.to > to) to = s.to;
    };
    if (fields.extra) {
      const v = (answersOf(p)[fields.extra.id] || '').trim();
      const m = answerMap[v];
      if (m && m.nights > 0) take(answerStays.filter(s => s.id === `auto_${m.after ? 'a' : 'b'}_${m.nights}`)[0]);
    }
    const bySub = wRules.bySub || {};
    for (const c of subsOf(p)) {
      const sid = (bySub[c.id] || {}).stayId;
      if (sid) take(allStays.filter(s => s.id === sid)[0]);
    }
    if (from && to) {
      // Der Haupt-Zeitraum ist die Basis — Abweichungen erweitern ihn nur.
      take(mainStay);
      return { from, to };
    }
    const exFrom = toDay(p.HotelFrom); const exTo = toDay(p.HotelTo);
    if (exFrom && exTo) return { from: exFrom, to: exTo };
    return mainStay ? { from: mainStay.from, to: mainStay.to } : { from: '', to: '' };
  }, [fields.extra, answersOf, answerMap, answerStays, wRules, subsOf, allStays, mainStay]);

  const forcedHotel = React.useCallback((p: SPRegistration): string => {
    const bySub = wRules.bySub || {};
    for (const c of subsOf(p)) {
      const h = (bySub[c.id] || {}).hotel;
      if (h) return h;
    }
    return '';
  }, [wRules, subsOf]);

  interface IPlanRow {
    hotel: DexHotel;
    people: SPRegistration[];
    base: DexHotelStay | null;
    extraBefore: number;
    extraAfter: number;
    nights: number;
  }

  const plan = React.useMemo<{
    rows: IPlanRow[]; unplaced: number; candidates: number;
    excludedAssigned: number; excludedNoWish: number;
    assignments: Array<{ reg: SPRegistration; hotel: string; from: string; to: string }>;
  }>(() => {
    // Priorität bestimmt die Reihenfolge — Hotels OHNE Kontingent kommen aber
    // immer zuletzt. Sonst saugt ein „unbegrenztes" Haus, das zufällig oben
    // steht, die gesamte Gruppe auf und die echten Kontingente verfallen.
    const ordered = wHotels.slice().sort((a, b) => {
      const ua = (a.capacity || 0) <= 0 ? 1 : 0;
      const ub = (b.capacity || 0) <= 0 ? 1 : 0;
      if (ua !== ub) return ua - ub;
      return (a.priority || 0) - (b.priority || 0);
    });
    const rows: IPlanRow[] = ordered.map(h => ({
      hotel: h,
      people: [],
      base: allStays.filter(s => s.id === h.capacityStayId)[0] || mainStay,
      extraBefore: 0, extraAfter: 0, nights: 0,
    }));
    const byName: Record<string, IPlanRow> = {};
    for (const r of rows) byName[r.hotel.name] = r;

    // Wer faellt raus — und warum? Ohne diese Zahlen wirkt eine Verteilung,
    // die nur eine von zwei Personen anfasst, wie ein Fehler.
    let excludedAssigned = 0; let excludedNoWish = 0;
    const candidates = people.filter(p => {
      if (!overwrite && (p.Hotel || '').trim()) { excludedAssigned++; return false; }
      if (wRules.skipNoWish && needsRoom(p) === false) { excludedNoWish++; return false; }
      return true;
    });

    const used: Record<string, number> = {};
    if (!overwrite) {
      for (const p of people) {
        const h = (p.Hotel || '').trim();
        if (h && byName[h]) used[h] = (used[h] || 0) + 1;
      }
    }
    const left: Record<string, number> = {};
    for (const r of rows) {
      const cap = r.hotel.capacity || 0;
      left[r.hotel.name] = cap > 0 ? Math.max(0, cap - (used[r.hotel.name] || 0)) : Number.MAX_SAFE_INTEGER;
    }

    const place = (p: SPRegistration, r: IPlanRow): void => {
      r.people.push(p);
      if (left[r.hotel.name] !== Number.MAX_SAFE_INTEGER) left[r.hotel.name] -= 1;
    };

    let unplaced = 0;
    const rest: SPRegistration[] = [];

    // 1. Feste Sub-Event-Regeln zuerst — sie sind eine Zusage, kein Wunsch.
    for (const p of candidates) {
      const forced = forcedHotel(p);
      if (forced && byName[forced]) place(p, byName[forced]);
      else rest.push(p);
    }

    // 2. Der Rest nach Priorität, Gruppen zusammen.
    const blocks: SPRegistration[][] = [];
    if (wRules.keepGroups) {
      const groups: Record<string, SPRegistration[]> = {};
      for (const p of rest) {
        const key = subsOf(p).map(c => c.id).sort().join('|') || '__none';
        (groups[key] = groups[key] || []).push(p);
      }
      for (const k of Object.keys(groups)) blocks.push(groups[k]);
      blocks.sort((a, b) => b.length - a.length);
    } else {
      for (const p of rest) blocks.push([p]);
    }

    for (const block of blocks) {
      let queue = block.slice();
      if (wRules.keepGroups) {
        const whole = rows.filter(r => left[r.hotel.name] >= queue.length)[0];
        if (whole) { for (const p of queue) place(p, whole); continue; }
      }
      for (const r of rows) {
        if (queue.length === 0) break;
        const take = Math.min(left[r.hotel.name], queue.length);
        if (take <= 0) continue;
        for (const p of queue.slice(0, take)) place(p, r);
        queue = queue.slice(take);
      }
      unplaced += queue.length;
    }

    // 3. Extranächte gegen den Kontingent-Zeitraum des Hotels rechnen.
    const assignments: Array<{ reg: SPRegistration; hotel: string; from: string; to: string }> = [];
    for (const r of rows) {
      for (const p of r.people) {
        const s = stayFor(p);
        assignments.push({ reg: p, hotel: r.hotel.name, from: s.from, to: s.to });
        r.nights += nightsBetween(s.from, s.to);
        if (!r.base) continue;
        r.extraBefore += nightsBetween(s.from, r.base.from);
        r.extraAfter += nightsBetween(r.base.to, s.to);
      }
    }
    return { rows, unplaced, candidates: candidates.length, excludedAssigned, excludedNoWish, assignments };
  }, [wHotels, allStays, mainStay, wRules, people, overwrite, forcedHotel, subsOf, stayFor, needsRoom]);

  const needBeds = React.useMemo(
    () => people.filter(p => !(wRules.skipNoWish && needsRoom(p) === false)).length,
    [people, wRules.skipNoWish, needsRoom],
  );
  const totalCap = wHotels.reduce((n, h) => n + (h.capacity || 0), 0);

  /* ------------------------------------------------------------------ *
   * Darstellung
   * ------------------------------------------------------------------ */

  const box: React.CSSProperties = { border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 12, marginTop: 12 };
  const smallInp: React.CSSProperties = { height: 34, fontSize: '0.84rem', padding: '0 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, minWidth: 0 };
  const th: React.CSSProperties = { textAlign: 'left', padding: '6px 8px', fontSize: '0.74rem', color: 'var(--dex-gray-600)', borderBottom: '1px solid var(--dex-gray-200)', whiteSpace: 'nowrap' };
  const td: React.CSSProperties = { padding: '6px 8px', fontSize: '0.82rem', borderBottom: '1px solid var(--dex-gray-100)' };
  const question: React.CSSProperties = { fontSize: '0.95rem', fontWeight: 700, color: 'var(--dex-gray-800)' };
  const explain: React.CSSProperties = { fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, margin: '4px 0 10px' };

  // Die Schritte tragen die Frage, die sie stellen — nicht den Fachbegriff.
  // „Zeiträume" sagt beim ersten Mal niemandem, was zu tun ist; „Wann?" schon.
  const STEPS = isDe
    ? ['Wann?', 'Wo?', 'Wer wohin?', 'Ergebnis']
    : ['When?', 'Where?', 'Who goes where?', 'Result'];

  const canNext = (): boolean => {
    if (step === 1) return !!mainStay && nightsBetween(mainStay.from, mainStay.to) > 0;
    if (step === 2) return wHotels.length > 0 && wHotels.every(h => (h.name || '').trim() !== '');
    return true;
  };

  /* ---- Überblick vor dem ersten Schritt ---- */
  const renderIntro = (): JSX.Element => {
    const asked = people.filter(p => needsRoom(p) === true).length;
    const declined = people.filter(p => needsRoom(p) === false).length;
    const withExtra = extraAnswers.filter(a => (answerMap[a.value] || { nights: 0 }).nights > 0).reduce((n, a) => n + a.count, 0);
    const items = isDe ? [
      { q: 'Wann braucht ihr Zimmer?', a: 'Ein Zeitraum für die meisten — plus Ausnahmen für alle, die früher anreisen oder länger bleiben.' },
      { q: 'Welche Hotels habt ihr?', a: 'Name und das Kontingent, das ihr geblockt habt — samt der Angabe, für welche Nächte das Kontingent gilt.' },
      { q: 'Wer kommt wohin?', a: 'Optionale Regeln, z.B. „alle vom Vorabend-Dinner ins Hotel A". Alle übrigen verteilt die App nach Kontingent.' },
      { q: 'Was heißt das für die Buchung?', a: 'Wer in welchem Hotel liegt — und wie viele Extranächte über das Kontingent hinaus gebucht werden müssen.' },
    ] : [
      { q: 'When do you need rooms?', a: 'One period for most people — plus exceptions for anyone arriving earlier or staying longer.' },
      { q: 'Which hotels do you have?', a: 'Name and the capacity you blocked — including which nights that capacity covers.' },
      { q: 'Who goes where?', a: 'Optional rules, e.g. „everyone from the prior-evening dinner into hotel A". The rest is filled by capacity.' },
      { q: 'What does that mean for booking?', a: 'Who stays where — and how many extra nights have to be booked beyond the capacity.' },
    ];
    return (
      <div>
        <div style={{ ...box, marginTop: 14, background: 'var(--dex-gray-50, #f7f7f5)' }}>
          <div style={{ fontSize: '0.86rem', lineHeight: 1.6 }}>
            {isDe
              ? <><strong>{people.length}</strong> Personen sind angemeldet.{fields.main
                ? <> Die Frage „{fields.main.label}" haben <strong>{asked}</strong> mit Ja beantwortet{declined > 0 ? <>, <strong>{declined}</strong> mit Nein</> : ''}.</>
                : <> Im Anmeldeformular gibt es keine Hotel-Frage — der Assistent geht davon aus, dass grundsätzlich alle ein Zimmer brauchen können.</>}</>
              : <><strong>{people.length}</strong> people are registered.{fields.main
                ? <> <strong>{asked}</strong> answered „{fields.main.label}" with yes{declined > 0 ? <>, <strong>{declined}</strong> with no</> : ''}.</>
                : <> There is no hotel question in the registration form — the wizard assumes anyone may need a room.</>}</>}
          </div>
          {fields.extra && (
            <div style={{ fontSize: '0.86rem', lineHeight: 1.6, marginTop: 6, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {isDe
                ? <>Außerdem fragt euer Formular unter „{fields.extra.label}" nach Zusatznächten — <strong>{withExtra}</strong> Person(en) haben dort etwas angegeben. Der Assistent liest das aus und rechnet die Extranächte daraus.</>
                : <>Your form also asks about additional nights under „{fields.extra.label}" — <strong>{withExtra}</strong> person(s) answered there. The wizard reads that and derives the extra nights from it.</>}
            </div>
          )}
        </div>

        <div style={{ fontSize: '0.86rem', fontWeight: 700, color: 'var(--dex-gray-800)', marginTop: 16 }}>
          {isDe ? 'Der Assistent stellt dir vier Fragen:' : 'The wizard asks you four questions:'}
        </div>
        <div style={{ marginTop: 8 }}>
          {items.map((it, i) => (
            <div key={it.q} style={{ display: 'flex', gap: 10, alignItems: 'flex-start', padding: '8px 0', borderTop: i === 0 ? 'none' : '1px solid var(--dex-gray-100)' }}>
              <span style={{
                flex: '0 0 24px', width: 24, height: 24, borderRadius: '50%', background: 'var(--dex-green, #86bc25)',
                color: '#fff', fontSize: '0.76rem', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center',
              }}>{i + 1}</span>
              <div>
                <div style={{ fontSize: '0.86rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{it.q}</div>
                <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, marginTop: 2 }}>{it.a}</div>
              </div>
            </div>
          ))}
        </div>

        <div style={{ ...box, background: 'rgba(134,188,37,0.07)', borderColor: 'var(--dex-green, #86bc25)' }}>
          <div style={{ fontSize: '0.82rem', lineHeight: 1.55 }}>
            {isDe
              ? <>Du kannst jederzeit zurückgehen und alles ändern. <strong>Gespeichert wird nichts</strong>, bevor du am Ende auf „Übernehmen" klickst — und die Anzeige für die Teilnehmer bleibt aus, bis du sie separat freigibst.</>
              : <>You can go back and change anything at any time. <strong>Nothing is saved</strong> before you click „Apply" at the end — and attendee visibility stays off until you release it separately.</>}
          </div>
        </div>
      </div>
    );
  };

  const renderStepper = (): JSX.Element => (
    <div style={{ display: 'flex', gap: 6, margin: '14px 0 4px', flexWrap: 'wrap' }}>
      {STEPS.map((s, i) => (
        <div key={s} style={{
          flex: '1 1 120px', padding: '6px 10px', borderRadius: 8, fontSize: '0.76rem', fontWeight: 600,
          background: i + 1 === step ? 'var(--dex-green, #86bc25)' : (i + 1 < step ? 'rgba(134,188,37,0.15)' : 'var(--dex-gray-50, #f7f7f5)'),
          color: i + 1 === step ? '#fff' : 'var(--dex-gray-700)',
          border: `1px solid ${i + 1 <= step ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
          textAlign: 'center',
        }}>
          {i + 1}. {s}
        </div>
      ))}
    </div>
  );

  /**
   * Kalender-Streifen für Schritt 1.
   *
   * Ein Zeitraum aus zwei Datumsfeldern bleibt abstrakt — vor allem der
   * Unterschied zwischen „Tag" und „Nacht" (eine Anreise am 06. mit Abreise am
   * 07. ist EINE Nacht, nicht zwei Tage). Der Streifen zeigt deshalb je Tag
   * eine Spalte: oben die Event-Tage, darunter je Zeitraum die belegten
   * Nächte. Nächte, die über den Standard-Zeitraum hinausgehen, sind orange —
   * das sind genau die Extranächte, die später über das Kontingent hinaus
   * gebucht werden müssen.
   */
  const renderCalendar = (): JSX.Element | null => {
    if (!mainStay) return null;
    const evFrom = toDay(event.startDate) || mainStay.from;
    const evTo = toDay(event.endDate) || evFrom;
    let min = mainStay.from; let max = mainStay.to;
    for (const s of allStays) {
      if (s.from && s.from < min) min = s.from;
      if (s.to && s.to > max) max = s.to;
    }
    if (evFrom < min) min = evFrom;
    if (evTo > max) max = evTo;
    min = addDays(min, -1);
    max = addDays(max, 1);
    const days: string[] = [];
    for (let d = min; d <= max && days.length < 21; d = addDays(d, 1)) days.push(d);
    if (days.length === 0) return null;

    const rows = [mainStay].concat(allStays.filter(s => s.id !== mainStay.id));
    const colTpl = `132px repeat(${days.length}, minmax(38px, 1fr))`;
    const cellBase: React.CSSProperties = {
      height: 26, display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: '0.68rem', borderRight: '1px solid #fff',
    };
    const rowLabel: React.CSSProperties = {
      fontSize: '0.74rem', color: 'var(--dex-gray-700)', display: 'flex', alignItems: 'center',
      paddingRight: 8, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
    };

    return (
      <div style={{ ...box, padding: 10 }}>
        <div style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--dex-gray-800)', marginBottom: 6 }}>
          {isDe ? 'Im Kalender' : 'On the calendar'}
        </div>
        <div style={{ overflowX: 'auto' }}>
          <div style={{ minWidth: 132 + days.length * 38 }}>
            {/* Kopfzeile: Wochentag + Datum */}
            <div style={{ display: 'grid', gridTemplateColumns: colTpl, alignItems: 'end' }}>
              <div />
              {days.map(d => {
                const dt = new Date(`${d}T00:00:00Z`);
                const wd = dt.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'short', timeZone: 'UTC' });
                const we = dt.getUTCDay() === 0 || dt.getUTCDay() === 6;
                return (
                  <div key={d} style={{ textAlign: 'center', fontSize: '0.66rem', color: we ? 'var(--dex-gray-500)' : 'var(--dex-gray-600)', lineHeight: 1.25, paddingBottom: 3 }}>
                    <div>{wd}</div>
                    <div style={{ fontWeight: 700 }}>{dt.getUTCDate()}.{dt.getUTCMonth() + 1}.</div>
                  </div>
                );
              })}
            </div>

            {/* Event-Tage */}
            <div style={{ display: 'grid', gridTemplateColumns: colTpl, marginBottom: 4 }}>
              <div style={rowLabel}><strong>{isDe ? 'Event' : 'Event'}</strong></div>
              {days.map(d => {
                const on = d >= evFrom && d <= evTo;
                return (
                  <div key={d} style={{
                    ...cellBase,
                    background: on ? 'rgba(134,188,37,0.28)' : 'var(--dex-gray-50, #f4f4f2)',
                    color: 'var(--dex-gray-700)', fontWeight: 700,
                  }}>{on ? (isDe ? 'Event' : 'Event') : ''}</div>
                );
              })}
            </div>

            {/* Zeiträume — je Zeile die belegten Nächte */}
            {rows.map(s => (
              <div key={s.id} style={{ display: 'grid', gridTemplateColumns: colTpl, marginBottom: 3 }}>
                <div style={rowLabel} title={s.label}>
                  {s.id === mainStay.id ? <strong>{s.label}</strong> : s.label}
                </div>
                {days.map(d => {
                  const night = d >= s.from && d < s.to;              // Nacht vom Tag d auf d+1
                  const inMain = d >= mainStay.from && d < mainStay.to;
                  const extra = night && !inMain;                      // über den Standard hinaus
                  const arrival = d === s.from;
                  const departure = d === s.to;
                  return (
                    <div key={d}
                      title={night
                        ? (extra
                          ? (isDe ? `Extranacht ${fmtDay(d, isDe)} → ${fmtDay(addDays(d, 1), isDe)}` : `Extra night ${fmtDay(d, isDe)} → ${fmtDay(addDays(d, 1), isDe)}`)
                          : (isDe ? `Übernachtung ${fmtDay(d, isDe)} → ${fmtDay(addDays(d, 1), isDe)}` : `Night ${fmtDay(d, isDe)} → ${fmtDay(addDays(d, 1), isDe)}`))
                        : (departure ? (isDe ? 'Abreise' : 'Departure') : '')}
                      style={{
                        ...cellBase,
                        background: night
                          ? (extra ? '#f3a83c' : 'var(--dex-green, #86bc25)')
                          : (departure ? 'rgba(134,188,37,0.14)' : 'var(--dex-gray-50, #f4f4f2)'),
                        color: night ? '#fff' : 'var(--dex-gray-500)',
                        fontWeight: 700,
                        borderLeft: arrival ? '2px solid var(--dex-gray-800)' : `1px solid #fff`,
                      }}>
                      {night ? '' : (departure ? '⇤' : '')}
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
        </div>

        <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap', marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-600)', alignItems: 'center' }}>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, background: 'rgba(134,188,37,0.28)', borderRadius: 2, marginRight: 5, verticalAlign: -2 }} />{isDe ? 'Event-Tag' : 'Event day'}</span>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, background: 'var(--dex-green, #86bc25)', borderRadius: 2, marginRight: 5, verticalAlign: -2 }} />{isDe ? 'Übernachtung im Standard-Zeitraum' : 'Night within the standard period'}</span>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, background: '#f3a83c', borderRadius: 2, marginRight: 5, verticalAlign: -2 }} />{isDe ? 'Extranacht (über das Kontingent hinaus)' : 'Extra night (beyond the contingent)'}</span>
          <span><span style={{ display: 'inline-block', width: 2, height: 12, background: 'var(--dex-gray-800)', marginRight: 5, verticalAlign: -2 }} />{isDe ? 'Anreise' : 'Arrival'}</span>
          <span>⇤ {isDe ? 'Abreise (keine Nacht mehr)' : 'Departure (no night)'}</span>
        </div>
      </div>
    );
  };

  /* ---- Schritt 1: Wann? ---- */
  const renderStays = (): JSX.Element => {
    const main = mainStay;
    const core = coreStay(event);
    const coreFits = !!main && main.from === core.from && main.to === core.to;

    const toggleVariant = (from: string, to: string, label: string): void => {
      const hit = wStays.filter(s => s.from === from && s.to === to)[0];
      if (hit) {
        if (main && hit.id === main.id) return;
        setWStays(wStays.filter(s => s.id !== hit.id));
        return;
      }
      setWStays(wStays.concat([{ id: uid('s'), label, from, to, isDefault: false }]));
    };

    const variantDefs = main ? [
      { from: addDays(main.from, -1), to: main.to, label: isDe ? 'Einen Tag früher anreisen' : 'Arrive a day earlier' },
      { from: main.from, to: addDays(main.to, 1), label: isDe ? 'Einen Tag länger bleiben' : 'Stay a day longer' },
      { from: addDays(main.from, -1), to: addDays(main.to, 1), label: isDe ? 'Früher und länger' : 'Earlier and longer' },
    ] : [];

    return (
      <div>
        <div style={question}>{isDe ? 'Von wann bis wann braucht ihr Zimmer?' : 'From when to when do you need rooms?'}</div>
        <p style={explain}>
          {isDe
            ? <>Trag den Zeitraum ein, den die <strong>meisten</strong> Teilnehmer brauchen — meist die Nacht bzw. Nächte des Events. Jede Person bekommt ihn automatisch; Abweichungen regelst du darunter.</>
            : <>Enter the period <strong>most</strong> attendees need — usually the night(s) of the event itself. Everyone gets it automatically; exceptions are handled below.</>}
        </p>

        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
          <div className="form-group" style={{ marginBottom: 0, flex: '1 1 165px' }}>
            <label className="form-label" style={{ fontSize: '0.8rem', marginBottom: 4 }}>{isDe ? 'Anreise' : 'Arrival'}</label>
            <DatePicker selected={main && main.from ? new Date(`${main.from}T00:00:00`) : null}
              onChange={(d: Date | null) => { if (d) setMainRange(toLocalDay(d), main ? main.to : ''); }}
              dateFormat="dd.MM.yyyy" locale={isDe ? 'de' : undefined} placeholderText={isDe ? 'TT.MM.JJJJ' : 'dd/mm/yyyy'}
              className="form-input" wrapperClassName="dex-datepicker-wrapper" calendarClassName="dex-datepicker-calendar"
              popperPlacement="bottom-start" autoComplete="off" />
          </div>
          <div className="form-group" style={{ marginBottom: 0, flex: '1 1 165px' }}>
            <label className="form-label" style={{ fontSize: '0.8rem', marginBottom: 4 }}>{isDe ? 'Abreise' : 'Departure'}</label>
            <DatePicker selected={main && main.to ? new Date(`${main.to}T00:00:00`) : null}
              onChange={(d: Date | null) => { if (d) setMainRange(main ? main.from : '', toLocalDay(d)); }}
              dateFormat="dd.MM.yyyy" locale={isDe ? 'de' : undefined} placeholderText={isDe ? 'TT.MM.JJJJ' : 'dd/mm/yyyy'}
              minDate={main && main.from ? new Date(`${main.from}T00:00:00`) : undefined}
              className="form-input" wrapperClassName="dex-datepicker-wrapper" calendarClassName="dex-datepicker-calendar"
              popperPlacement="bottom-start" autoComplete="off" />
          </div>
          <div style={{ paddingBottom: 13, fontSize: '0.86rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
            = {main ? nightLabel(nightsBetween(main.from, main.to), isDe) : '—'}
          </div>
        </div>

        {!coreFits && (
          <div style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
            {isDe ? 'Aus dem Event-Datum' : 'From the event dates'}:{' '}
            <button type="button" onClick={() => setMainRange(core.from, core.to)}
              style={{ border: 'none', background: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline', fontSize: '0.78rem' }}>
              {fmtDay(core.from, isDe)} – {fmtDay(core.to, isDe)} {isDe ? 'übernehmen' : 'apply'}
            </button>
          </div>
        )}

        {renderCalendar()}

        {/* ---- Zusatznächte aus dem Anmeldeformular ---- */}
        {fields.extra && extraAnswers.length > 0 && (
          <div style={{ ...box, borderColor: 'var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.05)' }}>
            <div style={{ fontSize: '0.88rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Zusatznächte — aus dem Anmeldeformular gelesen' : 'Additional nights — read from the registration form'}
            </div>
            <p style={{ ...explain, marginBottom: 8 }}>
              {isDe
                ? <>Eure Teilnehmer haben unter „<strong>{fields.extra.label}</strong>" selbst angegeben, ob sie länger brauchen. Der Assistent hat die Antworten gezählt und den passenden Zeitraum vorgeschlagen — prüf die Zuordnung und korrigiere sie, wo sie nicht stimmt.</>
                : <>Your attendees stated under „<strong>{fields.extra.label}</strong>" whether they need longer. The wizard counted the answers and proposed a matching period — check and correct it where needed.</>}
            </p>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 560 }}>
                <thead>
                  <tr>
                    <th style={th}>{isDe ? 'Antwort im Formular' : 'Answer in the form'}</th>
                    <th style={th}>{isDe ? 'Personen' : 'People'}</th>
                    <th style={th}>{isDe ? 'Zusätzliche Nächte' : 'Additional nights'}</th>
                    <th style={th}>{isDe ? 'Ergibt den Zeitraum' : 'Resulting period'}</th>
                  </tr>
                </thead>
                <tbody>
                  {extraAnswers.map(a => {
                    const m = answerMap[a.value] || { nights: 0, after: false };
                    const st = m.nights > 0 ? answerStays.filter(s => s.id === `auto_${m.after ? 'a' : 'b'}_${m.nights}`)[0] : mainStay;
                    return (
                      <tr key={a.value || '__empty'}>
                        <td style={{ ...td, maxWidth: 300 }}>
                          {a.value || <span style={{ color: 'var(--dex-gray-500)' }}>{isDe ? '(keine Angabe)' : '(no answer)'}</span>}
                        </td>
                        <td style={{ ...td, whiteSpace: 'nowrap' }}>{a.count}</td>
                        <td style={td}>
                          <span style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
                            <input type="number" min={0} max={14} style={{ ...smallInp, width: 64 }} value={m.nights}
                              onChange={e => setAnswerMap({ ...answerMap, [a.value]: { ...m, nights: Math.max(0, parseInt(e.target.value, 10) || 0) } })} />
                            {m.nights > 0 && (
                              <select style={{ ...smallInp, width: 110 }} value={m.after ? 'after' : 'before'}
                                onChange={e => setAnswerMap({ ...answerMap, [a.value]: { ...m, after: e.target.value === 'after' } })}>
                                <option value="before">{isDe ? 'vorher' : 'before'}</option>
                                <option value="after">{isDe ? 'danach' : 'after'}</option>
                              </select>
                            )}
                          </span>
                        </td>
                        <td style={{ ...td, color: 'var(--dex-gray-600)', whiteSpace: 'nowrap' }}>
                          {st ? `${fmtDay(st.from, isDe)} – ${fmtDay(st.to, isDe)} · ${nightLabel(nightsBetween(st.from, st.to), isDe)}` : '—'}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ---- Manuelle Ausnahmen ---- */}
        <div style={box}>
          <div style={{ fontSize: '0.88rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>
            {fields.extra
              ? (isDe ? 'Weitere Ausnahmen' : 'Further exceptions')
              : (isDe ? 'Reisen einzelne früher an oder bleiben länger?' : 'Do some arrive earlier or stay longer?')}
            <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · {isDe ? 'optional' : 'optional'}</span>
          </div>
          <p style={{ ...explain, marginBottom: 8 }}>
            {isDe
              ? <>Zeiträume, die du hier anlegst, kannst du beim Zuordnen pro Person mit einem Klick vergeben. Alles, was über den Kontingent-Zeitraum des Hotels hinausgeht, erscheint in Schritt 4 als <strong>Extranacht</strong>.</>
              : <>Periods you add here can be applied per person with one click when assigning. Anything beyond a hotel’s capacity period shows up as an <strong>extra night</strong> in step 4.</>}
          </p>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {variantDefs.map(v => {
              const on = wStays.some(s => s.from === v.from && s.to === v.to && (!main || s.id !== main.id));
              return (
                <button key={v.label} type="button" onClick={() => toggleVariant(v.from, v.to, v.label)} style={{
                  textAlign: 'left', cursor: 'pointer', padding: '8px 12px', borderRadius: 999, fontSize: '0.8rem',
                  border: `1.5px solid ${on ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                  background: on ? 'rgba(134,188,37,0.10)' : '#fff',
                }}>
                  {on ? '✓ ' : '+ '}{v.label}
                  <span style={{ color: 'var(--dex-gray-500)' }}> · {fmtDay(v.from, isDe)}–{fmtDay(v.to, isDe)}</span>
                </button>
              );
            })}
          </div>

          {/* v28.61: ALLE Zeiträume in einer Liste — auch der Standard. Der ist
              zwar der Anker für die Extranacht-Rechnung, aber wenn bei eurem
              Event schlicht niemand nur eine Nacht bleibt, muss er weg können.
              Löschen macht den nächsten Zeitraum zum Standard; nur der letzte
              verbleibende lässt sich nicht entfernen. */}
          {wStays.length > 0 && (
            <div style={{ marginTop: 12 }}>
              <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 4 }}>
                {isDe ? 'Eure Zeiträume' : 'Your periods'}
              </div>
              {wStays.map(s => {
                const isMain = !!main && s.id === main.id;
                return (
                  <div key={s.id} style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', padding: '5px 0', borderTop: '1px solid var(--dex-gray-100)' }}>
                    {isMain ? (
                      <span style={{
                        fontSize: '0.7rem', fontWeight: 700, padding: '3px 8px', borderRadius: 999,
                        background: 'var(--dex-green, #86bc25)', color: '#fff', whiteSpace: 'nowrap',
                      }}>{isDe ? 'Standard' : 'Standard'}</span>
                    ) : (
                      <button type="button" title={isDe ? 'Diesen Zeitraum zum Standard machen' : 'Make this the standard'}
                        onClick={() => setWStays(wStays.map(x => ({ ...x, isDefault: x.id === s.id })))}
                        style={{
                          fontSize: '0.7rem', padding: '3px 8px', borderRadius: 999, cursor: 'pointer',
                          border: '1px solid var(--dex-gray-300)', background: '#fff', color: 'var(--dex-gray-600)', whiteSpace: 'nowrap',
                        }}>{isDe ? 'als Standard' : 'make standard'}</button>
                    )}
                    <input style={{ ...smallInp, flex: '2 1 170px' }} value={s.label}
                      onChange={e => setWStays(wStays.map(x => x.id === s.id ? { ...x, label: e.target.value } : x))} />
                    <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
                      {fmtDay(s.from, isDe)} – {fmtDay(s.to, isDe)} · {nightLabel(nightsBetween(s.from, s.to), isDe)}
                    </span>
                    <button type="button" disabled={wStays.length <= 1}
                      title={wStays.length <= 1
                        ? (isDe ? 'Mindestens ein Zeitraum muss bleiben.' : 'At least one period must remain.')
                        : (isDe ? 'Zeitraum entfernen' : 'Remove period')}
                      onClick={() => {
                        const next = wStays.filter(x => x.id !== s.id);
                        if (next.length > 0 && !next.some(x => x.isDefault)) next[0] = { ...next[0], isDefault: true };
                        setWStays(next);
                      }}
                      style={{
                        marginLeft: 'auto', border: 'none', background: 'none', fontSize: '0.95rem',
                        cursor: wStays.length <= 1 ? 'default' : 'pointer',
                        color: 'var(--dex-red, #c00)', opacity: wStays.length <= 1 ? 0.3 : 1,
                      }}>×</button>
                  </div>
                );
              })}
            </div>
          )}

          {!showCustom && (
            <button type="button" onClick={() => setShowCustom(true)}
              style={{ marginTop: 10, border: 'none', background: 'none', padding: 0, cursor: 'pointer', fontSize: '0.8rem', color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline' }}>
              {isDe ? '+ Anderer Zeitraum mit eigenen Daten' : '+ Other period with custom dates'}
            </button>
          )}
          {showCustom && (
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'flex-end', marginTop: 12 }}>
              <div className="form-group" style={{ marginBottom: 0, flex: '2 1 170px' }}>
                <label className="form-label" style={{ fontSize: '0.78rem', marginBottom: 4 }}>{isDe ? 'Bezeichnung' : 'Label'}</label>
                <input className="form-input" placeholder={isDe ? 'z.B. „Nur Sonntagnacht"' : 'e.g. „Sunday night only"'} value={newStay.label}
                  onChange={e => setNewStay({ ...newStay, label: e.target.value })} />
              </div>
              <div className="form-group" style={{ marginBottom: 0, flex: '1 1 145px' }}>
                <label className="form-label" style={{ fontSize: '0.78rem', marginBottom: 4 }}>{isDe ? 'Anreise' : 'Arrival'}</label>
                <DatePicker selected={newStay.from ? new Date(`${newStay.from}T00:00:00`) : null}
                  onChange={(d: Date | null) => setNewStay({ ...newStay, from: d ? toLocalDay(d) : '' })}
                  dateFormat="dd.MM.yyyy" locale={isDe ? 'de' : undefined} placeholderText={isDe ? 'TT.MM.JJJJ' : 'dd/mm/yyyy'}
                  className="form-input" wrapperClassName="dex-datepicker-wrapper" calendarClassName="dex-datepicker-calendar"
                  popperPlacement="bottom-start" isClearable autoComplete="off" />
              </div>
              <div className="form-group" style={{ marginBottom: 0, flex: '1 1 145px' }}>
                <label className="form-label" style={{ fontSize: '0.78rem', marginBottom: 4 }}>{isDe ? 'Abreise' : 'Departure'}</label>
                <DatePicker selected={newStay.to ? new Date(`${newStay.to}T00:00:00`) : null}
                  onChange={(d: Date | null) => setNewStay({ ...newStay, to: d ? toLocalDay(d) : '' })}
                  dateFormat="dd.MM.yyyy" locale={isDe ? 'de' : undefined} placeholderText={isDe ? 'TT.MM.JJJJ' : 'dd/mm/yyyy'}
                  minDate={newStay.from ? new Date(`${newStay.from}T00:00:00`) : undefined}
                  className="form-input" wrapperClassName="dex-datepicker-wrapper" calendarClassName="dex-datepicker-calendar"
                  popperPlacement="bottom-start" isClearable autoComplete="off" />
              </div>
              <button type="button" className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '11px 18px' }}
                onClick={() => {
                  const n = nightsBetween(newStay.from, newStay.to);
                  if (n <= 0) return;
                  setWStays(wStays.concat([{
                    id: uid('s'),
                    label: (newStay.label || '').trim() || nightLabel(n, isDe),
                    from: newStay.from, to: newStay.to,
                    isDefault: wStays.length === 0,
                  }]));
                  setNewStay({ ...newStay, label: '' });
                  setShowCustom(false);
                }}>
                {isDe ? '+ Hinzufügen' : '+ Add'}
              </button>
            </div>
          )}
        </div>
      </div>
    );
  };

  /* ---- Schritt 2: Wo? ---- */
  const renderHotels = (): JSX.Element => (
    <div>
      <div style={question}>{isDe ? 'Welche Hotels habt ihr?' : 'Which hotels do you have?'}</div>
      <p style={explain}>
        {isDe
          ? <>Ein Hotel je Zeile. <strong>Kontingent</strong> = die Zimmer, die ihr dort geblockt habt (leer lassen, wenn es keine feste Obergrenze gibt). Weil ein Kontingent fast immer nur für bestimmte Nächte gilt, sag rechts, für welche — alles darüber hinaus wird in Schritt 4 als Extranacht ausgewiesen.</>
          : <>One hotel per row. <strong>Capacity</strong> = the rooms you blocked there (leave empty if there is no fixed limit). Because a contingent usually covers specific nights only, state which on the right — anything beyond shows as an extra night in step 4.</>}
      </p>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 640 }}>
          <thead>
            <tr>
              <th style={th}>{isDe ? 'Hotel' : 'Hotel'}</th>
              <th style={th}>{isDe ? 'Adresse (optional)' : 'Address (optional)'}</th>
              <th style={th}>{isDe ? 'Kontingent' : 'Capacity'}</th>
              <th style={th}>{isDe ? 'Kontingent gilt für' : 'Capacity applies to'}</th>
              <th style={th} />
            </tr>
          </thead>
          <tbody>
            {wHotels.map(h => (
              <tr key={h.id}>
                <td style={td}>
                  <input style={{ ...smallInp, width: '100%' }} value={h.name} placeholder={isDe ? 'Hotelname' : 'Hotel name'}
                    onChange={e => setWHotels(wHotels.map(x => x.id === h.id ? { ...x, name: e.target.value } : x))} />
                </td>
                <td style={td}>
                  <input style={{ ...smallInp, width: '100%' }} value={h.address || ''}
                    onChange={e => setWHotels(wHotels.map(x => x.id === h.id ? { ...x, address: e.target.value } : x))} />
                </td>
                <td style={td}>
                  <input style={{ ...smallInp, width: 90 }} type="number" min={0} value={h.capacity || ''}
                    placeholder={isDe ? 'offen' : 'open'}
                    onChange={e => setWHotels(wHotels.map(x => x.id === h.id ? { ...x, capacity: parseInt(e.target.value, 10) || 0 } : x))} />
                </td>
                <td style={td}>
                  <select style={{ ...smallInp, width: '100%' }} value={h.capacityStayId || (mainStay ? mainStay.id : '')}
                    onChange={e => setWHotels(wHotels.map(x => x.id === h.id ? { ...x, capacityStayId: e.target.value } : x))}>
                    {allStays.map(s => (
                      <option key={s.id} value={s.id}>
                        {s.label} ({fmtDay(s.from, isDe)}–{fmtDay(s.to, isDe)})
                      </option>
                    ))}
                  </select>
                </td>
                <td style={{ ...td, textAlign: 'right' }}>
                  <button type="button" onClick={() => setWHotels(wHotels.filter(x => x.id !== h.id))}
                    style={{ border: 'none', background: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', fontSize: '0.95rem' }}>×</button>
                </td>
              </tr>
            ))}
            {wHotels.length === 0 && (
              <tr><td style={{ ...td, color: 'var(--dex-gray-500)' }} colSpan={5}>
                {isDe ? 'Noch kein Hotel — leg unten das erste an.' : 'No hotel yet — add the first one below.'}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
      <button type="button" className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 16px', marginTop: 10 }}
        onClick={() => setWHotels(wHotels.concat([{ id: uid('h'), name: '', address: '', capacity: 0, notes: '', priority: wHotels.length, capacityStayId: mainStay ? mainStay.id : '' }]))}>
        {isDe ? '+ Hotel' : '+ Hotel'}
      </button>

      <div style={{
        ...box,
        background: totalCap > 0 && totalCap < needBeds ? '#fff6e5' : 'var(--dex-gray-50, #f7f7f5)',
        borderColor: totalCap > 0 && totalCap < needBeds ? '#e0a300' : 'var(--dex-gray-200)',
      }}>
        <div style={{ fontSize: '0.84rem' }}>
          {isDe
            ? <>Kontingent gesamt: <strong>{totalCap > 0 ? totalCap : '—'}</strong> · Personen mit Bettenbedarf: <strong>{needBeds}</strong></>
            : <>Total capacity: <strong>{totalCap > 0 ? totalCap : '—'}</strong> · people needing a bed: <strong>{needBeds}</strong></>}
        </div>
        {totalCap > 0 && totalCap < needBeds && (
          <div style={{ fontSize: '0.8rem', color: '#8a5a00', marginTop: 4 }}>
            {isDe
              ? `Es fehlen ${needBeds - totalCap} Plätze. Du kannst trotzdem weiter — die Vorschau zeigt, wer übrig bleibt.`
              : `${needBeds - totalCap} places short. You can continue — the preview shows who is left over.`}
          </div>
        )}
      </div>
    </div>
  );

  /* ---- Schritt 3: Wer wohin? ---- */
  const renderRules = (): JSX.Element => (
    <div>
      <div style={question}>{isDe ? 'Wer kommt in welches Hotel?' : 'Who goes into which hotel?'}</div>
      <p style={explain}>
        {isDe
          ? <>Alles hier ist <strong>optional</strong>. Ohne Regel füllt die App die Hotels der Reihe nach auf, bis das Kontingent erreicht ist. Regeln brauchst du nur, wenn eine bestimmte Gruppe zusammen in ein bestimmtes Haus soll.</>
          : <>Everything here is <strong>optional</strong>. Without rules the app fills the hotels in order until capacity is reached. You only need rules if a specific group has to go to a specific house.</>}
      </p>

      {childEvents.length > 0 && (
        <div style={box}>
          <div style={{ fontSize: '0.84rem', fontWeight: 700, marginBottom: 2 }}>
            {isDe ? 'Feste Zuordnung je Sub-Event' : 'Fixed assignment per sub-event'}
          </div>
          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 8, lineHeight: 1.45 }}>
            {isDe
              ? 'Z.B. „alle vom Vorabend-Dinner ins Hotel A, mit 2 Nächten". Wer in mehreren Sub-Events ist, bekommt die Hülle aus allen Zeiträumen — früheste Anreise, späteste Abreise.'
              : 'E.g. „everyone from the prior-evening dinner into hotel A, 2 nights". People in several sub-events get the envelope of all periods — earliest arrival, latest departure.'}
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 520 }}>
              <thead>
                <tr>
                  <th style={th}>{isDe ? 'Sub-Event' : 'Sub-event'}</th>
                  <th style={th}>{isDe ? 'Personen' : 'People'}</th>
                  <th style={th}>{isDe ? 'Hotel' : 'Hotel'}</th>
                  <th style={th}>{isDe ? 'Zeitraum' : 'Period'}</th>
                </tr>
              </thead>
              <tbody>
                {childEvents.map(c => {
                  const r = (wRules.bySub || {})[c.id] || {};
                  const n = (subEmails[c.id] || []).length;
                  const set = (patch: { hotel?: string; stayId?: string }): void =>
                    setWRules({ ...wRules, bySub: { ...(wRules.bySub || {}), [c.id]: { ...r, ...patch } } });
                  return (
                    <tr key={c.id}>
                      <td style={td}>{shortTitle(c.title || '')}</td>
                      <td style={{ ...td, color: 'var(--dex-gray-600)' }}>{n || '—'}</td>
                      <td style={td}>
                        <select style={{ ...smallInp, width: '100%' }} value={r.hotel || ''} onChange={e => set({ hotel: e.target.value })}>
                          <option value="">{isDe ? '— automatisch —' : '— automatic —'}</option>
                          {wHotels.filter(h => (h.name || '').trim()).map(h => <option key={h.id} value={h.name}>{h.name}</option>)}
                        </select>
                      </td>
                      <td style={td}>
                        <select style={{ ...smallInp, width: '100%' }} value={r.stayId || ''} onChange={e => set({ stayId: e.target.value })}>
                          <option value="">{isDe ? '— Standard —' : '— default —'}</option>
                          {allStays.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                        </select>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div style={box}>
        <div style={{ fontSize: '0.84rem', fontWeight: 700, marginBottom: 2 }}>
          {isDe ? 'Füll-Reihenfolge' : 'Fill order'}
        </div>
        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 8 }}>
          {isDe ? 'Welches Hotel soll zuerst voll werden? Oben zuerst. Häuser ohne Kontingent kommen immer zuletzt.' : 'Which hotel fills up first? Top first. Houses without capacity always come last.'}
        </div>
        {wHotels.slice().sort((a, b) => (a.priority || 0) - (b.priority || 0)).map((h, i, arr) => (
          <div key={h.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '5px 0', borderBottom: '1px solid var(--dex-gray-100)' }}>
            <span style={{ width: 20, fontWeight: 700, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{i + 1}.</span>
            <span style={{ flex: 1, fontSize: '0.85rem' }}>{h.name || (isDe ? '(ohne Namen)' : '(unnamed)')}</span>
            <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
              {h.capacity ? `${h.capacity} ${isDe ? 'Plätze' : 'places'}` : (isDe ? 'kein Kontingent' : 'no capacity')}
            </span>
            <button type="button" disabled={i === 0} onClick={() => {
              const order = arr.map(x => x.id);
              const tmp = order[i - 1]; order[i - 1] = order[i]; order[i] = tmp;
              setWHotels(wHotels.map(x => ({ ...x, priority: order.indexOf(x.id) })));
            }} style={{ border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 6, cursor: i === 0 ? 'default' : 'pointer', opacity: i === 0 ? 0.35 : 1, padding: '1px 7px' }}>↑</button>
            <button type="button" disabled={i === arr.length - 1} onClick={() => {
              const order = arr.map(x => x.id);
              const tmp = order[i + 1]; order[i + 1] = order[i]; order[i] = tmp;
              setWHotels(wHotels.map(x => ({ ...x, priority: order.indexOf(x.id) })));
            }} style={{ border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 6, cursor: i === arr.length - 1 ? 'default' : 'pointer', opacity: i === arr.length - 1 ? 0.35 : 1, padding: '1px 7px' }}>↓</button>
          </div>
        ))}
      </div>

      <div style={box}>
        <label style={{ display: 'flex', gap: 8, alignItems: 'flex-start', fontSize: '0.84rem', cursor: 'pointer', marginBottom: 8 }}>
          <input type="checkbox" checked={!!wRules.keepGroups} onChange={e => setWRules({ ...wRules, keepGroups: e.target.checked })}
            style={{ width: 16, height: 16, marginTop: 2, accentColor: 'var(--dex-green, #86bc25)' }} />
          <span>
            {isDe ? 'Gruppen zusammen unterbringen' : 'Keep groups together'}
            <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Personen mit derselben Sub-Event-Kombination kommen möglichst ins selbe Haus.' : 'People with the same sub-event combination go into the same house where possible.'}
            </span>
          </span>
        </label>
        <label style={{ display: 'flex', gap: 8, alignItems: 'flex-start', fontSize: '0.84rem', cursor: 'pointer', marginBottom: 8 }}>
          <input type="checkbox" checked={!!wRules.skipNoWish} onChange={e => setWRules({ ...wRules, skipNoWish: e.target.checked })}
            style={{ width: 16, height: 16, marginTop: 2, accentColor: 'var(--dex-green, #86bc25)' }} />
          <span>
            {isDe ? '„Kein Hotel nötig" überspringen' : 'Skip „no accommodation"'}
            <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
              {fields.main
                ? (isDe ? `Wer „${fields.main.label}" mit Nein beantwortet hat, bekommt kein Zimmer zugeteilt.` : `Anyone who answered „${fields.main.label}" with no is not assigned a room.`)
                : (isDe ? 'Wer im Anmeldeformular ausdrücklich kein Hotel wollte, bekommt keins zugeteilt.' : 'Anyone who explicitly declined accommodation is not assigned.')}
            </span>
          </span>
        </label>
        <label style={{ display: 'flex', gap: 8, alignItems: 'flex-start', fontSize: '0.84rem', cursor: 'pointer' }}>
          <input type="checkbox" checked={overwrite} onChange={e => setOverwrite(e.target.checked)}
            style={{ width: 16, height: 16, marginTop: 2, accentColor: 'var(--dex-green, #86bc25)' }} />
          <span>
            {isDe ? 'Bestehende Zuordnungen überschreiben' : 'Overwrite existing assignments'}
            <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Aus — bereits zugeordnete Personen bleiben, wo sie sind (empfohlen).' : 'Off — already assigned people stay where they are (recommended).'}
            </span>
          </span>
        </label>
      </div>
    </div>
  );

  /* ---- Schritt 4: Ergebnis ---- */
  const renderPreview = (): JSX.Element => {
    const extraTotal = plan.rows.reduce((n, r) => n + r.extraBefore + r.extraAfter, 0);
    return (
      <div>
        <div style={question}>{isDe ? 'Passt das so?' : 'Does this look right?'}</div>
        <p style={explain}>
          {isDe
            ? <>So sähe die Verteilung aus. <strong>Extranächte</strong> sind Nächte außerhalb des Kontingent-Zeitraums des jeweiligen Hotels — die deckt euer Kontingent nicht ab, die müssen zusätzlich gebucht werden.</>
            : <>This is how the distribution would look. <strong>Extra nights</strong> fall outside each hotel’s capacity period — your contingent does not cover them, they have to be booked on top.</>}
        </p>

        {/* Wer wird ueberhaupt angefasst? */}
        <div style={{ ...box, marginTop: 0, marginBottom: 12, background: 'var(--dex-gray-50, #f7f7f5)' }}>
          <div style={{ fontSize: '0.82rem', lineHeight: 1.6 }}>
            {isDe
              ? <><strong>{people.length}</strong> aktive Teilnehmer · davon werden <strong>{plan.candidates}</strong> jetzt verteilt.</>
              : <><strong>{people.length}</strong> active attendees · <strong>{plan.candidates}</strong> will be distributed now.</>}
          </div>
          {(plan.excludedAssigned > 0 || plan.excludedNoWish > 0) && (
            <ul style={{ margin: '4px 0 0', paddingLeft: 18, fontSize: '0.79rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {plan.excludedAssigned > 0 && (
                <li>{isDe
                  ? <><strong>{plan.excludedAssigned}</strong> haben bereits ein Hotel und bleiben unberührt — setz oben in Schritt 3 „Bestehende Zuordnungen überschreiben", wenn sie mit verteilt werden sollen.</>
                  : <><strong>{plan.excludedAssigned}</strong> already have a hotel and stay untouched — tick „Overwrite existing assignments" in step 3 to include them.</>}</li>
              )}
              {plan.excludedNoWish > 0 && (
                <li>{isDe
                  ? <><strong>{plan.excludedNoWish}</strong> haben im Anmeldeformular kein Zimmer gewünscht.</>
                  : <><strong>{plan.excludedNoWish}</strong> declined a room in the registration form.</>}</li>
              )}
            </ul>
          )}
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 660 }}>
            <thead>
              <tr>
                <th style={th}>{isDe ? 'Hotel' : 'Hotel'}</th>
                <th style={th}>{isDe ? 'Personen' : 'People'}</th>
                <th style={th}>{isDe ? 'Kontingent-Zeitraum' : 'Capacity period'}</th>
                <th style={th}>{isDe ? 'Extranächte vorab' : 'Extra nights before'}</th>
                <th style={th}>{isDe ? 'Extranächte danach' : 'Extra nights after'}</th>
                <th style={th}>{isDe ? 'Übernachtungen' : 'Room nights'}</th>
              </tr>
            </thead>
            <tbody>
              {plan.rows.map(r => {
                const over = !!r.hotel.capacity && r.people.length > (r.hotel.capacity || 0);
                return (
                  <tr key={r.hotel.id} style={{ background: over ? '#fef3f2' : 'transparent' }}>
                    <td style={{ ...td, fontWeight: 600 }}>{r.hotel.name || (isDe ? '(ohne Namen)' : '(unnamed)')}</td>
                    <td style={td}>
                      {r.people.length}{r.hotel.capacity ? ` / ${r.hotel.capacity}` : ''}
                      {over && <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 700 }}> · {isDe ? 'über' : 'over'}</span>}
                    </td>
                    <td style={{ ...td, color: 'var(--dex-gray-600)' }}>
                      {r.base ? `${fmtDay(r.base.from, isDe)}–${fmtDay(r.base.to, isDe)}` : '—'}
                    </td>
                    <td style={{ ...td, fontWeight: r.extraBefore > 0 ? 700 : 400, color: r.extraBefore > 0 ? '#b35a00' : 'var(--dex-gray-500)' }}>
                      {r.extraBefore > 0 ? `${r.extraBefore}×` : '—'}
                    </td>
                    <td style={{ ...td, fontWeight: r.extraAfter > 0 ? 700 : 400, color: r.extraAfter > 0 ? '#b35a00' : 'var(--dex-gray-500)' }}>
                      {r.extraAfter > 0 ? `${r.extraAfter}×` : '—'}
                    </td>
                    <td style={td}>{r.nights}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {extraTotal > 0 && (
          <div style={{ ...box, background: '#fff6e5', borderColor: '#e0a300' }}>
            <div style={{ fontSize: '0.84rem', fontWeight: 700, marginBottom: 4 }}>
              {isDe ? `${extraTotal} Extranacht/Extranächte über das Kontingent hinaus` : `${extraTotal} extra night(s) beyond the contingent`}
            </div>
            <ul style={{ margin: 0, paddingLeft: 18, fontSize: '0.82rem', lineHeight: 1.6 }}>
              {plan.rows.filter(r => r.extraBefore + r.extraAfter > 0).map(r => (
                <li key={r.hotel.id}>
                  {isDe
                    ? <><strong>{r.hotel.name}</strong>: {r.extraBefore > 0 ? `${r.extraBefore}× eine Nacht früher (ab ${r.base ? fmtDay(addDays(r.base.from, -1), isDe) : '—'})` : ''}{r.extraBefore > 0 && r.extraAfter > 0 ? ', ' : ''}{r.extraAfter > 0 ? `${r.extraAfter}× eine Nacht länger` : ''}</>
                    : <><strong>{r.hotel.name}</strong>: {r.extraBefore > 0 ? `${r.extraBefore}× one night earlier` : ''}{r.extraBefore > 0 && r.extraAfter > 0 ? ', ' : ''}{r.extraAfter > 0 ? `${r.extraAfter}× one night longer` : ''}</>}
                </li>
              ))}
            </ul>
          </div>
        )}

        {plan.unplaced > 0 && (
          <div style={{ ...box, background: '#fef3f2', borderColor: 'var(--dex-red, #c00)' }}>
            <div style={{ fontSize: '0.84rem' }}>
              {isDe
                ? `${plan.unplaced} Person(en) bleiben ohne Hotel — das Kontingent reicht nicht. Geh zurück zu Schritt 2 und erhöhe ein Kontingent oder leg ein weiteres Hotel an.`
                : `${plan.unplaced} person(s) stay without a hotel — capacity is not sufficient. Go back to step 2 and raise a capacity or add another hotel.`}
            </div>
          </div>
        )}

        <div style={box}>
          <label style={{ display: 'flex', gap: 8, alignItems: 'flex-start', fontSize: '0.84rem', cursor: 'pointer' }}>
            <input type="checkbox" checked={doAssign} onChange={e => setDoAssign(e.target.checked)}
              style={{ width: 16, height: 16, marginTop: 2, accentColor: 'var(--dex-green, #86bc25)' }} />
            <span>
              {isDe ? `Zuordnung jetzt schreiben (${plan.assignments.length} Person(en))` : `Write assignments now (${plan.assignments.length} person(s))`}
              <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                {isDe
                  ? 'Aus — es werden nur Zeiträume, Hotels und Regeln gespeichert. Die Verteilung kannst du später jederzeit über „Automatisch verteilen" auslösen.'
                  : 'Off — only periods, hotels and rules are saved. You can run the distribution later via „Auto-distribute".'}
              </span>
            </span>
          </label>
        </div>

        <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginTop: 10, lineHeight: 1.55 }}>
          {isDe
            ? <>Gespeichert werden: <strong>{allStays.length}</strong> Zeitraum/Zeiträume, <strong>{wHotels.length}</strong> Hotel(s){doAssign ? <>, <strong>{plan.assignments.length}</strong> Zuordnung(en)</> : ''}. Die Hotel-Anzeige für Teilnehmer bleibt unverändert — die gibst du separat frei.</>
            : <>Will be saved: <strong>{allStays.length}</strong> period(s), <strong>{wHotels.length}</strong> hotel(s){doAssign ? <>, <strong>{plan.assignments.length}</strong> assignment(s)</> : ''}. Attendee visibility stays as it is — you release that separately.</>}
        </div>
      </div>
    );
  };

  return (
    // v28.61: `dismissable={false}` — ein Klick neben den Dialog (oder Escape)
    // schliesst ihn nicht mehr. Im Assistenten stecken vier Schritte Eingabe,
    // die beim versehentlichen Schliessen komplett weg waren. Raus geht es nur
    // noch bewusst ueber „Abbrechen".
    <Modal open={open} onClose={onClose} maxWidth={900} dismissable={false}
      ariaLabel={isDe ? 'Hotel-Planung einrichten' : 'Set up hotel planning'}>
      <h2 style={{ margin: 0, fontSize: '1.25rem', color: 'var(--dex-gray-800)' }}>
        {isDe ? 'Hotel-Planung einrichten' : 'Set up hotel planning'}
      </h2>
      <p style={{ margin: '4px 0 0', fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
        {isDe
          ? 'Wann braucht ihr Zimmer, welche Hotels habt ihr, wer kommt wohin — daraus ergeben sich die Verteilung und die Extranächte. Geschrieben wird erst am Ende.'
          : 'When do you need rooms, which hotels do you have, who goes where — that gives you the distribution and the extra nights. Nothing is written before the end.'}
      </p>

      {!showIntro && renderStepper()}

      <div style={{ maxHeight: '58vh', overflowY: 'auto', paddingRight: 4, marginTop: 10 }}>
        {showIntro && renderIntro()}
        {!showIntro && step === 1 && renderStays()}
        {!showIntro && step === 2 && renderHotels()}
        {!showIntro && step === 3 && renderRules()}
        {!showIntro && step === 4 && renderPreview()}
      </div>

      {/* Zwischenstand: eine Zeile mit dem, was der Assistent bis hierhin
          verstanden hat. Ohne das weiß man beim ersten Mal nicht, ob die
          Eingabe angekommen ist. */}
      {!showIntro && (
        <div style={{ marginTop: 12, padding: '8px 12px', borderRadius: 8, background: 'var(--dex-gray-50, #f7f7f5)', fontSize: '0.78rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
          <strong>{isDe ? 'Stand: ' : 'So far: '}</strong>
          {mainStay
            ? `${isDe ? 'Zimmer' : 'Rooms'} ${fmtDay(mainStay.from, isDe)}–${fmtDay(mainStay.to, isDe)} (${nightLabel(nightsBetween(mainStay.from, mainStay.to), isDe)})`
            : (isDe ? 'kein Zeitraum' : 'no period')}
          {allStays.length > 1 ? ` + ${allStays.length - 1} ${isDe ? 'Ausnahme(n)' : 'exception(s)'}` : ''}
          {' · '}
          {wHotels.filter(h => (h.name || '').trim()).length > 0
            ? `${wHotels.filter(h => (h.name || '').trim()).length} ${isDe ? 'Hotel(s)' : 'hotel(s)'}${totalCap > 0 ? `, ${totalCap} ${isDe ? 'Plätze' : 'places'}` : ''}`
            : (isDe ? 'noch kein Hotel' : 'no hotel yet')}
          {' · '}
          {isDe ? `${needBeds} Person(en) mit Bettenbedarf` : `${needBeds} person(s) needing a bed`}
        </div>
      )}

      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 16, flexWrap: 'wrap' }}>
        <button type="button" className="btn btn-secondary" disabled={busy} onClick={onClose}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        {showIntro && (
          <button type="button" className="btn btn-primary" disabled={busy} onClick={() => setShowIntro(false)}>
            {isDe ? 'Los geht’s' : 'Let’s go'}
          </button>
        )}
        {!showIntro && step > 1 && (
          <button type="button" className="btn btn-secondary" disabled={busy} onClick={() => setStep(step - 1)}>
            {isDe ? 'Zurück' : 'Back'}
          </button>
        )}
        {!showIntro && step < 4 && (
          <button type="button" className="btn btn-primary" disabled={busy || !canNext()} onClick={() => setStep(step + 1)}>
            {isDe ? 'Weiter' : 'Next'}
          </button>
        )}
        {!showIntro && step === 4 && (
          <button type="button" className="btn btn-primary" disabled={busy}
            onClick={() => {
              void onApply({
                stays: allStays,
                // Den Kontingent-Zeitraum festschreiben: In der Auswahl stand
                // der Haupt-Zeitraum nur als Vorbelegung — ohne Wert wäre die
                // Extranacht-Rechnung später vom Standard abhängig, und der
                // kann sich ändern.
                hotels: wHotels
                  .filter(h => (h.name || '').trim() !== '')
                  .map((h, i) => ({
                    ...h,
                    name: h.name.trim(),
                    priority: typeof h.priority === 'number' ? h.priority : i,
                    capacityStayId: h.capacityStayId || (mainStay ? mainStay.id : ''),
                  })),
                rules: {
                  ...wRules,
                  byAnswer: fields.extra
                    ? {
                      fieldId: fields.extra.id,
                      map: Object.keys(answerMap).reduce((acc: Record<string, string>, k: string) => {
                        const m = answerMap[k];
                        acc[k] = m && m.nights > 0 ? `auto_${m.after ? 'a' : 'b'}_${m.nights}` : (mainStay ? mainStay.id : '');
                        return acc;
                      }, {}),
                    }
                    : undefined,
                },
                assignments: doAssign ? plan.assignments : [],
              });
            }}>
            {busy ? (isDe ? 'Wird übernommen…' : 'Applying…') : (isDe ? 'Übernehmen' : 'Apply')}
          </button>
        )}
      </div>
    </Modal>
  );
};

export default HotelSetupWizard;
