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
 * Geführte Erst-Einrichtung der Hotel-Planung. Vorher musste man sich die
 * Reihenfolge selbst zusammensuchen: Zeitraum anlegen → Hotel anlegen →
 * Kontingent eintragen → verteilen → hoffen, dass es passt. Der Assistent
 * führt genau diese Kette in vier Schritten und rechnet am Ende aus, was das
 * für die Buchung bedeutet.
 *
 * Der eigentliche Grund für den Assistenten steckt in Schritt 4: Hotels geben
 * ihr **Kontingent fast immer für eine bestimmte Nacht** heraus („40 Zimmer für
 * die Nacht 24./25."). Wer früher anreist oder später abreist, ist damit NICHT
 * abgedeckt — diese Nächte müssen separat gebucht werden. Genau diese Zahl
 * („Hotel A: 10× Extranacht vorab") war bisher Handarbeit in Excel.
 *
 * Rechenmodell:
 *  - Jede Person bekommt einen **Zeitraum**. Er kommt (in dieser Reihenfolge)
 *    aus einer Sub-Event-Regel, aus einer bereits gesetzten Zuordnung oder aus
 *    dem Standard-Zeitraum.
 *  - Jedes Hotel hat ein **Kontingent** und einen **Kontingent-Zeitraum**
 *    (`capacityStayId`). Eine Person belegt einen Kontingent-Platz, egal wie
 *    lange sie bleibt.
 *  - **Extranächte** = Nächte der Person außerhalb des Kontingent-Zeitraums,
 *    getrennt nach „vorab" (frühere Anreise) und „danach" (spätere Abreise).
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

/** Tage zwischen zwei Tagen, nie negativ (für die Extranacht-Rechnung). */
const daysGap = (from: string, to: string): number => {
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

export const HotelSetupWizard: React.FC<IHotelSetupWizardProps> = (props: IHotelSetupWizardProps) => {
  const { open, event, people, childEvents, subEmails, hotels, stays, rules, isDe, busy, wishOf, onClose, onApply } = props;

  /* ------------------------------------------------------------------ *
   * Schritt 1 — Zeiträume
   * ------------------------------------------------------------------ */

  /** Drei Vorschläge aus dem Event-Datum: Kern, Kern + Vorabend,
   *  Kern + Vorabend + Extranacht. Bei einem Tagesevent ergibt das genau die
   *  klassischen 1 / 2 / 3 Nächte. */
  const suggestions = React.useMemo<DexHotelStay[]>(() => {
    const core = coreStay(event);
    const n = Math.max(1, nightsBetween(core.from, core.to));
    const mk = (from: string, to: string, extra: string): DexHotelStay => ({
      id: `sug_${from}_${to}`,
      label: `${nightLabel(nightsBetween(from, to), isDe)}${extra ? ` · ${extra}` : ''}`,
      from, to,
    });
    return [
      mk(core.from, core.to, isDe ? 'Standard' : 'standard'),
      mk(addDays(core.from, -1), core.to, isDe ? 'mit Vorabend' : 'with prior evening'),
      mk(addDays(core.from, -1), addDays(core.to, 1), isDe ? 'Vorabend + Abreisetag' : 'prior evening + extra day'),
    ].filter((s, i, arr) => arr.findIndex(x => x.from === s.from && x.to === s.to) === i && n > 0);
  }, [event, isDe]);

  const [wStays, setWStays] = React.useState<DexHotelStay[]>([]);
  const [wHotels, setWHotels] = React.useState<DexHotel[]>([]);
  const [wRules, setWRules] = React.useState<DexHotelRules>({});
  const [step, setStep] = React.useState(1);
  const [doAssign, setDoAssign] = React.useState(true);
  const [overwrite, setOverwrite] = React.useState(false);
  const [newStay, setNewStay] = React.useState<{ label: string; from: string; to: string }>({ label: '', from: '', to: '' });

  // Beim Öffnen frisch aufsetzen: Vorhandenes übernehmen, sonst die Vorschläge.
  React.useEffect(() => {
    if (!open) return;
    setStep(1);
    setWStays(stays.length > 0
      ? stays.map(s => ({ ...s }))
      : suggestions.map((s, i) => ({ ...s, id: uid('s'), isDefault: i === 0 })));
    setWHotels(hotels.length > 0
      ? hotels.map((h, i) => ({ ...h, priority: typeof h.priority === 'number' ? h.priority : i }))
      : []);
    setWRules({
      bySub: (rules && rules.bySub) ? { ...rules.bySub } : {},
      keepGroups: rules && typeof rules.keepGroups === 'boolean' ? rules.keepGroups : true,
      skipNoWish: rules && typeof rules.skipNoWish === 'boolean' ? rules.skipNoWish : true,
    });
    setDoAssign(true);
    setOverwrite(false);
    const core = coreStay(event);
    setNewStay({ label: '', from: core.from, to: core.to });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const defaultStay = React.useMemo<DexHotelStay | null>(
    () => wStays.filter(s => s.isDefault)[0] || wStays[0] || null,
    [wStays],
  );

  const toggleSuggestion = (s: DexHotelStay): void => {
    const hit = wStays.filter(x => x.from === s.from && x.to === s.to)[0];
    if (hit) {
      const next = wStays.filter(x => x.id !== hit.id);
      if (next.length > 0 && !next.some(x => x.isDefault)) next[0] = { ...next[0], isDefault: true };
      setWStays(next);
      return;
    }
    setWStays(wStays.concat([{ ...s, id: uid('s'), isDefault: wStays.length === 0 }]));
  };

  /* ------------------------------------------------------------------ *
   * Schritt 4 — Verteilung + Extranächte
   * ------------------------------------------------------------------ */

  /** Sub-Events, für die diese Person angemeldet ist. */
  const subsOf = React.useCallback((p: SPRegistration): DeloitteEvent[] => {
    const em = (p.ParticipantEmail || '').trim().toLowerCase();
    if (!em) return [];
    return childEvents.filter(c => (subEmails[c.id] || []).indexOf(em) >= 0);
  }, [childEvents, subEmails]);

  /** Der Zeitraum, den diese Person bekommen soll.
   *  Sub-Event-Regeln gewinnen und werden zur Hülle vereinigt: früheste Anreise,
   *  späteste Abreise. Genau daraus entstehen die Extranächte — wer zum
   *  Vorabend-Dinner kommt, reist eben einen Tag früher an. */
  const stayFor = React.useCallback((p: SPRegistration): { from: string; to: string } => {
    const bySub = wRules.bySub || {};
    let from = ''; let to = '';
    for (const c of subsOf(p)) {
      const sid = (bySub[c.id] || {}).stayId;
      if (!sid) continue;
      const st = wStays.filter(s => s.id === sid)[0];
      if (!st) continue;
      if (!from || st.from < from) from = st.from;
      if (!to || st.to > to) to = st.to;
    }
    if (from && to) return { from, to };
    const exFrom = toDay(p.HotelFrom); const exTo = toDay(p.HotelTo);
    if (exFrom && exTo) return { from: exFrom, to: exTo };
    return defaultStay ? { from: defaultStay.from, to: defaultStay.to } : { from: '', to: '' };
  }, [wRules, wStays, subsOf, defaultStay]);

  /** Festes Hotel aus einer Sub-Event-Regel (erste Regel gewinnt). */
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
    /** Kontingent-Zeitraum, gegen den die Extranächte gerechnet werden. */
    base: DexHotelStay | null;
    extraBefore: number;
    extraAfter: number;
    nights: number;
  }

  const plan = React.useMemo<{ rows: IPlanRow[]; unplaced: number; candidates: number; assignments: Array<{ reg: SPRegistration; hotel: string; from: string; to: string }> }>(() => {
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
      base: wStays.filter(s => s.id === h.capacityStayId)[0] || defaultStay,
      extraBefore: 0, extraAfter: 0, nights: 0,
    }));
    const byName: Record<string, IPlanRow> = {};
    for (const r of rows) byName[r.hotel.name] = r;

    const candidates = people.filter(p => {
      if (!overwrite && (p.Hotel || '').trim()) return false;
      if (wRules.skipNoWish && wishOf(p) === false) return false;
      return true;
    });

    // Restkontingent: bereits zugeordnete Personen zählen mit, sonst
    // überbuchen wir beim zweiten Durchlauf.
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
        r.extraBefore += daysGap(s.from, r.base.from);
        r.extraAfter += daysGap(r.base.to, s.to);
      }
    }
    return { rows, unplaced, candidates: candidates.length, assignments };
  }, [wHotels, wStays, wRules, people, overwrite, defaultStay, forcedHotel, subsOf, stayFor, wishOf]);

  const needBeds = React.useMemo(
    () => people.filter(p => !(wRules.skipNoWish && wishOf(p) === false)).length,
    [people, wRules.skipNoWish, wishOf],
  );
  const totalCap = wHotels.reduce((n, h) => n + (h.capacity || 0), 0);

  /* ------------------------------------------------------------------ *
   * Darstellung
   * ------------------------------------------------------------------ */

  const box: React.CSSProperties = { border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 12, marginTop: 12 };
  const smallInp: React.CSSProperties = { height: 34, fontSize: '0.84rem', padding: '0 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, minWidth: 0 };
  const th: React.CSSProperties = { textAlign: 'left', padding: '6px 8px', fontSize: '0.74rem', color: 'var(--dex-gray-600)', borderBottom: '1px solid var(--dex-gray-200)', whiteSpace: 'nowrap' };
  const td: React.CSSProperties = { padding: '6px 8px', fontSize: '0.82rem', borderBottom: '1px solid var(--dex-gray-100)' };

  const STEPS = isDe
    ? ['Zeiträume', 'Hotels & Kontingente', 'Regeln', 'Vorschau']
    : ['Stay periods', 'Hotels & capacity', 'Rules', 'Preview'];

  const canNext = (): boolean => {
    if (step === 1) return wStays.length > 0;
    if (step === 2) return wHotels.length > 0 && wHotels.every(h => (h.name || '').trim() !== '');
    return true;
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

  /* ---- Schritt 1 ---- */
  const renderStays = (): JSX.Element => (
    <div>
      <p style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, margin: '0 0 8px' }}>
        {isDe
          ? <>Welche Aufenthalts-Zeiträume braucht ihr? Die Vorschläge kommen aus dem Event-Datum ({fmtDay(toDay(event.startDate), true)}{toDay(event.endDate) && toDay(event.endDate) !== toDay(event.startDate) ? `–${fmtDay(toDay(event.endDate), true)}` : ''}). Ein Klick nimmt einen Vorschlag auf oder wieder heraus.</>
          : <>Which stay periods do you need? The suggestions are derived from the event dates. Click to add or remove one.</>}
      </p>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(190px, 1fr))', gap: 8 }}>
        {suggestions.map(s => {
          const on = wStays.some(x => x.from === s.from && x.to === s.to);
          return (
            <button key={s.id} type="button" onClick={() => toggleSuggestion(s)} style={{
              textAlign: 'left', cursor: 'pointer', padding: '10px 12px', borderRadius: 10,
              border: `1.5px solid ${on ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
              background: on ? 'rgba(134,188,37,0.08)' : '#fff',
            }}>
              <div style={{ fontWeight: 700, fontSize: '0.86rem' }}>
                {on ? '✓ ' : ''}{s.label}
              </div>
              <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                {fmtDay(s.from, isDe)} – {fmtDay(s.to, isDe)}
              </div>
            </button>
          );
        })}
      </div>

      <div style={box}>
        <div style={{ fontSize: '0.8rem', fontWeight: 600, marginBottom: 8 }}>
          {isDe ? 'Ausgewählte Zeiträume' : 'Selected periods'}
        </div>
        {wStays.length === 0 && (
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-red, #c00)' }}>
            {isDe ? 'Bitte mindestens einen Zeitraum wählen.' : 'Please select at least one period.'}
          </div>
        )}
        {wStays.map(s => (
          <div key={s.id} style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', padding: '5px 0', borderBottom: '1px solid var(--dex-gray-100)' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: '0.78rem', cursor: 'pointer' }}
              title={isDe ? 'Standard für neue Zuordnungen' : 'Default for new assignments'}>
              <input type="radio" checked={!!s.isDefault} onChange={() => setWStays(wStays.map(x => ({ ...x, isDefault: x.id === s.id })))}
                style={{ accentColor: 'var(--dex-green, #86bc25)' }} />
              {isDe ? 'Standard' : 'Default'}
            </label>
            <input style={{ ...smallInp, flex: '2 1 160px' }} value={s.label}
              onChange={e => setWStays(wStays.map(x => x.id === s.id ? { ...x, label: e.target.value } : x))} />
            <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
              {fmtDay(s.from, isDe)} – {fmtDay(s.to, isDe)} · {nightLabel(nightsBetween(s.from, s.to), isDe)}
            </span>
            <button type="button" onClick={() => {
              const next = wStays.filter(x => x.id !== s.id);
              if (next.length > 0 && !next.some(x => x.isDefault)) next[0] = { ...next[0], isDefault: true };
              setWStays(next);
            }} style={{ marginLeft: 'auto', border: 'none', background: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', fontSize: '0.95rem' }}>×</button>
          </div>
        ))}

        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'flex-end', marginTop: 10 }}>
          <div className="form-group" style={{ marginBottom: 0, flex: '2 1 170px' }}>
            <label className="form-label" style={{ fontSize: '0.78rem', marginBottom: 4 }}>{isDe ? 'Eigener Zeitraum' : 'Custom period'}</label>
            <input className="form-input" placeholder={isDe ? 'Bezeichnung' : 'Label'} value={newStay.label}
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
            }}>
            {isDe ? '+ Hinzufügen' : '+ Add'}
          </button>
        </div>
      </div>
    </div>
  );

  /* ---- Schritt 2 ---- */
  const renderHotels = (): JSX.Element => (
    <div>
      <p style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, margin: '0 0 8px' }}>
        {isDe
          ? <>Trag die Hotels mit ihrem Kontingent ein. <strong>Wichtig:</strong> Das Kontingent gilt fast immer nur für einen bestimmten Zeitraum — wähle rechts, für welchen. Alles darüber hinaus rechnet der Assistent im letzten Schritt als Extranacht aus.</>
          : <>Add the hotels with their capacity. <strong>Important:</strong> capacity usually applies to one specific period — pick it on the right. Anything beyond becomes an extra night in the last step.</>}
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
                  <select style={{ ...smallInp, width: '100%' }} value={h.capacityStayId || (defaultStay ? defaultStay.id : '')}
                    onChange={e => setWHotels(wHotels.map(x => x.id === h.id ? { ...x, capacityStayId: e.target.value } : x))}>
                    {wStays.map(s => (
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
          </tbody>
        </table>
      </div>
      <button type="button" className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 16px', marginTop: 10 }}
        onClick={() => setWHotels(wHotels.concat([{ id: uid('h'), name: '', address: '', capacity: 0, notes: '', priority: wHotels.length, capacityStayId: defaultStay ? defaultStay.id : '' }]))}>
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

  /* ---- Schritt 3 ---- */
  const renderRules = (): JSX.Element => (
    <div>
      <p style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, margin: '0 0 8px' }}>
        {isDe
          ? 'Regeln bestimmen, wer wohin kommt und wie lange. Alles ist optional — ohne Regel verteilt der Assistent nach Reihenfolge und Kontingent.'
          : 'Rules decide who goes where and for how long. Everything is optional — without rules the wizard fills by order and capacity.'}
      </p>

      {childEvents.length > 0 && (
        <div style={box}>
          <div style={{ fontSize: '0.84rem', fontWeight: 700, marginBottom: 2 }}>
            {isDe ? 'Feste Zuordnung je Sub-Event' : 'Fixed assignment per sub-event'}
          </div>
          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 8, lineHeight: 1.45 }}>
            {isDe
              ? 'Z.B. „alle vom Vorabend-Dinner ins Hotel A, mit 2 Nächten". Wer in mehreren Sub-Events ist, bekommt die Hülle aus allen Zeiträumen — früheste Anreise, späteste Abreise. Genau daraus entstehen die Extranächte.'
              : 'E.g. „everyone from the prior-evening dinner into hotel A, 2 nights". People in several sub-events get the envelope of all periods — earliest arrival, latest departure. That is where extra nights come from.'}
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
                          {wStays.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                        </select>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {!childEvents.every(c => !!subEmails[c.id]) && (
            <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', marginTop: 6 }}>
              {isDe ? 'Hinweis: Für Sub-Events ohne Personenzahl ist die Teilnehmerliste noch nicht geladen — die Regel greift trotzdem, sobald sie da ist.'
                : 'Note: sub-events without a count have not loaded their list yet — the rule still applies once it has.'}
            </div>
          )}
        </div>
      )}

      <div style={box}>
        <div style={{ fontSize: '0.84rem', fontWeight: 700, marginBottom: 2 }}>
          {isDe ? 'Füll-Reihenfolge' : 'Fill order'}
        </div>
        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 8 }}>
          {isDe ? 'Welches Hotel soll zuerst voll werden? Oben zuerst.' : 'Which hotel fills up first? Top first.'}
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
            {isDe ? '„Keine Unterkunft" überspringen' : 'Skip „no accommodation"'}
            <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Wer im Anmeldeformular ausdrücklich kein Hotel wollte, bekommt keins zugeteilt.' : 'Anyone who explicitly declined accommodation is not assigned.'}
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

  /* ---- Schritt 4 ---- */
  const renderPreview = (): JSX.Element => {
    const extraTotal = plan.rows.reduce((n, r) => n + r.extraBefore + r.extraAfter, 0);
    return (
      <div>
        <p style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, margin: '0 0 10px' }}>
          {isDe
            ? <>So sähe die Verteilung aus. <strong>Extranächte</strong> sind Nächte außerhalb des Kontingent-Zeitraums des jeweiligen Hotels — die musst du zusätzlich buchen, das Kontingent deckt sie nicht ab.</>
            : <>This is how the distribution would look. <strong>Extra nights</strong> fall outside each hotel’s capacity period — you have to book those on top, the contingent does not cover them.</>}
        </p>

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
              {isDe ? `${extraTotal} Extranacht/Extranächte zusätzlich buchen` : `${extraTotal} extra night(s) to book on top`}
            </div>
            <ul style={{ margin: 0, paddingLeft: 18, fontSize: '0.82rem', lineHeight: 1.6 }}>
              {plan.rows.filter(r => r.extraBefore + r.extraAfter > 0).map(r => (
                <li key={r.hotel.id}>
                  {isDe
                    ? <><strong>{r.hotel.name}</strong>: {r.extraBefore > 0 ? `${r.extraBefore}× eine Nacht früher (Anreise ab ${r.base ? fmtDay(addDays(r.base.from, -1), isDe) : '—'})` : ''}{r.extraBefore > 0 && r.extraAfter > 0 ? ', ' : ''}{r.extraAfter > 0 ? `${r.extraAfter}× eine Nacht länger` : ''}</>
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
                ? `${plan.unplaced} Person(en) bleiben ohne Hotel — das Kontingent reicht nicht. Erhöhe ein Kontingent oder lege ein weiteres Hotel an.`
                : `${plan.unplaced} person(s) stay without a hotel — capacity is not sufficient. Raise a capacity or add another hotel.`}
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
            ? <>Gespeichert werden: <strong>{wStays.length}</strong> Zeitraum/Zeiträume, <strong>{wHotels.length}</strong> Hotel(s){doAssign ? <>, <strong>{plan.assignments.length}</strong> Zuordnung(en)</> : ''}. Die Hotel-Anzeige für Teilnehmer bleibt unverändert — die gibst du separat frei.</>
            : <>Will be saved: <strong>{wStays.length}</strong> period(s), <strong>{wHotels.length}</strong> hotel(s){doAssign ? <>, <strong>{plan.assignments.length}</strong> assignment(s)</> : ''}. Attendee visibility stays as it is — you release that separately.</>}
        </div>
      </div>
    );
  };

  return (
    <Modal open={open} onClose={onClose} maxWidth={900} dismissable={!busy}
      ariaLabel={isDe ? 'Hotel-Planung einrichten' : 'Set up hotel planning'}>
      <h2 style={{ margin: 0, fontSize: '1.25rem', color: 'var(--dex-gray-800)' }}>
        {isDe ? 'Hotel-Planung einrichten' : 'Set up hotel planning'}
      </h2>
      <p style={{ margin: '4px 0 0', fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
        {isDe
          ? 'In vier Schritten von den Zeiträumen bis zur fertigen Verteilung — inklusive der Rechnung, wie viele Extranächte ihr über das Kontingent hinaus buchen müsst. Geschrieben wird erst am Ende.'
          : 'Four steps from stay periods to a finished distribution — including how many extra nights you have to book beyond the contingent. Nothing is written before the last step.'}
      </p>

      {renderStepper()}

      <div style={{ maxHeight: '58vh', overflowY: 'auto', paddingRight: 4, marginTop: 10 }}>
        {step === 1 && renderStays()}
        {step === 2 && renderHotels()}
        {step === 3 && renderRules()}
        {step === 4 && renderPreview()}
      </div>

      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 16, flexWrap: 'wrap' }}>
        <button type="button" className="btn btn-secondary" disabled={busy} onClick={onClose}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        {step > 1 && (
          <button type="button" className="btn btn-secondary" disabled={busy} onClick={() => setStep(step - 1)}>
            {isDe ? 'Zurück' : 'Back'}
          </button>
        )}
        {step < 4 && (
          <button type="button" className="btn btn-primary" disabled={busy || !canNext()} onClick={() => setStep(step + 1)}>
            {isDe ? 'Weiter' : 'Next'}
          </button>
        )}
        {step === 4 && (
          <button type="button" className="btn btn-primary" disabled={busy}
            onClick={() => {
              void onApply({
                stays: wStays,
                // Den Kontingent-Zeitraum festschreiben: In der Auswahl stand
                // der Standard nur als Vorbelegung — ohne Wert wäre die
                // Extranacht-Rechnung später vom Standard abhängig, und der
                // kann sich ändern.
                hotels: wHotels
                  .filter(h => (h.name || '').trim() !== '')
                  .map((h, i) => ({
                    ...h,
                    name: h.name.trim(),
                    priority: typeof h.priority === 'number' ? h.priority : i,
                    capacityStayId: h.capacityStayId || (defaultStay ? defaultStay.id : ''),
                  })),
                rules: wRules,
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
