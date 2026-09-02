/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import DatePicker, { registerLocale } from 'react-datepicker';
import { de } from 'date-fns/locale';
import 'react-datepicker/dist/react-datepicker.css';

registerLocale('de', de);

/**
 * StayRangePicker (v28.63)
 * ------------------------
 * Das Eingabe-Element für den Feldtyp `daterange` („Übernachtungs-Zeitraum").
 *
 * Warum ein eigener Feldtyp: Bisher haben Organizer die Übernachtung über zwei
 * Auswahllisten abgefragt — „Hotel (24-25 Sept): Ja/Nein" plus „Hotel
 * (additional nights): Yes, I need ONE additional night from 23 - 24 Sept.".
 * Das funktioniert, hat aber drei Nachteile: die Daten stecken als Fließtext in
 * der Antwort, jede Variante muss der Organizer von Hand anlegen, und die
 * Hotel-Planung muss den Text hinterher wieder deuten. Mit diesem Feld sagt der
 * Teilnehmer direkt, von wann bis wann er ein Zimmer braucht — die Nächte
 * ergeben sich daraus, und die Hotel-Planung übernimmt sie ohne Umweg.
 *
 * Gespeichert wird ein einziger, stabil parsebarer String:
 *   - `''`                        → noch nicht beantwortet
 *   - `'-'`                       → ausdrücklich KEIN Hotel nötig
 *   - `'2026-09-23 – 2026-09-25'` → Anreise – Abreise (ISO, en-dash)
 * ISO deshalb, weil der Wert unverändert in Exporte und Listen läuft und dort
 * eindeutig sein muss (24/09 vs. 09/24).
 */

/** Wert = „kein Hotel nötig". */
export const STAY_NONE = '-';

const pad = (n: number): string => String(n).padStart(2, '0');

export const toLocalDay = (d: Date): string =>
  `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

export const dayToDate = (day: string): Date | null => {
  if (!day || day.length < 10) return null;
  const d = new Date(`${day}T00:00:00`);
  return isNaN(d.getTime()) ? null : d;
};

export const stayNights = (from: string, to: string): number => {
  if (!from || !to) return 0;
  const a = Date.parse(`${from}T00:00:00Z`);
  const b = Date.parse(`${to}T00:00:00Z`);
  if (isNaN(a) || isNaN(b) || b <= a) return 0;
  return Math.round((b - a) / 86400000);
};

/** Gespeicherten Wert zerlegen. */
export const parseStayValue = (value: string): { none: boolean; from: string; to: string } => {
  const v = (value || '').trim();
  if (!v) return { none: false, from: '', to: '' };
  if (v === STAY_NONE) return { none: true, from: '', to: '' };
  const m = v.match(/(\d{4}-\d{2}-\d{2})\s*[–-]\s*(\d{4}-\d{2}-\d{2})/);
  if (m) return { none: false, from: m[1], to: m[2] };
  return { none: false, from: '', to: '' };
};

export const formatStayValue = (from: string, to: string): string =>
  (from && to) ? `${from} – ${to}` : '';

/** Für die Anzeige in Listen, Mails und „Meine Events". */
export const formatStayLabel = (value: string, isDe: boolean): string => {
  const p = parseStayValue(value);
  if (p.none) return isDe ? 'Kein Hotel nötig' : 'No hotel needed';
  if (!p.from || !p.to) return value || '';
  const fmt = (d: string): string =>
    new Date(`${d}T00:00:00Z`).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const n = stayNights(p.from, p.to);
  return `${fmt(p.from)} – ${fmt(p.to)} (${n} ${isDe ? (n === 1 ? 'Nacht' : 'Nächte') : (n === 1 ? 'night' : 'nights')})`;
};

export interface IStayRangePickerProps {
  value: string;
  onChange: (next: string) => void;
  isDe: boolean;
  /** Frühester wählbarer Tag (ISO). Leer = keine Grenze. */
  rangeStart?: string;
  /** Spätester wählbarer Tag (ISO). Leer = keine Grenze. */
  rangeEnd?: string;
  /** Obergrenze für die Nächte. 0/leer = keine. */
  maxNights?: number;
  /** Pflichtfeld — dann muss aktiv „kein Hotel" gewählt werden. */
  required?: boolean;
  disabled?: boolean;
  compact?: boolean;
}

export const StayRangePicker: React.FC<IStayRangePickerProps> = (props: IStayRangePickerProps) => {
  const { value, onChange, isDe, rangeStart, rangeEnd, maxNights, required, disabled, compact } = props;
  const parsed = parseStayValue(value);
  const nights = stayNights(parsed.from, parsed.to);
  const overMax = !!maxNights && maxNights > 0 && nights > maxNights;

  const minD = dayToDate(rangeStart || '');
  const maxD = dayToDate(rangeEnd || '');

  const setFrom = (d: Date | null): void => {
    if (!d) { onChange(formatStayValue('', parsed.to)); return; }
    const from = toLocalDay(d);
    // Abreise nachziehen, wenn sie nicht mehr passt — sonst steht dort ein
    // Datum vor der Anreise und das Feld zeigt 0 Nächte.
    let to = parsed.to;
    if (!to || stayNights(from, to) <= 0) {
      // v30.67: Kalendertag statt +24 h. `d` ist LOKALE Mitternacht, und in
      // der Nacht, in der die Sommerzeit endet, hat der Tag 25 Stunden —
      // +86400000 ms landete dann auf 23:00 desselben Datums, die Abreise
      // fiel auf den Anreisetag und das Feld zeigte „0 Nächte". setDate
      // rechnet über die Kalender-Komponenten und überspringt den Sprung.
      const cand = new Date(d);
      cand.setDate(cand.getDate() + 1);
      to = (maxD && cand > maxD) ? toLocalDay(maxD) : toLocalDay(cand);
    }
    onChange(formatStayValue(from, to));
  };

  const setTo = (d: Date | null): void => {
    if (!d) { onChange(formatStayValue(parsed.from, '')); return; }
    onChange(formatStayValue(parsed.from, toLocalDay(d)));
  };

  const dp = {
    dateFormat: 'dd.MM.yyyy',
    locale: isDe ? 'de' : undefined,
    placeholderText: isDe ? 'TT.MM.JJJJ' : 'dd/mm/yyyy',
    className: 'form-input',
    wrapperClassName: 'dex-datepicker-wrapper',
    calendarClassName: 'dex-datepicker-calendar',
    popperPlacement: 'bottom-start' as const,
    autoComplete: 'off',
    disabled: disabled || parsed.none,
  };

  const fmtShort = (d: string): string =>
    d ? new Date(`${d}T00:00:00Z`).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '';

  return (
    <div>
      <div style={{ display: 'flex', gap: compact ? 8 : 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
        <div style={{ flex: '1 1 150px', minWidth: 0 }}>
          <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>
            {isDe ? 'Anreise' : 'Arrival'}
          </div>
          <DatePicker {...dp}
            selected={dayToDate(parsed.from)}
            onChange={setFrom}
            minDate={minD || undefined}
            maxDate={maxD || undefined}
          />
        </div>
        <div style={{ flex: '1 1 150px', minWidth: 0 }}>
          <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>
            {isDe ? 'Abreise' : 'Departure'}
          </div>
          <DatePicker {...dp}
            selected={dayToDate(parsed.to)}
            onChange={setTo}
            minDate={dayToDate(parsed.from) || minD || undefined}
            maxDate={maxD || undefined}
          />
        </div>
        <div style={{
          paddingBottom: 12, fontSize: '0.88rem', fontWeight: 700, whiteSpace: 'nowrap',
          color: overMax ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)',
        }}>
          {parsed.none
            ? ''
            : `= ${nights} ${isDe ? (nights === 1 ? 'Nacht' : 'Nächte') : (nights === 1 ? 'night' : 'nights')}`}
        </div>
      </div>

      {(rangeStart || rangeEnd) && !parsed.none && (
        <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
          {isDe
            ? `Buchbar ${rangeStart ? `ab ${fmtShort(rangeStart)}` : ''}${rangeStart && rangeEnd ? ' ' : ''}${rangeEnd ? `bis ${fmtShort(rangeEnd)}` : ''}${maxNights ? ` · maximal ${maxNights} ${maxNights === 1 ? 'Nacht' : 'Nächte'}` : ''}`
            : `Bookable ${rangeStart ? `from ${fmtShort(rangeStart)}` : ''}${rangeStart && rangeEnd ? ' ' : ''}${rangeEnd ? `until ${fmtShort(rangeEnd)}` : ''}${maxNights ? ` · max. ${maxNights} ${maxNights === 1 ? 'night' : 'nights'}` : ''}`}
        </div>
      )}

      {overMax && (
        <div style={{ fontSize: '0.76rem', color: 'var(--dex-red, #c00)', marginTop: 4 }}>
          {isDe
            ? `Es sind höchstens ${maxNights} ${maxNights === 1 ? 'Nacht' : 'Nächte'} vorgesehen.`
            : `At most ${maxNights} ${maxNights === 1 ? 'night' : 'nights'} are allowed.`}
        </div>
      )}

      {/* Ausdrücklich „kein Hotel" — ohne diese Option wäre ein leeres Feld
          zweideutig: nicht ausgefüllt oder nicht gebraucht? */}
      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, marginTop: 8, fontSize: '0.82rem', cursor: disabled ? 'default' : 'pointer', color: 'var(--dex-gray-700)' }}>
        <input
          type="checkbox"
          checked={parsed.none}
          disabled={disabled}
          onChange={e => onChange(e.target.checked ? STAY_NONE : '')}
          style={{ width: 16, height: 16, accentColor: 'var(--dex-green, #86bc25)' }}
        />
        {isDe ? 'Ich brauche kein Hotel' : 'I don’t need a hotel'}
        {required && <span style={{ color: 'var(--dex-gray-500)' }}> ({isDe ? 'oder Zeitraum wählen' : 'or pick a period'})</span>}
      </label>
    </div>
  );
};

export default StayRangePicker;
