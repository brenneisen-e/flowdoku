/* DateOptionsPicker — v32.3. Antwortmöglichkeiten einer Auswahl-Frage per
 * Klick in einen Kalender setzen (Schalter „Als Kalender"). Nutzer-Ansage
 * 28.09.2026: „einen Schalter ‚als Kalender', und dann kann man selbst im
 * Event-Assistenten den Kalender anklicken — vorbefüllt mit der Auswahl".
 *
 * Die Optionen bleiben Text im Format „19.11.2026", sortiert. So lesen Mails,
 * Export und Organizer Center sie unverändert, und die Anmeldeseite erkennt
 * sie über utils/optionDates wieder als Daten. */
import * as React from 'react';
import { dayKey } from '../../utils/optionDates';
import { ChevronLeft, ChevronRight } from '../Icons';

export interface DateOptionsPickerProps {
  dates: Date[];
  /** Monat, der ohne gewählte Tage zuerst erscheint (z. B. Event-Start). */
  startHint?: string;
  onChange: (options: string[]) => void;
  isDe: boolean;
  disabled?: boolean;
}

export const formatOptionDate = (d: Date): string =>
  `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;

export const DateOptionsPicker: React.FC<DateOptionsPickerProps> = ({ dates, startHint, onChange, isDe, disabled }) => {
  const first = React.useMemo(() => {
    if (dates.length > 0) return new Date(Math.min(...dates.map(d => d.getTime())));
    const h = startHint ? new Date(startHint) : null;
    return h && isFinite(h.getTime()) ? h : new Date();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps -- nur der Startmonat
  const [month, setMonth] = React.useState<{ y: number; m: number }>({ y: first.getFullYear(), m: first.getMonth() });
  const selected = new Set(dates.map(dayKey));
  const toggle = (d: Date): void => {
    const k = dayKey(d);
    const next = selected.has(k) ? dates.filter(x => dayKey(x) !== k) : dates.concat([d]);
    next.sort((a, b) => a.getTime() - b.getTime());
    onChange(next.map(formatOptionDate));
  };
  const shift = (delta: number): void => setMonth(prev => {
    const d = new Date(prev.y, prev.m + delta, 1);
    return { y: d.getFullYear(), m: d.getMonth() };
  });
  const firstOfMonth = new Date(month.y, month.m, 1);
  const lead = (firstOfMonth.getDay() + 6) % 7;
  const daysInMonth = new Date(month.y, month.m + 1, 0).getDate();
  const cells: Array<Date | null> = [];
  for (let i = 0; i < lead; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(new Date(month.y, month.m, d));
  const weekdays = isDe ? ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'] : ['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su'];
  const navBtn: React.CSSProperties = { width: 28, height: 28, padding: 0, justifyContent: 'center' };
  return (
    <div style={{ maxWidth: 340 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
        <button type="button" className="dex-ui-iconbtn" style={navBtn} onClick={() => shift(-1)} disabled={disabled}
          aria-label={isDe ? 'Vorheriger Monat' : 'Previous month'}><ChevronLeft size={16} /></button>
        <strong style={{ fontSize: '0.88rem' }}>
          {firstOfMonth.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { month: 'long', year: 'numeric' })}
        </strong>
        <button type="button" className="dex-ui-iconbtn" style={navBtn} onClick={() => shift(1)} disabled={disabled}
          aria-label={isDe ? 'Nächster Monat' : 'Next month'}><ChevronRight size={16} /></button>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
        {weekdays.map(w => (
          <div key={w} style={{ fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-gray-400)', textAlign: 'center' }}>{w}</div>
        ))}
        {cells.map((d, i) => {
          if (!d) return <div key={`e${i}`} />;
          const on = selected.has(dayKey(d));
          return (
            <button
              key={dayKey(d)}
              type="button"
              disabled={disabled}
              aria-pressed={on}
              onClick={() => toggle(d)}
              style={{
                padding: '6px 0', borderRadius: 8, cursor: disabled ? 'default' : 'pointer', fontFamily: 'inherit',
                fontSize: '0.82rem', fontWeight: on ? 700 : 500,
                border: `1px solid ${on ? 'var(--dex-green, #86BC25)' : 'var(--dex-gray-200, #e5e5e5)'}`,
                background: on ? 'var(--dex-green, #86BC25)' : 'var(--dex-white, #fff)',
                color: on ? '#fff' : 'var(--dex-gray-700, #444)',
              }}
            >{d.getDate()}</button>
          );
        })}
      </div>
      <div className="dex-ui-muted" style={{ fontSize: '0.78rem', marginTop: 6 }}>
        {dates.length === 0
          ? (isDe ? 'Tippe die Tage an, die zur Auswahl stehen sollen.' : 'Tap the days participants can choose.')
          : (isDe ? `${dates.length} ${dates.length === 1 ? 'Tag' : 'Tage'} gewählt: ` : `${dates.length} day${dates.length === 1 ? '' : 's'} selected: `)
            + dates.map(formatOptionDate).join(', ')}
      </div>
    </div>
  );
};

export default DateOptionsPicker;
