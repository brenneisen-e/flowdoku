/* OptionDateCalendar — v32.2.2. Auswahl-Feld mit Daten als Antworten wird
 * als Monatsraster gezeigt (Erkennung: utils/optionDates). Nur die
 * angebotenen Tage sind klickbar; der gespeicherte Wert bleibt der
 * Options-Text („19.11.2026", bei Mehrfachauswahl „A | B") — wie beim
 * Dropdown, damit Mails, Export und Organizer Center unverändert lesen.
 * Optik wie der Termin-Kalender der Anmeldeseite (v28.91). */
import * as React from 'react';
import { dayKey } from '../../utils/optionDates';
import { Check } from '../Icons';

export interface OptionDateCalendarProps {
  options: string[];
  dates: Date[];
  /** Anzeige-Text je Option (EN-Variante), gleiche Reihenfolge wie `options`. */
  labels?: string[];
  multi: boolean;
  value: string;
  onChange: (next: string) => void;
  isDe: boolean;
  error?: boolean;
}

const SEP = ' | ';

export const OptionDateCalendar: React.FC<OptionDateCalendarProps> = ({ options, dates, labels, multi, value, onChange, isDe, error }) => {
  const opts = options.map(o => (o || '').trim()).filter(Boolean);
  const byDay: Record<string, { opt: string; idx: number }> = {};
  dates.forEach((d, i) => { if (opts[i] && !byDay[dayKey(d)]) byDay[dayKey(d)] = { opt: opts[i], idx: i }; });
  const selected = (value || '').split(SEP).map(s => s.trim()).filter(Boolean);
  const monthKeys = Array.from(new Set(dates.map(d => `${d.getFullYear()}-${d.getMonth() + 1}`)))
    .sort((a, b) => {
      const [ay, am] = a.split('-').map(Number); const [by, bm] = b.split('-').map(Number);
      return ay !== by ? ay - by : am - bm;
    });
  const weekdays = isDe ? ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'] : ['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su'];
  const toggle = (opt: string): void => {
    if (!multi) { onChange(selected[0] === opt ? '' : opt); return; }
    const has = selected.indexOf(opt) >= 0;
    // Reihenfolge der Optionen beibehalten, nicht die der Klicks.
    const next = opts.filter(o => (o === opt ? !has : selected.indexOf(o) >= 0));
    onChange(next.join(SEP));
  };
  const labelOf = (idx: number, opt: string): string => (labels && labels[idx] && labels[idx].trim()) || opt;
  return (
    <div role="group" style={{
      border: `1px solid ${error ? 'var(--dex-red, #c00)' : 'var(--dex-gray-200, #e5e5e5)'}`,
      borderRadius: 10, padding: '10px 12px', background: 'var(--dex-white, #fff)',
    }}>
      {monthKeys.map(mk => {
        const [my, mm] = mk.split('-').map(n => parseInt(n, 10));
        const first = new Date(my, mm - 1, 1);
        const daysInMonth = new Date(my, mm, 0).getDate();
        const lead = (first.getDay() + 6) % 7;
        const cells: Array<Date | null> = [];
        for (let i = 0; i < lead; i++) cells.push(null);
        for (let d = 1; d <= daysInMonth; d++) cells.push(new Date(my, mm - 1, d));
        while (cells.length % 7 !== 0) cells.push(null);
        // Nur die Wochen von der ersten bis zur letzten angebotenen — ein
        // ganzer Monat für einen einzelnen Tag ist Scroll-Weg ohne Nutzen.
        const rows: Array<Array<Date | null>> = [];
        for (let i = 0; i < cells.length; i += 7) rows.push(cells.slice(i, i + 7));
        const hasOpt = (r: Array<Date | null>): boolean => r.some(c => !!c && !!byDay[dayKey(c)]);
        const firstRow = rows.findIndex(hasOpt);
        let lastRow = rows.length - 1;
        while (lastRow > firstRow && !hasOpt(rows[lastRow])) lastRow--;
        const shown = rows.slice(Math.max(0, firstRow), lastRow + 1).reduce<Array<Date | null>>((acc, r) => acc.concat(r), []);
        return (
          <div key={mk} style={{ marginBottom: 8 }}>
            <div style={{ fontWeight: 700, fontSize: '0.86rem', marginBottom: 6, color: 'var(--dex-gray-700, #444)' }}>
              {first.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { month: 'long', year: 'numeric' })}
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
              {weekdays.map(w => (
                <div key={w} style={{ fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-gray-400)', textAlign: 'center', padding: '2px 0' }}>{w}</div>
              ))}
              {shown.map((d, i) => {
                if (!d) return <div key={`e${i}`} />;
                const k = dayKey(d);
                const entry = byDay[k];
                if (!entry) {
                  return <div key={k} style={{ textAlign: 'center', padding: '8px 0', fontSize: '0.8rem', color: 'var(--dex-gray-300, #ccc)' }}>{d.getDate()}</div>;
                }
                const isSel = selected.indexOf(entry.opt) >= 0;
                return (
                  <button
                    key={k}
                    type="button"
                    aria-pressed={isSel}
                    title={labelOf(entry.idx, entry.opt)}
                    onClick={() => toggle(entry.opt)}
                    style={{
                      position: 'relative', textAlign: 'center', padding: '8px 0', borderRadius: 8, cursor: 'pointer',
                      fontSize: '0.84rem', fontWeight: 700, fontFamily: 'inherit',
                      border: `1.5px solid ${isSel ? 'var(--dex-green, #86BC25)' : 'var(--dex-green-light, #cfe6a4)'}`,
                      background: isSel ? 'var(--dex-green, #86BC25)' : 'var(--dex-green-50, #f4f9ea)',
                      color: isSel ? '#fff' : 'var(--dex-gray-800, #222)',
                    }}
                  >
                    {d.getDate()}
                    {isSel && <span style={{ position: 'absolute', top: 2, right: 3, lineHeight: 0 }}><Check size={10} /></span>}
                  </button>
                );
              })}
            </div>
          </div>
        );
      })}
      <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500, #777)', marginTop: 2 }}>
        {selected.length === 0
          ? (multi
            ? (isDe ? 'Tippe auf einen oder mehrere hervorgehobene Tage.' : 'Tap one or more highlighted days.')
            : (isDe ? 'Tippe auf einen hervorgehobenen Tag.' : 'Tap a highlighted day.'))
          : (isDe ? 'Gewählt: ' : 'Selected: ') + selected.map(s => {
            const idx = opts.indexOf(s);
            return idx >= 0 ? labelOf(idx, s) : s;
          }).join(', ')}
      </div>
    </div>
  );
};

export default OptionDateCalendar;
