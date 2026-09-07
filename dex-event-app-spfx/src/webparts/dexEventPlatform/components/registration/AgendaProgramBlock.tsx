/* AgendaProgramBlock — v30.86.
 * Das Programm eines Events auf der Anmeldeseite: die Agenda-Punkte aus
 * „Ort & Programm", chronologisch, mit Tages-Zwischenüberschriften. Nur
 * lesen — wer sich anmeldet, meldet sich fürs Event an, nicht für Punkte.
 *
 * Warum jetzt: Der Wizard-Tooltip versprach seit v22, die Agenda erscheine
 * „auf der Anmelde-Seite und in Meine Events" — gerendert wurde sie nur in
 * Meine Events. Mit Programmpunkten (Check-in je Punkt) ist die Liste die
 * Gliederung des Events, und die muss der Teilnehmer VOR der Anmeldung sehen.
 * Bei Events mit Programmpunkte-Modus trägt die Überschrift die vom
 * Organizer gewählte Bezeichnung, sonst „Programm". */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { AgendaItem, DeloitteEvent } from '../../types';
import { Locale } from '../../context/LanguageContext';

export interface AgendaProgramBlockProps {
  event: DeloitteEvent;
  locale: Locale;
}

export const AgendaProgramBlock: React.FC<AgendaProgramBlockProps> = ({ event, locale }) => {
  const items = (event.agenda || []).filter(a => a && (a.title || a.time));
  if (items.length === 0) return null;
  const isDe = locale === 'de';
  const sorted = items.slice().sort((a, b) => ((a.date || '') + (a.time || '')).localeCompare((b.date || '') + (b.time || '')));
  const byDay: Array<[string, AgendaItem[]]> = [];
  sorted.forEach(it => {
    const key = it.date || '';
    const last = byDay[byDay.length - 1];
    if (last && last[0] === key) last[1].push(it); else byDay.push([key, [it]]);
  });
  const heading = event.agendaCheckIn
    ? (event.agendaTermPlural || (isDe ? 'Programmpunkte' : 'Agenda items'))
    : (isDe ? 'Programm' : 'Schedule');
  const fmtDay = (d: string): string => {
    if (!d) return isDe ? 'Termin folgt' : 'Date to be announced';
    try { return new Date(d + 'T00:00').toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' }); } catch { return d; }
  };
  return (
    <div style={{ padding: '12px 16px', borderTop: '1px solid var(--dex-gray-200)', background: '#fff' }}>
      <div style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--dex-gray-700)', marginBottom: 8, display: 'flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="BulletedList2" style={{ fontSize: 14, color: 'var(--dex-green-dark, #6b9a1e)' }} />
        {heading}
        <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}>· {sorted.length}</span>
        {event.agendaCheckIn && (
          <span style={{ fontWeight: 400, fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
            {isDe ? '— die Anmeldung gilt fürs ganze Event' : '— registration covers the whole event'}
          </span>
        )}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: byDay.length > 1 ? 'repeat(auto-fit, minmax(240px, 1fr))' : '1fr', gap: 12 }}>
        {byDay.map(([day, list]) => (
          <div key={day || 'tbd'} style={{ background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 10, minWidth: 0 }}>
            {(byDay.length > 1 || day) && (
              <div style={{ fontSize: '0.76rem', fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)', marginBottom: 6 }}>{fmtDay(day)}</div>
            )}
            {list.map(it => (
              <div key={it.id} style={{ display: 'flex', gap: 10, padding: '5px 0', borderLeft: '2px solid var(--dex-green, #86bc25)', paddingLeft: 10, marginLeft: 2 }}>
                <div style={{ flexShrink: 0, minWidth: 92, fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-700)', fontVariantNumeric: 'tabular-nums' }}>
                  {it.time || '—'}{it.endTime ? ` – ${it.endTime}` : ''}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: '0.82rem', fontWeight: 600, wordBreak: 'break-word' }}>
                    {sorted.indexOf(it) + 1}. {it.title || (isDe ? '(ohne Titel)' : '(untitled)')}
                  </div>
                  {(it.location || it.description) && (
                    <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 1, wordBreak: 'break-word' }}>
                      {it.location && <span>{it.location}</span>}
                      {it.location && it.description && <span> · </span>}
                      {it.description && <span>{it.description}</span>}
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
};

export default AgendaProgramBlock;
