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
 * Organizer gewählte Bezeichnung, sonst „Programm".
 *
 * v31.9: Optik nach `docs/ui-leitfaden.md` — jeder Cluster ist ein
 * `dex-ui-section` mit Abschnitts-Kopf, die Punkte sind `dex-ui-row`-Zeilen
 * ohne Hover (`--static`: hier gibt es nichts zu klicken, und ein Hover
 * verspricht eine Aktion, Grundsatz 3). Die Cluster stehen untereinander
 * statt in einem mehrspaltigen Raster: Auf dem Handy war die zweite Spalte
 * ohnehin ein Umbruch, und untereinander bleibt die Reihenfolge lesbar.
 * Die Gruppierung selbst kommt unverändert aus `agendaGroups()`. */
import * as React from 'react';
import { DeloitteEvent } from '../../types';
import { Locale } from '../../context/LanguageContext';
import { FileText } from '../Icons';
import { agendaGroups, sortAgenda, groupLabel, groupDateLabel } from '../../utils/agendaGroups';

export interface AgendaProgramBlockProps {
  event: DeloitteEvent;
  locale: Locale;
}

export const AgendaProgramBlock: React.FC<AgendaProgramBlockProps> = ({ event, locale }) => {
  const items = (event.agenda || []).filter(a => a && (a.title || a.time));
  if (items.length === 0) return null;
  const isDe = locale === 'de';
  const sorted = sortAgenda(items);
  // v30.94: Cluster (utils/agendaGroups) — benannte Gruppe, sonst der Tag.
  const byDay = agendaGroups(items);
  const heading = event.agendaCheckIn
    ? (event.agendaTermPlural || (isDe ? 'Programmpunkte' : 'Agenda items'))
    : (isDe ? 'Programm' : 'Schedule');
  return (
    <div style={{ padding: '12px 16px', borderTop: '1px solid var(--dex-gray-200)', background: '#fff' }}>
      <div className="dex-ui-section-title">
        <FileText size={14} />
        <span>{heading}</span>
        <span className="dex-ui-pill dex-ui-pill--gray dex-ui-pill--sm">{sorted.length}</span>
      </div>
      {event.agendaCheckIn && (
        <div className="dex-ui-section-desc">
          {isDe ? 'Die Anmeldung gilt fürs ganze Event.' : 'Registration covers the whole event.'}
        </div>
      )}
      {byDay.map((g, gi) => (
        <div key={g.key} className="dex-ui-section">
          {(byDay.length > 1 || g.date || g.cluster) && (
            <div className="dex-ui-section-title">
              <span>{groupLabel(g, gi, isDe)}</span>
              <span style={{ fontWeight: 500, textTransform: 'none', letterSpacing: 0 }}>
                {g.dates.length ? groupDateLabel(g, isDe) : (isDe ? 'Termin folgt' : 'Date to be announced')}
              </span>
            </div>
          )}
          <div className="dex-ui-card dex-ui-card--soft dex-ui-card--list">
            {g.items.map(it => (
              <div key={it.id} className="dex-ui-row dex-ui-row--static">
                <span style={{ flexShrink: 0, minWidth: 92, fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-700)', fontVariantNumeric: 'tabular-nums' }}>
                  {it.time || '—'}{it.endTime ? ` – ${it.endTime}` : ''}
                </span>
                <div className="dex-ui-row-main">
                  <div className="dex-ui-row-title dex-ui-row-title--wrap">
                    {/* v30.99: ohne laufende Nummer — sie stand vor jedem Titel
                        („17. Post Merger …") und störte (Nutzer 07.09.2026); die
                        Reihenfolge ergibt sich aus Cluster und Uhrzeit. */}
                    {it.title || (isDe ? '(ohne Titel)' : '(untitled)')}
                  </div>
                  {(it.location || it.description) && (
                    <div className="dex-ui-row-sub" style={{ whiteSpace: 'normal', wordBreak: 'break-word' }}>
                      {it.location && <span>{it.location}</span>}
                      {it.location && it.description && <span> · </span>}
                      {it.description && <span>{it.description}</span>}
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
};

export default AgendaProgramBlock;
