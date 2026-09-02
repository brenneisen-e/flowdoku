/* DuplicateInSubEventHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10737-10772 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';

export interface DuplicateInSubEventHintBoxProps {
  isDe: boolean;
  subEventDupGroups: { sectionTitle: string; email: string; name: string; count: number; rowsInfo: string; }[];
}

export const DuplicateInSubEventHintBox: React.FC<DuplicateInSubEventHintBoxProps> = (p) => {
  const { isDe, subEventDupGroups } = p;
          // v28.21: ECHTE Doppel-Anmeldung im Klammer-Modus — dieselbe Person
          // zweimal aktiv IM SELBEN Sub-Event. Nur das belegt zwei Plätze und
          // gehört rot gemeldet; die doppelten Klammer-Schatten-Zeilen laufen
          // über den neutralen Hinweis oben.
          if (subEventDupGroups.length === 0) return null;
          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-red, #c00)', background: 'rgba(200,0,0,0.06)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                <Icon iconName="Warning" style={{ fontSize: 18, color: 'var(--dex-red, #c00)' }} />
                <strong style={{ color: 'var(--dex-red, #c00)', fontSize: '0.95rem' }}>
                  {isDe ? `Doppel-Anmeldungen erkannt (${subEventDupGroups.length})` : `Duplicate registrations detected (${subEventDupGroups.length})`}
                </strong>
                <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Dieselbe Person ist im selben Sub-Event mehrfach angemeldet und belegt dort zwei Plätze. Achtung: Die doppelten Zeilen können auch auf der WARTELISTE des Sub-Events stehen (dann fehlen sie in der Teilnehmer-Tabelle) — Status und Zeitpunkt stehen hinter jedem Eintrag. Im jeweiligen Sub-Event-Tab entfernst du die doppelte Zeile über „Abmelden" bzw. „Entfernen" (Warteliste) still.'
                    : 'The same person is registered more than once in the same sub-event and occupies two seats there. Note: the duplicate rows may sit on the sub-event’s WAITLIST (then they are missing from the attendee table) — status and time are shown behind each entry. Remove the duplicate row silently via „Cancel" or „Remove" (waitlist) in that sub-event tab.'}
                </span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {subEventDupGroups.map(g => (
                  <div key={`${g.sectionTitle}::${g.email}`} style={{ fontSize: '0.84rem', color: 'var(--dex-gray-800)', display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline' }}>
                    <strong>{g.name}</strong>
                    <span style={{ color: 'var(--dex-gray-500)' }}>{g.email}</span>
                    <span style={{ padding: '1px 8px', borderRadius: 999, background: 'var(--dex-red, #c00)', color: '#fff', fontSize: '0.72rem', fontWeight: 700 }}>
                      {isDe ? `${g.count}× angemeldet` : `${g.count}× registered`}
                    </span>
                    <span style={{ color: 'var(--dex-gray-600)' }}>— {g.sectionTitle}</span>
                    {/* v30.14: Status + Zeitpunkt je Zeile — Warteliste-Duplikate sind sonst nicht auffindbar. */}
                    <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>({g.rowsInfo})</span>
                  </div>
                ))}
              </div>
            </div>
          );
};

