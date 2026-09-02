/* DuplicateRegHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10632-10735 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import { DUP_ACTIVE_STATI } from '../../admin/adminConstants';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../../types';

export interface DuplicateRegHintBoxProps {
  cleanupShadowDuplicates: () => Promise<void>;
  duplicateEmails: Set<string>;
  isAdmin: boolean;
  isConsolidatedMode: boolean;
  isDe: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  shadowDupBusy: boolean;
}

export const DuplicateRegHintBox: React.FC<DuplicateRegHintBoxProps> = (p) => {
  const { cleanupShadowDuplicates, duplicateEmails, isAdmin, isConsolidatedMode, isDe, isOrganizerFor, registrations, selectedEvent, shadowDupBusy } = p;
          // v23.2: Doppel-Anmelde-Hinweis. Listet jede Person, die mit
          // derselben E-Mail ≥2 nicht-abgemeldete Zeilen hat (z.B. dieselbe
          // Person in zwei Teams). Die betroffenen Zeilen sind in der Tabelle
          // zusätzlich rot markiert; pro Person kann der Organizer über den
          // „Abmelden"-Button das Duplikat still entfernen.
          if (duplicateEmails.size === 0 || !selectedEvent) return null;
          // Pro betroffener E-Mail die aktiven Zeilen sammeln (Name + Teams).
          const dupGroups: Array<{ email: string; rows: SPRegistration[] }> = [];
          duplicateEmails.forEach(em => {
            const rows = registrations.filter(r => DUP_ACTIVE_STATI.indexOf(r.Status || '') >= 0 && (r.ParticipantEmail || '').trim().toLowerCase() === em);
            if (rows.length > 1) dupGroups.push({ email: em, rows });
          });
          if (dupGroups.length === 0) return null;
          // v28.21: Klammer-Modus — doppelte Zeilen auf der KLAMMER sind keine
          // Doppel-Anmeldungen. Die Klammer-Zeile ist nur eine Schatten-Zeile
          // zur Datenvollständigkeit (kein Platz, keine Mail, kein Outlook);
          // die echten Anmeldungen liegen in den Sub-Events, und alle Zähler
          // rechnen ohnehin pro Person entdoppelt. Solche Zeilen entstehen
          // z.B., wenn zwei verschiedene Assistenzen dieselbe Person nach-
          // einander anmelden — die Vorab-Prüfung sieht die fremde Zeile
          // wegen der Zeilen-Berechtigungen nicht. Deshalb hier ein neutraler
          // technischer Hinweis statt der roten Doppel-Anmelde-Warnung.
          if (isConsolidatedMode) {
            return (
              <div style={{ marginBottom: 20, padding: 14, borderRadius: 12, border: '1px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #f7f7f7)' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 8 }}>
                  <Icon iconName="Info" style={{ fontSize: 16, color: 'var(--dex-gray-500)' }} />
                  <strong style={{ fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>
                    {isDe ? `Doppelte Klammer-Zeilen (${dupGroups.length})` : `Duplicate overall-event rows (${dupGroups.length})`}
                  </strong>
                </div>
                <p style={{ margin: '0 0 8px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? <>Diese Personen haben zwei Zeilen auf dem Gesamt-Event. Das ist <strong>keine Doppel-Anmeldung</strong>: Die Klammer-Zeile hält nur die Antworten auf die Hauptevent-Felder — sie belegt keinen Platz und löst weder Mail noch Outlook-Termin aus. Die echten Anmeldungen stehen in den {selectedEvent.childEventTermPlural || 'Sub-Events'}, und alle Zähler rechnen pro Person entdoppelt. Typische Ursache: zwei verschiedene Assistenzen haben dieselbe Person nacheinander angemeldet. Aufräumen ist optional.</>
                    : <>These people have two rows on the overall event. This is <strong>not a duplicate registration</strong>: the overall-event row only holds the answers to the main-event fields — it takes no seat and triggers neither email nor calendar invite. The real registrations are in the sub-events, and all counters de-duplicate per person. Typical cause: two different assistants registered the same person one after another. Cleaning up is optional.</>}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {dupGroups.map(g => {
                    const first = g.rows[0];
                    const dispName = (first.Vorname && first.Nachname) ? `${first.Vorname} ${first.Nachname}` : (first.ParticipantName || g.email);
                    return (
                      <div key={g.email} style={{ fontSize: '0.8rem', color: 'var(--dex-gray-700)', display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline' }}>
                        <strong>{dispName}</strong>
                        <span style={{ color: 'var(--dex-gray-500)' }}>{g.email}</span>
                        <span style={{ color: 'var(--dex-gray-500)' }}>— {isDe ? `${g.rows.length} Klammer-Zeilen` : `${g.rows.length} overall-event rows`}</span>
                      </div>
                    );
                  })}
                </div>
                {(isAdmin || isOrganizerFor(selectedEvent)) && (
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ marginTop: 12, fontSize: '0.8rem', padding: '6px 14px' }}
                    disabled={shadowDupBusy}
                    onClick={() => { cleanupShadowDuplicates().catch(() => { /* */ }); }}
                    title={isDe
                      ? 'Entfernt je Person die überzählige Klammer-Zeile — still, ohne Mail/Outlook/Nachrücken.'
                      : 'Removes the surplus overall-event row per person — silently, no email/Outlook/promotion.'}
                  >
                    {shadowDupBusy
                      ? (isDe ? 'Wird bereinigt…' : 'Cleaning up…')
                      : (isDe ? 'Doppelte Zeilen bereinigen' : 'Clean up duplicate rows')}
                  </button>
                )}
              </div>
            );
          }
          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-red, #c00)', background: 'rgba(200,0,0,0.06)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                <Icon iconName="Warning" style={{ fontSize: 18, color: 'var(--dex-red, #c00)' }} />
                <strong style={{ color: 'var(--dex-red, #c00)', fontSize: '0.95rem' }}>
                  {isDe ? `Doppel-Anmeldungen erkannt (${dupGroups.length})` : `Duplicate registrations detected (${dupGroups.length})`}
                </strong>
                <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Dieselbe Person ist mehrfach angemeldet. Die betroffenen Zeilen sind unten rot markiert — über „Abmelden" kannst du die doppelte Zeile still entfernen.'
                    : 'The same person is registered more than once. The affected rows are marked red below — use „Cancel" to silently remove the duplicate row.'}
                </span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {dupGroups.map(g => {
                  const first = g.rows[0];
                  const dispName = (first.Vorname && first.Nachname) ? `${first.Vorname} ${first.Nachname}` : (first.ParticipantName || g.email);
                  const teamList = g.rows
                    .map(r => r.TeamName ? `„${r.TeamName}"` : (r.TeamId ? (isDe ? '(Team ohne Namen)' : '(unnamed team)') : (isDe ? '(Einzel-Anmeldung)' : '(individual)')))
                    .join(', ');
                  return (
                    <div key={g.email} style={{ fontSize: '0.84rem', color: 'var(--dex-gray-800)', display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline' }}>
                      <strong>{dispName}</strong>
                      <span style={{ color: 'var(--dex-gray-500)' }}>{g.email}</span>
                      <span style={{ padding: '1px 8px', borderRadius: 999, background: 'var(--dex-red, #c00)', color: '#fff', fontSize: '0.72rem', fontWeight: 700 }}>
                        {isDe ? `${g.rows.length}× angemeldet` : `${g.rows.length}× registered`}
                      </span>
                      <span style={{ color: 'var(--dex-gray-600)' }}>— {teamList}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
};

