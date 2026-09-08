/* InactiveAccountsBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6852-6928 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a/5b umgebaut — `dex-ui-callout--warn` statt
 * eigener Orange-Fläche, jede Person eine Zeile mit Personen-Zelle
 * (`dex-ui-person`) und ihren Aktionen darin statt einer Aufzählung mit
 * angehängten Mini-Knöpfen. Die ganze Zeile springt zur Person (Leitfaden 2a′),
 * „Abmelden" stoppt die Weitergabe; der Erklärtext steht auf zwei Zeilen, ohne
 * dass eine Aussage entfällt.
 */
import * as React from 'react';
import { AlertCircle, Search, Trash2 } from '../../Icons';
import { ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { ConsolidatedRow } from '../../admin/adminTypes';

export interface InactiveAccountsBoxProps {
  consolidatedRows: ConsolidatedRow[];
  inactiveAccounts: string[];
  isAdmin: boolean;
  isConsolidatedMode: boolean;
  isDe: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  jumpToParticipant: (email: string) => void;
  openDeregModal: (row: ConsolidatedRow) => void;
  orgPastLock: boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
}

export const InactiveAccountsBox: React.FC<InactiveAccountsBoxProps> = (p) => {
  const { consolidatedRows, inactiveAccounts, isAdmin, isConsolidatedMode, isDe, isOrganizerFor, jumpToParticipant, openDeregModal, orgPastLock, registrations, selectedEvent } = p;
  // Idempotent — der Kasten steht ohne Modal/WizardFormShell in der Event-Seite.
  ensureDexUiStyles();
        const items = inactiveAccounts.map(email => {
          const reg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === email);
          // v29.30: Bei einer Klammer steht die Person oft NUR in den
          // Sub-Event-Listen — dann blieb hier nur die nackte Adresse stehen.
          const cRow = consolidatedRows.find(r => r.emailKey === email);
          const name = reg
            ? ((reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || email))
            : (cRow ? (`${cRow.vorname} ${cRow.nachname}`.trim() || email) : email);
          return { email, name, cRow };
        });
        // v31.3: Initialen statt Foto — dieser Kasten hat keine Profilbilder,
        // und eine Personen-Zelle ohne Kachel links bricht die Zeilenoptik.
        const initials = (n: string): string =>
          n.split(/[\s.@_-]+/).filter(Boolean).slice(0, 2).map(s => s.charAt(0).toUpperCase()).join('') || '?';
        // Unverändert gegenüber v29.30, nur einmal statt je Zeile ausgewertet:
        // dieselbe Reihenfolge, also auch dasselbe Kurzschluss-Verhalten.
        const canDereg = isConsolidatedMode && (isAdmin || isOrganizerFor(selectedEvent)) && !orgPastLock;
        return (
          <div className="dex-ui-callout dex-ui-callout--warn" style={{ flexDirection: 'column', marginBottom: 20 }}>
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={18} /></span>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: 700 }}>
                  {isDe
                    ? (items.length === 1 ? 'Eine Person hat womöglich Deloitte verlassen' : `${items.length} Personen haben womöglich Deloitte verlassen`)
                    : `${items.length === 1 ? 'One person' : `${items.length} people`} may have left Deloitte`}
                </div>
                <div style={{ marginTop: 3 }}>
                  {isDe
                    ? 'Zu diesen Adressen gibt es kein aktives Deloitte-Konto mehr (deaktiviert oder gelöscht) — Mails und Outlook-Termine kommen dort womöglich nicht an. Schau sie dir an und melde ab, wer nicht mehr dabei ist.'
                    : 'There is no active Deloitte account for these addresses any more (disabled or deleted) — emails and Outlook invites may not arrive. Review them and deregister whoever is no longer with us.'}
                </div>
              </div>
            </div>
            <ul className="dex-ui-card dex-ui-card--list" style={{ margin: 0, listStyle: 'none', background: '#fff' }}>
              {items.map(it => (
                // Klick auf die Zeile tut dasselbe wie „Zur Person springen".
                <li key={it.email} className="dex-ui-row dex-ui-row--bordered" style={{ cursor: 'pointer' }}
                  onClick={() => jumpToParticipant(it.email)}
                  title={isDe ? 'In der Teilnehmerliste anzeigen' : 'Show in participant list'}>
                  <div className="dex-ui-person" style={{ flex: 1 }}>
                    <span className="dex-ui-avatar" aria-hidden="true">{initials(it.name)}</span>
                    <div style={{ minWidth: 0 }}>
                      <div className="dex-ui-person-name">{it.name}</div>
                      <div className="dex-ui-person-sub">{it.email}</div>
                    </div>
                  </div>
                  <div className="dex-ui-row-actions">
                    <button type="button" className="dex-ui-textbtn"
                      onClick={e => { e.stopPropagation(); jumpToParticipant(it.email); }}>
                      <Search size={14} />{isDe ? 'Zur Person springen' : 'Jump to person'}
                    </button>
                    {/* v29.30: Abmelden DIREKT aus der Box. Der Text darüber
                        fordert dazu auf („melde ab, wer nicht mehr dabei ist"),
                        führte aber nur zur Zeile — und dort steckt das
                        Abmelden im generischen „Aktionen…"-Klappmenü, das
                        man erst öffnen muss. Der Dialog kommt mit still +
                        Klammer vorausgewählt (v29.29). */}
                    {canDereg && it.cRow && (
                      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--danger"
                        onClick={e => { e.stopPropagation(); if (it.cRow) openDeregModal(it.cRow); }}
                        title={isDe
                          ? 'Abmelde-Dialog öffnen — ohne Mail und Outlook-Absage, inklusive Klammer und aller Sub-Events'
                          : 'Open the cancellation dialog — no mail or Outlook withdrawal, including umbrella and all sub-events'}>
                        <Trash2 size={13} />{isDe ? 'Abmelden' : 'Deregister'}
                      </button>
                    )}
                  </div>
                </li>
              ))}
            </ul>
          </div>
        );
};

