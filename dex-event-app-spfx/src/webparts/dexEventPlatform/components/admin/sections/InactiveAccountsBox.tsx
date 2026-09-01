/* InactiveAccountsBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6852-6928 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { AlertCircle, Trash2 } from '../../Icons';
import { Icon } from '@fluentui/react/lib/Icon';
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
        return (
          <div style={{
            display: 'flex', alignItems: 'flex-start', gap: 12,
            padding: '14px 16px', marginBottom: 20,
            background: '#fff3e0', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 12,
            color: 'var(--dex-gray-800)',
          }}>
            <AlertCircle size={20} />
            <div style={{ minWidth: 0 }}>
              <div style={{ fontWeight: 700, marginBottom: 4, color: 'var(--dex-orange-dark, #b35a00)' }}>
                {isDe
                  ? `${items.length === 1 ? 'Eine Person' : `${items.length} Personen`} hat womöglich Deloitte verlassen`
                  : `${items.length === 1 ? 'One person' : `${items.length} people`} may have left Deloitte`}
              </div>
              <div style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, marginBottom: 8 }}>
                {isDe
                  ? 'Zu folgenden Teilnehmern wurde kein aktives Deloitte-Konto mehr gefunden — das Konto ist deaktiviert oder existiert nicht mehr. Mails/Outlook-Termine an diese Adressen kommen ggf. nicht an. Bitte prüfen und ggf. abmelden.'
                  : 'No active Deloitte account was found for the following participants — the account is disabled or no longer exists. Emails/Outlook invites to these addresses may not arrive. Please review and deregister if needed.'}
              </div>
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>
                {items.map(it => (
                  <li key={it.email} style={{ marginBottom: 4 }}>
                    <strong>{it.name}</strong> <span style={{ color: 'var(--dex-gray-500)' }}>({it.email})</span>
                    <button
                      type="button"
                      onClick={() => jumpToParticipant(it.email)}
                      title={isDe ? 'In der Teilnehmerliste anzeigen' : 'Show in participant list'}
                      style={{
                        marginLeft: 8, display: 'inline-flex', alignItems: 'center', gap: 4,
                        background: 'transparent', border: '1px solid var(--dex-orange, #ed8b00)', color: 'var(--dex-orange-dark, #b35a00)',
                        borderRadius: 6, padding: '1px 8px', cursor: 'pointer', fontSize: '0.74rem', fontWeight: 600, fontFamily: 'inherit',
                      }}
                    >
                      <Icon iconName="Search" style={{ fontSize: 11 }} />
                      {isDe ? 'Zur Person springen' : 'Jump to person'}
                    </button>
                    {/* v29.30: Abmelden DIREKT aus der Box. Der Text darüber
                        fordert dazu auf („Bitte prüfen und ggf. abmelden"),
                        führte aber nur zur Zeile — und dort steckt das
                        Abmelden im generischen „Aktionen…"-Klappmenü, das
                        man erst öffnen muss. Der Dialog kommt mit still +
                        Klammer vorausgewählt (v29.29). */}
                    {isConsolidatedMode && it.cRow && (isAdmin || isOrganizerFor(selectedEvent)) && !orgPastLock && (
                      <button
                        type="button"
                        onClick={() => { if (it.cRow) openDeregModal(it.cRow); }}
                        title={isDe
                          ? 'Abmelde-Dialog öffnen — ohne Mail und Outlook-Absage, inklusive Klammer und aller Sub-Events'
                          : 'Open the cancellation dialog — no mail or Outlook withdrawal, including umbrella and all sub-events'}
                        style={{
                          marginLeft: 6, display: 'inline-flex', alignItems: 'center', gap: 4,
                          background: 'var(--dex-red, #c00)', border: '1px solid var(--dex-red, #c00)', color: '#fff',
                          borderRadius: 6, padding: '1px 8px', cursor: 'pointer', fontSize: '0.74rem', fontWeight: 600, fontFamily: 'inherit',
                        }}
                      >
                        <Trash2 size={11} />
                        {isDe ? 'Abmelden' : 'Deregister'}
                      </button>
                    )}
                  </li>
                ))}
              </ul>
            </div>
          </div>
        );
};

