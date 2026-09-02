/* AssignAssistModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14432-14487 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Users } from '../../Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import { UserFieldPicker } from '../../UserFieldPicker';
import { ConsolidatedRow } from '../../admin/adminTypes';

export interface AssignAssistModalProps {
  assignAssistBusy: boolean;
  assignAssistRow: ConsolidatedRow;
  assignAssistValue: string;
  isDe: boolean;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  setAssignAssistRow: React.Dispatch<React.SetStateAction<ConsolidatedRow>>;
  setAssignAssistValue: React.Dispatch<React.SetStateAction<string>>;
  submitAssignAssistant: () => Promise<void>;
}

export const AssignAssistModal: React.FC<AssignAssistModalProps> = (p) => {
  const { assignAssistBusy, assignAssistRow, assignAssistValue, isDe, searchUser, searchUsers, setAssignAssistRow, setAssignAssistValue, submitAssignAssistant } = p;
  return (
          <Modal
            open={true}
            onClose={() => { if (!assignAssistBusy) { setAssignAssistRow(null); setAssignAssistValue(''); } }}
            maxWidth={520}
            dismissable={!assignAssistBusy}
            ariaLabel={isDe ? 'Assistenz zuordnen' : 'Assign assistant'}
          >
            <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Users size={18} /> {isDe ? 'Assistenz zuordnen' : 'Assign assistant'}
            </h3>
            <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', marginTop: 10 }}>
              {isDe
                ? <>Person: <strong>{`${assignAssistRow.vorname || ''} ${assignAssistRow.nachname || ''}`.trim() || assignAssistRow.email}</strong> · {assignAssistRow.email}</>
                : <>Person: <strong>{`${assignAssistRow.vorname || ''} ${assignAssistRow.nachname || ''}`.trim() || assignAssistRow.email}</strong> · {assignAssistRow.email}</>}
            </p>
            <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start', background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 8, padding: '10px 12px', margin: '4px 0 14px' }}>
              <Icon iconName="Info" style={{ fontSize: 16, color: 'var(--dex-orange, #ed8b00)', marginTop: 1, flexShrink: 0 }} />
              <div style={{ fontSize: '0.8rem', color: '#5a4a20', lineHeight: 1.5 }}>
                {isDe
                  ? <>Die gewählte Assistenz übernimmt die Verwaltung dieser Anmeldung — die <strong>Klammer-Anmeldung UND alle aktiven Sub-Event-Anmeldungen</strong> werden ihr zugeschrieben. Danach erscheint die Person in der <strong>„Assistenz“-Kachel</strong> der Assistenz und kann dort Angaben anpassen, ab- und anmelden. Es werden <strong>keine Mails</strong> versendet.</>
                  : <>The selected assistant takes over managing this registration — the <strong>umbrella registration AND all active sub-event registrations</strong> are attributed to them. The person then appears in the assistant&apos;s <strong>„Assistant“ tile</strong>. <strong>No emails</strong> are sent.</>}
              </div>
            </div>
            <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginBottom: 6 }}>
              {isDe ? 'Assistenz (Person suchen)' : 'Assistant (search person)'}
            </label>
            <UserFieldPicker
              value={assignAssistValue}
              onChange={setAssignAssistValue}
              searchUsers={searchUsers}
              searchUserByEmail={searchUser}
              placeholder={isDe ? 'Name oder E-Mail der Assistenz…' : 'Assistant name or email…'}
              errorStyle={{}}
              forcedIsDe={isDe}
            />
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20 }}>
              <button
                type="button"
                className="btn btn-secondary"
                disabled={assignAssistBusy}
                onClick={() => { setAssignAssistRow(null); setAssignAssistValue(''); }}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                disabled={assignAssistBusy || !assignAssistValue}
                onClick={() => { void submitAssignAssistant(); }}
              >
                {assignAssistBusy ? '…' : (isDe ? 'Zuordnen' : 'Assign')}
              </button>
            </div>
          </Modal>
  );
};

