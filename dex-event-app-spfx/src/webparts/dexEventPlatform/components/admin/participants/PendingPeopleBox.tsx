/* PendingPeopleBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10452-10523 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Users } from '../../Icons';
import { PersonContactHover } from '../../PersonContactHover';
import { AudiencePerson } from '../../admin/adminTypes';

export interface PendingPeopleBoxProps {
  isDe: boolean;
  openInviteModal: () => void;
  pendingPeople: { people: AudiencePerson[]; reachable: number; };
  setInviteAudienceOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setInviteCustomEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setInviteTarget: React.Dispatch<React.SetStateAction<"organizer" | "audience" | "pending" | "uninvited">>;
  setPendingPeople: React.Dispatch<React.SetStateAction<{ people: AudiencePerson[]; reachable: number; }>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export const PendingPeopleBox: React.FC<PendingPeopleBoxProps> = (p) => {
  const { isDe, openInviteModal, pendingPeople, setInviteAudienceOpen, setInviteCustomEmails, setInviteTarget, setPendingPeople, showAlert } = p;
  return (
          <Modal open={true} onClose={() => setPendingPeople(null)} maxWidth={720} padding={0} ariaLabel={isDe ? 'Wer hat noch nicht geantwortet' : 'Who has not responded yet'}>
            <div style={{ padding: '20px 24px 12px' }}>
              <h3 style={{ margin: '0 0 6px', fontSize: '1.15rem', display: 'flex', alignItems: 'center', gap: 8 }}>
                <Users size={18} />
                {isDe ? 'Noch keine Rückmeldung' : 'No response yet'}
              </h3>
              <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? <>Von <strong>{pendingPeople.reachable}</strong> Personen, die dieses Event sehen können, haben sich <strong>{pendingPeople.people.length}</strong> noch gar nicht geäußert: weder angemeldet noch abgemeldet und auch keine Absage. Das Organizer-Team ist nicht mitgezählt.</>
                  : <>Of <strong>{pendingPeople.reachable}</strong> people who can see this event, <strong>{pendingPeople.people.length}</strong> have not responded at all: neither registered nor cancelled nor declined. The organizer team is not counted.</>}
              </div>
            </div>
            <div style={{ maxHeight: '48vh', overflowY: 'auto', borderTop: '1px solid var(--dex-gray-200)', borderBottom: '1px solid var(--dex-gray-200)' }}>
              {pendingPeople.people.map((p, i) => {
                const nm = (p.displayName || '').trim() || p.email;
                const sub = [p.jobTitle, p.location].filter(Boolean).join(' · ');
                return (
                  <div
                    key={p.email}
                    style={{
                      display: 'flex', alignItems: 'center', gap: 12, padding: '8px 24px',
                      background: i % 2 === 0 ? '#fff' : 'var(--dex-gray-50, #fafafa)',
                    }}
                  >
                    <PersonContactHover email={p.email} name={nm} size={34} subline={sub} isDe={isDe} />
                    <div style={{ minWidth: 0, flex: 1 }}>
                      <div style={{ fontSize: '0.88rem', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{nm}</div>
                      <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-600)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {sub ? `${sub} · ${p.email}` : p.email}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
            <div style={{ padding: '14px 24px 20px', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
              <button type="button" className="btn btn-secondary" style={{ fontSize: '0.85rem' }} onClick={() => setPendingPeople(null)}>
                {isDe ? 'Schließen' : 'Close'}
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                style={{ fontSize: '0.85rem' }}
                onClick={() => {
                  const list = pendingPeople.people.map(p => p.email).join('; ');
                  void navigator.clipboard.writeText(list).then(
                    () => showAlert(isDe ? `${pendingPeople.people.length} Adressen kopiert.` : `${pendingPeople.people.length} addresses copied.`, { variant: 'success' }),
                    () => showAlert(isDe ? 'Kopieren nicht möglich.' : 'Copying failed.', { variant: 'error' })
                  );
                }}
              >
                {isDe ? 'Adressen kopieren' : 'Copy addresses'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                style={{ marginLeft: 'auto', fontSize: '0.85rem' }}
                onClick={() => {
                  const emails = pendingPeople.people.map(p => p.email);
                  setPendingPeople(null);
                  openInviteModal();
                  setInviteTarget('pending');
                  setInviteCustomEmails(emails);
                  setInviteAudienceOpen(true);
                }}
              >
                {isDe ? `Erinnerung schreiben (${pendingPeople.people.length})` : `Write reminder (${pendingPeople.people.length})`}
              </button>
            </div>
          </Modal>
  );
};

