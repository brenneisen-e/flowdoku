/* PendingPeopleBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10452-10523 des
 * Stands vor dem Schnitt). Wer hier steht, entscheidet der Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md (Abschnitt „Modale") umgebaut — Kopf und
 * Fuss kommen jetzt aus `Modal` (title/subtitle/icon/footer) statt aus einem
 * eigenen <h3> mit `padding: 0`; die Personen stehen als `dex-ui-person`-Zelle
 * (Foto, Name, darunter Position und Adresse) statt in Zebrastreifen, und der
 * Primärknopf sagt, was danach passiert. Die drei Aktionen sind unverändert.
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
          <Modal
            open={true}
            onClose={() => setPendingPeople(null)}
            maxWidth={720}
            ariaLabel={isDe ? 'Wer hat noch nicht geantwortet' : 'Who has not responded yet'}
            title={isDe ? 'Noch keine Rückmeldung' : 'No response yet'}
            icon={<Users size={20} />}
            subtitle={isDe
              ? <>Von <strong>{pendingPeople.reachable}</strong> Personen, die dieses Event sehen können, haben sich <strong>{pendingPeople.people.length}</strong> noch gar nicht geäußert: weder angemeldet noch abgemeldet und auch keine Absage. Das Organizer-Team ist nicht mitgezählt.</>
              : <>Of <strong>{pendingPeople.reachable}</strong> people who can see this event, <strong>{pendingPeople.people.length}</strong> have not responded at all: neither registered nor cancelled nor declined. The organizer team is not counted.</>}
            footer={<>
              {/* v31.3: Die beiden Nebenaktionen links, der Primärknopf rechts
                  aussen (Leitfaden „Modale") — vorher schob ein margin-left:auto
                  den Knopf weg und die Fusszeile hatte keine Trennlinie. */}
              <span className="dex-ui-modal-foot-left">
                <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setPendingPeople(null)}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
                <button
                  type="button"
                  className="btn btn-secondary dex-ui-btn-sm"
                  onClick={() => {
                    const list = pendingPeople.people.map(x => x.email).join('; ');
                    void navigator.clipboard.writeText(list).then(
                      () => showAlert(isDe ? `${pendingPeople.people.length} Adressen kopiert.` : `${pendingPeople.people.length} addresses copied.`, { variant: 'success' }),
                      () => showAlert(isDe ? 'Kopieren nicht möglich.' : 'Copying failed.', { variant: 'error' })
                    );
                  }}
                >
                  {isDe ? 'Adressen kopieren' : 'Copy addresses'}
                </button>
              </span>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  const emails = pendingPeople.people.map(x => x.email);
                  setPendingPeople(null);
                  openInviteModal();
                  setInviteTarget('pending');
                  setInviteCustomEmails(emails);
                  setInviteAudienceOpen(true);
                }}
              >
                {isDe ? `Erinnerung an alle ${pendingPeople.people.length} schreiben` : `Write reminder to all ${pendingPeople.people.length}`}
              </button>
            </>}
          >
            {/* v31.3: Personen-Zelle statt Zebrastreifen (Leitfaden 5b). Die
                Zeile ist bewusst OHNE Hover — klickbar ist nur das Foto, das
                die Kontaktkarte öffnet. */}
            <div style={{ maxHeight: '48vh', overflowY: 'auto' }}>
              {pendingPeople.people.map(person => {
                const nm = (person.displayName || '').trim() || person.email;
                const sub = [person.jobTitle, person.location].filter(Boolean).join(' · ');
                return (
                  <div key={person.email} className="dex-ui-person dex-ui-row--bordered" style={{ display: 'flex', padding: '8px 0' }}>
                    <PersonContactHover email={person.email} name={nm} size={34} subline={sub} isDe={isDe} />
                    <div style={{ minWidth: 0, flex: 1 }}>
                      <div className="dex-ui-person-name">{nm}</div>
                      <div className="dex-ui-person-sub">{sub ? `${sub} · ${person.email}` : person.email}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          </Modal>
  );
};

