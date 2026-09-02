/* WaitlistPositionModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16563-16639 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';

export interface WaitlistPositionModalProps {
  eventServiceRef: EventService;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
  setWlPosBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setWlPosModal: React.Dispatch<React.SetStateAction<{ reg: SPRegistration; currentPos: number; total: number; }>>;
  setWlPosValue: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  wlPosBusy: boolean;
  wlPosModal: { reg: SPRegistration; currentPos: number; total: number; };
  wlPosValue: string;
}

export const WaitlistPositionModal: React.FC<WaitlistPositionModalProps> = (p) => {
  const { eventServiceRef, getAllRegistrations, isDe, selectedEvent, setRegistrations, setWlPosBusy, setWlPosModal, setWlPosValue, showAlert, wlPosBusy, wlPosModal, wlPosValue } = p;
        const reg = wlPosModal.reg;
        const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail || '');
        const parsed = parseInt(wlPosValue, 10);
        const valid = !isNaN(parsed) && parsed >= 1 && parsed <= wlPosModal.total;
        const close = (): void => { setWlPosModal(null); setWlPosBusy(false); };
        const apply = async (): Promise<void> => {
          if (!eventServiceRef || !selectedEvent.subsiteUrl || !valid) return;
          setWlPosBusy(true);
          try {
            // v30.67: Gruppe mitgeben — bei geteilten Kapazitaeten sortierte
            // "Platz aendern" sonst ueber die GESAMTE Warteliste und schob die
            // Person in die falsche Gruppen-Reihenfolge. Aber NUR bei getrennten
            // Wartelisten: bei splitSharedWaitlist ist es ein Topf (CLAUDE.md),
            // dort waere eine Gruppen-Sortierung genauso falsch.
            const perGroup = !!(selectedEvent.durchstarterCapacity || selectedEvent.funstarterCapacity) && !selectedEvent.splitSharedWaitlist;
            const group = perGroup ? (reg.PreferredStarterType || undefined) : undefined;
            const res = await eventServiceRef.setWaitlistPosition(selectedEvent.subsiteUrl, reg.Id, parsed, group);
            if (!res.ok) {
              showAlert(res.error || (isDe ? 'Der Platz konnte nicht geändert werden.' : 'The position could not be changed.'), { variant: 'error' });
              return;
            }
            const allRegs = await getAllRegistrations(selectedEvent.id);
            setRegistrations(allRegs);
            showAlert(res.changed === 0
              ? (isDe ? `${name} steht bereits auf Platz ${res.to}.` : `${name} is already at position ${res.to}.`)
              : (isDe ? `${name} steht jetzt auf Platz ${res.to} (vorher ${res.from}).` : `${name} is now at position ${res.to} (previously ${res.from}).`),
              { variant: 'success' });
            close();
          } finally {
            setWlPosBusy(false);
          }
        };
        return (
          <Modal
            open={true}
            onClose={() => { if (!wlPosBusy) close(); }}
            dismissable={!wlPosBusy}
            maxWidth={520}
            padding={24}
            ariaLabel={isDe ? 'Wartelisten-Platz ändern' : 'Change waitlist position'}
          >
            <div>
              <h3 style={{ marginTop: 0 }}>{isDe ? 'Wartelisten-Platz ändern' : 'Change waitlist position'}</h3>
              <p style={{ fontSize: '0.88rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
                {isDe
                  ? <><strong>{name}</strong> steht aktuell auf <strong>Platz {wlPosModal.currentPos}</strong> von {wlPosModal.total}. Auf welchen Platz soll die Person?</>
                  : <><strong>{name}</strong> is currently at <strong>position {wlPosModal.currentPos}</strong> of {wlPosModal.total}. Which position should they get?</>}
              </p>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center', margin: '12px 0' }}>
                <input
                  className="form-input"
                  type="number"
                  min={1}
                  max={wlPosModal.total}
                  value={wlPosValue}
                  onChange={e => setWlPosValue(e.target.value)}
                  disabled={wlPosBusy}
                  style={{ width: 110 }}
                />
                <button type="button" className="btn btn-secondary" style={{ fontSize: '0.8rem', padding: '6px 12px' }} disabled={wlPosBusy} onClick={() => setWlPosValue('1')}>
                  {isDe ? 'Ganz nach oben' : 'To the top'}
                </button>
              </div>
              <div style={{
                padding: '8px 10px', borderRadius: 6, fontSize: '0.78rem', lineHeight: 1.5,
                background: 'var(--dex-gray-50, #f8f9fa)', border: '1px solid var(--dex-gray-200)', color: 'var(--dex-gray-700)',
              }}>
                {isDe
                  ? <>Die anderen Wartenden rücken entsprechend auf oder nach. Angemeldete Teilnehmer sind nicht betroffen, es geht <strong>keine Mail</strong> raus und niemand wird dadurch angemeldet — die Person rückt nur früher nach, sobald ein Platz frei wird. Die Teilnehmer-Nummern <strong>innerhalb der Warteliste</strong> werden dabei neu vergeben.</>
                  : <>The other waitlisted people shift accordingly. Registered attendees are unaffected, <strong>no email</strong> is sent and nobody gets registered by this — the person is simply promoted earlier once a seat frees up. Attendee numbers <strong>within the waitlist</strong> are reassigned.</>}
              </div>
              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 18 }}>
                <button type="button" className="btn btn-secondary" onClick={close} disabled={wlPosBusy}>
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button type="button" className="btn btn-primary" onClick={() => { void apply(); }} disabled={!valid || wlPosBusy}>
                  {wlPosBusy ? (isDe ? 'Wird gesetzt…' : 'Applying…') : (isDe ? 'Platz setzen' : 'Set position')}
                </button>
              </div>
            </div>
          </Modal>
        );
};

