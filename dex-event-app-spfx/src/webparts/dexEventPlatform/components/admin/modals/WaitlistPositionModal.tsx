/* WaitlistPositionModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16563-16639 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Hash, Info } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';

export interface WaitlistPositionModalProps {
  eventServiceRef: EventService;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setWlPosBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setWlPosModal: React.Dispatch<React.SetStateAction<{ reg: SPRegistration; currentPos: number; total: number; }>>;
  setWlPosValue: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  wlPosBusy: boolean;
  wlPosModal: { reg: SPRegistration; currentPos: number; total: number; };
  wlPosValue: string;
}

export const WaitlistPositionModal: React.FC<WaitlistPositionModalProps> = (p) => {
  const { eventServiceRef, isDe, reloadRegistrations, selectedEvent, setWlPosBusy, setWlPosModal, setWlPosValue, showAlert, wlPosBusy, wlPosModal, wlPosValue } = p;
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
            // v30.67 (Review): '' = Tabelle „Warteliste ohne Gruppe" — der
            // Service trennt das von undefined (= Gesamtliste).
            const group = perGroup
              ? ((reg.PreferredStarterType === 'Durchstarter' || reg.PreferredStarterType === 'Funstarter') ? reg.PreferredStarterType : '')
              : undefined;
            const res = await eventServiceRef.setWaitlistPosition(selectedEvent.subsiteUrl, reg.Id, parsed, group);
            if (!res.ok) {
              showAlert(res.error || (isDe ? 'Der Platz konnte nicht geändert werden.' : 'The position could not be changed.'), { variant: 'error' });
              return;
            }
            // v30.67 (Review): gemeinsamer Nachlade-Pfad statt `[]` bei 429.
            await reloadRegistrations();
            showAlert(res.changed === 0
              ? (isDe ? `${name} steht bereits auf Platz ${res.to}.` : `${name} is already at position ${res.to}.`)
              : (isDe ? `${name} steht jetzt auf Platz ${res.to} (vorher ${res.from}).` : `${name} is now at position ${res.to} (previously ${res.from}).`),
              { variant: 'success' });
            close();
          } finally {
            setWlPosBusy(false);
          }
        };
        // v31.2: Nur für die Anzeige — `valid` und `apply` bleiben die Wahrheit.
        // Eine Eingabe außerhalb 1..total zeigte vorher nur einen grauen Knopf;
        // der Organizer sah nicht, WARUM. Ein leeres Feld ist kein Fehler.
        const outOfRange = wlPosValue.trim() !== '' && !valid;
        const unchanged = valid && parsed === wlPosModal.currentPos;
        return (
          <Modal
            open={true}
            onClose={() => { if (!wlPosBusy) close(); }}
            dismissable={!wlPosBusy}
            maxWidth={520}
            ariaLabel={isDe ? 'Wartelisten-Platz ändern' : 'Change waitlist position'}
            title={isDe ? 'Wartelisten-Platz ändern' : 'Change waitlist position'}
            subtitle={isDe
              ? <><strong>{name}</strong> steht aktuell auf <strong>Platz {wlPosModal.currentPos}</strong> von {wlPosModal.total}.</>
              : <><strong>{name}</strong> is currently at <strong>position {wlPosModal.currentPos}</strong> of {wlPosModal.total}.</>}
            icon={<Hash size={20} />}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={close} disabled={wlPosBusy}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button type="button" className="btn btn-primary" onClick={() => { void apply(); }} disabled={!valid || wlPosBusy}>
                {wlPosBusy ? (isDe ? 'Wird gesetzt…' : 'Applying…') : (isDe ? 'Platz setzen' : 'Set position')}
              </button>
            </>}
          >
            <div className="dex-ui-modal-body">
              <div className="dex-ui-field">
                <label className="dex-ui-label" htmlFor="dex-wlpos-input">
                  {isDe ? 'Auf welchen Platz soll die Person?' : 'Which position should they get?'}
                </label>
                <div className="dex-ui-inline">
                  <input
                    id="dex-wlpos-input"
                    className="dex-ui-input"
                    type="number"
                    min={1}
                    max={wlPosModal.total}
                    value={wlPosValue}
                    onChange={e => setWlPosValue(e.target.value)}
                    disabled={wlPosBusy}
                    aria-invalid={outOfRange || undefined}
                    style={{ width: 110, borderColor: outOfRange ? 'var(--dex-red, #da291c)' : undefined }}
                  />
                  <button type="button" className="dex-ui-textbtn" disabled={wlPosBusy} onClick={() => setWlPosValue('1')}>
                    {isDe ? 'Ganz nach oben' : 'To the top'}
                  </button>
                  {unchanged && <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'schon dort' : 'already there'}</span>}
                </div>
                <div className="dex-ui-help" style={outOfRange ? { color: 'var(--dex-red, #da291c)' } : undefined}>
                  {outOfRange
                    ? (isDe ? `Bitte eine Zahl zwischen 1 und ${wlPosModal.total}.` : `Enter a number between 1 and ${wlPosModal.total}.`)
                    : (isDe ? `1 bis ${wlPosModal.total} — Platz 1 rückt als Nächstes nach.` : `1 to ${wlPosModal.total} — position 1 is promoted next.`)}
                </div>
              </div>
              <div className="dex-ui-callout dex-ui-callout--neutral">
                <span className="dex-ui-callout-icon" aria-hidden="true"><Info size={16} /></span>
                <div>
                  {isDe
                    ? <>Die anderen Wartenden rücken auf oder nach; die Teilnehmer-Nummern <strong>innerhalb der Warteliste</strong> werden neu vergeben. Es geht <strong>keine Mail</strong> raus, niemand wird dadurch angemeldet — die Person rückt nur früher nach, sobald ein Platz frei wird. Angemeldete Teilnehmer sind nicht betroffen.</>
                    : <>The other waitlisted people shift up or down; attendee numbers <strong>within the waitlist</strong> are reassigned. <strong>No email</strong> is sent and nobody gets registered by this — the person is simply promoted earlier once a seat frees up. Registered attendees are unaffected.</>}
                </div>
              </div>
            </div>
          </Modal>
        );
};

