/* DupCancelModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16367-16434 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface DupCancelModalProps {
  dupCancelBusy: boolean;
  dupCancelReg: SPRegistration;
  isDe: boolean;
  performSilentDuplicateDelete: (reg: SPRegistration) => Promise<void>;
  performStandardCancel: (reg: SPRegistration) => Promise<void>;
  selectedEvent: DeloitteEvent;
  setDupCancelBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setDupCancelReg: React.Dispatch<React.SetStateAction<SPRegistration>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export const DupCancelModal: React.FC<DupCancelModalProps> = (p) => {
  const { dupCancelBusy, dupCancelReg, isDe, performSilentDuplicateDelete, performStandardCancel, selectedEvent, setDupCancelBusy, setDupCancelReg, showAlert } = p;
        const reg = dupCancelReg;
        const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
        const teamLabel = reg.TeamName ? `„${reg.TeamName}"` : (reg.TeamId ? (isDe ? 'Team ohne Namen' : 'unnamed team') : (isDe ? 'Einzel-Anmeldung' : 'individual registration'));
        return (
          <Modal
            open={true}
            onClose={() => { if (!dupCancelBusy) setDupCancelReg(null); }}
            dismissable={!dupCancelBusy}
            maxWidth={540}
            padding={24}
            ariaLabel={isDe ? 'Doppel-Anmeldung entfernen' : 'Remove duplicate registration'}
          >
            <h3 style={{ marginTop: 0, marginBottom: 8, color: 'var(--dex-red, #c00)' }}>
              {isDe ? 'Doppel-Anmeldung entfernen' : 'Remove duplicate registration'}
            </h3>
            <p style={{ marginTop: 0, fontSize: '0.88rem', lineHeight: 1.5 }}>
              {isDe
                ? <><strong>{name}</strong> ({reg.ParticipantEmail}) ist mehrfach für dieses Event angemeldet. Du entfernst gerade die Zeile <strong>{teamLabel}</strong>.</>
                : <><strong>{name}</strong> ({reg.ParticipantEmail}) is registered more than once for this event. You are removing the row <strong>{teamLabel}</strong>.</>}
            </p>
            <p style={{ fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? 'Bei einer Dublette ist die stille Entfernung richtig: Die Zeile wird gelöscht, ohne dass eine Abmelde-Mail oder eine Outlook-Absage rausgeht und ohne dass jemand von der Warteliste nachrückt — die Person bleibt über ihre andere Anmeldung regulär dabei. Nur falls es KEINE Dublette ist, sondern eine echte Abmeldung, wähle „Normal abmelden".'
                : 'For a duplicate, silent removal is the right choice: the row is deleted without a cancellation email or Outlook removal and without promoting anyone from the waitlist — the person stays registered via their other entry. Only if this is NOT a duplicate but a real cancellation, choose „Cancel normally".'}
            </p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 18 }}>
              <button
                type="button"
                className="btn btn-primary"
                style={{ background: 'var(--dex-red, #c00)', borderColor: 'var(--dex-red, #c00)' }}
                disabled={dupCancelBusy}
                onClick={async () => {
                  setDupCancelBusy(true);
                  await performSilentDuplicateDelete(reg);
                  setDupCancelBusy(false);
                  setDupCancelReg(null);
                  showAlert(isDe ? 'Doppelte Anmeldung still entfernt.' : 'Duplicate registration silently removed.', { variant: 'success' });
                }}
              >
                {dupCancelBusy ? (isDe ? 'Wird entfernt…' : 'Removing…') : (isDe ? 'Duplikat still entfernen (empfohlen)' : 'Silently remove duplicate (recommended)')}
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                disabled={dupCancelBusy}
                onClick={async () => {
                  setDupCancelBusy(true);
                  await performStandardCancel(reg);
                  setDupCancelBusy(false);
                  setDupCancelReg(null);
                }}
              >
                {isDe ? 'Normal abmelden (mit Mail & Nachrücken)' : 'Cancel normally (with email & promotion)'}
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                style={{ border: 'none' }}
                disabled={dupCancelBusy}
                onClick={() => setDupCancelReg(null)}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            </div>
          </Modal>
        );
};

