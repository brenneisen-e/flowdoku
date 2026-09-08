/* DupCancelModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16367-16434 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 * v31.2: Aus drei gestapelten Knöpfen mit einem Absatz Erklärung wurden zwei
 * Optionen mit je einer Zeile Folge und dem Knopf daneben — der Organizer
 * sieht, was jede Wahl auslöst, bevor er klickt. Handler unverändert.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { SPRegistration } from '../../../services/EventService';
import { AlertCircle } from '../../Icons';

export interface DupCancelModalProps {
  dupCancelBusy: boolean;
  dupCancelReg: SPRegistration;
  isDe: boolean;
  performSilentDuplicateDelete: (reg: SPRegistration) => Promise<boolean>;
  performStandardCancel: (reg: SPRegistration) => Promise<void>;
  setDupCancelBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setDupCancelReg: React.Dispatch<React.SetStateAction<SPRegistration>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export const DupCancelModal: React.FC<DupCancelModalProps> = (p) => {
  const { dupCancelBusy, dupCancelReg, isDe, performSilentDuplicateDelete, performStandardCancel, setDupCancelBusy, setDupCancelReg, showAlert } = p;
  const reg = dupCancelReg;
  const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
  const teamLabel = reg.TeamName ? `„${reg.TeamName}"` : (reg.TeamId ? (isDe ? 'Team ohne Namen' : 'unnamed team') : (isDe ? 'Einzel-Anmeldung' : 'individual registration'));
  const close = (): void => setDupCancelReg(null);
  const optionStyle: React.CSSProperties = { display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' };
  return (
    <Modal
      open={true}
      onClose={() => { if (!dupCancelBusy) close(); }}
      dismissable={!dupCancelBusy}
      maxWidth={580}
      ariaLabel={isDe ? 'Doppelte Anmeldung entfernen' : 'Remove duplicate registration'}
      title={isDe ? 'Doppelte Anmeldung entfernen' : 'Remove duplicate registration'}
      subtitle={isDe
        ? <><strong>{name}</strong> ({reg.ParticipantEmail}) ist mehrfach für dieses Event angemeldet.</>
        : <><strong>{name}</strong> ({reg.ParticipantEmail}) is registered more than once for this event.</>}
      icon={<AlertCircle size={20} />}
      footer={<button type="button" className="btn btn-secondary" disabled={dupCancelBusy} onClick={close}>{isDe ? 'Abbrechen' : 'Cancel'}</button>}
    >
      <div className="dex-ui-callout dex-ui-callout--warn">
        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
        <span>{isDe ? <>Du entfernst die Zeile <strong>{teamLabel}</strong>.</> : <>You are removing the row <strong>{teamLabel}</strong>.</>}</span>
      </div>
      <div className="dex-ui-stack">
        <div className="dex-ui-card dex-ui-card--hover dex-ui-card--accent" style={optionStyle}>
          <div className="dex-ui-choice-body">
            <div className="dex-ui-choice-title dex-ui-inline">
              {isDe ? 'Still entfernen' : 'Remove silently'}
              <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Empfohlen bei Dubletten' : 'Recommended for duplicates'}</span>
            </div>
            <div className="dex-ui-choice-desc">
              {isDe
                ? 'Die Zeile wird gelöscht — ohne Abmelde-Mail, ohne Outlook-Absage, und niemand rückt von der Warteliste nach. Die Person bleibt über ihre andere Anmeldung regulär dabei.'
                : 'The row is deleted — no cancellation email, no Outlook removal, and nobody is promoted from the waitlist. The person stays registered via their other entry.'}
            </div>
          </div>
          <button
            type="button"
            className="btn btn-primary"
            disabled={dupCancelBusy}
            onClick={async () => {
              setDupCancelBusy(true);
              // v30.67: Rückgabewert auswerten — `deleteRegistration` meldet
              // Misserfolg über false, nicht über eine Ausnahme.
              const okDel = await performSilentDuplicateDelete(reg);
              setDupCancelBusy(false);
              setDupCancelReg(null);
              showAlert(okDel
                ? (isDe ? 'Doppelte Anmeldung still entfernt.' : 'Duplicate registration silently removed.')
                : (isDe ? 'Die Zeile konnte NICHT entfernt werden (fehlende Rechte oder Drosselung) — sie ist noch da.' : 'The row could NOT be removed (missing permissions or throttling) — it is still there.'),
                { variant: okDel ? 'success' : 'error' });
            }}
          >
            {dupCancelBusy ? (isDe ? 'Wird entfernt…' : 'Removing…') : (isDe ? 'Still entfernen' : 'Remove silently')}
          </button>
        </div>
        <div className="dex-ui-card dex-ui-card--hover" style={optionStyle}>
          <div className="dex-ui-choice-body">
            <div className="dex-ui-choice-title">{isDe ? 'Normal abmelden' : 'Cancel normally'}</div>
            <div className="dex-ui-choice-desc">
              {isDe
                ? 'Nur wenn es KEINE Dublette ist, sondern eine echte Abmeldung: Abmelde-Mail und Outlook-Absage gehen raus, die Warteliste rückt nach.'
                : 'Only if this is NOT a duplicate but a real cancellation: the cancellation email and Outlook removal go out, and the waitlist is promoted.'}
            </div>
          </div>
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
            {isDe ? 'Normal abmelden' : 'Cancel normally'}
          </button>
        </div>
      </div>
    </Modal>
  );
};
