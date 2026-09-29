/* MailTypeModal — v32.26. Erste Frage vor jeder Mail aus dem Organizer
 * Center: WAS für eine Mail? Nutzer-Ansage 29.09.2026: „davor kommt noch die
 * Frage, was für eine Mail — Einladungsmail, Remindermail oder Info /
 * Sonstiges. Danach die Auswahl wer und dann die Mail."
 *
 * Die drei Wege sind die bestehenden: Einladung = Einladungs-Editor,
 * Reminder = Empfängerwahl mit vorgewählter „Erinnerung", Info = Empfänger-
 * wahl wie bisher (aktive Teilnehmer). Kein eigener Versandweg. */
import * as React from 'react';
import Modal from '../../Modal';
import { Mail, Send, Info } from '../../Icons';

export interface MailTypeModalProps {
  open: boolean;
  isDe: boolean;
  onClose: () => void;
  onInvite: () => void;
  onReminder: () => void;
  onInfo: () => void;
}

export const MailTypeModal: React.FC<MailTypeModalProps> = ({ open, isDe, onClose, onInvite, onReminder, onInfo }) => {
  if (!open) return null;
  const kachel = (icon: React.ReactNode, title: string, desc: string, go: () => void): React.ReactElement => (
    <button type="button" className="dex-ui-choice" style={{ width: '100%', textAlign: 'left' }}
      onClick={() => { onClose(); go(); }}>
      <span className="dex-ui-choice-icon">{icon}</span>
      <span className="dex-ui-choice-body">
        <span className="dex-ui-choice-title" style={{ display: 'block' }}>{title}</span>
        <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{desc}</span>
      </span>
    </button>
  );
  return (
    <Modal open={true} onClose={onClose} maxWidth={560}
      ariaLabel={isDe ? 'Was für eine Mail?' : 'What kind of email?'}
      icon={<Mail size={20} />}
      title={isDe ? 'Was für eine Mail möchtest du schicken?' : 'What kind of email do you want to send?'}
      subtitle={isDe ? 'Danach wählst du, wer sie bekommt, und schreibst den Text.' : 'Next you choose who gets it, then write the text.'}
      footer={<button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>}
    >
      <div className="dex-ui-stack" style={{ gap: 10 }}>
        {kachel(<Send size={18} />,
          isDe ? 'Einladungsmail' : 'Invitation email',
          isDe ? 'Mit Anmelde-Link — an dich zum Weiterleiten, an den Mailverteiler oder an einzelne Personen.' : 'With the registration link — to yourself for forwarding, to the distribution list or to individual people.',
          onInvite)}
        {kachel(<Mail size={18} />,
          isDe ? 'Reminder' : 'Reminder',
          isDe ? 'Erinnerung an alle, die das Event sehen, aber noch nicht geantwortet haben.' : 'A reminder to everyone who can see the event but has not responded yet.',
          onReminder)}
        {kachel(<Info size={18} />,
          isDe ? 'Info / Sonstiges' : 'Info / other',
          isDe ? 'Nachricht an Teilnehmer — z. B. Anreise, Programm, Änderungen. Auch mit Umfrage.' : 'A message to attendees — e.g. travel, programme, changes. Polls possible too.',
          onInfo)}
      </div>
    </Modal>
  );
};

export default MailTypeModal;
