/* MassmailPickModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14939-15018 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { MassmailAudience, AudiencePerson } from '../adminTypes';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { Users } from '../../Icons';
import { useLocaleSafe } from '../../../context/LanguageContext';
import { SPRegistration } from '../../../services/EventService';
import { DeloitteEvent } from '../../../types';
import { MassmailZielChips, MassmailExtra, massmailEmpfaenger, massmailZielLabel } from './MassmailZielChips';

export interface MassmailPickModalProps {
  massmailAudience: MassmailAudience;
  massmailStatuses: Set<string>;
  registrations: SPRegistration[];
  setMassmailAudience: React.Dispatch<React.SetStateAction<MassmailAudience>>;
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setMassmailStatuses: React.Dispatch<React.SetStateAction<Set<string>>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
  /** v31.72: Wer das Event sieht, aber noch nicht geantwortet hat (s. AdminPage).
   *  undefined = wird gerechnet, null = keine Liste (Standort-Sichtbarkeit). */
  massmailOffene?: AudiencePerson[] | null;
  /** v32.30: additive Zusätze (Offene, ich, Organizer, Test-Team). */
  massmailExtras: Set<MassmailExtra>;
  setMassmailExtras: React.Dispatch<React.SetStateAction<Set<MassmailExtra>>>;
  selectedEvent: DeloitteEvent;
  myEmail: string;
}

export const MassmailPickModal: React.FC<MassmailPickModalProps> = (p) => {
  const { massmailAudience, massmailStatuses, registrations, setMassmailAudience, setMassmailMode, setMassmailPasteRaw, setMassmailStatuses, setShowEmailModal, massmailOffene, massmailExtras, setMassmailExtras, selectedEvent, myEmail } = p;
  // v31.2: Die Props kennen kein isDe (Schnittstelle bleibt) — die Sprache
  // kommt wie in Modal.tsx aus dem Kontext.
  const isDe = useLocaleSafe() === 'de';
  const closeAll = (): void => { setMassmailMode('closed'); setMassmailPasteRaw(''); };
  // v32.30: Mehrfachauswahl über Chips (MassmailZielChips); die beiden
  // Einfüge-Modi (Nachrücker, Erinnerung ohne bekannte Offene) bleiben
  // eigene Zeilen, weil nur sie einen Zwischenschritt haben.
  const einfuegen = massmailAudience === 'nachruecker' || massmailAudience === 'reminder';
  const empfaenger = einfuegen ? [] : massmailEmpfaenger({
    audience: massmailAudience, statuses: massmailStatuses, extras: massmailExtras, registrations,
    offene: massmailOffene, pasteRaw: '', ev: selectedEvent, myEmail,
  });
  const proceed = (): void => {
    if (einfuegen) { setMassmailMode('paste'); return; }
    if (empfaenger.length === 0) return;
    setShowEmailModal(true); setMassmailMode('editor');
  };
  return (
    <Modal open={true} onClose={closeAll} maxWidth={860}
      ariaLabel={isDe ? 'Empfänger wählen' : 'Choose recipients'}
      title={isDe ? 'An wen soll die Mail gehen?' : 'Who should get the mail?'}
      subtitle={isDe ? 'Wähle eine oder mehrere Gruppen — den Text schreibst du danach im Mail-Editor.' : 'Pick one or more groups — you write the text in the mail editor afterwards.'}
      icon={<Users size={20} />}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={closeAll}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={proceed} disabled={!einfuegen && empfaenger.length === 0}>
          {einfuegen ? (isDe ? 'Weiter: Verteiler einfügen' : 'Next: paste list') : (isDe ? `Weiter zum Mail-Editor (${empfaenger.length})` : `Continue to mail editor (${empfaenger.length})`)}
        </button>
      </>}>
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{isDe ? 'Gruppen' : 'Groups'}</div>
        <MassmailZielChips
          isDe={isDe} ev={selectedEvent} myEmail={myEmail} registrations={registrations}
          audience={massmailAudience} setAudience={setMassmailAudience}
          statuses={massmailStatuses} setStatuses={setMassmailStatuses}
          extras={massmailExtras} setExtras={setMassmailExtras} offene={massmailOffene}
        />
        {!einfuegen && (
          <div className="dex-ui-inline" style={{ gap: 6 }}>
            <span className={cx('dex-ui-pill', empfaenger.length > 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--red')}><Users size={12} /> {empfaenger.length} {isDe ? 'Empfänger' : 'recipients'}</span>
            <span className="dex-ui-pill dex-ui-pill--gray dex-ui-pill--wrap">{massmailZielLabel(isDe, massmailAudience, massmailStatuses, massmailExtras)}</span>
            {empfaenger.length === 0 && <span className="dex-ui-help" style={{ margin: 0 }}>{isDe ? 'Wähle mindestens eine Gruppe mit Personen.' : 'Pick at least one group with people in it.'}</span>}
          </div>
        )}
      </div>
    </Modal>
  );
};
