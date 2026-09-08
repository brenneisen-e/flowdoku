/* DangerZoneModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 4380-4509 des Stands
 * vor dem Schnitt). Die Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3 (docs/ui-leitfaden.md): Backdrop, Kopf und Fußzeile macht jetzt
 * `Modal`, die Folgen stehen in einem roten Hinweiskasten, die Tipp-
 * Bestätigung in einem eigenen Feld mit Hilfetext (vorher gab der Dialog
 * keinen Hinweis, WARUM der Knopf grau bleibt). Die Bestätigungstexte und
 * die Bedingung (`matches`, `canDelete`) sind wortgleich geblieben — sie sind
 * das Einzige, was zwischen einem Klick und einem gelöschten Event steht.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { Trash2 } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { useDialog } from '../../../context/DialogContext';
import { useEvents } from '../../../context/EventContext';

export interface DangerZoneModalProps {
  confirmDeleteEvent: DeloitteEvent;
  confirmDeleteText: string;
  deleteEvent: (eventId: string) => Promise<boolean>;
  deletePolicy: { loading: true; } | { loading: false; allowed: boolean; requiresTitle: boolean; externalCount: number; reason?: string; };
  isDe: boolean;
  isDeleting: boolean;
  setConfirmDeleteEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  setConfirmDeleteText: React.Dispatch<React.SetStateAction<string>>;
  setDeletingId: React.Dispatch<React.SetStateAction<string>>;
  setIsDeleting: React.Dispatch<React.SetStateAction<boolean>>;
}

export const DangerZoneModal: React.FC<DangerZoneModalProps> = (p) => {
  const { confirmDeleteEvent, confirmDeleteText, deleteEvent, deletePolicy, isDe, isDeleting, setConfirmDeleteEvent, setConfirmDeleteText, setDeletingId, setIsDeleting } = p;
  // v30.67: deleteEvent bricht seit diesem Release ab, wenn ein Termin nicht
  // gelöscht werden konnte, und liefert false. Ohne Auswertung schloss sich
  // das Modal kommentarlos — der Admin hielt das Event für gelöscht.
  const { showAlert } = useDialog();
  const { getLastEventDeleteError } = useEvents();
  const expected = (confirmDeleteEvent.title || '').trim().toLowerCase();
  const typed = confirmDeleteText.trim().toLowerCase();
  const matches = !!expected && expected === typed;
  const close = (): void => { setConfirmDeleteEvent(null); setConfirmDeleteText(''); };
  // v24.0: Narrowing in Primitive auflösen — sonst verliert TS die
  // Discriminated-Union-Verengung in den verschachtelten JSX-Closures.
  const pol = deletePolicy;
  const polLoaded = pol && pol.loading === false ? pol : null;
  const polLoading = !pol || pol.loading === true;
  const polAllowed = !!polLoaded && polLoaded.allowed;
  const polRequiresTitle = !!polLoaded && polLoaded.requiresTitle;
  const polExternalCount = polLoaded ? polLoaded.externalCount : 0;
  const polReason = polLoaded ? polLoaded.reason : undefined;
  const canDelete = polAllowed && (polRequiresTitle ? matches : true);
  const runDelete = async (): Promise<void> => {
    if (!canDelete || !confirmDeleteEvent) return;
    setIsDeleting(true);
    setDeletingId(confirmDeleteEvent.id);
    try {
      const ok = await deleteEvent(confirmDeleteEvent.id);
      if (!ok) {
        const why = getLastEventDeleteError(isDe ? 'de' : 'en');
        showAlert(why || (isDe ? 'Das Event konnte nicht gelöscht werden.' : 'The event could not be deleted.'), { variant: 'error' });
      }
    } finally {
      setIsDeleting(false);
      setDeletingId(null);
      close();
    }
  };
  return (
    <Modal
      open onClose={close} maxWidth={560} dismissable={!isDeleting}
      ariaLabel={isDe ? 'Event löschen' : 'Delete event'}
      title={isDe ? 'Event löschen' : 'Delete event'}
      subtitle={confirmDeleteEvent.title}
      // v31.3: Rotes Symbol-Feld statt des grünen Standard-Kreises — das
      // Modal trägt die einzige unumkehrbare Aktion des Organizer Centers.
      // Das innere Feld deckt `dex-ui-modal-head-icon` vollflächig ab, weil
      // es dafür (noch) keinen Modifier `--danger` in dexUi.ts gibt.
      icon={<span style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '100%', height: '100%', borderRadius: 12, background: 'var(--dex-red-light, #fce8e6)', color: 'var(--dex-red, #da291c)' }}><Trash2 size={20} /></span>}
      footer={polLoading ? null : !polAllowed ? (
        <button type="button" className="btn btn-secondary" onClick={close}>{isDe ? 'Schließen' : 'Close'}</button>
      ) : (
        <>
          <button type="button" className="btn btn-secondary" onClick={close} disabled={isDeleting}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
          {/* v31.3: Bewusst `btn` OHNE `btn-danger` plus Inline-Rot — wie vor dem
              Umbau. `.dex-modal-overlay .btn-danger` färbt mit `!important` GRAU;
              die einzige unumkehrbare Aktion des Organizer Centers sähe dann aus
              wie „Abbrechen". Nacktes `.btn` setzt im Overlay nur Layout. */}
          <button type="button" className="btn" disabled={!canDelete || isDeleting} onClick={runDelete}
            style={{ background: canDelete && !isDeleting ? 'var(--dex-red, #da291c)' : 'var(--dex-gray-300, #d1d1d1)', color: '#fff' }}>
            <Trash2 size={14} /> {isDeleting ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Endgültig löschen' : 'Delete')}
          </button>
        </>
      )}
    >
      {polLoading ? (
        <p className="dex-ui-muted" style={{ margin: 0, padding: '8px 0' }}>
          {isDe ? 'Prüfe, ob das Event gelöscht werden darf …' : 'Checking whether this event may be deleted …'}
        </p>
      ) : !polAllowed ? (
        <div className="dex-ui-callout dex-ui-callout--danger">
          <span className="dex-ui-callout-icon"><Trash2 size={16} /></span>
          <div>
            <strong>{isDe ? 'Löschen nicht möglich' : 'Deletion not possible'}</strong>
            <p style={{ margin: '6px 0 0' }}>{polReason}</p>
          </div>
        </div>
      ) : (
        <>
          {/* v31.3: Was passiert — Satz und Folgen in EINEM roten Kasten statt
              als Fließtext plus loser Liste. Die Texte selbst sind unverändert. */}
          <div className="dex-ui-callout dex-ui-callout--danger">
            <span className="dex-ui-callout-icon"><Trash2 size={16} /></span>
            <div>
              {isDe
                ? <>Du bist dabei das Event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong> zu löschen.</>
                : <>You are about to delete the event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong>.</>}
              <ul style={{ margin: '6px 0 0', paddingLeft: 18, lineHeight: 1.55 }}>
                <li>{isDe ? 'Subsite (inkl. Teilnehmerliste) und Event-Item wandern in den SharePoint-Papierkorb.' : 'Subsite (incl. attendee list) and event item move to the SharePoint recycle bin.'}</li>
                <li>{isDe ? 'Wiederherstellung durch einen Admin innerhalb von 93 Tagen möglich (zweistufig).' : 'A site collection admin can restore within 93 days (two-stage).'}</li>
                <li>{isDe ? 'Outlook-Termin wird über den Power-Automate-Flow gelöscht.' : 'Outlook calendar event will be deleted via the Power Automate flow.'}</li>
                <li>{isDe ? 'Diese Aktion wird im DEX_ChangeLog mit deinem Namen + Datum protokolliert.' : 'This action is logged in DEX_ChangeLog with your name + date.'}</li>
              </ul>
            </div>
          </div>
          {polRequiresTitle ? (
            <div className="dex-ui-field">
              <p style={{ margin: '0 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
                {isDe
                  ? <>Dieses Event hatte <strong>{polExternalCount}</strong> Anmeldung(en) über das Organizer-Team hinaus und ist älter als ein Jahr. Zur Sicherheit:</>
                  : <>This event had <strong>{polExternalCount}</strong> registration(s) beyond the organizer team and is older than a year. For safety:</>}
              </p>
              <label className="dex-ui-label" htmlFor="dex-danger-confirm">
                {isDe
                  ? <>Tippe zur Bestätigung den Event-Titel <strong>kleingeschrieben</strong> ein:</>
                  : <>Type the event title <strong>in lowercase</strong> to confirm:</>}
              </label>
              <code style={{ display: 'inline-block', padding: '4px 8px', background: 'var(--dex-gray-100, #f5f5f5)', borderRadius: 6, fontSize: '0.85rem', marginBottom: 8, wordBreak: 'break-all' }}>{expected}</code>
              <input
                id="dex-danger-confirm"
                className={cx('dex-ui-input', typed && !matches && 'dex-ui-input--error')}
                value={confirmDeleteText}
                onChange={e => setConfirmDeleteText(e.target.value)}
                placeholder={isDe ? 'Event-Titel kleingeschrieben…' : 'Event title in lowercase…'}
                disabled={isDeleting}
                autoFocus
              />
              {/* v31.3: Sagt, warum der rote Knopf noch grau ist — vorher war das
                  nur zu erraten. */}
              <div className="dex-ui-help">
                {matches
                  ? (isDe ? 'Der Titel stimmt — du kannst jetzt löschen.' : 'The title matches — you can delete now.')
                  : (isDe ? 'Der Knopf unten wird erst aktiv, wenn der Titel genau übereinstimmt.' : 'The button below stays inactive until the title matches exactly.')}
              </div>
            </div>
          ) : (
            <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
              {isDe ? 'Dieses Event hat keine Anmeldungen über das Organizer-Team hinaus. Wirklich löschen?' : 'This event has no registrations beyond the organizer team. Delete it?'}
            </p>
          )}
        </>
      )}
    </Modal>
  );
};
