/* DangerZoneModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 4380-4509 des Stands
 * vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { Trash2, X } from '../../Icons';
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
  // geloescht werden konnte, und liefert false. Ohne Auswertung schloss sich
  // das Modal kommentarlos — der Admin hielt das Event fuer geloescht.
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
    return (
      <div
        style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
          display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
        }}
        onClick={() => { if (!isDeleting) close(); }}
      >
        <div
          className="card"
          style={{ width: '100%', maxWidth: 560, padding: 24, borderRadius: 16, background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.3)', borderTop: '4px solid var(--dex-red, #c00)' }}
          onClick={e => e.stopPropagation()}
        >
          <div className="flex-between mb-16">
            <h3 style={{ margin: 0, color: 'var(--dex-red, #c00)', display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Trash2 size={20} /> {isDe ? 'Danger Zone — Event löschen' : 'Danger Zone — Delete event'}
            </h3>
            <button
              onClick={close}
              disabled={isDeleting}
              style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: isDeleting ? 'not-allowed' : 'pointer', color: 'var(--dex-gray-500)' }}
            ><X size={20} /></button>
          </div>
          {polLoading ? (
            <div style={{ padding: '16px 0', fontSize: '0.88rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Prüfe, ob das Event gelöscht werden darf …' : 'Checking whether this event may be deleted …'}
            </div>
          ) : !polAllowed ? (
            <>
              <div style={{ background: 'rgba(218,41,28,0.06)', border: '1px solid var(--dex-red, #c00)', padding: 14, borderRadius: 8, marginBottom: 16, fontSize: '0.85rem', lineHeight: 1.55 }}>
                <strong>{isDe ? 'Löschen nicht möglich' : 'Deletion not possible'}</strong>
                <p style={{ margin: '6px 0 0' }}>{polReason}</p>
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                <button type="button" className="btn btn-outline" onClick={close}>{isDe ? 'Schließen' : 'Close'}</button>
              </div>
            </>
          ) : (
            <>
              <p style={{ margin: '0 0 12px', fontSize: '0.88rem', lineHeight: 1.55 }}>
                {isDe
                  ? <>Du bist dabei das Event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong> zu löschen.</>
                  : <>You are about to delete the event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong>.</>}
              </p>
              <ul style={{ margin: '0 0 16px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.55, paddingLeft: 18 }}>
                <li>{isDe ? 'Subsite (inkl. Teilnehmerliste) und Event-Item wandern in den SharePoint-Papierkorb.' : 'Subsite (incl. attendee list) and event item move to the SharePoint recycle bin.'}</li>
                <li>{isDe ? 'Wiederherstellung durch einen Admin innerhalb von 93 Tagen möglich (zweistufig).' : 'A site collection admin can restore within 93 days (two-stage).'}</li>
                <li>{isDe ? 'Outlook-Termin wird über den Power-Automate-Flow gelöscht.' : 'Outlook calendar event will be deleted via the Power Automate flow.'}</li>
                <li>{isDe ? 'Diese Aktion wird im DEX_ChangeLog mit deinem Namen + Datum protokolliert.' : 'This action is logged in DEX_ChangeLog with your name + date.'}</li>
              </ul>
              {polRequiresTitle ? (
                <div style={{ background: 'rgba(218,41,28,0.06)', border: '1px solid var(--dex-red, #c00)', padding: 12, borderRadius: 8, marginBottom: 16 }}>
                  <p style={{ margin: '0 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
                    {isDe
                      ? <>Dieses Event hatte <strong>{polExternalCount}</strong> Anmeldung(en) über das Organizer-Team hinaus und ist älter als ein Jahr. Zur Sicherheit:</>
                      : <>This event had <strong>{polExternalCount}</strong> registration(s) beyond the organizer team and is older than a year. For safety:</>}
                  </p>
                  <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, marginBottom: 6 }}>
                    {isDe
                      ? <>Tippe zur Bestätigung den Event-Titel <strong>kleingeschrieben</strong> ein:</>
                      : <>Type the event title <strong>in lowercase</strong> to confirm:</>}
                  </label>
                  <code style={{ display: 'inline-block', padding: '4px 8px', background: '#fff', borderRadius: 4, fontSize: '0.85rem', marginBottom: 8, wordBreak: 'break-all' }}>{expected}</code>
                  <input
                    className="form-input"
                    value={confirmDeleteText}
                    onChange={e => setConfirmDeleteText(e.target.value)}
                    placeholder={isDe ? 'Event-Titel kleingeschrieben…' : 'Event title in lowercase…'}
                    disabled={isDeleting}
                    autoFocus
                    style={{ width: '100%' }}
                  />
                </div>
              ) : (
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  {isDe ? 'Dieses Event hat keine Anmeldungen über das Organizer-Team hinaus. Wirklich löschen?' : 'This event has no registrations beyond the organizer team. Delete it?'}
                </p>
              )}
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
                <button
                  type="button"
                  className="btn btn-outline"
                  onClick={close}
                  disabled={isDeleting}
                >{isDe ? 'Abbrechen' : 'Cancel'}</button>
                <button
                  type="button"
                  className="btn btn-danger"
                  disabled={!canDelete || isDeleting}
                  style={{
                    background: canDelete && !isDeleting ? 'var(--dex-red, #c00)' : 'var(--dex-gray-300)',
                    color: '#fff', border: 'none',
                    cursor: canDelete && !isDeleting ? 'pointer' : 'not-allowed',
                    padding: '8px 16px',
                  }}
                  onClick={async () => {
                    if (!canDelete || !confirmDeleteEvent) return;
                    setIsDeleting(true);
                    setDeletingId(confirmDeleteEvent.id);
                    try {
                      const ok = await deleteEvent(confirmDeleteEvent.id);
                      if (!ok) {
                        const why = getLastEventDeleteError();
                        showAlert(why || (isDe ? 'Das Event konnte nicht gelöscht werden.' : 'The event could not be deleted.'), { variant: 'error' });
                      }
                    } finally {
                      setIsDeleting(false);
                      setDeletingId(null);
                      close();
                    }
                  }}
                >
                  <Trash2 size={14} /> {isDeleting ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Endgültig löschen' : 'Delete')}
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    );
};

