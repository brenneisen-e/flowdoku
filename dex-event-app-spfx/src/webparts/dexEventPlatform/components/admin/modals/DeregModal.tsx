/* DeregModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14620-14773 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { AlertCircle, Calendar, Mail, Trash2, Users } from '../../Icons';
import { cx } from '../../dexUi';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { translateStatus } from '../../../utils/eventStatus';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface DeregModalProps {
  closeDeregModal: () => void;
  deregBusy: boolean;
  deregModal: { emailKey: string; name: string; email: string; items: { child: DeloitteEvent; reg: SPRegistration; isParent?: boolean; }[]; };
  deregSelected: Set<string>;
  deregSilent: boolean;
  inactiveAccounts: string[];
  isDe: boolean;
  runDeregModal: () => Promise<void>;
  selectedEvent: DeloitteEvent;
  setDeregSelected: React.Dispatch<React.SetStateAction<Set<string>>>;
  setDeregSilent: React.Dispatch<React.SetStateAction<boolean>>;
}

export const DeregModal: React.FC<DeregModalProps> = (p) => {
  const { closeDeregModal, deregBusy, deregModal, deregSelected, deregSilent, inactiveAccounts, isDe, runDeregModal, selectedEvent, setDeregSelected, setDeregSilent } = p;
        const allChecked = deregModal.items.length > 0 && deregModal.items.every(i => deregSelected.has(i.child.id));
        const selectedCount = deregModal.items.filter(i => deregSelected.has(i.child.id)).length;
        const toggleAll = (): void => {
          if (allChecked) setDeregSelected(new Set());
          else setDeregSelected(new Set(deregModal.items.map(i => i.child.id)));
        };
        const toggleOne = (cid: string): void => {
          setDeregSelected(prev => {
            const next = new Set(prev);
            if (next.has(cid)) next.delete(cid); else next.add(cid);
            return next;
          });
        };
        // v31.2: Die Folgen-Zeilen unten sprechen die Klammer nur an, wenn sie
        // in der Auswahl steckt — und sagen dann vorher, was runDeregModal tut:
        // Klammer ZULETZT und nur, wenn kein Termin mehr aktiv bleibt (v30.68).
        // Vorher erfuhr der Organizer das erst aus der Fehlermeldung danach.
        const total = deregModal.items.length;
        const parentSelected = deregModal.items.some(i => !!i.isParent && deregSelected.has(i.child.id)) && deregModal.items.some(i => !i.isParent);
        const subsRemainActive = deregModal.items.filter(i => !i.isParent && !deregSelected.has(i.child.id)).length;
        const parentTerm = selectedEvent.subEventsOnlyMode ? (isDe ? 'Klammer' : 'Umbrella') : (isDe ? 'Haupt-Event' : 'Main event');
        const accountInactive = inactiveAccounts.indexOf(deregModal.emailKey) >= 0;
        const consequence = (icon: React.ReactNode, text: React.ReactNode, warn?: boolean): React.ReactElement => (
          <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start', fontSize: '0.82rem', lineHeight: 1.5, color: warn ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-700, #444)' }}>
            <span style={{ flexShrink: 0, marginTop: 2, color: warn ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-500)' }}>{icon}</span>
            <span>{text}</span>
          </div>
        );
        return (
          <Modal
            open={!!deregModal}
            onClose={() => { if (!deregBusy) closeDeregModal(); }}
            maxWidth={640}
            dismissable={!deregBusy}
            ariaLabel={isDe ? 'Teilnehmer abmelden' : 'Deregister attendee'}
            icon={<Trash2 size={20} />}
            title={isDe ? <>Abmelden: <span style={{ color: 'var(--dex-green-dark)' }}>{deregModal.name}</span></> : <>Deregister: <span style={{ color: 'var(--dex-green-dark)' }}>{deregModal.name}</span></>}
            subtitle={isDe ? 'Wähle die Anmeldungen, die verbindlich abgemeldet werden sollen. Dieser Schritt lässt sich nicht automatisch rückgängig machen.' : 'Pick the registrations to cancel for good. This step cannot be undone automatically.'}
            footer={<>
              {total > 0 && (
                <span className="dex-ui-modal-foot-left dex-ui-muted">
                  {selectedCount === 0
                    ? (isDe ? 'Wähle mindestens eine Anmeldung aus.' : 'Select at least one registration.')
                    : (isDe ? `${selectedCount} von ${total} Anmeldungen — ${deregSilent ? 'ohne Benachrichtigung' : 'mit Mail und Outlook-Absage'}.` : `${selectedCount} of ${total} registrations — ${deregSilent ? 'without notification' : 'with email and Outlook withdrawal'}.`)}
                </span>
              )}
              <button type="button" className="btn btn-secondary" onClick={closeDeregModal} disabled={deregBusy}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button type="button" className="btn btn-danger" onClick={runDeregModal} disabled={deregBusy || selectedCount === 0}>
                {deregBusy
                  ? (isDe ? 'Melde ab…' : 'Cancelling…')
                  : (isDe ? `Jetzt abmelden (${selectedCount})` : `Deregister now (${selectedCount})`)}
              </button>
            </>}
          >
            {total === 0 ? (
              <div className="dex-ui-empty">
                <span className="dex-ui-empty-icon"><Users size={20} /></span>
                <div className="dex-ui-empty-title">{isDe ? 'Nichts mehr abzumelden' : 'Nothing left to cancel'}</div>
                {isDe ? 'Diese Person hat keine aktive Anmeldung mehr — weder auf dem Haupt-Event noch in einem Sub-Event.' : 'This person no longer has an active registration — neither on the main event nor in any sub-event.'}
              </div>
            ) : (
              <>
                {/* 1) Pflicht: Welche Anmeldungen? „Alle auswählen" als Textknopf
                    neben der Überschrift statt eigener Checkbox-Zeile. */}
                <div className="dex-ui-section" style={{ marginTop: 0 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
                    <div className="dex-ui-section-title" style={{ flex: 1, margin: 0 }}>{isDe ? 'Welche Anmeldungen?' : 'Which registrations?'}</div>
                    <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={toggleAll} disabled={deregBusy}>
                      {allChecked ? (isDe ? 'Auswahl aufheben' : 'Clear selection') : (isDe ? 'Alle Anmeldungen auswählen' : 'Select all registrations')}
                    </button>
                  </div>
                  <div className="dex-ui-stack" style={{ gap: 6, maxHeight: 320, overflowY: 'auto' }}>
                    {deregModal.items.map(({ child, reg, isParent }) => {
                      const checked = deregSelected.has(child.id);
                      return (
                        <label key={child.id} className={cx('dex-ui-toggle-row', checked && 'is-active', deregBusy && 'is-disabled')} style={{ alignItems: 'center', padding: '9px 12px' }}>
                          <input type="checkbox" checked={checked} onChange={() => toggleOne(child.id)} disabled={deregBusy} style={{ marginTop: 0 }} />
                          <span className="dex-ui-toggle-row-body">
                            <span className="dex-ui-toggle-row-title">
                              {isParent ? child.title : (shortSubEventTitle(child.title, selectedEvent.title) || child.title)}
                              {/* v29.29: Die Klammer-Zeile ausweisen — sonst liest
                                  sie sich wie ein weiteres Sub-Event. */}
                              {isParent && <span className="dex-ui-pill dex-ui-pill--green">{parentTerm}</span>}
                            </span>
                          </span>
                          <span className={cx('dex-ui-pill', reg.Status === 'Eingecheckt' ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>{translateStatus(reg.Status, isDe)}</span>
                        </label>
                      );
                    })}
                  </div>
                </div>
                {/* 2) Optional: still? — v29.29: der Regelfall bei einer Person,
                    die das Unternehmen verlassen hat: Das Postfach existiert
                    nicht mehr, Mail und Outlook-Absage laufen ins Leere. Das
                    Nachrücken von der Warteliste läuft trotzdem. */}
                <div className="dex-ui-section" style={{ marginTop: 0 }}>
                  <div className="dex-ui-section-title">{isDe ? 'Soll die Person benachrichtigt werden?' : 'Should the person be notified?'}</div>
                  <label className={cx('dex-ui-toggle-row', deregSilent && 'is-active', deregBusy && 'is-disabled')}>
                    <input type="checkbox" checked={deregSilent} onChange={e => setDeregSilent(e.target.checked)} disabled={deregBusy} />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">{isDe ? 'Still abmelden — ohne E-Mail und ohne Outlook-Absage' : 'Cancel silently — no email, no Outlook withdrawal'}</span>
                      <span className="dex-ui-toggle-row-desc">
                        {isDe
                          ? 'Für Personen, die das Unternehmen verlassen haben: Das Postfach existiert nicht mehr, die Abmelde-Mail käme als Unzustellbarkeits-Meldung zurück. Frei werdende Plätze rücken trotzdem nach, und die nachrückende Person bekommt ihre Mail wie immer.'
                          : 'For people who have left the company: the mailbox no longer exists, so the cancellation email would bounce. Freed seats are still filled from the waitlist, and the promoted person receives their mail as usual.'}
                      </span>
                      {accountInactive && (
                        <span className="dex-ui-toggle-row-desc" style={{ fontWeight: 600, color: 'var(--dex-orange-dark, #b35a00)' }}>
                          {isDe
                            ? 'Für diese Adresse wurde kein aktives Deloitte-Konto gefunden — deshalb ist die stille Abmeldung vorausgewählt.'
                            : 'No active Deloitte account was found for this address — that is why silent cancellation is preselected.'}
                        </span>
                      )}
                    </span>
                  </label>
                </div>
                {/* 3) Folgen — je eine Zeile, abhängig vom Schalter oben. Vorher
                    stand alles als ein Absatz im orangen Kasten ganz oben. */}
                <div className="dex-ui-section" style={{ marginTop: 0 }}>
                  <div className="dex-ui-section-title">{isDe ? 'Was beim Abmelden passiert' : 'What happens on cancellation'}</div>
                  <div className="dex-ui-card dex-ui-card--soft dex-ui-stack" style={{ gap: 8, padding: '12px 14px' }}>
                    {consequence(<Mail size={16} />, deregSilent
                      ? (isDe ? 'Keine Abmelde-Bestätigung und keine Outlook-Absage — die Person erfährt nichts.' : 'No cancellation confirmation and no Outlook withdrawal — the person is not informed.')
                      : (isDe ? <>Pro Anmeldung eine <strong>Abmelde-Bestätigung per Mail</strong>, der Outlook-Termin wird zurückgezogen — außer Mails bzw. Outlook sind event-weit deaktiviert.</> : <>One <strong>cancellation confirmation by email</strong> per registration; the Outlook invite is withdrawn — unless emails or Outlook are disabled event-wide.</>))}
                    {consequence(<Users size={16} />, isDe
                      ? 'Frei werdende Plätze rücken von der Warteliste nach, und die Teilnehmer-IDs werden neu vergeben.'
                      : 'Freed seats are filled from the waitlist, and participant IDs are reassigned.')}
                    {parentSelected && consequence(<Calendar size={16} strokeWidth={2} />, subsRemainActive > 0
                      ? (isDe ? `${parentTerm}-Zeile bleibt bestehen: ${subsRemainActive === 1 ? 'Ein Termin bleibt' : `${subsRemainActive} Termine bleiben`} aktiv. Sie wird erst abgemeldet, wenn kein Termin mehr aktiv ist.` : `${parentTerm} row remains: ${subsRemainActive === 1 ? 'one date stays' : `${subsRemainActive} dates stay`} active. It is only cancelled once no date is active anymore.`)
                      : (isDe ? `${parentTerm}-Zeile wird zuletzt abgemeldet — nach den Terminen.` : `${parentTerm} row is cancelled last — after the dates.`), subsRemainActive > 0)}
                    {consequence(<AlertCircle size={16} />, isDe ? 'Nicht automatisch rückgängig zu machen.' : 'Cannot be undone automatically.', true)}
                  </div>
                </div>
              </>
            )}
          </Modal>
        );
};

