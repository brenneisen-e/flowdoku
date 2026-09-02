/* DeregModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14620-14773 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { AlertCircle, Trash2, X } from '../../Icons';
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
        const someChecked = deregModal.items.some(i => deregSelected.has(i.child.id));
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
        return (
          <Modal
            open={!!deregModal}
            onClose={() => { if (!deregBusy) closeDeregModal(); }}
            maxWidth={640}
            dismissable={!deregBusy}
            ariaLabel={isDe ? 'Teilnehmer abmelden' : 'Deregister attendee'}
          >
            <div className="flex-between">
              <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                <Trash2 size={18} />{' '}
                {isDe ? 'Abmelden' : 'Deregister'}
                {' — '}
                <span style={{ color: 'var(--dex-green-dark)' }}>{deregModal.name}</span>
              </h3>
              <button
                onClick={closeDeregModal}
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}
                aria-label={isDe ? 'Schließen' : 'Close'}
                disabled={deregBusy}
              ><X size={20} /></button>
            </div>
            {/* Orange Sicherheits-Hinweis */}
            <div style={{ display: 'flex', gap: 10, padding: '10px 12px', background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-gray-800)' }}>
              <span style={{ color: 'var(--dex-orange, #ed8b00)', flexShrink: 0, marginTop: 1 }}><AlertCircle size={18} /></span>
              <span>
                {isDe
                  ? <>Die ausgewählten Anmeldungen werden <strong>verbindlich abgemeldet</strong>. Pro Anmeldung bekommt die Person (sofern nicht unten still abgemeldet oder event-weit deaktiviert) eine Abmelde-Bestätigung, der Outlook-Termin wird zurückgezogen, frei werdende Plätze rücken nach und die Teilnehmer-IDs werden neu vergeben. Dieser Schritt lässt sich nicht automatisch rückgängig machen.</>
                  : <>The selected registrations will be <strong>cancelled for good</strong>. For each registration the person receives (unless cancelled silently below or disabled event-wide) a cancellation confirmation, the Outlook invite is withdrawn, freed seats are filled from the waitlist and participant IDs are reassigned. This step cannot be undone automatically.</>}
              </span>
            </div>
            {deregModal.items.length === 0 ? (
              <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
                {isDe ? 'Diese Person hat keine aktive Anmeldung mehr — weder auf dem Haupt-Event noch in einem Sub-Event.' : 'This person no longer has an active registration — neither on the main event nor in any sub-event.'}
              </p>
            ) : (
              <>
                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-800)', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={allChecked}
                    ref={el => { if (el) el.indeterminate = !allChecked && someChecked; }}
                    onChange={toggleAll}
                    style={{ accentColor: 'var(--dex-green)' }}
                    disabled={deregBusy}
                  />
                  {isDe ? 'Alle Anmeldungen auswählen' : 'Select all registrations'}
                </label>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 320, overflowY: 'auto' }}>
                  {deregModal.items.map(({ child, reg, isParent }) => {
                    const checked = deregSelected.has(child.id);
                    return (
                      <label
                        key={child.id}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 10, padding: '8px 10px',
                          background: checked ? 'rgba(218,41,28,0.05)' : 'var(--dex-gray-50)',
                          border: `1px solid ${checked ? 'rgba(218,41,28,0.35)' : 'var(--dex-gray-200)'}`,
                          borderRadius: 8, cursor: deregBusy ? 'default' : 'pointer', fontSize: '0.85rem',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleOne(child.id)}
                          style={{ accentColor: 'var(--dex-red, #c00)' }}
                          disabled={deregBusy}
                        />
                        <span style={{ flex: 1, fontWeight: 500 }}>
                          {isParent ? child.title : (shortSubEventTitle(child.title, selectedEvent.title) || child.title)}
                          {/* v29.29: Die Klammer-Zeile ausweisen — sonst liest
                              sie sich wie ein weiteres Sub-Event. */}
                          {isParent && (
                            <span style={{ marginLeft: 8, fontSize: '0.7rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                              {selectedEvent.subEventsOnlyMode
                                ? (isDe ? 'KLAMMER' : 'UMBRELLA')
                                : (isDe ? 'HAUPT-EVENT' : 'MAIN EVENT')}
                            </span>
                          )}
                        </span>
                        <span className={`badge ${reg.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>{translateStatus(reg.Status, isDe)}</span>
                      </label>
                    );
                  })}
                </div>
                {/* v29.29: Stille Abmeldung — der Regelfall bei einer Person,
                    die das Unternehmen verlassen hat: Das Postfach existiert
                    nicht mehr, Mail und Outlook-Absage laufen ins Leere. Das
                    Nachrücken von der Warteliste läuft trotzdem. */}
                <label style={{
                  display: 'flex', alignItems: 'flex-start', gap: 10, padding: '10px 12px', borderRadius: 8,
                  border: `1px solid ${deregSilent ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-200)'}`,
                  background: deregSilent ? 'rgba(237,139,0,0.06)' : 'var(--dex-gray-50)',
                  cursor: deregBusy ? 'default' : 'pointer',
                }}>
                  <input
                    type="checkbox"
                    checked={deregSilent}
                    onChange={e => setDeregSilent(e.target.checked)}
                    style={{ marginTop: 2, accentColor: 'var(--dex-orange, #ed8b00)' }}
                    disabled={deregBusy}
                  />
                  <span style={{ fontSize: '0.84rem', lineHeight: 1.5 }}>
                    <strong>{isDe ? 'Still abmelden — ohne E-Mail und ohne Outlook-Absage' : 'Cancel silently — no email, no Outlook withdrawal'}</strong>
                    <span style={{ display: 'block', marginTop: 3, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                      {isDe
                        ? 'Für Personen, die das Unternehmen verlassen haben: Das Postfach existiert nicht mehr, die Abmelde-Mail käme als Unzustellbarkeits-Meldung zurück. Frei werdende Plätze rücken trotzdem nach, und die nachrückende Person bekommt ihre Mail wie immer.'
                        : 'For people who have left the company: the mailbox no longer exists, so the cancellation email would bounce. Freed seats are still filled from the waitlist, and the promoted person receives their mail as usual.'}
                    </span>
                    {inactiveAccounts.indexOf(deregModal.emailKey) >= 0 && (
                      <span style={{ display: 'block', marginTop: 4, fontSize: '0.76rem', fontWeight: 600, color: 'var(--dex-orange-dark, #b35a00)' }}>
                        {isDe
                          ? 'Für diese Adresse wurde kein aktives Deloitte-Konto gefunden — deshalb ist die stille Abmeldung vorausgewählt.'
                          : 'No active Deloitte account was found for this address — that is why silent cancellation is preselected.'}
                      </span>
                    )}
                  </span>
                </label>
              </>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
              <button type="button" className="btn btn-outline" onClick={closeDeregModal} disabled={deregBusy}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={runDeregModal}
                disabled={deregBusy || selectedCount === 0}
                style={{ opacity: (deregBusy || selectedCount === 0) ? 0.6 : 1, background: 'var(--dex-red, #c00)', borderColor: 'var(--dex-red, #c00)' }}
              >
                {deregBusy
                  ? (isDe ? 'Melde ab…' : 'Cancelling…')
                  : (isDe ? `Abmelden (${selectedCount})` : `Deregister (${selectedCount})`)}
              </button>
            </div>
          </Modal>
        );
};

