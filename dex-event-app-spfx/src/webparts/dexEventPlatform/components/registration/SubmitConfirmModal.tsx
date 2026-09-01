/* SubmitConfirmModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Sicherheitshinweis-Dialog vor dem Absenden (v18.75), inklusive Nachwahl der
 * Sub-Events. Inhalt zeichengleich uebernommen; die Anzeige-Bedingung
 * (`confirmDialogOpen && event`) ist beim Aufrufer geblieben. */
import * as React from 'react';
import Modal from '../Modal';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';

/** Sicherheitshinweis-Dialog vor dem Absenden (v18.75). */
export interface SubmitConfirmModalProps {
  childEvents: DeloitteEvent[];
  childTermPlural: string;
  confirmDialogAck: boolean;
  confirmDialogConfirmedRef: React.MutableRefObject<boolean>;
  confirmDialogOpen: boolean;
  confirmDraftParent: boolean;
  confirmDraftSessions: Set<string>;
  event: DeloitteEvent;
  handleSubmit: () => Promise<void>;
  locale: Locale;
  registerForOther: boolean;
  resolveMainEventLabel: (defaultLabel: string) => string | null;
  selectedSessions: Set<string>;
  sessionFieldValues: Record<string, Record<string, string>>;
  setConfirmDialogAck: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDialogOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDraftParent: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDraftSessions: React.Dispatch<React.SetStateAction<Set<string>>>;
  setPendingSubEventModal: React.Dispatch<React.SetStateAction<{ subEventId: string; draftValues: Record<string, string>; }>>;
  setRegisterForParent: React.Dispatch<React.SetStateAction<boolean>>;
  setSelectedSessions: React.Dispatch<React.SetStateAction<Set<string>>>;
  willRegisterParent: boolean;
}
export const SubmitConfirmModal: React.FC<SubmitConfirmModalProps> = (p) => {
  const { childEvents, childTermPlural, confirmDialogAck, confirmDialogConfirmedRef, confirmDialogOpen, confirmDraftParent, confirmDraftSessions, event, handleSubmit, locale, registerForOther, resolveMainEventLabel, selectedSessions, sessionFieldValues, setConfirmDialogAck, setConfirmDialogOpen, setConfirmDraftParent, setConfirmDraftSessions, setPendingSubEventModal, setRegisterForParent, setSelectedSessions, willRegisterParent } = p;
        const isFree = event.confirmDialogMode === 'freetext';
        // v18.76: ALLE Sub-Events zeigen (auch nicht ausgewählte), damit der
        // Teilnehmer im Dialog ab- UND zuwählen kann.
        const allChildren = childEvents;
        // v23.9: Im Klammer-Modus (subEventsOnlyMode) ist das Hauptevent nicht
        // buchbar — es wird nur als Schatten mitgeführt. Deshalb NICHT als
        // wählbare „(Haupt-Event)"-Zeile im Bestätigungs-Dialog zeigen (auch
        // nicht im Stellvertreter-Modus, wo es sonst über registerForOther
        // fälschlich auftauchte).
        const showParent = (willRegisterParent || registerForOther) && !(event && event.subEventsOnlyMode);
        const parentEditable = willRegisterParent && !registerForOther; // proxy: Parent fix
        const canConfirm = isFree
          ? confirmDialogAck
          : (confirmDraftParent || confirmDraftSessions.size > 0 || (showParent && !parentEditable));
        // v18.76: Datum + Uhrzeit pro Eintrag anzeigen.
        const fmtDT = (iso?: string): string => {
          if (!iso) return '';
          try { return new Date(iso).toLocaleString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch { return ''; }
        };
        const dtRange = (s?: string, e?: string): string => {
          const a = fmtDT(s); const b = fmtDT(e);
          return a && b ? `${a} – ${b}` : (a || b);
        };
        return (
          <Modal
            open={confirmDialogOpen}
            onClose={() => setConfirmDialogOpen(false)}
            maxWidth={560}
            padding={24}
            ariaLabel={locale === 'de' ? 'Anmeldung bestätigen' : 'Confirm registration'}
          >
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {locale === 'de' ? 'Bitte bestätigen' : 'Please confirm'}
            </h3>
            {isFree ? (
              <>
                <div style={{
                  margin: '0 0 14px', padding: '12px 14px', whiteSpace: 'pre-wrap',
                  background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)',
                  borderRadius: 8, fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-800)',
                }}>
                  {(event.confirmDialogText || '').trim() || (locale === 'de' ? 'Bitte bestätige deine Anmeldung.' : 'Please confirm your registration.')}
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginBottom: 4 }}>
                  <input type="checkbox" checked={confirmDialogAck} onChange={e => setConfirmDialogAck(e.target.checked)} style={{ marginTop: 3 }} />
                  <span style={{ flex: 1, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                    {locale === 'de' ? 'Ich habe den Hinweis gelesen und bestätige.' : 'I have read and acknowledge the note.'}
                  </span>
                </label>
              </>
            ) : (() => {
              // v19.0: statt generischem „Punkte/items" den konfigurierten
              // Section-Begriff verwenden (Default „Event-Sections").
              const sectionTerm = childTermPlural || (locale === 'de' ? 'Event-Sections' : 'event sections');
              return (
              <>
                <p style={{ margin: '0 0 12px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
                  {locale === 'de'
                    ? `Du meldest dich für die angehakten ${sectionTerm} an. Du kannst vor dem Absenden einzelne ${sectionTerm} ab- oder zuwählen:`
                    : `You are registering for the checked ${sectionTerm}. You can de-/select ${sectionTerm} before submitting:`}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 14 }}>
                  {showParent && (
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: parentEditable ? 'pointer' : 'default', padding: '8px 10px', background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)', borderRadius: 6 }}>
                      <input
                        type="checkbox"
                        checked={parentEditable ? confirmDraftParent : true}
                        disabled={!parentEditable}
                        onChange={e => setConfirmDraftParent(e.target.checked)}
                        style={{ marginTop: 2 }}
                      />
                      <span style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.88rem', fontWeight: 600, display: 'block' }}>{event.title}{(() => { const lbl = resolveMainEventLabel(locale === 'de' ? 'Haupt-Event' : 'main event'); return lbl ? <> <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>({lbl})</span></> : null; })()}</span>
                        {dtRange(event.startDate, event.endDate) && (
                          <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', display: 'block', marginTop: 1 }}>{dtRange(event.startDate, event.endDate)}</span>
                        )}
                      </span>
                    </label>
                  )}
                  {allChildren.map(ce => (
                    <label key={ce.id} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', padding: '8px 10px', background: confirmDraftSessions.has(ce.id) ? 'rgba(134,188,37,0.06)' : 'var(--dex-gray-50, #f7f7f5)', border: `1px solid ${confirmDraftSessions.has(ce.id) ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, borderRadius: 6 }}>
                      <input
                        type="checkbox"
                        checked={confirmDraftSessions.has(ce.id)}
                        style={{ marginTop: 2 }}
                        onChange={e => {
                          if (e.target.checked) {
                            // v18.76: Sub-Event mit eigenen Pflichtfeldern erst über
                            // das Sub-Event-Modal erfassen, damit keine leeren
                            // Pflicht-Antworten entstehen. Dialog schließen, Modal
                            // öffnen; nach dem Ausfüllen erscheint der Dialog erneut.
                            const hasCF = (ce.eventSpecificFields || []).length > 0;
                            if (hasCF && !sessionFieldValues[ce.id] && !selectedSessions.has(ce.id)) {
                              confirmDialogConfirmedRef.current = false;
                              setConfirmDialogOpen(false);
                              setPendingSubEventModal({ subEventId: ce.id, draftValues: { ...(sessionFieldValues[ce.id] || {}) } });
                            } else {
                              setConfirmDraftSessions(prev => { const n = new Set(prev); n.add(ce.id); return n; });
                            }
                          } else {
                            setConfirmDraftSessions(prev => { const n = new Set(prev); n.delete(ce.id); return n; });
                          }
                        }}
                      />
                      <span style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.88rem', display: 'block' }}>{ce.title}</span>
                        {dtRange(ce.startDate, ce.endDate) && (
                          <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', display: 'block', marginTop: 1 }}>{dtRange(ce.startDate, ce.endDate)}</span>
                        )}
                      </span>
                    </label>
                  ))}
                </div>
                {!canConfirm && (
                  <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-red, #c00)' }}>
                    {locale === 'de' ? `Bitte mindestens eine ${sectionTerm} auswählen.` : `Please select at least one of the ${sectionTerm}.`}
                  </p>
                )}
              </>
            );
            })()}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button className="btn btn-secondary" onClick={() => setConfirmDialogOpen(false)} style={{ fontSize: '0.85rem' }}>
                {locale === 'de' ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                className="btn btn-primary"
                disabled={!canConfirm}
                onClick={() => {
                  // v18.75: In der Auswahl-Übersicht die (ggf. angepasste)
                  // Auswahl in den echten State übernehmen, dann Submit erneut
                  // anstoßen (Ref überspringt den Dialog).
                  if (!isFree) {
                    if (parentEditable) setRegisterForParent(confirmDraftParent);
                    setSelectedSessions(new Set(confirmDraftSessions));
                  }
                  confirmDialogConfirmedRef.current = true;
                  setConfirmDialogOpen(false);
                  setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 60);
                }}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Anmeldung bestätigen' : 'Confirm registration'}
              </button>
            </div>
          </Modal>
        );
};
