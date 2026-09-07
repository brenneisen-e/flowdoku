/* SubmitConfirmModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Sicherheitshinweis-Dialog vor dem Absenden (v18.75), inklusive Nachwahl der
 * Sub-Events. Inhalt zeichengleich uebernommen; die Anzeige-Bedingung
 * (`confirmDialogOpen && event`) ist beim Aufrufer geblieben. */
import * as React from 'react';
import Modal from '../Modal';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';
// v31.2: Gemeinsame Klassen (Toggle-Zeilen, Callouts, Pills) statt
// Inline-Styles — nur so bekommen die Zeilen einen Hover (docs/ui-leitfaden.md).
import { cx } from '../dexUi';
import { AlertCircle, Check } from '../Icons';

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
  /** v30.67: Ref statt Funktion — s. RegistrationPage.handleSubmitRef. */
  handleSubmitRef: React.MutableRefObject<() => Promise<void>>;
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
  const { childEvents, childTermPlural, confirmDialogAck, confirmDialogConfirmedRef, confirmDialogOpen, confirmDraftParent, confirmDraftSessions, event, handleSubmitRef, locale, registerForOther, resolveMainEventLabel, selectedSessions, sessionFieldValues, setConfirmDialogAck, setConfirmDialogOpen, setConfirmDraftParent, setConfirmDraftSessions, setPendingSubEventModal, setRegisterForParent, setSelectedSessions, willRegisterParent } = p;
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
        const isDe = locale === 'de';
        // v19.0: statt generischem „Punkte/items" den konfigurierten
        // Section-Begriff verwenden (Default „Event-Sections").
        const sectionTerm = childTermPlural || (isDe ? 'Event-Sections' : 'event sections');
        // v31.2: Zähler „gewählt / gesamt" über der Liste — die Person sieht
        // auf einen Blick, wofür sie sich gleich anmeldet; eine fest gebuchte
        // Klammer-Zeile (Stellvertreter-Modus) zählt dabei mit.
        const parentChecked = parentEditable ? confirmDraftParent : true;
        const selectedCount = (showParent && parentChecked ? 1 : 0) + confirmDraftSessions.size;
        const totalCount = (showParent ? 1 : 0) + allChildren.length;
        return (
          <Modal
            open={confirmDialogOpen}
            onClose={() => setConfirmDialogOpen(false)}
            maxWidth={560}
            ariaLabel={isDe ? 'Anmeldung bestätigen' : 'Confirm registration'}
            // v31.2: Kopf und Fuß über die Modal-Props — Erklärsatz als Untertitel,
            // Aktionen in der Fußzeile, damit alle Dialoge gleich aufgebaut sind.
            title={isDe ? 'Bitte bestätigen' : 'Please confirm'}
            subtitle={isFree
              ? (isDe ? 'Lies den Hinweis kurz durch und bestätige ihn — dann geht deine Anmeldung raus.' : 'Read the note briefly and acknowledge it — then your registration goes out.')
              : (isDe ? `Du meldest dich für die angehakten ${sectionTerm} an. Vor dem Absenden kannst du hier noch ab- oder zuwählen.` : `You are registering for the checked ${sectionTerm}. You can still select or deselect here before submitting.`)}
            icon={<Check size={20} />}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={() => setConfirmDialogOpen(false)}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
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
                  // v30.67: Über den Ref — das beim Öffnen übergebene handleSubmit
                  // schloss über die Auswahl von damals; die eben gesetzten Haken
                  // kamen nie im Submit an (B blieb angemeldet, C fehlte, das
                  // abgewählte Haupt-Event wurde trotzdem gebucht).
                  setTimeout(() => { handleSubmitRef.current().catch(() => { /* */ }); }, 60);
                }}
              >
                {isDe ? 'Anmeldung bestätigen' : 'Confirm registration'}
              </button>
            </>}
          >
            {isFree ? (
              <div className="dex-ui-stack">
                {/* v31.2: Der Organizer-Text als ruhiger Kasten; die Bestätigung
                    als Toggle-Zeile mit Hover und einer Zeile, was der Haken bewirkt. */}
                <div className="dex-ui-callout dex-ui-callout--neutral" style={{ whiteSpace: 'pre-wrap', color: 'var(--dex-gray-800, #333)', fontSize: '0.9rem' }}>
                  {(event.confirmDialogText || '').trim() || (isDe ? 'Bitte bestätige deine Anmeldung.' : 'Please confirm your registration.')}
                </div>
                <label className={cx('dex-ui-toggle-row', confirmDialogAck && 'is-active')}>
                  <input type="checkbox" checked={confirmDialogAck} onChange={e => setConfirmDialogAck(e.target.checked)} />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">{isDe ? 'Ich habe den Hinweis gelesen und bestätige.' : 'I have read and acknowledge the note.'}</span>
                    <span className="dex-ui-toggle-row-desc">{isDe ? 'Mit dem Haken wird der Knopf „Anmeldung bestätigen“ frei.' : 'Ticking this unlocks the “Confirm registration” button.'}</span>
                  </span>
                </label>
              </div>
            ) : (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  {isDe ? 'Deine Auswahl' : 'Your selection'}
                  <span className={cx('dex-ui-pill', selectedCount > 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>
                    {isDe ? `${selectedCount} von ${totalCount} gewählt` : `${selectedCount} of ${totalCount} selected`}
                  </span>
                </div>
                <div className="dex-ui-stack" style={{ gap: 8 }}>
                  {showParent && (() => {
                    const lbl = resolveMainEventLabel(isDe ? 'Haupt-Event' : 'main event');
                    const when = dtRange(event.startDate, event.endDate);
                    return (
                      // v31.2: Im Stellvertreter-Modus ist die Klammer fest gebucht —
                      // dann kein Zeiger-Cursor und eine Zeile, warum der Haken nicht geht.
                      <label className={cx('dex-ui-toggle-row', parentChecked && 'is-active')} style={parentEditable ? undefined : { cursor: 'default' }}>
                        <input
                          type="checkbox"
                          checked={parentChecked}
                          disabled={!parentEditable}
                          onChange={e => setConfirmDraftParent(e.target.checked)}
                        />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            {event.title}
                            {lbl ? <span className="dex-ui-pill dex-ui-pill--gray">{lbl}</span> : null}
                          </span>
                          {when && <span className="dex-ui-toggle-row-desc">{when}</span>}
                          {!parentEditable && (
                            <span className="dex-ui-toggle-row-desc">{isDe ? 'Fester Teil dieser Anmeldung — lässt sich hier nicht abwählen.' : 'A fixed part of this registration — cannot be deselected here.'}</span>
                          )}
                        </span>
                      </label>
                    );
                  })()}
                  {allChildren.map(ce => {
                    const on = confirmDraftSessions.has(ce.id);
                    const when = dtRange(ce.startDate, ce.endDate);
                    return (
                      <label key={ce.id} className={cx('dex-ui-toggle-row', on && 'is-active')}>
                        <input
                          type="checkbox"
                          checked={on}
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
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{ce.title}</span>
                          {when && <span className="dex-ui-toggle-row-desc">{when}</span>}
                        </span>
                      </label>
                    );
                  })}
                </div>
                {!canConfirm && (
                  <div className="dex-ui-callout dex-ui-callout--warn" role="alert" style={{ marginTop: 10 }}>
                    <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                    <span>{isDe ? `Wähl mindestens eine der ${sectionTerm} aus — sonst gibt es nichts anzumelden.` : `Select at least one of the ${sectionTerm} — otherwise there is nothing to register.`}</span>
                  </div>
                )}
              </div>
            )}
          </Modal>
        );
};
