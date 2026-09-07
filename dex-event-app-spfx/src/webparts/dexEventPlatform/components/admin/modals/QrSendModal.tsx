/* QrSendModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13627-13805 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { isExternalEmail } from '../../../utils/deloitteDomain';
// v31.2: Gemeinsame Klassen statt qrDisclosureStyle/Inline-Styles — der
// Dialog soll aussehen wie alle anderen Modale, und Inline-Styles können
// kein :hover (CLAUDE.md).
import { cx } from '../../dexUi';
// v31.2: ChevronDown, nicht ChevronRight — die Klasse dex-ui-disclosure dreht
// den Chevron geöffnet um 180°; ein Rechts-Pfeil zeigte dann nach links.
import { AlertCircle, ChevronDown, Pencil, QrCode } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { QrEmailOverride } from '../../../services/EmailTemplates';

export interface QrSendModalProps {
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  currentUser: import("../../../types/index").User;
  getQrMailOverride: (ev: DeloitteEvent | null) => QrEmailOverride | undefined;
  isDe: boolean;
  isSendingQR: boolean;
  openQrMailEditor: (target?: DeloitteEvent) => Promise<void>;
  qrFullSendAction: () => Promise<void>;
  qrHelpOpen: boolean;
  qrPreviewAction: () => Promise<void>;
  qrPreviewLoading: boolean;
  qrSendModalOpen: boolean;
  qrSendResult: string;
  qrSentCount: number;
  qrSubMailsOpen: boolean;
  qrTestSendAction: (liveOverride?: QrEmailOverride, target?: DeloitteEvent) => Promise<void>;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setQrHelpOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setQrSendModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setQrSubMailsOpen: React.Dispatch<React.SetStateAction<boolean>>;
}

export const QrSendModal: React.FC<QrSendModalProps> = (p) => {
  const { childEventsOf, currentUser, getQrMailOverride, isDe, isSendingQR, openQrMailEditor, qrFullSendAction, qrHelpOpen, qrPreviewAction, qrPreviewLoading, qrSendModalOpen, qrSendResult, qrSentCount, qrSubMailsOpen, qrTestSendAction, registrations, selectedEvent, setQrHelpOpen, setQrSendModalOpen, setQrSubMailsOpen } = p;
  return (
        <Modal
          open={qrSendModalOpen}
          onClose={() => setQrSendModalOpen(false)}
          dismissable={!isSendingQR}
          maxWidth={640}
          ariaLabel="QR-Codes versenden"
          title={isDe ? 'QR-Codes an Teilnehmer' : 'QR codes to attendees'}
          subtitle={isDe
            ? 'Jede angemeldete Person bekommt ihren persönlichen Code per Mail — mit Name und Teilnehmer-ID daneben.'
            : 'Every registered person receives their personal code by email — with name and attendee ID beside it.'}
          icon={<QrCode size={20} />}
          footer={
            <button type="button" className="btn btn-secondary" onClick={() => setQrSendModalOpen(false)} disabled={isSendingQR}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          }
        >
            {/* v30.36: Entschlackt. Vorher standen hier fuenf konkurrierende
                Aktionen, zwei Erklaerkaesten, eine Warnung und eine zweite
                Spalte mit dem Self-Check-in — beim Oeffnen musste man erst
                lesen, um handeln zu koennen. Jetzt: drei nummerierte Schritte,
                bei denen der KNOPF der Schritt ist (vorher waren Nummern und
                Knoepfe getrennt und mussten im Kopf zugeordnet werden), und
                alles Erklaerende hinter Aufklappern. Der Self-Check-in ist
                ganz raus: Er ist seit v30.36 eine gleichrangige Wahl im
                Einstiegs-Modal davor — zweimal dieselbe Entscheidung
                anzubieten ist genau die Falle, die CLAUDE.md beschreibt.
                v31.2: Kopf/Fuß über die Modal-Props, Klassen aus dexUi.ts
                (dex-ui-step, -pill, -disclosure, -callout) statt Inline-
                Styles — damit der Dialog aussieht wie die anderen 40. */}

            {/* Stand der Liste: Wer hat schon einen Code, wer noch nicht? */}
            {(() => {
              const without = registrations.filter(r => r.Status === 'Angemeldet').length;
              const withQr = registrations.filter(r => r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
              const externalCount = registrations.filter(r => r.Status === 'Angemeldet').filter(r => isExternalEmail(r.ParticipantEmail)).length;
              return (
                <div className="dex-ui-inline">
                  <span className={cx('dex-ui-pill', without > 0 ? 'dex-ui-pill--orange' : 'dex-ui-pill--green')}><strong>{without}</strong> {isDe ? 'noch ohne Code' : 'still without a code'}</span>
                  <span className="dex-ui-pill dex-ui-pill--gray"><strong>{withQr}</strong> {isDe ? 'mit Code' : 'with code'}</span>
                  {externalCount > 0 && (
                    // v31.2: InfoTooltip statt `title` — ein title auf einem span ist nur per Maus erreichbar.
                    <span className="dex-ui-pill dex-ui-pill--orange">
                      {isDe ? `${externalCount} extern → QR an Organizer` : `${externalCount} external → QR to organizer`}
                      <InfoTooltip text={isDe ? 'Externe Adressen bekommen keine Mail — ihr Code geht an die Organizer.' : 'External addresses get no email — their code goes to the organizers.'} />
                    </span>
                  )}
                </div>
              );
            })()}

            {/* Die drei Schritte — der Knopf IST der Schritt (dex-ui-step).
                Kein dex-ui-section: Das Modal-Gap von 14 px ist schon da. */}
            <div>
              <div className="dex-ui-section-title">{isDe ? 'So gehst du vor' : 'How to proceed'}</div>
              <div className="dex-ui-stack">
              {(() => {
                const stepRow = (n: number, label: React.ReactNode, hint: React.ReactNode, btn: React.ReactElement, state?: 'done' | 'pending'): React.ReactElement => (
                  <div className={cx('dex-ui-step', state === 'done' && 'is-done', state === 'pending' && 'is-pending')}>
                    <span className="dex-ui-step-num">{n}</span>
                    <div className="dex-ui-step-body">
                      <div className="dex-ui-step-title">{label}</div>
                      <div className="dex-ui-step-hint">{hint}</div>
                    </div>
                    <div className="dex-ui-step-action">{btn}</div>
                  </div>
                );
                const pending = registrations.filter(r => r.Status === 'Angemeldet').length;
                const textCustomized = !!getQrMailOverride(selectedEvent);
                // v31.2: Wer selbst nicht angemeldet ist, wundert sich beim
                // Test-Scan — der Hinweis steht deshalb IM Test-Schritt (als
                // zweite Zeile des Hints), nicht mehr am Ende des Dialogs und
                // auch nicht als eigener Block zwischen 2 und 3, der die Kette
                // unterbräche.
                const orgEmail = (currentUser.email || '').toLowerCase();
                const isOrgRegistered = !!orgEmail && registrations.some(r => (r.ParticipantEmail || '').toLowerCase() === orgEmail && (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt'));
                return (
                  <>
                    {stepRow(1,
                      isDe ? 'Vorschau ansehen' : 'Preview the email',
                      <>
                        {isDe ? 'So sieht die Mail aus, die rausgeht.' : 'How the email will look.'}{' '}
                        <button type="button" className="dex-ui-textbtn" disabled={isSendingQR} onClick={() => { openQrMailEditor().catch(() => { /* */ }); }} style={{ padding: '1px 6px', fontSize: '0.76rem' }}>
                          <Pencil size={12} />
                          {isDe ? 'Mail-Text anpassen' : 'Customize email text'}
                          {textCustomized && <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'angepasst' : 'customized'}</span>}
                        </button>
                      </>,
                      <button className="btn btn-outline dex-ui-btn-sm" disabled={isSendingQR || qrPreviewLoading} onClick={() => { qrPreviewAction().catch(() => { /* */ }); }} style={{ minWidth: 110 }}>
                        {qrPreviewLoading ? (isDe ? 'Lädt…' : 'Loading…') : (isDe ? 'Ansehen' : 'View')}
                      </button>)}
                    {stepRow(2,
                      isDe ? 'Testmail an dich schicken' : 'Send a test to yourself',
                      <>
                        {isDe ? 'Geht an alle Organizer dieses Events.' : 'Goes to all organizers of this event.'}
                        {!isOrgRegistered && (
                          <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" role="note" style={{ marginTop: 6 }}>
                            <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
                            <span>
                              {isDe
                                ? 'Du bist selbst nicht angemeldet — die Test-Mail kommt an, aber ein späterer Check-in-Scan findet dich nicht in der Liste.'
                                : 'You are not registered yourself — the test email arrives, but a later check-in scan will not find you in the list.'}
                            </span>
                          </div>
                        )}
                      </>,
                      <button className="btn btn-secondary dex-ui-btn-sm" disabled={isSendingQR} onClick={() => { qrTestSendAction().catch(() => { /* */ }); }} style={{ minWidth: 110 }}>
                        {isDe ? 'Test senden' : 'Send test'}
                      </button>)}
                    {stepRow(3,
                      isDe ? 'An alle ohne Code senden' : 'Send to everyone without a code',
                      isDe ? 'Danach bekommt jede neue Anmeldung ihren Code automatisch.' : 'Afterwards every new registration gets its code automatically.',
                      <button
                        className="btn btn-primary dex-ui-btn-sm"
                        onClick={() => { qrFullSendAction().catch(() => { /* */ }); }}
                        disabled={isSendingQR || pending === 0}
                        style={{ minWidth: 110, fontWeight: 700 }}
                      >
                        {isSendingQR
                          ? `${isDe ? 'Versende' : 'Sending'}… (${qrSentCount})`
                          : pending === 0
                            ? (isDe ? 'Erledigt' : 'Done')
                            : (isDe ? `An ${pending} senden` : `Send to ${pending}`)}
                      </button>,
                      pending === 0 && !isSendingQR ? 'done' : undefined)}
                  </>
                );
              })()}
              </div>
            </div>

            {qrSendResult && (
              <div className={cx('dex-ui-callout', qrSendResult.startsWith('Fehler') ? 'dex-ui-callout--danger' : 'dex-ui-callout--success')} role="status">
                {qrSendResult}
              </div>
            )}

            {/* Erklärendes und Feineinstellungen auf Abruf — beim Öffnen soll
                man handeln können, nicht erst lesen müssen. v31.2: Der Knopf
                „Mail-Text anpassen" steht jetzt bei Schritt 1 (Vorschau) —
                Text und Vorschau sind dasselbe Thema. Hier bleiben die
                Sub-Event-Texte (Feineinstellung) und die Dauer (Erklärung).
                Nackter div: dex-ui-section ohne Titel wäre nur ein Abstand. */}
            <div>
              {(() => {
                const kids = selectedEvent ? childEventsOf(selectedEvent.id) : [];
                if (kids.length === 0) return null;
                const term = (selectedEvent && selectedEvent.childEventTermPlural) || (isDe ? 'Sub-Events' : 'sub-events');
                const customized = kids.filter(ce => !!getQrMailOverride(ce)).length;
                return (
                  <>
                    <button type="button" className={cx('dex-ui-disclosure', qrSubMailsOpen && 'is-open')} aria-expanded={qrSubMailsOpen} onClick={() => setQrSubMailsOpen(v => !v)}>
                      <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                      {isDe ? `Mail-Texte der ${term} einzeln anpassen` : `Customize the ${term} emails individually`}
                      {/* Dieselbe grüne Pill wie an den Terminen darunter — ein Zeichen für dasselbe. */}
                      {customized > 0 && <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 'auto' }}>{isDe ? `${customized} angepasst` : `${customized} customized`}</span>}
                    </button>
                    {qrSubMailsOpen && (
                      <div className="dex-ui-disclosure-body">
                        {/* v31.2: Der QR-Override liegt je Event-Zeile; Versand und Editor
                            fallen NICHT auf das Haupt-Event zurück, sondern auf den Standardtext. */}
                        <div className="dex-ui-muted" style={{ marginBottom: 6 }}>
                          {isDe
                            ? 'Ohne eigenen Text geht der Standardtext der QR-Mail raus — nicht der Text des Haupt-Events.'
                            : 'Without its own text, the standard QR email text is sent — not the main event’s text.'}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                          {kids.map(ce => (
                            <button
                              key={ce.id}
                              type="button"
                              className="dex-ui-textbtn"
                              disabled={isSendingQR}
                              onClick={() => { openQrMailEditor(ce).catch(() => { /* */ }); }}
                              style={{ justifyContent: 'flex-start', textAlign: 'left', maxWidth: '100%' }}
                            >
                              <Pencil size={13} />
                              <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ce.title || (isDe ? 'Ohne Titel' : 'Untitled')}</span>
                              {getQrMailOverride(ce) && (
                                <span className="dex-ui-pill dex-ui-pill--green" style={{ flexShrink: 0 }}>{isDe ? 'angepasst' : 'customized'}</span>
                              )}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                  </>
                );
              })()}

              <button type="button" className={cx('dex-ui-disclosure', qrHelpOpen && 'is-open')} aria-expanded={qrHelpOpen} onClick={() => setQrHelpOpen(v => !v)}>
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                {isDe ? 'Wie lange dauert der Versand?' : 'How long does sending take?'}
              </button>
              {qrHelpOpen && (
                <div className="dex-ui-disclosure-body dex-ui-muted" style={{ lineHeight: 1.5 }}>
                  {isDe
                    ? <>Die Mails gehen <strong>einzeln nacheinander</strong> raus. Bei mehr als 100 Personen kann das <strong>über 10 Minuten</strong> dauern — das ist normal. Ihren Code sehen Teilnehmer ohnehin jederzeit in der App unter &bdquo;Meine Events&ldquo;.</>
                    : <>Emails are sent <strong>one by one</strong>. With more than 100 people this can take <strong>over 10 minutes</strong> — that is normal. Attendees can see their code anytime in the app under &bdquo;My Events&ldquo;.</>}
                </div>
              )}
            </div>

        </Modal>
  );
};

