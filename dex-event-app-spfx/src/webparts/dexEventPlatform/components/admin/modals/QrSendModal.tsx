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
import { DeniedSubEventList } from '../../admin/adminTypes';
import { bundledCommOf } from '../../../utils/bundledComm';

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
  /** v31.75: Ziel des Versands setzen — '' = das geöffnete Event, sonst ein Termin. */
  setQrSendTargetId: React.Dispatch<React.SetStateAction<string>>;
  qrSendTarget: DeloitteEvent | null;
  qrSendTargetRegs: SPRegistration[] | null;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
  deniedSubEventLists: DeniedSubEventList[];
}

export const QrSendModal: React.FC<QrSendModalProps> = (p) => {
  const { childEventsOf, currentUser, getQrMailOverride, isDe, isSendingQR, openQrMailEditor, qrFullSendAction, qrHelpOpen, qrSendModalOpen, qrSendResult, qrSentCount, qrSubMailsOpen, qrTestSendAction, registrations, selectedEvent, setQrHelpOpen, setQrSendModalOpen, setQrSubMailsOpen, setQrSendTargetId, qrSendTarget, qrSendTargetRegs, subEventRegsByEventId, deniedSubEventLists } = p;
  // v31.75: Alles unten rechnet mit der Liste des ZIELS. `null` = nicht
  // lesbar → der Versand-Knopf bleibt aus, der Dialog sagt warum.
  const regs: SPRegistration[] | null = qrSendTarget ? qrSendTargetRegs : registrations;
  const zielEv: DeloitteEvent = qrSendTarget || selectedEvent;
  const kids = selectedEvent ? childEventsOf(selectedEvent.id) : [];
  const listUnreadable = (ev: DeloitteEvent): boolean => deniedSubEventLists.some(d => d.title === (ev.title || ev.id));
  const ohneCode = (list: SPRegistration[] | null): number => (list || []).filter(r => r.Status === 'Angemeldet').length;
  const schliessen = (): void => { setQrSendModalOpen(false); setQrSendTargetId(''); };
  return (
        <Modal
          open={qrSendModalOpen}
          onClose={schliessen}
          dismissable={!isSendingQR}
          maxWidth={640}
          ariaLabel="QR-Codes versenden"
          title={isDe ? 'QR-Codes an Teilnehmer' : 'QR codes to attendees'}
          subtitle={isDe
            ? 'Jede angemeldete Person bekommt ihren persönlichen Code per Mail — mit Name und Teilnehmer-ID daneben.'
            : 'Every registered person receives their personal code by email — with name and attendee ID beside it.'}
          icon={<QrCode size={20} />}
          footer={
            <button type="button" className="btn btn-secondary" onClick={schliessen} disabled={isSendingQR}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          }
        >
            {/* v31.75: Für welches Event geht der Code raus? Nutzer-Ansage
                22.09.2026: „im Modal entscheiden, für welches Event der QR-Code
                versendet wird — Sub-Event oder Klammer-Event". Vorher galt
                stillschweigend das geöffnete Event; für einen Termin musste
                man ihn erst im Organizer Center öffnen. Der Code trägt die
                Nummer des gewählten Events — am Check-in wählt das Team dann
                genau dieses Event. */}
            {kids.length > 0 && (
              <div className="dex-ui-field">
                <div className="dex-ui-label">{isDe ? 'Für welches Event geht der Code raus?' : 'Which event is the code for?'}</div>
                <div className="dex-ui-inline">
                  {[selectedEvent, ...kids].map(ev => {
                    const isParent = ev.id === selectedEvent.id;
                    const on = isParent ? !qrSendTarget : (!!qrSendTarget && qrSendTarget.id === ev.id);
                    const unreadable = !isParent && listUnreadable(ev);
                    const n = isParent ? ohneCode(registrations) : (unreadable ? null : ohneCode(subEventRegsByEventId[ev.id] || null));
                    return (
                      <button
                        key={ev.id}
                        type="button"
                        className={cx('dex-ui-chip', on && 'is-active')}
                        aria-pressed={on}
                        disabled={isSendingQR}
                        title={unreadable ? (isDe ? 'Teilnehmerliste nicht lesbar' : 'Attendee list not readable') : ev.title}
                        onClick={() => setQrSendTargetId(isParent ? '' : ev.id)}
                      >
                        {isParent ? (isDe ? 'Hauptevent' : 'Main event') : (ev.title || (isDe ? 'Ohne Titel' : 'Untitled'))}
                        <span className={cx('dex-ui-pill', 'dex-ui-pill--sm', unreadable ? 'dex-ui-pill--orange' : (n && n > 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--gray'))} style={{ marginLeft: 6 }}>
                          {unreadable ? (isDe ? 'nicht lesbar' : 'unreadable') : `${n} ${isDe ? 'ohne Code' : 'without code'}`}
                        </span>
                      </button>
                    );
                  })}
                </div>
                {/* v31.76: zwei kurze Sätze statt eines Absatzes („das versteht
                    man nicht", 22.09.2026). Der Bündelungs-Hinweis nur, wenn er
                    zutrifft — als eigener Kasten, nicht als Nebensatz. */}
                <div className="dex-ui-help">
                  {isDe
                    ? <>Der Code gilt für <strong>{zielEv.title}</strong>. Am Check-in wählt das Team dasselbe Event.</>
                    : <>The code is for <strong>{zielEv.title}</strong>. At check-in the team picks the same event.</>}
                </div>
                {!qrSendTarget && !!selectedEvent.subEventsOnlyMode && !bundledCommOf(selectedEvent).qr && (
                  <div className="dex-ui-callout dex-ui-callout--neutral dex-ui-callout--sm" style={{ marginTop: 6 }}>
                    <span>{isDe
                      ? <>Wer sich <strong>später</strong> anmeldet, bekommt automatisch den Code seines Termins. Soll auch dann der Code des Hauptevents rausgehen: im Assistenten unter &bdquo;Kommunikation&ldquo; den Schalter &bdquo;Einen QR-Code fürs Gesamt-Event&ldquo; einschalten.</>
                      : <>Anyone registering <strong>later</strong> automatically gets the code of their session. To send the main-event code in that case too: in the wizard under “Communication”, enable “One QR code for the whole event”.</>}</span>
                  </div>
                )}
              </div>
            )}
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
              if (regs === null) {
                return (
                  <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" role="status">
                    <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
                    <span>{isDe
                      ? <>Die Teilnehmerliste von <strong>{zielEv.title}</strong> ist nicht lesbar — ob jemand ohne Code ist, lässt sich nicht sagen. Kein Versand, bis sie lesbar ist (Berechtigung oder Drosselung).</>
                      : <>The attendee list of <strong>{zielEv.title}</strong> is not readable — whether anyone lacks a code cannot be told. No sending until it can be read (permission or throttling).</>}</span>
                  </div>
                );
              }
              const without = regs.filter(r => r.Status === 'Angemeldet').length;
              const withQr = regs.filter(r => r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
              const externalCount = regs.filter(r => r.Status === 'Angemeldet').filter(r => isExternalEmail(r.ParticipantEmail)).length;
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
                // v31.75: Zähler und Text-Vermerk gehören zum ZIEL.
                const pending = (regs || []).filter(r => r.Status === 'Angemeldet').length;
                const textCustomized = !!getQrMailOverride(zielEv);
                // v31.2: Wer selbst nicht angemeldet ist, wundert sich beim
                // Test-Scan — der Hinweis steht deshalb IM Test-Schritt (als
                // zweite Zeile des Hints), nicht mehr am Ende des Dialogs und
                // auch nicht als eigener Block zwischen 2 und 3, der die Kette
                // unterbräche.
                const orgEmail = (currentUser.email || '').toLowerCase();
                const isOrgRegistered = !!orgEmail && (regs || []).some(r => (r.ParticipantEmail || '').toLowerCase() === orgEmail && (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt'));
                return (
                  <>
                    {/* v31.76: EIN Schritt „E-Mail anpassen" statt „Vorschau
                        ansehen" + Textlink „Mail-Text anpassen" (Nutzer-Frage
                        22.09.2026: „warum steht da nicht E-Mail anpassen statt
                        Ansehen?"). Der Editor zeigt die Mail rechts live —
                        Text, Kopfbild und Vorschau sind dort eine Sache; ein
                        zweiter Weg nur zum Ansehen war eine Bedienung doppelt. */}
                    {stepRow(1,
                      <>
                        {isDe ? 'E-Mail anpassen' : 'Customize the email'}
                        {textCustomized && <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 8 }}>{isDe ? 'angepasst' : 'customized'}</span>}
                      </>,
                      isDe ? 'Text und Kopfbild — rechts siehst du sofort, wie die Mail ankommt.' : 'Text and header image — the preview on the right shows the result immediately.',
                      <button className="btn btn-outline dex-ui-btn-sm" disabled={isSendingQR} onClick={() => { openQrMailEditor().catch(() => { /* */ }); }} style={{ minWidth: 110 }}>
                        <Pencil size={14} />
                        {isDe ? 'Anpassen' : 'Customize'}
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
                        disabled={isSendingQR || pending === 0 || regs === null}
                        style={{ minWidth: 110, fontWeight: 700 }}
                      >
                        {isSendingQR
                          ? `${isDe ? 'Versende' : 'Sending'}… (${qrSentCount})`
                          : regs === null
                            ? (isDe ? 'Nicht lesbar' : 'Unreadable')
                            : pending === 0
                              ? (isDe ? 'Erledigt' : 'Done')
                              : (isDe ? `An ${pending} senden` : `Send to ${pending}`)}
                      </button>,
                      pending === 0 && regs !== null && !isSendingQR ? 'done' : undefined)}
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
                if (kids.length === 0) return null;
                const term = (selectedEvent && selectedEvent.childEventTermPlural) || (isDe ? 'Sub-Events' : 'sub-events');
                const customized = kids.filter(ce => !!getQrMailOverride(ce)).length;
                return (
                  <>
                    <button type="button" className={cx('dex-ui-disclosure', qrSubMailsOpen && 'is-open')} aria-expanded={qrSubMailsOpen} onClick={() => setQrSubMailsOpen(v => !v)}>
                      <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                      {isDe ? `E-Mails der ${term} einzeln anpassen` : `Customize the ${term} emails individually`}
                      {/* Dieselbe grüne Pill wie an den Terminen darunter — ein Zeichen für dasselbe. */}
                      {customized > 0 && <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 'auto' }}>{isDe ? `${customized} angepasst` : `${customized} customized`}</span>}
                    </button>
                    {qrSubMailsOpen && (
                      <div className="dex-ui-disclosure-body">
                        {/* v31.2: Der QR-Override liegt je Event-Zeile; Versand und Editor
                            fallen NICHT auf das Haupt-Event zurück, sondern auf den Standardtext. */}
                        <div className="dex-ui-muted" style={{ marginBottom: 6 }}>
                          {isDe
                            ? 'Ohne eigene Anpassung geht die Standard-QR-Mail raus — nicht Text und Bild des Haupt-Events.'
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

