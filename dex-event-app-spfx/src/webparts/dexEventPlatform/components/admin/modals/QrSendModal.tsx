/* QrSendModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13627-13805 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { isExternalEmail } from '../../../utils/deloitteDomain';
import { qrDisclosureStyle } from '../../admin/adminStyles';
import { Pencil } from '../../Icons';
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
          maxWidth={860}
          padding={24}
          ariaLabel="QR-Codes versenden"
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
                anzubieten ist genau die Falle, die CLAUDE.md beschreibt. */}
            <h3 style={{ margin: '0 0 4px', fontSize: '1.15rem' }}>{isDe ? 'QR-Codes an Teilnehmer' : 'QR codes to attendees'}</h3>
            <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? 'Jede·r bekommt einen persönlichen Code per Mail — mit Name und Teilnehmer-ID daneben.'
                : 'Everyone receives a personal code by email — with name and attendee ID beside it.'}
            </p>

            {(() => {
              const without = registrations.filter(r => r.Status === 'Angemeldet').length;
              const withQr = registrations.filter(r => r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
              const externalCount = registrations.filter(r => r.Status === 'Angemeldet').filter(r => isExternalEmail(r.ParticipantEmail)).length;
              const pill = (bg: string, fg: string, content: React.ReactNode): React.ReactElement => (
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, padding: '5px 12px', borderRadius: 999, background: bg, color: fg, fontSize: '0.8rem', fontWeight: 600 }}>{content}</span>
              );
              return (
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 18 }}>
                  {pill(without > 0 ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.12)', without > 0 ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)', <><strong>{without}</strong> {isDe ? 'ohne Code' : 'without code'}</>)}
                  {pill('var(--dex-gray-100, #eee)', 'var(--dex-gray-600)', <><strong>{withQr}</strong> {isDe ? 'mit Code' : 'with code'}</>)}
                  {externalCount > 0 && pill('#fff3e0', '#7a4a00', isDe ? `${externalCount} extern → QR an Organizer` : `${externalCount} external → QR to organizer`)}
                </div>
              );
            })()}

            {/* Die drei Schritte — der Knopf IST der Schritt. */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              {(() => {
                const stepRow = (n: number, label: React.ReactNode, hint: string, btn: React.ReactElement): React.ReactElement => (
                  <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                    <span style={{ flexShrink: 0, width: 24, height: 24, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 700, fontSize: '0.78rem', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>{n}</span>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '0.86rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{label}</div>
                      <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>{hint}</div>
                    </div>
                    <div style={{ flexShrink: 0 }}>{btn}</div>
                  </div>
                );
                const pending = registrations.filter(r => r.Status === 'Angemeldet').length;
                return (
                  <>
                    {stepRow(1,
                      isDe ? 'Vorschau ansehen' : 'Preview',
                      isDe ? 'So sieht die Mail aus, die rausgeht.' : 'How the email will look.',
                      <button className="btn btn-outline" disabled={isSendingQR || qrPreviewLoading} onClick={() => { qrPreviewAction().catch(() => { /* */ }); }} style={{ fontSize: '0.85rem', minWidth: 110 }}>
                        {qrPreviewLoading ? (isDe ? 'Lädt…' : 'Loading…') : (isDe ? 'Ansehen' : 'View')}
                      </button>)}
                    {stepRow(2,
                      isDe ? 'Test an dich' : 'Test to yourself',
                      isDe ? 'Geht an alle Organisatoren dieses Events.' : 'Goes to all organizers of this event.',
                      <button className="btn btn-secondary" disabled={isSendingQR} onClick={() => { qrTestSendAction().catch(() => { /* */ }); }} style={{ fontSize: '0.85rem', minWidth: 110 }}>
                        {isDe ? 'Senden' : 'Send'}
                      </button>)}
                    {stepRow(3,
                      isDe ? 'An alle ohne Code senden' : 'Send to everyone without a code',
                      isDe ? 'Danach bekommt jede neue Anmeldung ihren Code automatisch.' : 'Afterwards every new registration gets its code automatically.',
                      <button
                        className="btn btn-primary"
                        onClick={() => { qrFullSendAction().catch(() => { /* */ }); }}
                        disabled={isSendingQR || pending === 0}
                        style={{ fontSize: '0.85rem', minWidth: 110, fontWeight: 700 }}
                      >
                        {isSendingQR
                          ? `${isDe ? 'Versende' : 'Sending'}… (${qrSentCount})`
                          : pending === 0
                            ? (isDe ? 'Erledigt' : 'Done')
                            : (isDe ? `An ${pending} senden` : `Send to ${pending}`)}
                      </button>)}
                  </>
                );
              })()}
            </div>

            {/* Erklärendes auf Abruf — beim Öffnen soll man handeln können,
                nicht erst lesen müssen. */}
            <div style={{ marginTop: 18, display: 'flex', flexDirection: 'column', gap: 6 }}>
              <button type="button" onClick={() => setQrHelpOpen(v => !v)} style={qrDisclosureStyle}>
                {qrHelpOpen ? '▾ ' : '▸ '}{isDe ? 'Wie lange dauert der Versand?' : 'How long does sending take?'}
              </button>
              {qrHelpOpen && (
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, padding: '2px 0 6px 14px' }}>
                  {isDe
                    ? <>Die Mails gehen <strong>einzeln nacheinander</strong> raus. Bei mehr als 100 Personen kann das <strong>über 10 Minuten</strong> dauern — das ist normal. Ihren Code sehen Teilnehmer ohnehin jederzeit in der App unter &bdquo;Meine Events&ldquo;.</>
                    : <>Emails are sent <strong>one by one</strong>. With more than 100 people this can take <strong>over 10 minutes</strong> — that is normal. Attendees can see their code anytime in the app under &bdquo;My Events&ldquo;.</>}
                </div>
              )}

              <button type="button" disabled={isSendingQR} onClick={() => { openQrMailEditor().catch(() => { /* */ }); }} style={qrDisclosureStyle}>
                {isDe
                  ? `✎ Mail-Text anpassen${getQrMailOverride(selectedEvent) ? ' (angepasst)' : ''}`
                  : `✎ Customize email text${getQrMailOverride(selectedEvent) ? ' (customized)' : ''}`}
              </button>

              {(() => {
                const kids = selectedEvent ? childEventsOf(selectedEvent.id) : [];
                if (kids.length === 0) return null;
                const term = (selectedEvent && selectedEvent.childEventTermPlural) || (isDe ? 'Sub-Events' : 'sub-events');
                return (
                  <>
                    <button type="button" onClick={() => setQrSubMailsOpen(v => !v)} style={qrDisclosureStyle}>
                      {qrSubMailsOpen ? '▾ ' : '▸ '}{isDe ? `Mail-Texte der ${term} einzeln anpassen` : `Customize the ${term} emails individually`}
                    </button>
                    {qrSubMailsOpen && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, padding: '2px 0 6px 14px' }}>
                        {kids.map(ce => (
                          <button
                            key={ce.id}
                            type="button"
                            className="btn btn-outline"
                            disabled={isSendingQR}
                            onClick={() => { openQrMailEditor(ce).catch(() => { /* */ }); }}
                            style={{ fontSize: '0.76rem', padding: '4px 10px', display: 'inline-flex', alignItems: 'center', gap: 6, justifyContent: 'flex-start', textAlign: 'left' }}
                          >
                            <Pencil size={12} />
                            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ce.title || (isDe ? 'Ohne Titel' : 'Untitled')}</span>
                            {getQrMailOverride(ce) && (
                              <span style={{ flexShrink: 0, fontSize: '0.66rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? '(angepasst)' : '(customized)'}</span>
                            )}
                          </button>
                        ))}
                      </div>
                    )}
                  </>
                );
              })()}
            </div>

            {/* Bleibt sichtbar statt hinter einem Aufklapper: Wer selbst nicht
                angemeldet ist, wundert sich sonst beim Test-Scan. */}
            {(() => {
              const orgEmail = (currentUser.email || '').toLowerCase();
              const isOrgRegistered = !!orgEmail && registrations.some(r => (r.ParticipantEmail || '').toLowerCase() === orgEmail && (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt'));
              if (isOrgRegistered) return null;
              return (
                <div style={{ marginTop: 12, fontSize: '0.76rem', color: '#7a5a00', background: '#fff8e1', border: '1px solid #f5b400', borderRadius: 6, padding: '7px 10px', lineHeight: 1.4 }}>
                  {isDe
                    ? 'Du bist selbst nicht angemeldet — die Test-Mail kommt an, aber ein späterer Check-in-Scan findet dich nicht in der Liste.'
                    : 'You are not registered yourself — the test email arrives, but a later check-in scan will not find you in the list.'}
                </div>
              );
            })()}

            {qrSendResult && (
              <div style={{
                padding: '8px 12px', marginBottom: 12, borderRadius: 6,
                background: qrSendResult.startsWith('Fehler') ? 'var(--dex-red-light, #ffe5e5)' : 'var(--dex-green-light, #f0f8e8)',
                fontSize: '0.85rem', color: qrSendResult.startsWith('Fehler') ? 'var(--dex-red-dark, #b00)' : 'var(--dex-green-dark)',
              }}>{qrSendResult}</div>
            )}

            <div style={{ marginTop: 16, textAlign: 'right' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setQrSendModalOpen(false)}
                disabled={isSendingQR}
                style={{ fontSize: '0.85rem' }}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            </div>
        </Modal>
  );
};

