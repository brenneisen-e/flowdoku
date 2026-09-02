/* CheckInHubModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13525-13623 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { hubChoiceStyle } from '../../admin/adminStyles';
import { QrCode, Send } from '../../Icons';
import { DeloitteEvent } from '../../../types';

export interface CheckInHubModalProps {
  checkInHubOpen: boolean;
  checkInHubStep: "choose" | "checkin";
  isDe: boolean;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  openSelfCheckInModal: () => Promise<void>;
  sciBusy: boolean;
  selectedEvent: DeloitteEvent;
  setCheckInHubOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setCheckInHubStep: React.Dispatch<React.SetStateAction<"choose" | "checkin">>;
  setQrSendModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setQrSendResult: React.Dispatch<React.SetStateAction<string>>;
}

export const CheckInHubModal: React.FC<CheckInHubModalProps> = (p) => {
  const { checkInHubOpen, checkInHubStep, isDe, navigate, openSelfCheckInModal, sciBusy, selectedEvent, setCheckInHubOpen, setCheckInHubStep, setQrSendModalOpen, setQrSendResult } = p;
  return (
        <Modal
          open={checkInHubOpen}
          onClose={() => setCheckInHubOpen(false)}
          maxWidth={640}
          padding={24}
          ariaLabel={isDe ? 'QR-Codes und Check-In' : 'QR codes and check-in'}
        >
          <h3 style={{ margin: '0 0 4px', fontSize: '1.15rem' }}>
            {checkInHubStep === 'choose'
              ? (isDe ? 'QR-Codes und Check-In' : 'QR codes and check-in')
              : (isDe ? 'Check-In am Event-Tag' : 'Check-in on event day')}
          </h3>
          <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {checkInHubStep === 'choose'
              ? (isDe ? 'Was möchtest du tun?' : 'What would you like to do?')
              : (isDe ? 'Wer scannt?' : 'Who scans?')}
          </p>

          {checkInHubStep === 'choose' ? (
            <div style={{ display: 'grid', gap: 12 }}>
              <button
                type="button"
                className="card"
                onClick={() => { setCheckInHubOpen(false); setQrSendResult(null); setQrSendModalOpen(true); }}
                style={hubChoiceStyle}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6 }}>
                  <span style={{ color: 'var(--dex-green)', display: 'inline-flex' }}><Send size={18} /></span>
                  {/* v30.61: „verschicken" las sich wie ein Auslöser — als ginge
                      mit dem Klick sofort alles raus. Es öffnet aber nur die
                      Einrichtung: Text anpassen, Vorschau, Test, und ERST DANN
                      der Versand. Die Beschriftung sagt das jetzt, statt es zu
                      verschweigen und auf Mut zu hoffen. */}
                  <span style={{ fontWeight: 700 }}>{isDe ? 'QR-Code-Versand einrichten' : 'Set up the QR code send-out'}</span>
                </div>
                <span style={{ fontSize: '0.83rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Hier wird noch nichts verschickt. Du stellst zuerst Text und Bild ein, siehst die Vorschau und kannst dir einen Test schicken — der Versand an alle ist der letzte Schritt und wird eigens bestätigt.'
                    : 'Nothing is sent yet. You first set the copy and image, see the preview and can send yourself a test — sending to everyone is the final step and is confirmed separately.'}
                </span>
              </button>
              <button
                type="button"
                className="card"
                onClick={() => setCheckInHubStep('checkin')}
                style={hubChoiceStyle}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6 }}>
                  <span style={{ color: 'var(--dex-green)', display: 'inline-flex' }}><QrCode size={18} /></span>
                  <span style={{ fontWeight: 700 }}>{isDe ? 'Check-In am Event-Tag' : 'Check-in on event day'}</span>
                </div>
                <span style={{ fontSize: '0.83rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Einlass öffnen — entweder dein Team scannt die Codes, oder die Teilnehmer checken sich selbst ein.'
                    : 'Open the entrance — either your team scans the codes, or attendees check themselves in.'}
                </span>
              </button>
            </div>
          ) : (
            <div style={{ display: 'grid', gap: 12 }}>
              <button
                type="button"
                className="card"
                onClick={() => { setCheckInHubOpen(false); navigate('check-in', selectedEvent.id); }}
                style={hubChoiceStyle}
              >
                <div style={{ fontWeight: 700, marginBottom: 6 }}>{isDe ? 'Unser Team scannt' : 'Our team scans'}</div>
                <span style={{ fontSize: '0.83rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Öffnet das Check-in-Werkzeug: scannen, per Teilnehmer-ID oder Namen einchecken, Live-Zahlen im Blick.'
                    : 'Opens the check-in tool: scan, check in by attendee ID or name, watch the live counts.'}
                </span>
              </button>
              <button
                type="button"
                className="card"
                onClick={() => { setCheckInHubOpen(false); openSelfCheckInModal().catch(() => { /* best-effort */ }); }}
                style={hubChoiceStyle}
                disabled={sciBusy}
              >
                <div style={{ fontWeight: 700, marginBottom: 6 }}>{isDe ? 'Teilnehmer checken sich selbst ein' : 'Attendees check themselves in'}</div>
                <span style={{ fontSize: '0.83rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Ein Event-QR am Eingang, den die Teilnehmer mit der Handy-Kamera scannen — ohne Scanner-Team. Dahinter: Zeitfenster einstellen, QR-PDF zum Aushängen und die rotierende Live-Anzeige.'
                    : 'One event QR at the entrance that attendees scan with their phone camera — no scanner team. Behind it: set the time window, get the printable QR PDF and the rotating live display.'}
                </span>
              </button>
              <button
                type="button"
                onClick={() => setCheckInHubStep('choose')}
                style={{ background: 'none', border: 'none', color: 'var(--dex-gray-600)', fontSize: '0.85rem', cursor: 'pointer', justifySelf: 'start', padding: 4 }}
              >
                {isDe ? '← Zurück' : '← Back'}
              </button>
            </div>
          )}
        </Modal>
  );
};

