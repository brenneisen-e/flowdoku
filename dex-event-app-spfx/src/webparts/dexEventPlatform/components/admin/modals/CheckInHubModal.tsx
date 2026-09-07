/* CheckInHubModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13525-13623 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Optik auf die gemeinsamen dex-ui-Klassen umgestellt (Leitfaden
 * docs/ui-leitfaden.md). Die Kacheln sind `dex-ui-choice` mit Symbol — sie
 * hatten vorher keinen Hover und lasen sich wie Beschriftungen. Kopf und Fuß
 * kommen aus `Modal` (title/subtitle/icon/footer); `hubChoiceStyle` aus
 * adminStyles wird hier nicht mehr gebraucht. Verhalten, Handler und
 * Reihenfolge der beiden Stufen sind unverändert.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Check, ChevronLeft, QrCode, Send, Users } from '../../Icons';
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

/** v31.2: Eine Auswahl-Kachel — Symbol, Titel, eine Zeile „was dann passiert". */
const HubChoice: React.FC<{ icon: React.ReactNode; title: string; desc: string; onClick: () => void; disabled?: boolean }> = ({ icon, title, desc, onClick, disabled }) => (
  <button type="button" className="dex-ui-choice" onClick={onClick} disabled={disabled}>
    <span className="dex-ui-choice-icon" aria-hidden="true">{icon}</span>
    <span className="dex-ui-choice-body">
      <span className="dex-ui-choice-title" style={{ display: 'block' }}>{title}</span>
      <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{desc}</span>
    </span>
  </button>
);

export const CheckInHubModal: React.FC<CheckInHubModalProps> = (p) => {
  const { checkInHubOpen, checkInHubStep, isDe, navigate, openSelfCheckInModal, sciBusy, selectedEvent, setCheckInHubOpen, setCheckInHubStep, setQrSendModalOpen, setQrSendResult } = p;
  const isChoose = checkInHubStep === 'choose';
  return (
        <Modal
          open={checkInHubOpen}
          onClose={() => setCheckInHubOpen(false)}
          maxWidth={640}
          padding={24}
          ariaLabel={isDe ? 'QR-Codes und Check-In' : 'QR codes and check-in'}
          icon={<QrCode size={20} />}
          title={isChoose
            ? (isDe ? 'QR-Codes und Check-In' : 'QR codes and check-in')
            : (isDe ? 'Check-In am Event-Tag' : 'Check-in on event day')}
          subtitle={isChoose
            ? (isDe ? 'Was möchtest du tun?' : 'What would you like to do?')
            : (isDe ? 'Wer scannt die Codes am Eingang?' : 'Who scans the codes at the entrance?')}
          // v31.2: „Zurück" steht im Fuß statt als loser Textknopf unter den
          // Kacheln — so ist die Navigation von den Aktionen getrennt.
          footer={!isChoose && (
            <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginRight: 'auto' }} onClick={() => setCheckInHubStep('choose')}>
              <ChevronLeft size={16} />{isDe ? 'Zurück' : 'Back'}
            </button>
          )}
        >
          {/* key: beim Stufenwechsel blendet der Inhalt neu ein, statt zu springen. */}
          {isChoose ? (
            <div key="choose" className="dex-ui-stack dex-ui-fade-in">
              {/* v30.61: „verschicken" las sich wie ein Auslöser — als ginge
                  mit dem Klick sofort alles raus. Es öffnet aber nur die
                  Einrichtung: Text anpassen, Vorschau, Test, und ERST DANN
                  der Versand. Die Beschriftung sagt das jetzt, statt es zu
                  verschweigen und auf Mut zu hoffen. */}
              <HubChoice
                icon={<Send size={18} />}
                title={isDe ? 'QR-Code-Versand einrichten' : 'Set up the QR code send-out'}
                desc={isDe
                  ? 'Hier wird noch nichts verschickt: Du stellst Text und Bild ein, siehst die Vorschau und schickst dir einen Test. Der Versand an alle ist der letzte Schritt und wird eigens bestätigt.'
                  : 'Nothing is sent yet: you set the copy and image, see the preview and send yourself a test. Sending to everyone is the final step and is confirmed separately.'}
                onClick={() => { setCheckInHubOpen(false); setQrSendResult(null); setQrSendModalOpen(true); }}
              />
              <HubChoice
                icon={<Check size={18} />}
                title={isDe ? 'Check-In am Event-Tag' : 'Check-in on event day'}
                desc={isDe
                  ? 'Einlass öffnen — entweder dein Team scannt die Codes, oder die Teilnehmer checken sich selbst ein. Du wählst im nächsten Schritt.'
                  : 'Open the entrance — either your team scans the codes, or attendees check themselves in. You choose in the next step.'}
                onClick={() => setCheckInHubStep('checkin')}
              />
            </div>
          ) : (
            <div key="checkin" className="dex-ui-stack dex-ui-fade-in">
              <HubChoice
                icon={<Users size={18} />}
                title={isDe ? 'Unser Team scannt' : 'Our team scans'}
                desc={isDe
                  ? 'Öffnet das Check-in-Werkzeug: scannen, per Teilnehmer-ID oder Namen einchecken, Live-Zahlen im Blick.'
                  : 'Opens the check-in tool: scan, check in by attendee ID or name, watch the live counts.'}
                onClick={() => { setCheckInHubOpen(false); navigate('check-in', selectedEvent.id); }}
              />
              <HubChoice
                icon={<QrCode size={18} />}
                title={isDe ? 'Teilnehmer checken sich selbst ein' : 'Attendees check themselves in'}
                desc={isDe
                  ? 'Ein Event-QR am Eingang, den die Teilnehmer mit der Handy-Kamera scannen — ohne Scanner-Team. Dort stellst du das Zeitfenster ein, holst das QR-PDF zum Aushängen und die rotierende Live-Anzeige.'
                  : 'One event QR at the entrance that attendees scan with their phone camera — no scanner team. There you set the time window, get the printable QR PDF and the rotating live display.'}
                onClick={() => { setCheckInHubOpen(false); openSelfCheckInModal().catch(() => { /* best-effort */ }); }}
                disabled={sciBusy}
              />
            </div>
          )}
        </Modal>
  );
};
