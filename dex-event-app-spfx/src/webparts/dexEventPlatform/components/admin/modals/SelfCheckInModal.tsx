/* SelfCheckInModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13417-13522 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Neu gegliedert nach dem UI-Leitfaden — Kopf/Fuß über die Modal-Props,
 * drei Fragen in der Reihenfolge, in der ein Organizer sie beantwortet:
 * Wann gilt der Self-Check-in? → Wie hängst du den Code aus? → Live-Anzeige.
 * Vorher stand der QR-Code oben und das Zeitfenster darunter — wer nur das
 * PDF zog, sah nie, dass ohne Speichern der Standard (2 h vor Start bis
 * Event-Ende) gilt. Die Knöpfe sitzen jetzt in den nummerierten Schritt-Zeilen.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { QrCode, Download, ExternalLink } from '../../Icons';
import { downloadSelfCheckInPdf } from '../../../utils/selfCheckInPdf';
import { DeloitteEvent } from '../../../types';

export interface SelfCheckInModalProps {
  isDe: boolean;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  saveSelfCheckInWindow: () => Promise<void>;
  sciBusy: boolean;
  sciFrom: string;
  sciModalOpen: boolean;
  sciModalQr: string;
  sciSaveMsg: string;
  sciTo: string;
  sciToken: string;
  selectedEvent: DeloitteEvent;
  setSciFrom: React.Dispatch<React.SetStateAction<string>>;
  setSciModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setSciTo: React.Dispatch<React.SetStateAction<string>>;
}

export const SelfCheckInModal: React.FC<SelfCheckInModalProps> = (p) => {
  const { isDe, navigate, saveSelfCheckInWindow, sciBusy, sciFrom, sciModalOpen, sciModalQr, sciSaveMsg, sciTo, sciToken, selectedEvent, setSciFrom, setSciModalOpen, setSciTo } = p;
  const saveFailed = sciSaveMsg.indexOf('fehlgeschlagen') >= 0 || sciSaveMsg.indexOf('failed') >= 0;
  return (
        <Modal
          open={sciModalOpen}
          onClose={() => setSciModalOpen(false)}
          dismissable={!sciBusy}
          maxWidth={560}
          padding={24}
          ariaLabel="Self-Check-in"
          icon={<QrCode size={20} />}
          title={isDe ? 'Self-Check-in einrichten' : 'Set up self check-in'}
          subtitle={isDe
            ? 'Teilnehmer scannen einen QR-Code am Eingang mit der Kamera ihres Firmenhandys und checken sich selbst ein — ganz ohne Scanner-Team. Jede Person kann nur sich selbst einchecken (Login-gebunden).'
            : 'Attendees scan a QR code at the entrance with their company phone camera and check themselves in — no scanner team needed. Each person can only check in themselves (login-bound).'}
          footer={
            <button type="button" className="btn btn-secondary" onClick={() => setSciModalOpen(false)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          }
        >
          {/* Zeitfenster: Von/Bis — verhindert verfrühte UND nachträgliche Check-ins. */}
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Wann gilt der Self-Check-in?' : 'When is self check-in open?'}</div>
            <p className="dex-ui-section-desc">
              {isDe
                ? 'Davor und danach ist kein Check-in möglich — auch kein nachträglicher. Vorbelegt ist der Standard: 2 Stunden vor Event-Start bis Event-Ende; er gilt, solange du nichts anderes speicherst.'
                : 'Before and after, no check-in is possible — including late ones. Prefilled with the default: 2 hours before event start until event end; it applies as long as you do not save anything else.'}
            </p>
            <div className="dex-ui-grid-2">
              <label className="dex-ui-field" style={{ display: 'block' }}>
                <span className="dex-ui-label">{isDe ? 'Von' : 'From'}</span>
                <input type="datetime-local" value={sciFrom} onChange={e => setSciFrom(e.target.value)} className="dex-ui-input" />
              </label>
              <label className="dex-ui-field" style={{ display: 'block' }}>
                <span className="dex-ui-label">{isDe ? 'Bis' : 'Until'}</span>
                <input type="datetime-local" value={sciTo} onChange={e => setSciTo(e.target.value)} className="dex-ui-input" />
              </label>
            </div>
            <div className={cx('dex-ui-step', sciBusy && 'is-pending')} style={{ marginTop: 12 }}>
              <span className="dex-ui-step-num">1</span>
              <div className="dex-ui-step-body">
                <div className="dex-ui-step-title">{isDe ? 'Zeitfenster speichern' : 'Save time window'}</div>
                <div className="dex-ui-step-hint">{isDe ? 'Gilt sofort für alle Scans dieses Events.' : 'Applies immediately to all scans for this event.'}</div>
              </div>
              <div className="dex-ui-step-action">
                <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={sciBusy} onClick={() => { saveSelfCheckInWindow().catch(() => { /* */ }); }}>
                  {sciBusy ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
                </button>
              </div>
            </div>
            {sciSaveMsg && (
              <div className={cx('dex-ui-callout', saveFailed ? 'dex-ui-callout--danger' : 'dex-ui-callout--success')} role="status" style={{ marginTop: 10 }}>
                {sciSaveMsg}
              </div>
            )}
          </div>

          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Wie hängst du den Code aus?' : 'How do you post the code?'}</div>
            {sciModalQr ? (
              <div style={{ textAlign: 'center', marginBottom: 12 }}>
                <img src={sciModalQr} alt="Self-Check-in QR" style={{ width: 200, maxWidth: '70%', height: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 10, background: '#fff' }} />
              </div>
            ) : (
              <div className="dex-ui-empty" style={{ marginBottom: 12 }}>
                <div className="dex-ui-empty-icon"><QrCode size={20} /></div>
                <div className="dex-ui-empty-title">{isDe ? 'QR-Code konnte nicht erzeugt werden.' : 'QR code could not be generated.'}</div>
              </div>
            )}
            <div className="dex-ui-step">
              <span className="dex-ui-step-num">2</span>
              <div className="dex-ui-step-body">
                <div className="dex-ui-step-title">{isDe ? 'QR-PDF drucken und aushängen' : 'Print the QR PDF and post it'}</div>
                <div className="dex-ui-step-hint">{isDe ? 'Ein Blatt mit Code, Event-Titel, Datum und Ort — für Eingang, Tisch oder Stellwand.' : 'One sheet with code, event title, date and location — for the entrance, a table or a display board.'}</div>
              </div>
              <div className="dex-ui-step-action">
                <button
                  type="button"
                  className="btn btn-primary dex-ui-btn-sm"
                  disabled={sciBusy}
                  onClick={() => {
                    (async () => {
                      await downloadSelfCheckInPdf({
                        eventTitle: selectedEvent.title || 'Event',
                        eventDateLabel: selectedEvent.startDate ? new Date(selectedEvent.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '',
                        locationLabel: selectedEvent.location || '',
                        token: sciToken,
                      });
                    })().catch(() => { /* best-effort */ });
                  }}
                >
                  <Download size={14} />{isDe ? 'PDF herunterladen' : 'Download PDF'}
                </button>
              </div>
            </div>
          </div>

          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Live-Anzeige (optional)' : 'Live display (optional)'}</div>
            <div className="dex-ui-step">
              <span className="dex-ui-step-num">3</span>
              <div className="dex-ui-step-body">
                <div className="dex-ui-step-title">{isDe ? 'Rotierenden QR-Code auf einem Bildschirm zeigen' : 'Show a rotating QR code on a screen'}</div>
                <div className="dex-ui-step-hint">{isDe ? 'Für Beamer oder Monitor am Eingang; der Code wechselt laufend. Öffnet die Live-Seite und schließt diesen Dialog.' : 'For a projector or monitor at the entrance; the code keeps changing. Opens the live page and closes this dialog.'}</div>
              </div>
              <div className="dex-ui-step-action">
                <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={sciBusy} onClick={() => { setSciModalOpen(false); navigate('self-checkin-display', selectedEvent.id); }}>
                  <ExternalLink size={14} />{isDe ? 'Live-QR öffnen' : 'Open live QR'}
                </button>
              </div>
            </div>
          </div>
        </Modal>
  );
};
