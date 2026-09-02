/* SelfCheckInModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13417-13522 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
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
  return (
        <Modal
          open={sciModalOpen}
          onClose={() => setSciModalOpen(false)}
          dismissable={!sciBusy}
          maxWidth={560}
          padding={24}
          ariaLabel="Self-Check-in"
        >
          <h3 style={{ margin: '0 0 8px', fontSize: '1.05rem' }}>
            {isDe ? 'Self-Check-in — QR-Code' : 'Self check-in — QR code'}
          </h3>
          <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? 'Diesen QR-Code kannst du am Eingang aushängen oder auf einem Bildschirm zeigen. Teilnehmer scannen ihn mit der Kamera ihres Firmenhandys und checken sich damit selbst ein — ganz ohne Scanner-Team. Jede Person kann nur sich selbst einchecken (Login-gebunden).'
              : 'Post this QR code at the entrance or show it on a screen. Attendees scan it with their company phone camera and check themselves in — no scanner team needed. Each person can only check in themselves (login-bound).'}
          </p>
          {sciModalQr ? (
            <div style={{ textAlign: 'center', margin: '0 0 14px' }}>
              <img src={sciModalQr} alt="Self-Check-in QR" style={{ width: 260, maxWidth: '80%', height: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 10, background: '#fff' }} />
            </div>
          ) : (
            <p style={{ textAlign: 'center', color: 'var(--dex-gray-400)', fontSize: '0.85rem' }}>
              {isDe ? 'QR-Code konnte nicht erzeugt werden.' : 'QR code could not be generated.'}
            </p>
          )}
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'center', marginBottom: 18 }}>
            <button
              className="btn btn-primary"
              disabled={sciBusy}
              style={{ fontSize: '0.88rem', padding: '10px 18px' }}
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
              {isDe ? 'QR-PDF herunterladen (drucken)' : 'Download QR PDF (print)'}
            </button>
            <button
              className="btn btn-secondary"
              disabled={sciBusy}
              style={{ fontSize: '0.88rem', padding: '10px 18px' }}
              onClick={() => { setSciModalOpen(false); navigate('self-checkin-display', selectedEvent.id); }}
            >
              {isDe ? 'Live-QR anzeigen (rotierend)' : 'Show live QR (rotating)'}
            </button>
          </div>
          {/* Zeitfenster: Von/Bis — verhindert verfrühte UND nachträgliche Check-ins. */}
          <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '12px 14px' }}>
            <div style={{ fontWeight: 700, fontSize: '0.88rem', marginBottom: 4 }}>
              {isDe ? 'Check-in-Zeitfenster' : 'Check-in time window'}
            </div>
            <p style={{ margin: '0 0 10px', fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? 'Von wann bis wann der Self-Check-in möglich ist. Vor „Von" und nach „Bis" sind keine Check-ins möglich — also auch keine nachträglichen. Vorbelegt mit dem Standard: 2 Stunden vor Event-Start bis Event-Ende (gilt auch, solange du nichts anderes speicherst).'
                : 'From when until when self check-in is possible. Before "from" and after "until" no check-ins are possible — including late ones. Prefilled with the default: 2 hours before event start until event end (which also applies as long as you do not save anything else).'}
            </p>
            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
              <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                {isDe ? 'Von' : 'From'}
                <input
                  type="datetime-local"
                  value={sciFrom}
                  onChange={e => setSciFrom(e.target.value)}
                  className="form-input"
                  style={{ padding: '8px 10px', fontSize: '0.85rem' }}
                />
              </label>
              <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                {isDe ? 'Bis' : 'Until'}
                <input
                  type="datetime-local"
                  value={sciTo}
                  onChange={e => setSciTo(e.target.value)}
                  className="form-input"
                  style={{ padding: '8px 10px', fontSize: '0.85rem' }}
                />
              </label>
              <button
                className="btn btn-secondary"
                disabled={sciBusy}
                style={{ fontSize: '0.85rem', padding: '9px 16px' }}
                onClick={() => { saveSelfCheckInWindow().catch(() => { /* */ }); }}
              >
                {sciBusy ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Zeitfenster speichern' : 'Save time window')}
              </button>
            </div>
            {sciSaveMsg && (
              <p style={{ margin: '8px 0 0', fontSize: '0.78rem', color: sciSaveMsg.indexOf('fehlgeschlagen') >= 0 || sciSaveMsg.indexOf('failed') >= 0 ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)' }}>
                {sciSaveMsg}
              </p>
            )}
          </div>
          <div style={{ textAlign: 'right', marginTop: 14 }}>
            <button className="btn btn-secondary" onClick={() => setSciModalOpen(false)} style={{ fontSize: '0.85rem' }}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </Modal>
  );
};

