/* QrPreviewModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14780-14828 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { X } from '../../Icons';

export interface QrPreviewModalProps {
  isDe: boolean;
  qrPreviewHtml: string;
  qrPreviewOpen: boolean;
  qrPreviewSubject: string;
  setQrPreviewOpen: React.Dispatch<React.SetStateAction<boolean>>;
}

export const QrPreviewModal: React.FC<QrPreviewModalProps> = (p) => {
  const { isDe, qrPreviewHtml, qrPreviewOpen, qrPreviewSubject, setQrPreviewOpen } = p;
  return (
        <Modal
          open={qrPreviewOpen}
          onClose={() => setQrPreviewOpen(false)}
          maxWidth={720}
          padding={0}
          ariaLabel="Vorschau: QR-Code-Mail"
        >
            <div style={{
              padding: '14px 18px', borderBottom: '1px solid var(--dex-gray-200)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
            }}>
              <div style={{ minWidth: 0, flex: 1 }}>
                <h3 style={{ margin: 0, fontSize: '1.05rem' }}>Vorschau: QR-Code-Mail</h3>
                <p style={{ margin: '2px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  So sieht die Mail aus, die jeder angemeldete Teilnehmer bekommt — der QR-Code in der Vorschau ist auf dich ausgestellt.
                </p>
                <p style={{ margin: '6px 0 0', fontSize: '0.8rem', color: 'var(--dex-gray-700)' }}>
                  <strong>Betreff:</strong> <span style={{ color: 'var(--dex-gray-600)' }}>{qrPreviewSubject}</span>
                </p>
              </div>
              <button
                type="button"
                onClick={() => setQrPreviewOpen(false)}
                style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', padding: 4 }}
                aria-label="Schließen"
              >
                <X size={22} />
              </button>
            </div>
            <div style={{ flex: 1, overflow: 'hidden', background: '#f5f5f5', padding: 12 }}>
              <iframe
                title={isDe ? 'QR-Code-Mail-Vorschau' : 'QR code email preview'}
                srcDoc={qrPreviewHtml}
                sandbox=""
                style={{ width: '100%', height: '100%', minHeight: 480, border: 'none', borderRadius: 6, background: '#fff' }}
              />
            </div>
            <div style={{ padding: '12px 18px', borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'flex-end' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setQrPreviewOpen(false)}
                style={{ fontSize: '0.85rem' }}
              >
                Schließen
              </button>
            </div>
        </Modal>
  );
};

