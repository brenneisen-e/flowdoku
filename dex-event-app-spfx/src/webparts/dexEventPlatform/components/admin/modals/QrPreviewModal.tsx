/* QrPreviewModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14780-14828 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Kopf (Titel, Untertitel, Schließen-X) und Fuß kommen aus `Modal`;
 * der eigene Kopf mit zweitem X entfällt. Texte zweisprachig — bis v31.1
 * war der Dialog fest deutsch. Der Betreff steht als eigene Zeile über der
 * Vorschau, weil er im Postfach das Erste ist, was die Person liest.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { QrCode } from '../../Icons';

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
          ariaLabel={isDe ? 'Vorschau: QR-Code-Mail' : 'Preview: QR code email'}
          title={isDe ? 'Vorschau: QR-Code-Mail' : 'Preview: QR code email'}
          subtitle={isDe
            ? 'So sieht die Mail aus, die jede angemeldete Person bekommt — der QR-Code hier ist auf dich ausgestellt.'
            : 'This is the email every registered person receives — the QR code shown here is issued to you.'}
          icon={<QrCode size={20} />}
          footer={(
            <button type="button" className="btn btn-primary" onClick={() => setQrPreviewOpen(false)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          )}
        >
            <div className="dex-ui-inline" style={{ fontSize: '0.85rem' }}>
              <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'Betreff' : 'Subject'}</span>
              <strong style={{ color: 'var(--dex-gray-800)', minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis' }} title={qrPreviewSubject}>
                {qrPreviewSubject}
              </strong>
            </div>
            <div className="dex-ui-card dex-ui-card--soft" style={{ padding: 10 }}>
              <iframe
                title={isDe ? 'QR-Code-Mail-Vorschau' : 'QR code email preview'}
                srcDoc={qrPreviewHtml}
                sandbox=""
                style={{ display: 'block', width: '100%', minHeight: 480, border: 'none', borderRadius: 8, background: '#fff' }}
              />
            </div>
        </Modal>
  );
};
