/* ReorderProgressOverlay — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16329-16357 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';

export interface ReorderProgressOverlayProps {
  reorderProgress: number;
  reorderProgressLabel: string;
}

export const ReorderProgressOverlay: React.FC<ReorderProgressOverlayProps> = (p) => {
  const { reorderProgress, reorderProgressLabel } = p;
  return (
        <Modal
          open={reorderProgress !== null}
          onClose={() => { /* progress overlay — nicht schließbar */ }}
          dismissable={false}
          maxWidth={420}
          padding={28}
          ariaLabel="ID-Neuvergabe"
        >
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontWeight: 600, marginBottom: 14 }}>{reorderProgressLabel || 'IDs werden neu vergeben…'}</div>
            <div style={{ height: 12, borderRadius: 6, background: 'var(--dex-gray-200, #e5e5e5)', overflow: 'hidden' }}>
              <div
                style={{
                  height: '100%', width: `${reorderProgress}%`,
                  background: 'var(--dex-green, #86bc25)', borderRadius: 6,
                  transition: 'width 0.25s ease',
                }}
              />
            </div>
            <div style={{ marginTop: 10, fontSize: '1.1rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {reorderProgress}%
            </div>
            <div style={{ marginTop: 6, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              Bitte warten — das Fenster nicht schließen.
            </div>
          </div>
        </Modal>
  );
};

