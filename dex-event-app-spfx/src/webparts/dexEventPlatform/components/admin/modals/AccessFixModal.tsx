/* AccessFixModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13362-13413 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';

export interface AccessFixModalProps {
  accessFixModal: { running: boolean; evIdx: number; evTotal: number; evTitle: string; itemDone: number; itemTotal: number; summary: string[]; };
  isDe: boolean;
  setAccessFixModal: React.Dispatch<React.SetStateAction<{ running: boolean; evIdx: number; evTotal: number; evTitle: string; itemDone: number; itemTotal: number; summary: string[]; }>>;
}

export const AccessFixModal: React.FC<AccessFixModalProps> = (p) => {
  const { accessFixModal, isDe, setAccessFixModal } = p;
  return (
        <Modal
          open={true}
          onClose={() => { if (!accessFixModal.running) setAccessFixModal(null); }}
          dismissable={!accessFixModal.running}
          maxWidth={520}
          padding={24}
          ariaLabel={isDe ? 'Zugriffs-Prüfung' : 'Access check'}
        >
          <h3 style={{ margin: '0 0 10px', fontSize: '1.05rem' }}>
            {isDe ? 'Fremd-Anmeldungen: Zugriff reparieren' : 'Proxy registrations: repair access'}
          </h3>
          {accessFixModal.running ? (
            <>
              <p style={{ margin: '0 0 6px', fontSize: '0.9rem', color: 'var(--dex-gray-700)' }}>
                {isDe ? 'Event' : 'Event'} {accessFixModal.evIdx}/{accessFixModal.evTotal}: <strong>{accessFixModal.evTitle || '…'}</strong>
              </p>
              <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
                {accessFixModal.itemTotal > 0
                  ? (isDe
                    ? `Eintrag ${accessFixModal.itemDone}/${accessFixModal.itemTotal} wird geprüft…`
                    : `Checking item ${accessFixModal.itemDone}/${accessFixModal.itemTotal}…`)
                  : (isDe ? 'Liste wird geladen…' : 'Loading list…')}
              </p>
              {(() => {
                const evBase = Math.max(0, accessFixModal.evIdx - 1);
                const inner = accessFixModal.itemTotal > 0 ? accessFixModal.itemDone / accessFixModal.itemTotal : 0;
                const pct = Math.min(100, Math.round(((evBase + inner) / Math.max(1, accessFixModal.evTotal)) * 100));
                return (
                  <div style={{ background: 'var(--dex-gray-100, #f0f0f0)', borderRadius: 999, height: 10, overflow: 'hidden' }}>
                    <div style={{ width: `${pct}%`, height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease' }} />
                  </div>
                );
              })()}
              <p style={{ margin: '10px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                {isDe ? 'Bitte das Fenster geöffnet lassen, bis die Prüfung abgeschlossen ist.' : 'Please keep this window open until the check completes.'}
              </p>
            </>
          ) : (
            <>
              <ul style={{ margin: '0 0 14px', paddingLeft: 18, fontSize: '0.88rem', color: 'var(--dex-gray-700)', lineHeight: 1.6 }}>
                {(accessFixModal.summary || []).map((s, i) => <li key={i}>{s}</li>)}
              </ul>
              <div style={{ textAlign: 'right' }}>
                <button className="btn btn-primary" onClick={() => setAccessFixModal(null)} style={{ fontSize: '0.88rem', padding: '9px 18px' }}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
            </>
          )}
        </Modal>
  );
};

