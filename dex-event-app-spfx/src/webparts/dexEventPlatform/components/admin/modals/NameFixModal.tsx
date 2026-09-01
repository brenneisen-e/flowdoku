/* NameFixModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13318-13360 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';

export interface NameFixModalProps {
  isDe: boolean;
  nameFixModal: { running: boolean; step: string; evIdx: number; evTotal: number; summary: string[]; };
  setNameFixModal: React.Dispatch<React.SetStateAction<{ running: boolean; step: string; evIdx: number; evTotal: number; summary: string[]; }>>;
}

export const NameFixModal: React.FC<NameFixModalProps> = (p) => {
  const { isDe, nameFixModal, setNameFixModal } = p;
  return (
        <Modal
          open={true}
          onClose={() => { if (!nameFixModal.running) setNameFixModal(null); }}
          dismissable={!nameFixModal.running}
          maxWidth={520}
          padding={24}
          ariaLabel={isDe ? 'Namen reparieren' : 'Repair names'}
        >
          <h3 style={{ margin: '0 0 10px', fontSize: '1.05rem' }}>
            {isDe ? 'Login-Tokens in Namen reparieren' : 'Repair login tokens in names'}
          </h3>
          {nameFixModal.running ? (
            <>
              <p style={{ margin: '0 0 6px', fontSize: '0.9rem', color: 'var(--dex-gray-700)' }}>
                <strong>{nameFixModal.step || '…'}</strong>
              </p>
              <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
                {nameFixModal.evTotal > 0
                  ? (isDe ? `Event ${nameFixModal.evIdx}/${nameFixModal.evTotal}` : `Event ${nameFixModal.evIdx}/${nameFixModal.evTotal}`)
                  : (isDe ? 'Wird geprüft…' : 'Checking…')}
              </p>
              <div style={{ background: 'var(--dex-gray-100, #f0f0f0)', borderRadius: 999, height: 10, overflow: 'hidden' }}>
                <div style={{
                  width: `${Math.min(100, Math.round((nameFixModal.evIdx / Math.max(1, nameFixModal.evTotal)) * 100))}%`,
                  height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease',
                }} />
              </div>
            </>
          ) : (
            <>
              <ul style={{ margin: '0 0 14px', paddingLeft: 18, fontSize: '0.88rem', lineHeight: 1.7 }}>
                {(nameFixModal.summary || []).map(l => <li key={l}>{l}</li>)}
              </ul>
              <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                <button type="button" className="btn btn-primary" onClick={() => setNameFixModal(null)}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
            </>
          )}
        </Modal>
  );
};

