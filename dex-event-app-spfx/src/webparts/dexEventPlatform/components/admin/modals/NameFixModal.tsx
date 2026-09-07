/* NameFixModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13318-13360 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Kopf/Fuß aus `Modal`, Fortschritt als Karte mit Zähler-Pill und
 * Balken, Ergebnis als Hinweiskasten — dieselbe Gestalt wie AccessFixModal,
 * damit beide Reparatur-Dialoge im Organizer Center gleich zu lesen sind.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Wrench } from '../../Icons';

export interface NameFixModalProps {
  isDe: boolean;
  nameFixModal: { running: boolean; step: string; evIdx: number; evTotal: number; summary: string[]; };
  setNameFixModal: React.Dispatch<React.SetStateAction<{ running: boolean; step: string; evIdx: number; evTotal: number; summary: string[]; }>>;
}

export const NameFixModal: React.FC<NameFixModalProps> = (p) => {
  const { isDe, nameFixModal, setNameFixModal } = p;
  const running = nameFixModal.running;
  const pct = Math.min(100, Math.round((nameFixModal.evIdx / Math.max(1, nameFixModal.evTotal)) * 100));
  const lines = nameFixModal.summary || [];
  return (
        <Modal
          open={true}
          onClose={() => { if (!nameFixModal.running) setNameFixModal(null); }}
          dismissable={!running}
          maxWidth={520}
          ariaLabel={isDe ? 'Namen reparieren' : 'Repair names'}
          title={isDe ? 'Login-Tokens in Namen reparieren' : 'Repair login tokens in names'}
          subtitle={running
            ? (isDe ? 'Anzeigenamen, die noch eine Login-Kennung tragen, werden nachgeschlagen und ersetzt.' : 'Display names that still carry a login token are looked up and replaced.')
            : (isDe ? 'Fertig — das wurde geprüft und korrigiert:' : 'Done — this is what was checked and fixed:')}
          icon={<Wrench size={20} />}
          footer={running ? undefined : (
            <button type="button" className="btn btn-primary" onClick={() => setNameFixModal(null)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          )}
        >
          {running ? (
            <div className="dex-ui-card dex-ui-card--soft dex-ui-stack">
              <div className="dex-ui-inline" style={{ justifyContent: 'space-between' }}>
                <strong style={{ fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>{nameFixModal.step || '…'}</strong>
                <span className="dex-ui-pill dex-ui-pill--gray">
                  {nameFixModal.evTotal > 0 ? `Event ${nameFixModal.evIdx}/${nameFixModal.evTotal}` : (isDe ? 'Wird geprüft…' : 'Checking…')}
                </span>
              </div>
              <div role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={pct}
                style={{ background: 'var(--dex-gray-200, #e8e8e8)', borderRadius: 999, height: 8, overflow: 'hidden' }}>
                <div style={{ width: `${pct}%`, height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease' }} />
              </div>
              <p className="dex-ui-muted" style={{ margin: 0 }}>
                {isDe ? 'Lass das Fenster offen, bis die Reparatur fertig ist.' : 'Keep this window open until the repair is done.'}
              </p>
            </div>
          ) : lines.length === 0 ? (
            <div className="dex-ui-empty">{isDe ? 'Keine Rückmeldung aus dem Lauf.' : 'No result reported from the run.'}</div>
          ) : (
            <div className="dex-ui-callout dex-ui-callout--neutral">
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: '0.88rem', lineHeight: 1.7, color: 'var(--dex-gray-800)' }}>
                {lines.map(l => <li key={l}>{l}</li>)}
              </ul>
            </div>
          )}
        </Modal>
  );
};
