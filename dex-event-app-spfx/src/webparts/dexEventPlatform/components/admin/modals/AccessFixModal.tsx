/* AccessFixModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13362-13413 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Kopf und Fuß aus `Modal`, Fortschritt als Karte (Event-Zähler als
 * Pill, Titel, Balken über alle Listen), Ergebnis als Hinweiskasten — gleiche
 * Gestalt wie NameFixModal. Die Prozentrechnung ist unverändert, nur aus der
 * IIFE in Konstanten gezogen.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Users } from '../../Icons';

export interface AccessFixModalProps {
  accessFixModal: { running: boolean; evIdx: number; evTotal: number; evTitle: string; itemDone: number; itemTotal: number; summary: string[]; };
  isDe: boolean;
  setAccessFixModal: React.Dispatch<React.SetStateAction<{ running: boolean; evIdx: number; evTotal: number; evTitle: string; itemDone: number; itemTotal: number; summary: string[]; }>>;
}

export const AccessFixModal: React.FC<AccessFixModalProps> = (p) => {
  const { accessFixModal, isDe, setAccessFixModal } = p;
  const running = accessFixModal.running;
  const evBase = Math.max(0, accessFixModal.evIdx - 1);
  const inner = accessFixModal.itemTotal > 0 ? accessFixModal.itemDone / accessFixModal.itemTotal : 0;
  const pct = Math.min(100, Math.round(((evBase + inner) / Math.max(1, accessFixModal.evTotal)) * 100));
  const lines = accessFixModal.summary || [];
  return (
        <Modal
          open={true}
          onClose={() => { if (!accessFixModal.running) setAccessFixModal(null); }}
          dismissable={!running}
          maxWidth={520}
          ariaLabel={isDe ? 'Zugriffs-Prüfung' : 'Access check'}
          title={isDe ? 'Zugriff auf Fremd-Anmeldungen reparieren' : 'Repair access to proxy registrations'}
          subtitle={running
            ? (isDe ? 'Wer von jemand anderem angemeldet wurde, bekommt Zugriff auf die eigene Anmeldung; die Listen-Sicherheit wird mitgeprüft.' : 'People registered by someone else get access to their own registration; list security is checked along the way.')
            : (isDe ? 'Fertig — das Ergebnis der Prüfung:' : 'Done — the result of the check:')}
          icon={<Users size={20} />}
          footer={running ? undefined : (
            <button type="button" className="btn btn-primary" onClick={() => setAccessFixModal(null)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          )}
        >
          {running ? (
            <div className="dex-ui-card dex-ui-card--soft dex-ui-stack">
              <div className="dex-ui-inline" style={{ justifyContent: 'space-between' }}>
                <strong style={{ fontSize: '0.9rem', color: 'var(--dex-gray-800)', minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {accessFixModal.evTitle || '…'}
                </strong>
                <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'Event' : 'Event'} {accessFixModal.evIdx}/{accessFixModal.evTotal}</span>
              </div>
              <div role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={pct}
                style={{ background: 'var(--dex-gray-200, #e8e8e8)', borderRadius: 999, height: 8, overflow: 'hidden' }}>
                <div style={{ width: `${pct}%`, height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease' }} />
              </div>
              <p className="dex-ui-muted" style={{ margin: 0 }}>
                {accessFixModal.itemTotal > 0
                  ? (isDe
                    ? `Eintrag ${accessFixModal.itemDone}/${accessFixModal.itemTotal} wird geprüft…`
                    : `Checking item ${accessFixModal.itemDone}/${accessFixModal.itemTotal}…`)
                  : (isDe ? 'Liste wird geladen…' : 'Loading list…')}
                {' '}{isDe ? 'Lass das Fenster offen, bis die Prüfung abgeschlossen ist.' : 'Keep this window open until the check completes.'}
              </p>
            </div>
          ) : lines.length === 0 ? (
            <div className="dex-ui-empty">{isDe ? 'Keine Rückmeldung aus dem Lauf.' : 'No result reported from the run.'}</div>
          ) : (
            <div className="dex-ui-callout dex-ui-callout--neutral">
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: '0.88rem', lineHeight: 1.6, color: 'var(--dex-gray-800)' }}>
                {lines.map((s, i) => <li key={i}>{s}</li>)}
              </ul>
            </div>
          )}
        </Modal>
  );
};
