/* DeclineCheckModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16104-16184 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { Copy } from '../../Icons';
import { SPRegistration } from '../../../services/EventService';

export interface DeclineCheckModalProps {
  declineCopied: boolean;
  declineResult: { declinedAndRegistered: { email: string; name: string; reg: SPRegistration; }[]; declinedTotal: number; error: string; };
  isDe: boolean;
  setDeclineCopied: React.Dispatch<React.SetStateAction<boolean>>;
  setShowDeclineModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export const DeclineCheckModal: React.FC<DeclineCheckModalProps> = (p) => {
  const { declineCopied, declineResult, isDe, setDeclineCopied, setShowDeclineModal, showAlert } = p;
  return (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 2000,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
          onClick={() => setShowDeclineModal(false)}
        >
          <div
            className="card"
            style={{ background: '#fff', maxWidth: 720, width: '100%', maxHeight: '80vh', overflow: 'auto', padding: 24 }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>Outlook-Absagen vs. Anmeldungen</h3>
              <button className="btn btn-secondary" style={{ padding: '4px 10px' }} onClick={() => setShowDeclineModal(false)}>
                Schließen
              </button>
            </div>
            {declineResult.error ? (
              <p style={{ color: 'var(--dex-red)', whiteSpace: 'pre-line' }}>{declineResult.error}</p>
            ) : declineResult.declinedAndRegistered.length === 0 ? (
              <p style={{ color: 'var(--dex-gray-600)' }}>
                Keine Diskrepanzen gefunden. {declineResult.declinedTotal > 0
                  ? `Es gibt ${declineResult.declinedTotal} Outlook-Absage(n), aber keiner davon ist in der Teilnehmerliste noch aktiv.`
                  : 'Niemand hat den Outlook-Termin abgelehnt.'}
              </p>
            ) : (
              <>
                <p style={{ color: 'var(--dex-gray-700)' }}>
                  <strong>{declineResult.declinedAndRegistered.length}</strong> Teilnehmer haben den Outlook-Termin abgelehnt,
                  stehen aber in der Teilnehmerliste noch als aktiv:
                </p>
                <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.8rem' }}
                    onClick={() => {
                      const emails = declineResult.declinedAndRegistered.map(d => d.email).join('; ');
                      navigator.clipboard.writeText(emails).then(() => {
                        setDeclineCopied(true);
                        setTimeout(() => setDeclineCopied(false), 2000);
                      }).catch(() => showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{emails}</span>, { title: isDe ? 'E-Mail-Adressen manuell kopieren' : 'Copy email addresses manually' }));
                    }}
                  >
                    <Copy size={14} /> {declineCopied ? 'Kopiert!' : 'E-Mails kopieren'}
                  </button>
                </div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ background: 'var(--dex-gray-50)' }}>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>ID</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>Name</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>E-Mail</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {declineResult.declinedAndRegistered.map(d => {
                      const displayName = (d.reg.Vorname && d.reg.Nachname)
                        ? `${d.reg.Vorname} ${d.reg.Nachname}`
                        : (d.reg.ParticipantName || d.name);
                      return (
                        <tr key={d.email}>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.reg.TeilnehmerID ?? '-'}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{displayName}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.email}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.reg.Status}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 12 }}>
                  Insgesamt {declineResult.declinedTotal} Outlook-Absage(n) erfasst.
                </p>
              </>
            )}
          </div>
        </div>
  );
};

