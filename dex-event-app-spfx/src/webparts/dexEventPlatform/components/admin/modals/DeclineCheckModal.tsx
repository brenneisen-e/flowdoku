/* DeclineCheckModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16104-16184 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { AlertCircle, Check, Copy } from '../../Icons';
import { SPRegistration } from '../../../services/EventService';

export interface DeclineCheckModalProps {
  declineCopied: boolean;
  declineResult: { declinedAndRegistered: { email: string; name: string; reg: SPRegistration; }[]; declinedTotal: number; error: string; };
  isDe: boolean;
  setDeclineCopied: React.Dispatch<React.SetStateAction<boolean>>;
  setShowDeclineModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

// v31.2: Status als Pill in derselben Lesart wie die Teilnehmerliste — grün hält
// einen Platz, blau hat den QR schon, orange wartet. Der rohe Text stand vorher
// ohne Gewicht in der Tabelle und ging neben ID und E-Mail unter.
const statusPill = (st: string | undefined): string => {
  if (st === 'Angemeldet' || st === 'Eingecheckt') return 'dex-ui-pill--green';
  if (st === 'QR versendet') return 'dex-ui-pill--blue';
  if (st === 'Warteliste') return 'dex-ui-pill--orange';
  return 'dex-ui-pill--gray';
};

export const DeclineCheckModal: React.FC<DeclineCheckModalProps> = (p) => {
  const { declineCopied, declineResult, isDe, setDeclineCopied, setShowDeclineModal, showAlert } = p;
  const close = (): void => setShowDeclineModal(false);
  const rows = declineResult.declinedAndRegistered;
  const copyEmails = (): void => {
    const emails = rows.map(d => d.email).join('; ');
    navigator.clipboard.writeText(emails).then(() => {
      setDeclineCopied(true);
      setTimeout(() => setDeclineCopied(false), 2000);
    }).catch(() => showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{emails}</span>, { title: isDe ? 'E-Mail-Adressen manuell kopieren' : 'Copy email addresses manually' }));
  };
  // v31.2: Vorher ein eigener Overlay-Klon (fester Rahmen, ohne Escape, nur
  // deutsch); jetzt das gemeinsame Modal. Kopieren ist die einzige Handlung im
  // Dialog und steht deshalb als Primär-Knopf im Fuß statt als Nebenknopf über
  // der Tabelle.
  return (
    <Modal open={true} onClose={close} maxWidth={720}
      ariaLabel={isDe ? 'Outlook-Absagen prüfen' : 'Check Outlook declines'}
      title={isDe ? 'Outlook-Absagen prüfen' : 'Check Outlook declines'}
      subtitle={isDe ? 'Wer hat den Termin in Outlook abgelehnt, steht aber in der Teilnehmerliste noch als aktiv?' : 'Who declined the Outlook invitation but is still active on the participant list?'}
      icon={<AlertCircle size={20} />}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={close}>{isDe ? 'Schließen' : 'Close'}</button>
        {!declineResult.error && rows.length > 0 && (
          <button type="button" className="btn btn-primary" onClick={copyEmails}>
            {declineCopied ? <Check size={14} /> : <Copy size={14} />}
            {declineCopied ? (isDe ? 'Kopiert!' : 'Copied!') : (isDe ? 'E-Mail-Adressen kopieren' : 'Copy email addresses')}
          </button>
        )}
      </>}>
      {declineResult.error ? (
        <div className="dex-ui-callout dex-ui-callout--danger" style={{ whiteSpace: 'pre-line' }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{declineResult.error}</span>
        </div>
      ) : rows.length === 0 ? (
        <div className="dex-ui-callout dex-ui-callout--success">
          <span className="dex-ui-callout-icon"><Check size={16} /></span>
          <span>
            <strong>{isDe ? 'Keine Diskrepanzen gefunden.' : 'No discrepancies found.'}</strong>{' '}
            {declineResult.declinedTotal > 0
              ? (isDe ? `${declineResult.declinedTotal} Outlook-Absage(n) erfasst — keine davon steht in der Teilnehmerliste noch als aktiv.` : `${declineResult.declinedTotal} Outlook decline(s) recorded — none of them is still active on the participant list.`)
              : (isDe ? 'Niemand hat den Outlook-Termin abgelehnt.' : 'Nobody declined the Outlook invitation.')}
          </span>
        </div>
      ) : (
        <>
          <div className="dex-ui-grid-2">
            <div className="dex-ui-kpi dex-ui-kpi--orange">
              <div className="dex-ui-kpi-value">{rows.length}</div>
              <div className="dex-ui-kpi-label">{isDe ? 'abgelehnt, aber noch aktiv' : 'declined, but still active'}</div>
            </div>
            <div className="dex-ui-kpi">
              <div className="dex-ui-kpi-value">{declineResult.declinedTotal}</div>
              <div className="dex-ui-kpi-label">{isDe ? 'Outlook-Absagen insgesamt erfasst' : 'Outlook declines recorded in total'}</div>
            </div>
          </div>
          <p className="dex-ui-muted" style={{ margin: 0 }}>
            {isDe ? 'Diese Personen haben den Outlook-Termin abgelehnt, halten aber noch ihren Platz. Kopiere die Adressen, um bei ihnen nachzufragen.' : 'These people declined the Outlook invitation but still hold their seat. Copy the addresses to follow up with them.'}
          </p>
          <div className="dex-ui-table-wrap">
            <table className="dex-ui-table">
              <thead>
                <tr><th>ID</th><th>Name</th><th>E-Mail</th><th>Status</th></tr>
              </thead>
              <tbody>
                {rows.map(d => {
                  const displayName = (d.reg.Vorname && d.reg.Nachname)
                    ? `${d.reg.Vorname} ${d.reg.Nachname}`
                    : (d.reg.ParticipantName || d.name);
                  return (
                    <tr key={d.email}>
                      <td>{d.reg.TeilnehmerID ?? '-'}</td>
                      <td style={{ fontWeight: 600 }}>{displayName}</td>
                      <td style={{ color: 'var(--dex-gray-600)' }}>{d.email}</td>
                      <td><span className={cx('dex-ui-pill', statusPill(d.reg.Status))}>{d.reg.Status}</span></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}
    </Modal>
  );
};

