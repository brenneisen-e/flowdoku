/* ParticipantDetailModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14378-14430 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 * v31.2: Optik auf die gemeinsamen dex-ui-Klassen umgestellt (Modal-Kopf/-Fuß,
 * Karte, Zeilen); die Daten und ihre Reihenfolge sind unverändert.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { translateStatus } from '../../../utils/eventStatus';
import { Icon } from '@fluentui/react/lib/Icon';
import { Users } from '../../Icons';

export interface ParticipantDetailModalProps {
  isDe: boolean;
  participantDetail: { name: string; email: string; jobTitle: string; location: string; company: string; department: string; phone: string; status: string; tid: number; };
  setParticipantDetail: React.Dispatch<React.SetStateAction<{ name: string; email: string; jobTitle: string; location: string; company: string; department: string; phone: string; status: string; tid: number; }>>;
}

export const ParticipantDetailModal: React.FC<ParticipantDetailModalProps> = (p) => {
  const { isDe, participantDetail, setParticipantDetail } = p;
  const d = participantDetail;
  const close = (): void => setParticipantDetail(null);
  const photo = d.email ? `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(d.email)}&size=L` : '';
  // v31.2: Initialen unter dem Foto — bis v31.1 verschwand das Bild bei einem
  // Ladefehler einfach, und der Kopf sah aus, als fehle etwas.
  const initials = (d.name || d.email || '?').split(/[\s.@]+/).filter(Boolean).slice(0, 2).map(s => s[0].toUpperCase()).join('');
  const rows: Array<[string, string]> = [];
  if (d.jobTitle) rows.push([isDe ? 'Position' : 'Job title', d.jobTitle]);
  if (d.department) rows.push([isDe ? 'Abteilung' : 'Department', d.department]);
  if (d.company) rows.push([isDe ? 'Unternehmen' : 'Company', d.company]);
  if (d.location) rows.push([isDe ? 'Standort' : 'Location', d.location]);
  if (d.phone) rows.push([isDe ? 'Telefon' : 'Phone', d.phone]);
  if (d.status) rows.push(['Status', translateStatus(d.status, isDe)]);
  if (d.tid) rows.push([isDe ? 'Teilnehmer-ID' : 'Participant ID', String(d.tid)]);
  return (
    <Modal
      open={true}
      onClose={close}
      maxWidth={460}
      ariaLabel={isDe ? 'Teilnehmer-Details' : 'Participant details'}
      title={isDe ? 'Teilnehmer-Details' : 'Participant details'}
      icon={<Users size={20} />}
      footer={<button type="button" className="btn btn-secondary" onClick={close}>{isDe ? 'Schließen' : 'Close'}</button>}
    >
      <div className="dex-ui-card dex-ui-card--soft" style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
        <span className="dex-ui-avatar" aria-hidden="true" style={{ width: 64, height: 64, fontSize: '1.2rem', position: 'relative' }}>
          {initials}
          {photo && <img src={photo} alt="" onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }} style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover' }} />}
        </span>
        <div style={{ minWidth: 0, flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: '1.05rem' }}>{d.name}</div>
          {d.email && <div className="dex-ui-muted" style={{ wordBreak: 'break-all' }}>{d.email}</div>}
          {d.email && (
            <div className="dex-ui-inline" style={{ marginTop: 10 }}>
              <a href={`mailto:${d.email}`} className="btn btn-secondary dex-ui-btn-sm" style={{ textDecoration: 'none' }}>
                <Icon iconName="Mail" /> {isDe ? 'E-Mail schreiben' : 'Send email'}
              </a>
              <a href={`msteams:/l/chat/0/0?users=${encodeURIComponent(d.email)}`} className="btn btn-secondary dex-ui-btn-sm" style={{ textDecoration: 'none', color: '#6264A7' }}>
                <Icon iconName="TeamsLogo" /> {isDe ? 'Teams-Chat' : 'Teams chat'}
              </a>
            </div>
          )}
        </div>
      </div>
      <div className="dex-ui-section" style={{ marginTop: 0 }}>
        <div className="dex-ui-section-title">{isDe ? 'Angaben' : 'Details'}</div>
        {rows.length === 0 ? (
          <div className="dex-ui-empty">{isDe ? 'Keine weiteren Angaben.' : 'No further details.'}</div>
        ) : rows.map(([k, v]) => (
          <div key={k} className="dex-ui-row dex-ui-row--bordered" style={{ padding: '8px 10px' }}>
            <span className="dex-ui-muted">{k}</span>
            <span style={{ marginLeft: 'auto', fontWeight: 600, fontSize: '0.86rem', textAlign: 'right' }}>
              {k === 'Status' ? <span className="dex-ui-pill dex-ui-pill--gray">{v}</span> : v}
            </span>
          </div>
        ))}
      </div>
    </Modal>
  );
};
