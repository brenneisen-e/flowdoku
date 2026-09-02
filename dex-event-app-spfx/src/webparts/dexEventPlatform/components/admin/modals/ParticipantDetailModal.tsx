/* ParticipantDetailModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14378-14430 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { translateStatus } from '../../../utils/eventStatus';
import { Icon } from '@fluentui/react/lib/Icon';

export interface ParticipantDetailModalProps {
  isDe: boolean;
  participantDetail: { name: string; email: string; jobTitle: string; location: string; company: string; department: string; phone: string; status: string; tid: number; };
  setParticipantDetail: React.Dispatch<React.SetStateAction<{ name: string; email: string; jobTitle: string; location: string; company: string; department: string; phone: string; status: string; tid: number; }>>;
}

export const ParticipantDetailModal: React.FC<ParticipantDetailModalProps> = (p) => {
  const { isDe, participantDetail, setParticipantDetail } = p;
  return (
          <Modal open={true} onClose={() => setParticipantDetail(null)} maxWidth={440} ariaLabel={isDe ? 'Teilnehmer-Details' : 'Participant details'}>
            {(() => {
              const p = participantDetail;
              const photo = p.email ? `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(p.email)}&size=L` : '';
              const rows: Array<[string, string]> = [];
              if (p.jobTitle) rows.push([isDe ? 'Position' : 'Job title', p.jobTitle]);
              if (p.department) rows.push([isDe ? 'Abteilung' : 'Department', p.department]);
              if (p.company) rows.push([isDe ? 'Unternehmen' : 'Company', p.company]);
              if (p.location) rows.push([isDe ? 'Standort' : 'Location', p.location]);
              if (p.phone) rows.push([isDe ? 'Telefon' : 'Phone', p.phone]);
              if (p.status) rows.push(['Status', translateStatus(p.status, isDe)]);
              if (p.tid) rows.push([isDe ? 'Teilnehmer-Nr.' : 'Participant no.', String(p.tid)]);
              return (
                <>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16 }}>
                    {photo && <img src={photo} alt="" onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }} style={{ width: 56, height: 56, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }} />}
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: '1.05rem' }}>{p.name}</div>
                      {p.email && <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', wordBreak: 'break-all' }}>{p.email}</div>}
                    </div>
                  </div>
                  {p.email && (
                    <div style={{ display: 'flex', gap: 10, marginBottom: 16, flexWrap: 'wrap' }}>
                      <a href={`mailto:${p.email}`} className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '6px 12px', textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="Mail" /> {isDe ? 'E-Mail' : 'Email'}
                      </a>
                      <a href={`msteams:/l/chat/0/0?users=${encodeURIComponent(p.email)}`} className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '6px 12px', textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 6, color: '#6264A7' }}>
                        <Icon iconName="TeamsLogo" /> {isDe ? 'Teams-Chat' : 'Teams chat'}
                      </a>
                    </div>
                  )}
                  <div style={{ fontWeight: 700, fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginBottom: 6 }}>{isDe ? 'Detailinfos' : 'Details'}</div>
                  {rows.length === 0 ? (
                    <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-400)' }}>{isDe ? 'Keine weiteren Angaben.' : 'No further details.'}</div>
                  ) : (
                    <div>
                      {rows.map(([k, v]) => (
                        <div key={k} style={{ display: 'flex', justifyContent: 'space-between', gap: 12, padding: '7px 0', borderBottom: '1px solid var(--dex-gray-100)', fontSize: '0.85rem' }}>
                          <span style={{ color: 'var(--dex-gray-500)' }}>{k}</span>
                          <span style={{ fontWeight: 600, textAlign: 'right' }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  )}
                  <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 18 }}>
                    <button className="btn btn-secondary" onClick={() => setParticipantDetail(null)}>{isDe ? 'Schließen' : 'Close'}</button>
                  </div>
                </>
              );
            })()}
          </Modal>
  );
};

