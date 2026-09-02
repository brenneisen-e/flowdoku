/* CommsLogModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14834-14935 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { ChevronDown, ChevronUp, Trash2, X } from '../../Icons';
import { formatDate } from '../../../utils/eventStatus';
import { DeloitteEvent } from '../../../types';
import { EventCommRow } from '../../../services/EventService';

export interface CommsLogModalProps {
  commsDeletingId: number;
  commsExpandedId: number;
  commsLoading: boolean;
  commsRows: EventCommRow[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  deleteCommRow: (row: EventCommRow) => Promise<void>;
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setCommsExpandedId: React.Dispatch<React.SetStateAction<number>>;
  setShowCommsModal: React.Dispatch<React.SetStateAction<boolean>>;
  showCommsModal: boolean;
}

export const CommsLogModal: React.FC<CommsLogModalProps> = (p) => {
  const { commsDeletingId, commsExpandedId, commsLoading, commsRows, confirmDialog, deleteCommRow, isDe, selectedEvent, setCommsExpandedId, setShowCommsModal, showCommsModal } = p;
  return (
        <Modal
          open={showCommsModal}
          onClose={() => setShowCommsModal(false)}
          maxWidth={760}
          padding={0}
          ariaLabel={isDe ? 'Gesendete Rundmails' : 'Sent broadcast emails'}
        >
          <div style={{
            padding: '14px 18px', borderBottom: '1px solid var(--dex-gray-200)',
            display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
          }}>
            <div style={{ minWidth: 0, flex: 1 }}>
              <h3 style={{ margin: 0, fontSize: '1.05rem' }}>{isDe ? 'Gesendete Rundmails' : 'Sent broadcast emails'}</h3>
              <p style={{ margin: '2px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{selectedEvent.title}</p>
            </div>
            <button
              type="button"
              onClick={() => setShowCommsModal(false)}
              style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', padding: 4 }}
              aria-label={isDe ? 'Schließen' : 'Close'}
            >
              <X size={22} />
            </button>
          </div>
          <div style={{ maxHeight: '65vh', overflowY: 'auto', padding: '10px 18px 18px' }}>
            {commsLoading ? (
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', padding: '16px 0' }}>{isDe ? 'Wird geladen…' : 'Loading…'}</p>
            ) : commsRows.length === 0 ? (
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', padding: '16px 0' }}>{isDe ? 'Es wurden noch keine Rundmails zu diesem Event versendet.' : 'No broadcast emails have been sent for this event yet.'}</p>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {commsRows.map(row => {
                  const expanded = commsExpandedId === row.id;
                  const typeLabel = row.emailType === 'Einladung'
                    ? (isDe ? 'Einladung' : 'Invitation')
                    : row.emailType === 'Massenmail'
                      ? (isDe ? 'Massenmail' : 'Mass mail')
                      : row.emailType;
                  return (
                    <div key={row.id} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, overflow: 'hidden' }}>
                      <div style={{ display: 'flex', alignItems: 'stretch' }}>
                        <button
                          type="button"
                          onClick={() => setCommsExpandedId(prev => prev === row.id ? null : row.id)}
                          style={{
                            flex: 1, minWidth: 0, textAlign: 'left', background: expanded ? 'var(--dex-gray-100)' : '#fff',
                            border: 'none', cursor: 'pointer', padding: '10px 12px',
                            display: 'flex', alignItems: 'flex-start', gap: 10,
                          }}
                        >
                          <span style={{ marginTop: 2, flex: '0 0 auto', color: 'var(--dex-gray-500)', display: 'inline-flex' }}>{expanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}</span>
                          <span style={{ minWidth: 0, flex: 1 }}>
                            <span style={{ display: 'block', fontWeight: 700, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>{row.subject || (isDe ? '(ohne Betreff)' : '(no subject)')}</span>
                            <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {typeLabel} · {formatDate(row.created)} · {isDe ? 'von' : 'from'} {row.sentByName || (isDe ? 'Unbekannt' : 'Unknown')}
                            </span>
                          </span>
                        </button>
                        {/* v26.69: Log-Eintrag löschen — z. B. ein versehentlich
                            protokollierter Eintrag, der den „Bereits versendete
                            Infos"-Hinweis fälschlich auslöst. */}
                        <button
                          type="button"
                          onClick={async () => {
                            if (!(await confirmDialog(isDe ? `Diesen Log-Eintrag „${row.subject || '(ohne Betreff)'}" wirklich löschen? Der Eintrag verschwindet aus den event-bezogenen Nachrichten der Teilnehmer, und der „Bereits versendete Infos"-Hinweis entfällt, falls dies der letzte Eintrag war, der keine Einladung ist.` : `Really delete this log entry „${row.subject || '(no subject)'}"? It disappears from participants’ event messages, and the „earlier updates" hint is removed if this was the last entry that is not an invitation.`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
                            void deleteCommRow(row);
                          }}
                          disabled={commsDeletingId === row.id}
                          title={isDe ? 'Diesen Log-Eintrag löschen' : 'Delete this log entry'}
                          style={{
                            flex: '0 0 auto', background: expanded ? 'var(--dex-gray-100)' : '#fff', border: 'none',
                            borderLeft: '1px solid var(--dex-gray-200)', cursor: 'pointer', padding: '0 14px',
                            color: 'var(--dex-red, #c00)', display: 'inline-flex', alignItems: 'center',
                          }}
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                      {expanded && (
                        <div style={{ borderTop: '1px solid var(--dex-gray-200)', background: '#f5f5f5', padding: 10 }}>
                          <iframe
                            title={isDe ? 'Mail-Vorschau' : 'Email preview'}
                            srcDoc={row.bodyHtml}
                            sandbox=""
                            style={{ width: '100%', height: 420, border: 'none', borderRadius: 6, background: '#fff' }}
                          />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
          <div style={{ padding: '12px 18px', borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'flex-end' }}>
            <button className="btn btn-secondary" onClick={() => setShowCommsModal(false)} style={{ fontSize: '0.85rem' }}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </Modal>
  );
};

