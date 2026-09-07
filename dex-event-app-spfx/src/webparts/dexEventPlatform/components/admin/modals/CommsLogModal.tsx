/* CommsLogModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14834-14935 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { ChevronDown, Mail, Trash2 } from '../../Icons';
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
  // v31.2: Zähler im Untertitel — der Organizer sieht auf einen Blick, wie
  // viel schon rausging, ohne die Liste zu zählen.
  const countLabel = commsLoading ? '' : ` · ${commsRows.length} ${isDe ? (commsRows.length === 1 ? 'Rundmail' : 'Rundmails') : (commsRows.length === 1 ? 'email' : 'emails')}`;
  return (
        <Modal
          open={showCommsModal}
          onClose={() => setShowCommsModal(false)}
          maxWidth={760}
          ariaLabel={isDe ? 'Gesendete Rundmails' : 'Sent broadcast emails'}
          title={isDe ? 'Gesendete Rundmails' : 'Sent broadcast emails'}
          subtitle={<>{selectedEvent.title}{countLabel}</>}
          icon={<Mail size={20} />}
          footer={
            <button type="button" className="btn btn-secondary" onClick={() => setShowCommsModal(false)}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          }
        >
          {/* v31.2: Eine Zeile, was das hier ist und was ein Klick tut — vorher
              stand die Liste ohne Einordnung da. */}
          <p className="dex-ui-muted" style={{ margin: 0 }}>
            {isDe
              ? 'Jede Rundmail, die aus dem Organizer Center an die Teilnehmer ging. Klicke einen Eintrag, um die Mail so zu sehen, wie sie ankam.'
              : 'Every broadcast email sent to attendees from the Organizer Center. Click an entry to see the email as it arrived.'}
          </p>
          <div style={{ maxHeight: '60vh', overflowY: 'auto' }}>
            {commsLoading ? (
              <p className="dex-ui-muted" style={{ padding: '16px 0', margin: 0 }}>{isDe ? 'Wird geladen…' : 'Loading…'}</p>
            ) : commsRows.length === 0 ? (
              <div className="dex-ui-empty">
                <span className="dex-ui-empty-icon"><Mail size={20} /></span>
                <div className="dex-ui-empty-title">{isDe ? 'Noch keine Rundmails' : 'No broadcast emails yet'}</div>
                <div>{isDe ? 'Zu diesem Event wurde noch keine Rundmail versendet.' : 'No broadcast emails have been sent for this event yet.'}</div>
              </div>
            ) : (
              <div className="dex-ui-stack" style={{ gap: 8 }}>
                {commsRows.map(row => {
                  const expanded = commsExpandedId === row.id;
                  const typeLabel = row.emailType === 'Einladung'
                    ? (isDe ? 'Einladung' : 'Invitation')
                    : row.emailType === 'Massenmail'
                      ? (isDe ? 'Massenmail' : 'Mass mail')
                      : row.emailType;
                  const subjectLabel = row.subject || (isDe ? '(ohne Betreff)' : '(no subject)');
                  return (
                    // v31.2: Karte statt Rahmen-Div; die grüne Kante markiert den
                    // aufgeklappten Eintrag, der Zeilen-Hover kommt aus dex-ui-row.
                    <div key={row.id} className={cx('dex-ui-card', expanded && 'dex-ui-card--accent')} style={{ padding: 0, overflow: 'hidden' }}>
                      <div className="dex-ui-row" style={{ padding: '4px 8px 4px 4px', borderRadius: 0 }}>
                        <button
                          type="button"
                          onClick={() => setCommsExpandedId(prev => prev === row.id ? null : row.id)}
                          aria-expanded={expanded}
                          className="dex-ui-row-main"
                          style={{
                            textAlign: 'left', background: 'none', border: 'none', cursor: 'pointer', padding: '8px 8px',
                            display: 'flex', alignItems: 'center', gap: 10, font: 'inherit', color: 'inherit',
                          }}
                        >
                          <span aria-hidden="true" style={{ flex: '0 0 auto', color: 'var(--dex-gray-500)', display: 'inline-flex', transition: 'transform 0.2s ease', transform: expanded ? 'rotate(180deg)' : 'none' }}><ChevronDown size={16} /></span>
                          <span style={{ minWidth: 0, flex: 1 }}>
                            <span className="dex-ui-row-title" style={{ display: 'block' }} title={subjectLabel}>{subjectLabel}</span>
                            <span className="dex-ui-row-sub dex-ui-inline" style={{ gap: 6 }}>
                              <span className={cx('dex-ui-pill', row.emailType === 'Einladung' ? 'dex-ui-pill--blue' : 'dex-ui-pill--gray')}>{typeLabel}</span>
                              <span>{formatDate(row.created)} · {isDe ? 'von' : 'from'} {row.sentByName || (isDe ? 'Unbekannt' : 'Unknown')}</span>
                            </span>
                          </span>
                        </button>
                        {/* v26.69: Log-Eintrag löschen — z. B. ein versehentlich
                            protokollierter Eintrag, der den „Bereits versendete
                            Infos"-Hinweis fälschlich auslöst. */}
                        <span className="dex-ui-row-actions">
                          <button
                            type="button"
                            className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                            onClick={async () => {
                              if (!(await confirmDialog(isDe ? `Diesen Log-Eintrag „${row.subject || '(ohne Betreff)'}" wirklich löschen? Der Eintrag verschwindet aus den event-bezogenen Nachrichten der Teilnehmer, und der „Bereits versendete Infos"-Hinweis entfällt, falls dies der letzte Eintrag war, der keine Einladung ist.` : `Really delete this log entry „${row.subject || '(no subject)'}"? It disappears from participants’ event messages, and the „earlier updates" hint is removed if this was the last entry that is not an invitation.`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
                              void deleteCommRow(row);
                            }}
                            disabled={commsDeletingId === row.id}
                            title={isDe ? 'Diesen Log-Eintrag löschen' : 'Delete this log entry'}
                            aria-label={isDe ? 'Diesen Log-Eintrag löschen' : 'Delete this log entry'}
                          >
                            <Trash2 size={16} />
                          </button>
                        </span>
                      </div>
                      {expanded && (
                        <div className="dex-ui-fade-in" style={{ borderTop: '1px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #fafafa)', padding: 10 }}>
                          <iframe
                            title={isDe ? 'Mail-Vorschau' : 'Email preview'}
                            srcDoc={row.bodyHtml}
                            sandbox=""
                            style={{ width: '100%', height: 420, border: 'none', borderRadius: 8, background: '#fff' }}
                          />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </Modal>
  );
};

