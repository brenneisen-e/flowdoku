/**
 * Meine Events - zeigt alle Events fuer die der User registriert ist.
 * Laedt Registrierungen aus den jeweiligen Teilnehmerlisten.
 * Ermoeglicht Abmeldung mit Zwei-Schritt-Bestaetigung.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SPHttpClient } from '@microsoft/sp-http';
import DocViewer, { DocViewerRenderers } from '@cyntler/react-doc-viewer';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { DeloitteEvent, EventSpecificField, AgendaItem, TransferTime } from '../types';
import { SPRegistration } from '../services/EventService';
import { useLanguage } from '../context/LanguageContext';

interface MyEventEntry {
  event: DeloitteEvent;
  registration: SPRegistration;
}

function formatDate(iso: string): string {
  if (!iso) return '-';
  return new Date(iso).toLocaleDateString('de-DE', {
    day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  });
}

function formatDateRange(start: string, end: string): string {
  if (!start) return '-';
  const s = new Date(start);
  const e = end ? new Date(end) : null;
  const sDate = s.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const sTime = s.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  if (!e) return `${sDate}, ${sTime}`;
  const eDate = e.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const eTime = e.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  // Gleicher Tag: "14.04.2026, 14:00 – 18:00"
  if (sDate === eDate) return `${sDate}, ${sTime} – ${eTime}`;
  // Verschiedene Tage
  return `${sDate}, ${sTime} – ${eDate}, ${eTime}`;
}

function getStatusBadgeClass(status: string): string {
  switch (status) {
    case 'Angemeldet': return 'badge-green';
    case 'QR versendet': return 'badge-green';
    case 'Warteliste': return 'badge-orange';
    case 'Abgemeldet': return 'badge-red';
    case 'Eingecheckt': return 'badge-green';
    default: return 'badge-gray';
  }
}

function getStatusLabel(status: string, t: (key: string) => string): string {
  switch (status) {
    case 'Angemeldet': return t('status.registered');
    case 'Warteliste': return t('status.waitlist');
    case 'Abgemeldet': return t('status.cancelled');
    case 'Eingecheckt': return t('status.checkedin');
    default: return status;
  }
}

function getDocIconName(name: string): string {
  const ext = name.split('.').pop()?.toLowerCase() || '';
  switch (ext) {
    case 'pdf': return 'PDF';
    case 'doc': case 'docx': return 'WordDocument';
    case 'xls': case 'xlsx': return 'ExcelDocument';
    case 'ppt': case 'pptx': return 'PowerPointDocument';
    case 'jpg': case 'jpeg': case 'png': case 'gif': return 'FileImage';
    default: return 'Page';
  }
}

function DocumentsViewer({ documents, t }: { documents: Array<{name: string; url: string; size?: number}>; t: (key: string) => string }): React.ReactElement {
  const [expandedDoc, setExpandedDoc] = React.useState<string | null>(null);
  const [blobUrl, setBlobUrl] = React.useState<string>('');
  const [loading, setLoading] = React.useState(false);

  const toggleDoc = async (doc: { url: string; name: string }): Promise<void> => {
    if (expandedDoc === doc.url) {
      setExpandedDoc(null);
      if (blobUrl) { URL.revokeObjectURL(blobUrl); setBlobUrl(''); }
      return;
    }
    setExpandedDoc(doc.url);
    setLoading(true);
    setBlobUrl('');

    try {
      // Datei per SPHttpClient laden (authentifiziert)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const origin = doc.url.match(/^https?:\/\/[^/]+/)?.[0] || '';
        const serverRelPath = doc.url.replace(origin, '');
        const resp = await ctx.spHttpClient.get(
          `${ctx.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelPath)}')/$value`,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': '*/*' } }
        );
        if (resp.ok) {
          const blob = await resp.blob();
          // MIME-Type korrigieren
          const ext = doc.name.split('.').pop()?.toLowerCase() || '';
          const mimeMap: Record<string, string> = { pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' };
          const correctBlob = mimeMap[ext] ? new Blob([blob], { type: mimeMap[ext] }) : blob;
          setBlobUrl(URL.createObjectURL(correctBlob));
        }
      }
    } catch { /* Blob-Laden fehlgeschlagen */ }
    setLoading(false);
  };

  // Cleanup blob URLs bei Unmount
  React.useEffect(() => {
    return () => { if (blobUrl) URL.revokeObjectURL(blobUrl); };
  }, []);

  return (
    <div style={{ marginTop: 12 }}>
      <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 6 }}>
        {t('myevents.documents')}
      </div>
      {documents.map((doc, i) => {
        const isExpanded = expandedDoc === doc.url;

        return (
          <div key={i} style={{ marginBottom: 6 }}>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 10, padding: '10px 14px',
              background: isExpanded ? 'var(--dex-green-light, #f0fdf4)' : 'var(--dex-gray-100)',
              borderRadius: isExpanded ? '8px 8px 0 0' : 8,
              cursor: 'pointer', fontSize: '0.85rem', color: 'var(--dex-gray-700)',
              transition: 'background 0.15s',
            }} onClick={() => toggleDoc(doc)}>
              <span style={{ width: 32, height: 32, borderRadius: '50%', background: 'var(--dex-green-dark, #6b9a1e)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                <Icon iconName={getDocIconName(doc.name)} style={{ fontSize: 16, color: '#fff' }} />
              </span>
              <span style={{ flex: 1, fontWeight: isExpanded ? 600 : 400 }}>{doc.name}</span>
              {doc.size ? <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.75rem' }}>{(doc.size / 1024).toFixed(0)} KB</span> : null}
              <a href={doc.url} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()} style={{ color: 'var(--dex-green-dark)', fontSize: '0.72rem', textDecoration: 'none' }}>
                <Icon iconName="Download" style={{ fontSize: 14 }} />
              </a>
              <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)' }}>{isExpanded ? '▲' : '▼'}</span>
            </div>
            {isExpanded && (
              <div style={{
                border: '1px solid var(--dex-gray-200)', borderTop: 'none',
                borderRadius: '0 0 8px 8px', overflow: 'hidden', background: '#fff',
              }}>
                {loading ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--dex-gray-400)' }}>
                    {t('myevents.agenda') === 'Programm' ? 'Vorschau wird geladen...' : 'Loading preview...'}
                  </div>
                ) : blobUrl ? (
                  <DocViewer
                    documents={[{ uri: blobUrl, fileName: doc.name }]}
                    pluginRenderers={DocViewerRenderers}
                    config={{ header: { disableHeader: true, disableFileName: true } }}
                    style={{ height: 500 }}
                  />
                ) : (
                  <div style={{ padding: 24, textAlign: 'center' }}>
                    <a href={doc.url} target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark)' }}>
                      {t('myevents.agenda') === 'Programm' ? 'Im Browser öffnen' : 'Open in browser'}
                    </a>
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

export default function MyEventsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { events, getMyRegistration, getMyEventNumbers, cancelRegistration, updateMyRegistration } = useEvents();
  const { t } = useLanguage();
  const [myEvents, setMyEvents] = React.useState<MyEventEntry[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [cancellingId, setCancellingId] = React.useState<string | null>(null);
  const [isCancelling, setIsCancelling] = React.useState(false);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [editData, setEditData] = React.useState<Record<string, string>>({});
  const [isSaving, setIsSaving] = React.useState(false);
  const [loadError, setLoadError] = React.useState('');

  React.useEffect(() => {
    loadMyRegistrations();
  }, [events]);

  async function loadMyRegistrations(): Promise<void> {
    setIsLoading(true);
    setLoadError('');
    const entries: MyEventEntry[] = [];

    // Schneller Pfad: DEX_Participants abfragen
    const myNumbers = await getMyEventNumbers();
    const allMyNumbers = [...myNumbers.registered, ...myNumbers.waitlisted];

    if (allMyNumbers.length > 0) {
      // Nur Events laden die in DEX_Participants stehen
      const relevantEvents = events.filter(e => e.eventNumber && allMyNumbers.indexOf(e.eventNumber) >= 0);
      for (const event of relevantEvents) {
        try {
          const reg = await getMyRegistration(event.id);
          if (reg) {
            entries.push({ event, registration: reg });
          }
        } catch { /* */ }
      }
    } else {
      // Fallback: Alter Weg fuer Altdaten ohne DEX_Participants-Eintrag
      for (const event of events) {
        try {
          const reg = await getMyRegistration(event.id);
          if (reg) {
            entries.push({ event, registration: reg });
          }
        } catch { /* */ }
      }
    }

    if (entries.length === 0 && allMyNumbers.length > 0) {
      setLoadError('Registrierungen konnten nicht geladen werden.');
    }
    setMyEvents(entries);
    setIsLoading(false);
  }

  const handleCancel = async (eventId: string): Promise<void> => {
    if (cancellingId === eventId) {
      setIsCancelling(true);

      // Check if this is a late cancellation
      const entry = myEvents.find(e => e.event.id === eventId);
      const isLateCancellation = entry?.event.lastDeregisterDate && new Date(entry.event.lastDeregisterDate) < new Date();

      const success = await cancelRegistration(eventId);
      if (success) {
        // If late cancellation, send notification to organizer
        if (isLateCancellation && entry) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx) {
              const { EventService } = await import('../services/EventService');
              const svc = new EventService(ctx);
              const userName = `${entry.registration.Vorname || ''} ${entry.registration.Nachname || ''}`.trim() || entry.registration.ParticipantEmail;
              await svc.queueEmail(
                `[DEX] Late cancellation: ${entry.event.title}`,
                entry.event.organizers[0] || '',
                entry.event.organizers[0] || '',
                `<p><strong>${userName}</strong> has cancelled their registration for <strong>${entry.event.title}</strong> after the cancellation deadline (${new Date(entry.event.lastDeregisterDate).toLocaleDateString('de-DE')}).</p><p>E-Mail: ${entry.registration.ParticipantEmail || entry.registration.Title}</p>`,
                'LateCancel',
                entry.event.title,
                eventId
              );
            }
          } catch { /* Email-Fehler ignorieren */ }
        }
        await loadMyRegistrations();
      }
      setCancellingId(null);
      setIsCancelling(false);
    } else {
      setCancellingId(eventId);
    }
  };

  const activeEntries = myEvents.filter(e => e.registration.Status !== 'Abgemeldet');
  const cancelledEntries = myEvents.filter(e => e.registration.Status === 'Abgemeldet');

  if (isLoading) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>{t('myevents.loading')}</p>
      </div>
    );
  }

  return (
    <div className="page-container">
      <h2 className="mb-16">{t('myevents.title')}</h2>

      {loadError && (
        <div className="card" style={{ padding: 16, marginBottom: 16, color: 'var(--dex-red)' }}>
          {loadError}
        </div>
      )}

      {activeEntries.length === 0 && cancelledEntries.length === 0 && !loadError && (
        <div className="card text-center" style={{ padding: 48 }}>
          <p style={{ color: 'var(--dex-gray-400)' }}>{t('myevents.empty')}</p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>{t('myevents.browse')}</button>
        </div>
      )}

      {activeEntries.length > 0 && (
        <div className="my-events-list">
          {activeEntries.map(({ event, registration }) => {
            // Custom Data parsen und IDs zu Labels mappen
            let customData: Record<string, string> = {};
            try {
              if (registration.CustomData) customData = JSON.parse(registration.CustomData);
            } catch { /* */ }

            // Feld-ID zu Label-Map aus den Event-Feldern erstellen
            const fieldLabelMap: Record<string, string> = {};
            for (const field of event.eventSpecificFields) {
              fieldLabelMap[field.id] = field.label;
            }

            // "salutation" überspringen (wird schon im Namen angezeigt)
            const displayData = Object.keys(customData)
              .filter(key => key !== 'salutation' && customData[key])
              .map(key => ({
                label: fieldLabelMap[key] || key,
                value: customData[key],
              }));

            return (
              <div key={event.id} className="card my-event-card" style={{ overflow: 'hidden' }}>
                {event.imageUrl && (
                  <div style={{
                    height: 100, background: `url(${event.imageUrl}) center/cover no-repeat`,
                    margin: '-16px -16px 0 -16px',
                  }} />
                )}

                {/* Header: Titel + Status Badge */}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginTop: event.imageUrl ? 16 : 0 }}>
                  <h3 style={{ margin: 0, fontSize: '1.1rem' }}>{event.title}</h3>
                  <span className={`badge ${getStatusBadgeClass(registration.Status)}`} style={{ flexShrink: 0, marginLeft: 12 }}>
                    {getStatusLabel(registration.Status, t)}
                  </span>
                </div>

                {/* Kompakte Info-Zeilen */}
                <div style={{ marginTop: 12, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 24px', fontSize: '0.88rem', color: 'var(--dex-gray-700)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}><Icon iconName="MapPin" style={{ fontSize: 14, color: 'var(--dex-gray-500)' }} /> {event.location || '-'}</div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}><Icon iconName="Calendar" style={{ fontSize: 14, color: 'var(--dex-gray-500)' }} /> {formatDateRange(event.startDate, event.endDate)}</div>
                </div>

                {/* Custom Fields als kompakte Tags */}
                {!editingId || editingId !== event.id ? (
                  displayData.length > 0 && (
                    <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                      {displayData.map(({ label, value }) => (
                        <span key={label} style={{
                          fontSize: '0.78rem', padding: '3px 10px', borderRadius: 12,
                          background: 'var(--dex-gray-100)', color: 'var(--dex-gray-700)',
                        }}>
                          {label}: <strong>{value}</strong>
                        </span>
                      ))}
                    </div>
                  )
                ) : (
                  <div style={{ marginTop: 12 }}>
                    {event.eventSpecificFields.map((field: EventSpecificField) => (
                      <div className="form-group" key={field.id} style={{ marginBottom: 10 }}>
                        <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 2 }}>
                          {field.required && <span className="required">*</span>}
                          {field.label}
                        </label>
                        {field.type === 'select' ? (
                          <select className="form-select" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })}>
                            <option value="">—</option>
                            {field.options && field.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        ) : (
                          <input className="form-input" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })} placeholder={field.label} type={field.type === 'number' ? 'number' : 'text'} />
                        )}
                      </div>
                    ))}
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button className="btn btn-primary" style={{ fontSize: '0.82rem' }} disabled={isSaving} onClick={async () => { setIsSaving(true); await updateMyRegistration(event.id, editData); await loadMyRegistrations(); setEditingId(null); setIsSaving(false); }}>
                        {isSaving ? t('myevents.saving') : t('myevents.save')}
                      </button>
                      <button className="btn btn-secondary" style={{ fontSize: '0.82rem' }} onClick={() => setEditingId(null)}>{t('general.cancel')}</button>
                    </div>
                  </div>
                )}

                {/* Agenda / Timeline - mehrspaltig bei mehreren Tagen */}
                {event.agenda && event.agenda.length > 0 && (() => {
                  const grouped = Object.entries(
                    event.agenda.reduce((groups: Record<string, AgendaItem[]>, item: AgendaItem) => {
                      const key = item.date || 'TBD';
                      if (!groups[key]) groups[key] = [];
                      groups[key].push(item);
                      return groups;
                    }, {} as Record<string, AgendaItem[]>)
                  ).sort(([a], [b]) => a.localeCompare(b));
                  const dayCount = grouped.length;
                  const cols = dayCount >= 3 ? 3 : dayCount >= 2 ? 2 : 1;

                  return (
                    <div style={{ marginTop: 12 }}>
                      <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                        {t('myevents.agenda')}
                      </div>
                      <div style={{
                        display: 'grid',
                        gridTemplateColumns: `repeat(${cols}, 1fr)`,
                        gap: 16,
                      }}>
                        {grouped.map(([date, items]) => (
                          <div key={date} style={{
                            background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, padding: 12,
                            border: '1px solid var(--dex-gray-200)',
                          }}>
                            <div style={{
                              fontSize: '0.78rem', fontWeight: 700, color: '#fff', marginBottom: 8,
                              background: 'var(--dex-green-dark, #6b9a1e)', borderRadius: 8, padding: '6px 12px',
                              textAlign: 'center',
                            }}>
                              {date !== 'TBD' ? new Date(date + 'T00:00').toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' }) : 'TBD'}
                            </div>
                            {items.sort((a: AgendaItem, b: AgendaItem) => (a.time || '').localeCompare(b.time || '')).map((item: AgendaItem) => (
                              <div key={item.id} style={{
                                display: 'flex', alignItems: 'flex-start', gap: 8, padding: '6px 0',
                                borderLeft: '2px solid var(--dex-green)', marginLeft: 4, paddingLeft: 10,
                              }}>
                                <span style={{ width: 26, height: 26, borderRadius: '50%', background: 'var(--dex-green-dark, #6b9a1e)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                  <Icon iconName={item.icon || 'Calendar'} style={{ fontSize: 12, color: '#fff' }} />
                                </span>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: '0.8rem', fontWeight: 600 }}>
                                    {item.time}{item.endTime ? ` – ${item.endTime}` : ''}
                                  </div>
                                  <div style={{ fontSize: '0.8rem' }}>{item.title}</div>
                                  {item.description && (
                                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 1 }}>{item.description}</div>
                                  )}
                                </div>
                              </div>
                            ))}
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}

                {/* Transferzeiten */}
                {event.transferTimes && event.transferTimes.length > 0 && (
                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 6 }}>
                      {t('myevents.transfers')}
                    </div>
                    {event.transferTimes.sort((a: TransferTime, b: TransferTime) => (a.date + a.departureTime).localeCompare(b.date + b.departureTime)).map((tr: TransferTime) => (
                      <div key={tr.id} style={{ display: 'flex', gap: 8, padding: '4px 0', fontSize: '0.82rem', borderLeft: '2px solid var(--dex-orange)', marginLeft: 8, paddingLeft: 12 }}>
                        <Icon iconName="Bus" style={{ fontSize: 14, color: 'var(--dex-orange)' }} />
                        <div>
                          <strong>{tr.location}</strong> – {new Date(tr.date + 'T00:00').toLocaleDateString('de-DE', {weekday: 'short', day: '2-digit', month: '2-digit'})}, {tr.departureTime}{tr.arrivalTime ? ` → ${tr.arrivalTime}` : ''}
                          {tr.description && <span style={{ color: 'var(--dex-gray-500)', marginLeft: 8 }}>{tr.description}</span>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Dokumente mit Viewer */}
                {event.documents && event.documents.length > 0 && (
                  <DocumentsViewer documents={event.documents} t={t} />
                )}

                {/* Registriert am + Aktionen */}
                <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                    {t('myevents.registeredon')}: {formatDate(registration.RegistrationDate)}
                  </span>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                    {cancellingId === event.id && !isCancelling && event.lastDeregisterDate && new Date(event.lastDeregisterDate) < new Date() && (
                      <span style={{ fontSize: '0.75rem', color: 'var(--dex-orange)', display: 'block', marginBottom: 4, width: '100%' }}>
                        {t('myevents.latecancel')}
                      </span>
                    )}
                    <button className="btn btn-secondary" style={{ fontSize: '0.78rem', padding: '4px 12px' }} onClick={() => { if (editingId === event.id) { setEditingId(null); } else { setEditData(customData); setEditingId(event.id); } }}>
                      {editingId === event.id ? t('general.cancel') : t('myevents.edit')}
                    </button>
                    <button className="btn" onClick={() => handleCancel(event.id)} disabled={isCancelling} style={{ fontSize: '0.78rem', padding: '4px 12px', background: cancellingId === event.id ? 'var(--dex-red)' : 'var(--dex-gray-200)', color: cancellingId === event.id ? '#fff' : 'var(--dex-gray-700)' }}>
                      {cancellingId === event.id ? (isCancelling ? '...' : t('myevents.confirmcancel')) : t('myevents.cancel')}
                    </button>
                    {cancellingId === event.id && !isCancelling && (
                      <button className="btn btn-secondary" onClick={() => setCancellingId(null)} style={{ fontSize: '0.78rem', padding: '4px 12px' }}>{t('myevents.keepreg')}</button>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {cancelledEntries.length > 0 && (
        <div>
          <h3 className="mt-24 mb-16" style={{ color: 'var(--dex-gray-400)' }}>{t('myevents.cancelledevents')}</h3>
          <div className="my-events-list">
            {cancelledEntries.map(({ event, registration }) => (
              <div key={event.id} className="card my-event-card" style={{ opacity: 0.6 }}>
                <div className="my-event-card__header">
                  <h3>{event.title}</h3>
                  <span className="badge badge-red">{t('status.cancelled')}</span>
                </div>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-400)' }}>
                  {t('myevents.cancelledon')}: {registration.CancellationDate ? formatDate(registration.CancellationDate) : '-'}
                </p>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
