/**
 * Meine Events - zeigt alle Events fuer die der User registriert ist.
 * Laedt Registrierungen aus den jeweiligen Teilnehmerlisten.
 * Ermoeglicht Abmeldung mit Zwei-Schritt-Bestaetigung.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';

import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { DeloitteEvent, EventSpecificField, AgendaItem, TransferTime, QuizQuestion } from '../types';
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

function QuizPlayer({ quiz, t }: { quiz: QuizQuestion[]; t: (key: string) => string }): React.ReactElement {
  const [currentQ, setCurrentQ] = React.useState(0);
  const [selectedAnswer, setSelectedAnswer] = React.useState<number | null>(null);
  const [score, setScore] = React.useState(0);
  const [finished, setFinished] = React.useState(false);
  const [showQuiz, setShowQuiz] = React.useState(false);

  const question = quiz[currentQ];

  const handleAnswer = (index: number): void => {
    if (selectedAnswer !== null) return; // Already answered
    setSelectedAnswer(index);
    if (index === question.correctIndex) setScore(s => s + 1);
  };

  const nextQuestion = (): void => {
    if (currentQ < quiz.length - 1) {
      setCurrentQ(currentQ + 1);
      setSelectedAnswer(null);
    } else {
      setFinished(true);
    }
  };

  const restart = (): void => {
    setCurrentQ(0);
    setSelectedAnswer(null);
    setScore(0);
    setFinished(false);
  };

  if (!showQuiz) {
    return (
      <div style={{ marginTop: 12 }}>
        <button
          onClick={() => setShowQuiz(true)}
          style={{
            width: '100%', padding: '14px 20px', borderRadius: 12,
            border: '2px solid var(--dex-green)', background: 'rgba(134,188,37,0.06)',
            cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
            fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-green-dark)',
          }}
        >
          <Icon iconName="Game" style={{ fontSize: 20 }} />
          Fun-Zone: Quiz ({quiz.length} {quiz.length === 1 ? 'Frage' : 'Fragen'})
        </button>
      </div>
    );
  }

  if (finished) {
    return (
      <div style={{ marginTop: 12, padding: 20, background: 'var(--dex-gray-50)', borderRadius: 12, textAlign: 'center' }}>
        <div style={{ fontSize: '2rem', marginBottom: 8 }}>{score === quiz.length ? '🎉' : score >= quiz.length / 2 ? '👏' : '💪'}</div>
        <div style={{ fontSize: '1.1rem', fontWeight: 700, marginBottom: 4 }}>
          {score} / {quiz.length} {t('myevents.agenda') === 'Programm' ? 'richtig' : 'correct'}
        </div>
        <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginBottom: 12 }}>
          {score === quiz.length
            ? (t('myevents.agenda') === 'Programm' ? 'Perfekt! Alle richtig!' : 'Perfect! All correct!')
            : (t('myevents.agenda') === 'Programm' ? 'Gut gemacht!' : 'Well done!')}
        </div>
        <div style={{ display: 'flex', gap: 8, justifyContent: 'center' }}>
          <button className="btn btn-primary" onClick={restart} style={{ fontSize: '0.82rem' }}>
            {t('myevents.agenda') === 'Programm' ? 'Nochmal spielen' : 'Play again'}
          </button>
          <button className="btn btn-secondary" onClick={() => setShowQuiz(false)} style={{ fontSize: '0.82rem' }}>
            {t('create.templates.close')}
          </button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ marginTop: 12, padding: 16, background: 'var(--dex-gray-50)', borderRadius: 12 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
        <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
          {t('myevents.agenda') === 'Programm' ? 'Frage' : 'Question'} {currentQ + 1} / {quiz.length}
        </span>
        <span style={{ fontSize: '0.78rem', color: 'var(--dex-green-dark)', fontWeight: 600 }}>
          {score} {t('myevents.agenda') === 'Programm' ? 'Punkte' : 'Points'}
        </span>
      </div>
      <div style={{ fontSize: '1rem', fontWeight: 600, marginBottom: 12 }}>{question.question}</div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {question.options.map((opt, i) => {
          let bg = 'var(--dex-white)';
          let border = '1px solid var(--dex-gray-200)';
          let color = 'var(--dex-gray-800)';
          if (selectedAnswer !== null) {
            if (i === question.correctIndex) { bg = 'rgba(134,188,37,0.15)'; border = '2px solid var(--dex-green)'; color = 'var(--dex-green-dark)'; }
            else if (i === selectedAnswer) { bg = 'rgba(218,41,28,0.1)'; border = '2px solid var(--dex-red)'; color = 'var(--dex-red)'; }
          }
          return (
            <button
              key={i}
              onClick={() => handleAnswer(i)}
              disabled={selectedAnswer !== null}
              style={{
                padding: '10px 14px', borderRadius: 10, border, background: bg, color,
                cursor: selectedAnswer !== null ? 'default' : 'pointer', textAlign: 'left',
                fontSize: '0.88rem', fontWeight: selectedAnswer !== null && i === question.correctIndex ? 700 : 400,
                transition: 'all 0.2s',
              }}
            >
              {opt}
            </button>
          );
        })}
      </div>
      {selectedAnswer !== null && (
        <button className="btn btn-primary" onClick={nextQuestion} style={{ marginTop: 12, fontSize: '0.82rem' }}>
          {currentQ < quiz.length - 1
            ? (t('myevents.agenda') === 'Programm' ? 'Nächste Frage' : 'Next question')
            : (t('myevents.agenda') === 'Programm' ? 'Ergebnis anzeigen' : 'Show result')}
        </button>
      )}
    </div>
  );
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) { setLoading(false); return; }

      // Datei per SPHttpClient REST API als Binary laden
      const siteUrl = ctx.pageContext.web.absoluteUrl;
      const origin = doc.url.match(/^https?:\/\/[^/]+/)?.[0] || '';
      const serverRelPath = decodeURIComponent(doc.url.replace(origin, ''));

      // Pfad-Segmente einzeln encoden (Leerzeichen, Klammern etc.)
      const encodedPath = serverRelPath.split('/').map(s => encodeURIComponent(s)).join('/');
      const apiUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/$value`;

      // XHR fuer Binary-Download (zuverlaessiger als fetch fuer SharePoint)
      const blob = await new Promise<Blob | null>((resolve) => {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', apiUrl, true);
        xhr.responseType = 'blob';
        xhr.withCredentials = true;
        xhr.setRequestHeader('Accept', '*/*');
        xhr.onload = () => {
          if (xhr.status === 200 && xhr.response) {
            resolve(xhr.response as Blob);
          } else {
            console.warn('[DEX] Doc XHR failed:', xhr.status, apiUrl);
            resolve(null);
          }
        };
        xhr.onerror = () => { console.warn('[DEX] Doc XHR error'); resolve(null); };
        xhr.send();
      });

      if (blob && blob.size > 0) {
        const ext = doc.name.split('.').pop()?.toLowerCase() || '';
        const mimeMap: Record<string, string> = { pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' };
        const correctBlob = (mimeMap[ext] && blob.type !== mimeMap[ext]) ? new Blob([blob], { type: mimeMap[ext] }) : blob;
        setBlobUrl(URL.createObjectURL(correctBlob));
      }
    } catch (err) { console.warn('[DEX] Doc viewer error:', err); }
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
              {doc.url && doc.url.startsWith('http') && (
                <a href={doc.url} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()} style={{ color: 'var(--dex-green-dark)', fontSize: '0.72rem', textDecoration: 'none' }}>
                  <Icon iconName="Download" style={{ fontSize: 14 }} />
                </a>
              )}
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
                  <iframe
                    src={blobUrl}
                    style={{ width: '100%', height: 500, border: 'none' }}
                    title={doc.name}
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
                    {registration.Status === 'Warteliste' && registration.TeilnehmerID && event.maxParticipants > 0
                      ? `${getStatusLabel(registration.Status, t)} #${registration.TeilnehmerID - event.maxParticipants}`
                      : getStatusLabel(registration.Status, t)}
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
                      <div key={tr.id} style={{
                        display: 'flex', gap: 10, padding: '8px 12px', marginBottom: 6, fontSize: '0.82rem',
                        background: 'var(--dex-gray-50, #fafafa)', borderRadius: 10,
                        borderLeft: '3px solid var(--dex-orange)',
                      }}>
                        <span style={{ width: 28, height: 28, borderRadius: '50%', background: 'var(--dex-orange)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          <Icon iconName="Bus" style={{ fontSize: 13, color: '#fff' }} />
                        </span>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: 600 }}>
                            {tr.location}{tr.meetingPoint ? ` – ${tr.meetingPoint}` : ''}
                          </div>
                          {tr.address && (
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                              <Icon iconName="MapPin" style={{ fontSize: 11, marginRight: 4 }} />{tr.address}
                            </div>
                          )}
                          <div style={{ marginTop: 2 }}>
                            <Icon iconName="Calendar" style={{ fontSize: 11, color: 'var(--dex-gray-400)', marginRight: 4 }} />
                            {new Date(tr.date + 'T00:00').toLocaleDateString('de-DE', {weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric'})},
                            {' '}{tr.departureTime}{tr.arrivalTime ? ` → ${tr.arrivalTime}` : ''} Uhr
                          </div>
                          {tr.description && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{tr.description}</div>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Dokumente mit Viewer */}
                {event.documents && event.documents.length > 0 && (
                  <DocumentsViewer documents={event.documents} t={t} />
                )}

                {/* Fun-Zone Quiz */}
                {event.quiz && event.quiz.length > 0 && (
                  <QuizPlayer quiz={event.quiz} t={t} />
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
