/**
 * Meine Events - zeigt alle Events fuer die der User registriert ist.
 * Laedt Registrierungen aus den jeweiligen Teilnehmerlisten.
 * Ermoeglicht Abmeldung mit Zwei-Schritt-Bestaetigung.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { DeloitteEvent, EventSpecificField } from '../types';
import { SPRegistration } from '../services/EventService';

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

function getStatusBadgeClass(status: string): string {
  switch (status) {
    case 'Angemeldet': return 'badge-green';
    case 'Warteliste': return 'badge-orange';
    case 'Abgemeldet': return 'badge-red';
    case 'Eingecheckt': return 'badge-green';
    default: return 'badge-gray';
  }
}

function getStatusLabel(status: string): string {
  switch (status) {
    case 'Angemeldet': return 'Angemeldet';
    case 'Warteliste': return 'Warteliste';
    case 'Abgemeldet': return 'Abgemeldet';
    case 'Eingecheckt': return 'Eingecheckt';
    default: return status;
  }
}

export default function MyEventsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { events, getMyRegistration, cancelRegistration, updateMyRegistration } = useEvents();
  const [myEvents, setMyEvents] = React.useState<MyEventEntry[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [cancellingId, setCancellingId] = React.useState<string | null>(null);
  const [isCancelling, setIsCancelling] = React.useState(false);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [editData, setEditData] = React.useState<Record<string, string>>({});
  const [isSaving, setIsSaving] = React.useState(false);

  React.useEffect(() => {
    loadMyRegistrations();
  }, [events]);

  async function loadMyRegistrations(): Promise<void> {
    setIsLoading(true);
    const entries: MyEventEntry[] = [];

    for (const event of events) {
      try {
        const reg = await getMyRegistration(event.id);
        if (reg) {
          entries.push({ event, registration: reg });
        }
      } catch {
        // Kein Zugriff auf diese Teilnehmerliste
      }
    }

    setMyEvents(entries);
    setIsLoading(false);
  }

  const handleCancel = async (eventId: string): Promise<void> => {
    if (cancellingId === eventId) {
      // Zweiter Klick = Bestaetigung
      setIsCancelling(true);
      const success = await cancelRegistration(eventId);
      if (success) {
        // Registrierung neu laden
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
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>Lade deine Registrierungen...</p>
      </div>
    );
  }

  return (
    <div className="page-container">
      <h2 className="mb-16">Meine Events</h2>

      {activeEntries.length === 0 && cancelledEntries.length === 0 && (
        <div className="card text-center" style={{ padding: 48 }}>
          <p style={{ color: 'var(--dex-gray-400)' }}>Du bist noch für kein Event registriert.</p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>Events durchsuchen</button>
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
              <div key={event.id} className="card my-event-card">
                {event.imageUrl && (
                  <div style={{
                    height: 120, borderRadius: 'var(--dex-radius) var(--dex-radius) 0 0',
                    background: `url(${event.imageUrl}) center/cover no-repeat`,
                    marginBottom: 0,
                  }} />
                )}
                <div className="my-event-card__header">
                  <h3>{event.title}</h3>
                  <span className={`badge ${getStatusBadgeClass(registration.Status)}`}>
                    {getStatusLabel(registration.Status)}
                  </span>
                </div>

                <div className="my-event-card__details">
                  <p><strong>Ort:</strong> {event.location || '-'}</p>
                  <p><strong>Datum:</strong> {formatDate(event.startDate)} - {formatDate(event.endDate)}</p>
                  <p><strong>Angemeldet am:</strong> {formatDate(registration.RegistrationDate)}</p>
                  {registration.Title && (
                    <p><strong>E-Mail-Adresse:</strong> {registration.Title}</p>
                  )}
                </div>

                {/* Anzeige oder Bearbeitung der Custom Data */}
                {editingId === event.id ? (
                  <div className="my-event-card__specific" style={{ padding: '12px 0' }}>
                    {event.eventSpecificFields.map((field: EventSpecificField) => (
                      <div className="form-group" key={field.id} style={{ marginBottom: 12 }}>
                        <label className="form-label" style={{ fontSize: '0.85rem' }}>
                          {field.required && <span className="required">*</span>}
                          {field.label}
                        </label>
                        {field.type === 'select' ? (
                          <select
                            className="form-select"
                            value={editData[field.id] || ''}
                            onChange={e => setEditData({ ...editData, [field.id]: e.target.value })}
                          >
                            <option value="">Bitte wählen</option>
                            {field.options && field.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        ) : (
                          <input
                            className="form-input"
                            value={editData[field.id] || ''}
                            onChange={e => setEditData({ ...editData, [field.id]: e.target.value })}
                            placeholder={field.label}
                            type={field.type === 'number' ? 'number' : 'text'}
                          />
                        )}
                      </div>
                    ))}
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button
                        className="btn btn-primary"
                        style={{ fontSize: '0.85rem' }}
                        disabled={isSaving}
                        onClick={async () => {
                          setIsSaving(true);
                          await updateMyRegistration(event.id, editData);
                          await loadMyRegistrations();
                          setEditingId(null);
                          setIsSaving(false);
                        }}
                      >
                        {isSaving ? 'Wird gespeichert...' : 'Speichern'}
                      </button>
                      <button className="btn btn-secondary" style={{ fontSize: '0.85rem' }} onClick={() => setEditingId(null)}>
                        Abbrechen
                      </button>
                    </div>
                  </div>
                ) : displayData.length > 0 ? (
                  <div className="my-event-card__specific" style={{ marginBottom: 12 }}>
                    <div style={{ display: 'grid', gridTemplateColumns: 'auto 1fr', gap: '4px 16px', fontSize: '0.9rem' }}>
                      {displayData.map(({ label, value }) => (
                        <React.Fragment key={label}>
                          <span style={{ fontWeight: 600, color: 'var(--dex-gray-700)' }}>{label}:</span>
                          <span style={{ color: 'var(--dex-gray-600)' }}>{value}</span>
                        </React.Fragment>
                      ))}
                    </div>
                  </div>
                ) : null}

                <div className="my-event-card__actions">
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.85rem' }}
                    onClick={() => {
                      if (editingId === event.id) {
                        setEditingId(null);
                      } else {
                        setEditData(customData);
                        setEditingId(event.id);
                      }
                    }}
                  >
                    {editingId === event.id ? 'Abbrechen' : 'Angaben bearbeiten'}
                  </button>
                  <button
                    className="btn"
                    onClick={() => handleCancel(event.id)}
                    disabled={isCancelling}
                    style={{
                      fontSize: '0.85rem',
                      background: 'var(--dex-orange)', color: '#fff',
                    }}
                  >
                    {cancellingId === event.id
                      ? (isCancelling ? 'Wird abgemeldet...' : 'Abmeldung bestätigen')
                      : 'Abmelden'}
                  </button>
                  {cancellingId === event.id && !isCancelling && (
                    <button className="btn btn-primary" onClick={() => setCancellingId(null)} style={{ fontSize: '0.85rem' }}>
                      Anmeldung behalten
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {cancelledEntries.length > 0 && (
        <div>
          <h3 className="mt-24 mb-16" style={{ color: 'var(--dex-gray-400)' }}>Abgemeldete Events</h3>
          <div className="my-events-list">
            {cancelledEntries.map(({ event, registration }) => (
              <div key={event.id} className="card my-event-card" style={{ opacity: 0.6 }}>
                <div className="my-event-card__header">
                  <h3>{event.title}</h3>
                  <span className="badge badge-red">Abgemeldet</span>
                </div>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-400)' }}>
                  Abgemeldet am: {registration.CancellationDate ? formatDate(registration.CancellationDate) : '-'}
                </p>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
