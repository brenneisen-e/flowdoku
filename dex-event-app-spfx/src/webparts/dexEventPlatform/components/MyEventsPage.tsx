/**
 * Meine Events - zeigt alle Events fuer die der User registriert ist.
 * Laedt Registrierungen aus den jeweiligen Teilnehmerlisten.
 * Ermoeglicht Abmeldung mit Zwei-Schritt-Bestaetigung.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { DeloitteEvent } from '../types';
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
    case 'Registered': return 'badge-green';
    case 'Waitlist': return 'badge-orange';
    case 'Cancelled': return 'badge-red';
    case 'Checked-In': return 'badge-green';
    default: return 'badge-gray';
  }
}

function getStatusLabel(status: string): string {
  switch (status) {
    case 'Registered': return 'Angemeldet';
    case 'Waitlist': return 'Warteliste';
    case 'Cancelled': return 'Abgemeldet';
    case 'Checked-In': return 'Eingecheckt';
    default: return status;
  }
}

export default function MyEventsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { events, getMyRegistration, cancelRegistration } = useEvents();
  const [myEvents, setMyEvents] = React.useState<MyEventEntry[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [cancellingId, setCancellingId] = React.useState<string | null>(null);
  const [isCancelling, setIsCancelling] = React.useState(false);

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

  const activeEntries = myEvents.filter(e => e.registration.Status !== 'Cancelled');
  const cancelledEntries = myEvents.filter(e => e.registration.Status === 'Cancelled');

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
          <p style={{ color: 'var(--dex-gray-400)' }}>Du bist noch fuer kein Event registriert.</p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>Events durchsuchen</button>
        </div>
      )}

      {activeEntries.length > 0 && (
        <div className="my-events-list">
          {activeEntries.map(({ event, registration }) => {
            // Custom Data parsen
            let customData: Record<string, string> = {};
            try {
              if (registration.CustomData) customData = JSON.parse(registration.CustomData);
            } catch { /* */ }

            return (
              <div key={event.id} className="card my-event-card">
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
                    <p><strong>Teilnehmer-Nr.:</strong> {registration.Title}</p>
                  )}
                </div>

                {Object.keys(customData).length > 0 && (
                  <div className="my-event-card__specific">
                    {Object.keys(customData).map(key => (
                      <span key={key} className="badge badge-gray" style={{ marginRight: 8, marginBottom: 4 }}>
                        {key}: {customData[key]}
                      </span>
                    ))}
                  </div>
                )}

                <div className="my-event-card__actions">
                  <button
                    className={`btn ${cancellingId === event.id ? 'btn-danger' : 'btn-secondary'}`}
                    onClick={() => handleCancel(event.id)}
                    disabled={isCancelling}
                    style={{ fontSize: '0.85rem' }}
                  >
                    {cancellingId === event.id
                      ? (isCancelling ? 'Wird abgemeldet...' : 'Abmeldung bestaetigen')
                      : 'Abmelden'}
                  </button>
                  {cancellingId === event.id && !isCancelling && (
                    <button className="btn btn-secondary" onClick={() => setCancellingId(null)} style={{ fontSize: '0.85rem' }}>
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
