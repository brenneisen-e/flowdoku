/**
 * Meine Events - zeigt alle Registrierungen des aktuellen Users.
 * Aufgeteilt in aktive und stornierte Registrierungen.
 *
 * TODO: Registrierungen ueber den EventContext verwalten
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { myRegistrations } from '../data/mockData';
import { Registration } from '../types';

function formatDate(iso: string): string {
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

export default function MyEventsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { events } = useEvents();
  const [registrations, setRegistrations] = React.useState<Registration[]>(myRegistrations);
  const [cancellingId, setCancellingId] = React.useState<string | null>(null);

  // Zwei-Schritt-Stornierung
  const handleCancel = (regId: string): void => {
    if (cancellingId === regId) {
      setRegistrations(prev =>
        prev.map(r =>
          r.id === regId ? { ...r, status: 'Cancelled' as const, cancellationDate: new Date().toISOString() } : r
        )
      );
      setCancellingId(null);
    } else {
      setCancellingId(regId);
    }
  };

  const activeRegs = registrations.filter(r => r.status !== 'Cancelled');
  const cancelledRegs = registrations.filter(r => r.status === 'Cancelled');

  return (
    <div className="page-container">
      <h2 className="mb-16">My Registrations</h2>

      {activeRegs.length === 0 && cancelledRegs.length === 0 && (
        <div className="card text-center" style={{ padding: 48 }}>
          <p style={{ color: 'var(--dex-gray-400)' }}>You are not registered for any events yet.</p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>Browse Events</button>
        </div>
      )}

      {activeRegs.length > 0 && (
        <div className="my-events-list">
          {activeRegs.map(reg => {
            const event = events.find(e => e.id === reg.eventId);
            return (
              <div key={reg.id} className="card my-event-card">
                <div className="my-event-card__header">
                  <h3>{reg.eventTitle}</h3>
                  <span className={`badge ${getStatusBadgeClass(reg.status)}`}>{reg.status}</span>
                </div>
                {event && (
                  <div className="my-event-card__details">
                    <p><strong>Location:</strong> {event.location}</p>
                    <p><strong>Date:</strong> {formatDate(event.startDate)} - {formatDate(event.endDate)}</p>
                    <p><strong>Registered on:</strong> {formatDate(reg.registrationDate)}</p>
                    {reg.status === 'Waitlist' && reg.waitlistPosition && (
                      <p className="text-red"><strong>Waitlist position:</strong> {reg.waitlistPosition}</p>
                    )}
                  </div>
                )}
                {Object.keys(reg.eventSpecificData).length > 0 && (
                  <div className="my-event-card__specific">
                    {Object.keys(reg.eventSpecificData).map(key => (
                      <span key={key} className="badge badge-gray" style={{ marginRight: 8, marginBottom: 4 }}>
                        {key}: {reg.eventSpecificData[key]}
                      </span>
                    ))}
                  </div>
                )}
                <div className="my-event-card__actions">
                  <button
                    className={`btn ${cancellingId === reg.id ? 'btn-danger' : 'btn-secondary'}`}
                    onClick={() => handleCancel(reg.id)}
                    style={{ fontSize: '0.85rem' }}
                  >
                    {cancellingId === reg.id ? 'Confirm Cancellation' : 'Cancel Registration'}
                  </button>
                  {cancellingId === reg.id && (
                    <button className="btn btn-secondary" onClick={() => setCancellingId(null)} style={{ fontSize: '0.85rem' }}>
                      Keep Registration
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {cancelledRegs.length > 0 && (
        <div>
          <h3 className="mt-24 mb-16" style={{ color: 'var(--dex-gray-400)' }}>Cancelled Registrations</h3>
          <div className="my-events-list">
            {cancelledRegs.map(reg => (
              <div key={reg.id} className="card my-event-card" style={{ opacity: 0.6 }}>
                <div className="my-event-card__header">
                  <h3>{reg.eventTitle}</h3>
                  <span className="badge badge-red">Cancelled</span>
                </div>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-400)' }}>
                  Cancelled on: {reg.cancellationDate ? formatDate(reg.cancellationDate) : 'N/A'}
                </p>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
