/**
 * Admin / Organizer Seite
 *
 * Zeigt alle Events des Admins. Nach Auswahl eines Events:
 * - Event bearbeiten (Daten ändern)
 * - Teilnehmerliste anzeigen
 * - Teilnehmerliste in SharePoint öffnen
 * - Neues Event erstellen
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { DeloitteEvent } from '../types';
import { SPRegistration } from '../services/EventService';
import { Plus, Users, FileText, Trash2, Copy, Mail } from './Icons';

function formatDate(iso: string): string {
  if (!iso) return '-';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '-';
  return d.toLocaleDateString('de-DE', {
    day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  });
}

function getStatusColor(status: string): string {
  switch (status) {
    case 'Active': return 'var(--dex-green)';
    case 'Completed': return 'var(--dex-gray-400)';
    case 'Cancelled': return 'var(--dex-red)';
    default: return 'var(--dex-orange)';
  }
}

export default function AdminPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { events, getAllRegistrations, deleteEvent } = useEvents();
  const { currentUser } = useCurrentUser();
  const { isAdmin, siteUrl } = useRoles();
  const [selectedEvent, setSelectedEvent] = React.useState<DeloitteEvent | null>(null);
  const [registrations, setRegistrations] = React.useState<SPRegistration[]>([]);
  const [isLoadingRegs, setIsLoadingRegs] = React.useState(false);
  const [deletingId, setDeletingId] = React.useState<string | null>(null);
  const [isDeleting, setIsDeleting] = React.useState(false);
  const [copiedEmails, setCopiedEmails] = React.useState(false);

  // SuperAdmin sieht alle Events, EventAdmin nur seine
  const adminEvents = isAdmin
    ? events
    : events.filter(e => e.organizers.some(o => o.toLowerCase().includes(currentUser.surname.toLowerCase())));

  const handleSelectEvent = async (event: DeloitteEvent): Promise<void> => {
    setSelectedEvent(event);
    setIsLoadingRegs(true);
    try {
      const regs = await getAllRegistrations(event.id);
      setRegistrations(regs);
    } catch {
      setRegistrations([]);
    }
    setIsLoadingRegs(false);
  };

  // Teilnehmerlisten-URL aus regListMap ableiten
  const getRegListUrl = (event: DeloitteEvent): string => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    const base = ctx ? ctx.pageContext.web.absoluteUrl : siteUrl;
    // Event-ID nutzen um den Listennamen zu finden
    // Der Listenname ist in der SPEvent gespeichert, hier nutzen wir die events aus dem Context
    return `${base}/Lists`;
  };

  if (!selectedEvent) {
    // Event-Auswahl
    return (
      <div className="page-container">
        <div className="flex-between mb-16">
          <h2>Admin / Organizer</h2>
          <div style={{ display: 'flex', gap: 8 }}>
            {isAdmin && (
              <button className="btn btn-secondary" onClick={() => navigate('participants')} style={{ fontSize: '0.85rem' }}>
                <Users size={16} /> Teilnehmer
              </button>
            )}
            {isAdmin && (
              <button className="btn btn-secondary" onClick={() => navigate('flowcharts')} style={{ fontSize: '0.85rem' }}>
                ↻ Prozesse
              </button>
            )}
            <a
              href={`${siteUrl}/Lists/DEX_Events/AllItems.aspx`}
              target="_blank"
              rel="noopener noreferrer"
              className="btn btn-secondary"
              style={{ fontSize: '0.85rem', textDecoration: 'none' }}
            >
              <FileText size={16} /> SharePoint-Liste
            </a>
            <button className="btn btn-primary" onClick={() => navigate('create-event')} style={{ fontSize: '0.85rem' }}>
              <Plus size={16} /> Neues Event erstellen
            </button>
          </div>
        </div>

        {adminEvents.length === 0 ? (
          <div className="card text-center" style={{ padding: 48 }}>
            <p style={{ color: 'var(--dex-gray-400)' }}>Keine Events gefunden.</p>
            <button className="btn btn-primary mt-24" onClick={() => navigate('create-event')}>
              Erstes Event erstellen
            </button>
          </div>
        ) : (
          <div className="my-events-list">
            {adminEvents.map(event => (
              <div
                key={event.id}
                className="card card-clickable"
                style={{ padding: '20px 24px', cursor: 'pointer' }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div onClick={() => handleSelectEvent(event)} style={{ flex: 1, display: 'flex', alignItems: 'center', gap: 16 }}>
                    {event.imageUrl && (
                      <div style={{
                        width: 60, height: 40, borderRadius: 'var(--dex-radius)', flexShrink: 0,
                        background: `url(${event.imageUrl}) center/cover no-repeat`,
                      }} />
                    )}
                    <div>
                    <h3 style={{ marginBottom: 4 }}>{event.title}</h3>
                    <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', margin: 0 }}>
                      {formatDate(event.startDate)} - {formatDate(event.endDate)}
                      {event.location ? ` · ${event.location}` : ''}
                    </p>
                    </div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
                    <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                      {event.currentParticipants}/{event.maxParticipants || '∞'} Teilnehmer
                    </span>
                    <span className="badge" style={{
                      background: getStatusColor(event.status) + '22',
                      color: getStatusColor(event.status),
                    }}>
                      {event.status}
                    </span>
                    <button
                      className={`btn ${deletingId === event.id ? 'btn-danger' : 'btn-secondary'}`}
                      style={{ fontSize: '0.8rem', padding: '6px 12px' }}
                      onClick={async (e) => {
                        e.stopPropagation();
                        if (deletingId === event.id) {
                          setIsDeleting(true);
                          await deleteEvent(event.id);
                          setDeletingId(null);
                          setIsDeleting(false);
                        } else {
                          setDeletingId(event.id);
                        }
                      }}
                      disabled={isDeleting}
                    >
                      <Trash2 size={14} />
                      {deletingId === event.id
                        ? (isDeleting ? 'Wird gelöscht...' : 'Wirklich löschen?')
                        : 'Löschen'}
                    </button>
                    {deletingId === event.id && !isDeleting && (
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 12px' }}
                        onClick={(e) => { e.stopPropagation(); setDeletingId(null); }}
                      >
                        Abbrechen
                      </button>
                    )}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  // Event ausgewählt - Detail-Ansicht
  const activeRegs = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'Eingecheckt');
  const waitlistRegs = registrations.filter(r => r.Status === 'Warteliste');
  const cancelledRegs = registrations.filter(r => r.Status === 'Abgemeldet');

  return (
    <div className="page-container">
      <div className="flex-between mb-16">
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <button
            className="back-btn"
            onClick={() => { setSelectedEvent(null); setRegistrations([]); }}
            aria-label="Zurück"
          >
            ←
          </button>
          <h2 style={{ margin: 0 }}>{selectedEvent.title}</h2>
          <span className="badge" style={{
            background: getStatusColor(selectedEvent.status) + '22',
            color: getStatusColor(selectedEvent.status),
          }}>
            {selectedEvent.status}
          </span>
        </div>
      </div>

      {/* Event-Bild */}
      {selectedEvent.imageUrl && (
        <div style={{
          height: 180, borderRadius: 'var(--dex-radius-lg)',
          background: `url(${selectedEvent.imageUrl}) center/cover no-repeat`,
          marginBottom: 24,
        }} />
      )}

      {/* Event-Info + Aktionen */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24, marginBottom: 24 }}>
        <div className="card" style={{ padding: 24 }}>
          <h3 className="mb-16">Event-Details</h3>
          <div className="settings-info">
            <div className="settings-info__row">
              <span className="settings-info__label">Zeitraum</span>
              <span>{formatDate(selectedEvent.startDate)} - {formatDate(selectedEvent.endDate)}</span>
            </div>
            <div className="settings-info__row">
              <span className="settings-info__label">Ort</span>
              <span>{selectedEvent.location || '-'}</span>
            </div>
            <div className="settings-info__row">
              <span className="settings-info__label">Max. Teilnehmer</span>
              <span>{selectedEvent.maxParticipants || 'Unbegrenzt'}</span>
            </div>
            <div className="settings-info__row">
              <span className="settings-info__label">Aktuell registriert</span>
              <span>{activeRegs.length}</span>
            </div>
            {waitlistRegs.length > 0 && (
              <div className="settings-info__row">
                <span className="settings-info__label">Warteliste</span>
                <span>{waitlistRegs.length}</span>
              </div>
            )}
          </div>
        </div>

        <div className="card" style={{ padding: 24 }}>
          <h3 className="mb-16">Aktionen</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <button
              className="btn btn-primary btn-block"
              onClick={() => navigate('edit-event', selectedEvent.id)}
            >
              <FileText size={16} /> Event bearbeiten
            </button>
            <a
              href={selectedEvent.subsiteUrl
                ? `${selectedEvent.subsiteUrl}/Lists/Teilnehmer/AllItems.aspx`
                : `${siteUrl}/Lists`}
              target="_blank"
              rel="noopener noreferrer"
              className="btn btn-primary btn-block"
              style={{ textDecoration: 'none', textAlign: 'center' }}
            >
              <FileText size={16} /> Teilnehmerliste in SharePoint öffnen
            </a>
            <button
              className="btn btn-secondary btn-block"
              onClick={() => {
                const emails = registrations
                  .filter(r => r.Status === 'Angemeldet' || r.Status === 'Eingecheckt')
                  .map(r => r.ParticipantEmail)
                  .join('; ');
                if (emails) {
                  navigator.clipboard.writeText(emails).then(() => {
                    setCopiedEmails(true);
                    setTimeout(() => setCopiedEmails(false), 2000);
                  }).catch(() => {
                    // Fallback: prompt
                    window.prompt('E-Mail-Adressen kopieren:', emails);
                  });
                }
              }}
            >
              <Copy size={16} /> {copiedEmails ? 'Kopiert!' : 'E-Mail-Adressen kopieren'}
            </button>
            <a
              href={`mailto:${registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'Eingecheckt').map(r => r.ParticipantEmail).join(';')}`}
              className="btn btn-secondary btn-block"
              style={{ textDecoration: 'none', textAlign: 'center' }}
            >
              <Mail size={16} /> E-Mail an alle Teilnehmer
            </a>
          </div>
        </div>
      </div>

      {/* Teilnehmerliste */}
      <div className="card" style={{ padding: 24 }}>
        <div className="flex-between mb-16">
          <h3 style={{ margin: 0 }}>
            <Users size={18} /> Teilnehmer ({activeRegs.length})
          </h3>
        </div>

        {isLoadingRegs ? (
          <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>Lade Teilnehmer...</p>
        ) : activeRegs.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-400)' }}>Noch keine Teilnehmer registriert.</p>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                  <th style={{ textAlign: 'left', padding: 8 }}>#</th>
                  <th style={{ textAlign: 'left', padding: 8 }}>Name</th>
                  <th style={{ textAlign: 'left', padding: 8 }}>Email</th>
                  <th style={{ textAlign: 'left', padding: 8 }}>Status</th>
                  <th style={{ textAlign: 'left', padding: 8 }}>Registriert am</th>
                </tr>
              </thead>
              <tbody>
                {activeRegs.map((reg, i) => (
                  <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                    <td style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{reg.Title || (i + 1)}</td>
                    <td style={{ padding: 8, fontWeight: 500 }}>{(reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName}</td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                    <td style={{ padding: 8 }}>
                      <span className="badge badge-green">{reg.Status}</span>
                    </td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {waitlistRegs.length > 0 && (
          <>
            <h4 style={{ marginTop: 24, color: 'var(--dex-orange)' }}>Warteliste ({waitlistRegs.length})</h4>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                <tbody>
                  {waitlistRegs.map(reg => (
                    <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                      <td style={{ padding: 8, fontWeight: 500 }}>{(reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName}</td>
                      <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                      <td style={{ padding: 8 }}>
                        <span className="badge badge-orange">Warteliste</span>
                      </td>
                      <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}

        {cancelledRegs.length > 0 && (
          <>
            <h4 style={{ marginTop: 24, color: 'var(--dex-gray-400)' }}>Abgemeldet ({cancelledRegs.length})</h4>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem', opacity: 0.6 }}>
                <tbody>
                  {cancelledRegs.map(reg => (
                    <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                      <td style={{ padding: 8 }}>{(reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName}</td>
                      <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                      <td style={{ padding: 8 }}>
                        <span className="badge badge-red">Abgemeldet</span>
                      </td>
                      <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.CancellationDate)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
