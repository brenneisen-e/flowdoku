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
import { useLanguage } from '../context/LanguageContext';
import { DeloitteEvent } from '../types';
import { SPRegistration } from '../services/EventService';
import { Plus, Users, FileText, Trash2, Copy, Mail } from './Icons';
import { EventService } from '../services/EventService';
import { qrCodeEmail, cancellationEmail } from '../services/EmailTemplates';
import * as QRCode from 'qrcode';

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
  const { events, isEventsLoading, getAllRegistrations, deleteEvent } = useEvents();
  const { currentUser } = useCurrentUser();
  const { isAdmin, siteUrl } = useRoles();
  const { t } = useLanguage();
  const [selectedEvent, setSelectedEvent] = React.useState<DeloitteEvent | null>(null);
  const [registrations, setRegistrations] = React.useState<SPRegistration[]>([]);
  const [isLoadingRegs, setIsLoadingRegs] = React.useState(false);
  const [deletingId, setDeletingId] = React.useState<string | null>(null);
  const [isDeleting, setIsDeleting] = React.useState(false);
  const [copiedEmails, setCopiedEmails] = React.useState(false);
  const [isSendingQR, setIsSendingQR] = React.useState(false);
  const [qrSentCount, setQrSentCount] = React.useState(0);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const spfxContext = (window as any).__dexSpfxContext;
  const eventServiceRef = React.useMemo(() => spfxContext ? new EventService(spfxContext) : null, []);

  // SuperAdmin sieht alle Events, EventAdmin nur seine
  const adminEvents = isAdmin
    ? events
    : events.filter(e => {
      const fullName = `${currentUser.firstName} ${currentUser.surname}`.toLowerCase();
      return e.organizers.some(o => o.toLowerCase().includes(fullName) || o.toLowerCase().includes(currentUser.surname.toLowerCase()));
    });

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
        <h2 className="mb-16">{t('admin.title')}</h2>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
          {isAdmin && (
            <button className="btn btn-secondary" onClick={() => navigate('participants')} style={{ fontSize: '0.85rem' }}>
              <Users size={16} /> {t('admin.participants')}
            </button>
          )}
          {isAdmin && (
            <button className="btn btn-secondary" onClick={() => navigate('flowcharts')} style={{ fontSize: '0.85rem' }}>
              ↻ {t('admin.processes')}
            </button>
          )}
          <a
            href={`${siteUrl}/Lists/DEX_Events/AllItems.aspx`}
            target="_blank"
            rel="noopener noreferrer"
            className="btn btn-secondary"
            style={{ fontSize: '0.85rem', textDecoration: 'none' }}
          >
            <FileText size={16} /> {t('admin.splist')}
          </a>
          <button className="btn btn-primary" onClick={() => navigate('create-event')} style={{ fontSize: '0.85rem' }}>
            <Plus size={16} /> {t('admin.newevent')}
          </button>
        </div>

        {isEventsLoading ? (
          <div className="card text-center" style={{ padding: 48 }}>
            <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 16 }}>
              <div style={{
                width: 40, height: 40, border: '4px solid var(--dex-gray-200)',
                borderTop: '4px solid var(--dex-green)', borderRadius: '50%',
                animation: 'dexOrbSpin 1s linear infinite',
              }} />
            </div>
            <p style={{ color: 'var(--dex-gray-400)' }}>Events werden geladen...</p>
          </div>
        ) : adminEvents.length === 0 ? (
          <div className="card text-center" style={{ padding: 48 }}>
            <p style={{ color: 'var(--dex-gray-400)' }}>{t('admin.noevents')}</p>
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
                <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
                  <div onClick={() => handleSelectEvent(event)} style={{ flex: '1 1 200px', display: 'flex', alignItems: 'center', gap: 16, cursor: 'pointer' }}>
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
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
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
  const activeRegs = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
  const waitlistRegs = registrations.filter(r => r.Status === 'Warteliste')
    .sort((a, b) => new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
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
              <FileText size={16} /> {t('admin.editbutton')}
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
              <FileText size={16} /> {t('admin.opensp')}
            </a>
            <button
              className="btn btn-secondary btn-block"
              onClick={() => {
                const emails = registrations
                  .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
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
              <Copy size={16} /> {copiedEmails ? t('admin.copied') : t('admin.copyemails')}
            </button>
            <a
              href={`mailto:${registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').map(r => r.ParticipantEmail).join(';')}`}
              className="btn btn-secondary btn-block"
              style={{ textDecoration: 'none', textAlign: 'center' }}
            >
              <Mail size={16} /> {t('admin.emailall')}
            </a>
          </div>
        </div>
      </div>

      {/* Zähler + QR/Check-in Aktionen */}
      <div className="admin-counters" style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 12, marginBottom: 24 }}>
        <div className="card" style={{ padding: 16, textAlign: 'center' }}>
          <div style={{ fontSize: '1.8rem', fontWeight: 700, color: '#1565c0' }}>
            {registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length}
          </div>
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.registered')}</div>
        </div>
        <div className="card" style={{ padding: 16, textAlign: 'center' }}>
          <div style={{ fontSize: '1.8rem', fontWeight: 700, color: '#6a1b9a' }}>
            {registrations.filter(r => r.Status === 'QR versendet').length}
          </div>
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.qrsent')}</div>
        </div>
        <div className="card" style={{ padding: 16, textAlign: 'center' }}>
          <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-green)' }}>
            {registrations.filter(r => r.Status === 'Eingecheckt').length}
          </div>
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.checkedin')}</div>
        </div>
        <div className="card" style={{ padding: 16, textAlign: 'center' }}>
          <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-orange)' }}>
            {registrations.filter(r => r.Status === 'Warteliste').length}
          </div>
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.waitlist')}</div>
        </div>
        <div className="card" style={{ padding: 16, textAlign: 'center' }}>
          <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-gray-400)' }}>
            {registrations.filter(r => r.Status === 'Abgemeldet').length}
          </div>
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.cancelled')}</div>
        </div>
      </div>

      <div className="admin-actions" style={{ display: 'flex', gap: 12, marginBottom: 24 }}>
        <button
          className="btn btn-primary"
          onClick={() => navigate('check-in', selectedEvent.id)}
          style={{ flex: 1 }}
        >
          {t('admin.checkin')}
        </button>
        <button
          className="btn btn-secondary"
          disabled={isSendingQR}
          onClick={async () => {
            if (!eventServiceRef || !selectedEvent) return;
            const eligible = registrations.filter(r => r.Status === 'Angemeldet');
            if (eligible.length === 0) return;
            if (!window.confirm(`QR-Codes an ${eligible.length} Teilnehmer versenden?`)) return;

            setIsSendingQR(true);
            setQrSentCount(0);
            let sent = 0;

            for (const reg of eligible) {
              const qrData = `DEX|${selectedEvent.eventNumber}|${reg.ParticipantEmail}`;
              const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
              // QR-Code als Base64-Bild generieren
              let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
              try {
                const qrDataUrl = await QRCode.toDataURL(qrData, { width: 300, margin: 2 });
                qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
              } catch { /* Fallback: Text */ }
              // QR-Code E-Mail im Deloitte-Template queuen
              const emailData = qrCodeEmail(name, selectedEvent.title, qrImageHtml);
              await eventServiceRef.queueEmail(
                emailData.subject,
                reg.ParticipantEmail,
                name,
                emailData.body,
                'QRCode',
                selectedEvent.title,
                selectedEvent.id
              );
              // Status auf 'QR versendet' setzen
              if (selectedEvent.subsiteUrl) {
                await eventServiceRef.setQRSentStatus(selectedEvent.subsiteUrl, reg.Id);
              }
              sent++;
              setQrSentCount(sent);
            }

            // Registrierungen neu laden
            const regs = await getAllRegistrations(selectedEvent.id);
            setRegistrations(regs);
            setIsSendingQR(false);
          }}
          style={{ flex: 1 }}
        >
          {isSendingQR ? `QR-Codes werden versendet... (${qrSentCount})` : `QR-Codes versenden (${registrations.filter(r => r.Status === 'Angemeldet').length})`}
        </button>
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
                  <th style={{ textAlign: 'left', padding: 8 }}>Aktion</th>
                </tr>
              </thead>
              <tbody>
                {activeRegs.map((reg, i) => (
                  <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                    <td style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{reg.TeilnehmerID || (i + 1)}</td>
                    <td style={{ padding: 8, fontWeight: 500 }}>{(reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName}</td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                    <td style={{ padding: 8 }}>
                      <span className={`badge ${reg.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>{reg.Status}</span>
                    </td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                    <td style={{ padding: 8, display: 'flex', gap: 4 }}>
                      {reg.Status === 'Eingecheckt' ? (
                        <button
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                          onClick={async () => {
                            if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                            await eventServiceRef.checkOutParticipant(selectedEvent.subsiteUrl, reg.Id);
                            const regs = await getAllRegistrations(selectedEvent.id);
                            setRegistrations(regs);
                          }}
                        >
                          Auschecken
                        </button>
                      ) : (
                        <button
                          className="btn btn-primary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                          onClick={async () => {
                            if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                            await eventServiceRef.checkInParticipant(selectedEvent.subsiteUrl, reg.Id);
                            const regs = await getAllRegistrations(selectedEvent.id);
                            setRegistrations(regs);
                          }}
                        >
                          Einchecken
                        </button>
                      )}
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
                        onClick={async () => {
                          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                          const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                          if (!confirm(`${name} (${reg.ParticipantEmail}) wirklich abmelden?`)) return;
                          await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id);
                          // Abmelde-Email und Outlook-Ausladen in Queue eintragen
                          if (reg.ParticipantEmail) {
                            const emailData = cancellationEmail(name, selectedEvent.title);
                            eventServiceRef.queueEmail(
                              emailData.subject, reg.ParticipantEmail, name, emailData.body,
                              'Abmeldung', selectedEvent.title, selectedEvent.id
                            ).catch(err => console.warn('[DEX]', err));
                            eventServiceRef.queueOutlookEvent(
                              reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
                            ).catch(err => console.warn('[DEX]', err));
                          }
                          // DEX_Participants aufraeumen
                          if (reg.ParticipantEmail && selectedEvent.eventNumber) {
                            eventServiceRef.removeParticipantEvent(
                              reg.ParticipantEmail, selectedEvent.eventNumber
                            ).catch(err => console.warn('[DEX]', err));
                          }
                          // IDReorder in Queue eintragen
                          if (selectedEvent.subsiteUrl) {
                            eventServiceRef.queueIDReorder(
                              selectedEvent.id, selectedEvent.eventNumber || 0,
                              selectedEvent.subsiteUrl, selectedEvent.title
                            ).catch(err => console.warn('[DEX]', err));
                          }
                          const regs = await getAllRegistrations(selectedEvent.id);
                          setRegistrations(regs);
                        }}
                      >
                        Abmelden
                      </button>
                    </td>
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
                <thead>
                  <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                    <th style={{ textAlign: 'left', padding: 8 }}>Platz</th>
                    <th style={{ textAlign: 'left', padding: 8 }}>Name</th>
                    <th style={{ textAlign: 'left', padding: 8 }}>Email</th>
                    <th style={{ textAlign: 'left', padding: 8 }}>Registriert am</th>
                    <th style={{ textAlign: 'left', padding: 8 }}>Aktion</th>
                  </tr>
                </thead>
                <tbody>
                  {waitlistRegs.map((reg, i) => (
                    <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                      <td style={{ padding: 8, fontWeight: 600, color: 'var(--dex-orange)' }}>{i + 1}</td>
                      <td style={{ padding: 8, fontWeight: 500 }}>{(reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName}</td>
                      <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                      <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                      <td style={{ padding: 8 }}>
                        <button
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
                          onClick={async () => {
                            if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                            const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                            if (!confirm(`${name} von der Warteliste entfernen?`)) return;
                            await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id);
                            if (reg.ParticipantEmail) {
                              const emailData = cancellationEmail(name, selectedEvent.title);
                              eventServiceRef.queueEmail(
                                emailData.subject, reg.ParticipantEmail, name, emailData.body,
                                'Abmeldung', selectedEvent.title, selectedEvent.id
                              ).catch(err => console.warn('[DEX]', err));
                            }
                            if (reg.ParticipantEmail && selectedEvent.eventNumber) {
                              eventServiceRef.removeParticipantEvent(reg.ParticipantEmail, selectedEvent.eventNumber).catch(err => console.warn('[DEX]', err));
                            }
                            const regs = await getAllRegistrations(selectedEvent.id);
                            setRegistrations(regs);
                          }}
                        >
                          Entfernen
                        </button>
                      </td>
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
