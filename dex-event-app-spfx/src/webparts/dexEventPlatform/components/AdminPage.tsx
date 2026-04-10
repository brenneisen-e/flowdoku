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
import { Plus, Users, FileText, Trash2, Copy, Mail, Send } from './Icons';
import { EventService } from '../services/EventService';
import { qrCodeEmail, cancellationEmail, wrapTemplate } from '../services/EmailTemplates';
import * as QRCode from 'qrcode';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';

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
  const [regLoadError, setRegLoadError] = React.useState('');
  const [deletingId, setDeletingId] = React.useState<string | null>(null);
  const [isDeleting, setIsDeleting] = React.useState(false);
  const [copiedEmails, setCopiedEmails] = React.useState(false);
  const [isSendingQR, setIsSendingQR] = React.useState(false);
  const [qrSentCount, setQrSentCount] = React.useState(0);
  const [searchQuery, setSearchQuery] = React.useState('');
  const [sortColumn, setSortColumn] = React.useState<'id' | 'anrede' | 'name' | 'email' | 'status' | 'date'>('id');
  const [sortAsc, setSortAsc] = React.useState(true);
  const [isReorderingIDs, setIsReorderingIDs] = React.useState(false);
  const [reorderResult, setReorderResult] = React.useState<string | null>(null);
  const [isFixingColumns, setIsFixingColumns] = React.useState(false);
  const [fixColumnsResult, setFixColumnsResult] = React.useState<string | null>(null);
  // Email Compose Modal
  const [showEmailModal, setShowEmailModal] = React.useState(false);
  const [emailSubject, setEmailSubject] = React.useState('');
  const [emailHeading, setEmailHeading] = React.useState('');
  const [emailBody, setEmailBody] = React.useState('');
  const [emailSending, setEmailSending] = React.useState(false);
  const [emailShowPreview, setEmailShowPreview] = React.useState(false);

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

  /**
   * CSV Export fuer Teilnehmerlisten.
   * - 'deloitte': alle internen Felder (Anrede, Name, Email, Department, Location, JobTitle, Phone, Status, ...)
   * - 'b2run': Format laut B2Run Excel-Template (Nr, Anrede, Vorname, Nachname, E-Mail, Startblock, Zustimmung AGB, Anonym, Gruppe, Strasse, PLZ, Stadt, Mobilnummer, Infoservice, Altersklasse)
   */
  const exportCsv = (mode: 'deloitte' | 'b2run'): void => {
    if (!selectedEvent) return;
    const activeRegsForExport = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
    if (activeRegsForExport.length === 0) { alert('Keine aktiven Teilnehmer zum Exportieren.'); return; }

    // CSV-Wert sicher escapen (Komma, Quotes, Newlines)
    const esc = (v: unknown): string => {
      const s = (v === null || v === undefined) ? '' : String(v);
      if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0 || s.indexOf(';') >= 0) {
        return `"${s.replace(/"/g, '""')}"`;
      }
      return s;
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parseCustom = (json: string): Record<string, any> => {
      try { return JSON.parse(json || '{}'); } catch { return {}; }
    };

    let headers: string[] = [];
    let rows: string[][] = [];

    if (mode === 'b2run') {
      // Reihenfolge exakt wie B2Run Excel
      headers = [
        'Nr.', 'Anrede', 'Vorname', 'Nachname', 'E-Mail',
        'Startblock', 'Zustimmung AGB & Datenschutzhinweise', 'Anonym',
        'Gruppe', 'Strasse und Hausnummer (privat)', 'PLZ (privat)', 'Stadt (privat)',
        'Mobilnummer', 'Verwendung Infoservice', 'Altersklasse',
      ];
      rows = activeRegsForExport.map(r => {
        const cd = parseCustom(r.CustomData || '{}');
        const vorname = r.Vorname || (r.ParticipantName || '').split(' ').slice(0, -1).join(' ') || '';
        const nachname = r.Nachname || (r.ParticipantName || '').split(' ').slice(-1).join(' ') || '';
        return [
          String(r.TeilnehmerID || ''),
          r.Anrede || '',
          vorname,
          nachname,
          r.ParticipantEmail || '',
          cd.b2run_startblock || '',
          cd.b2run_datenschutz ? 'Ja' : 'Nein',
          cd.b2run_anonym ? 'Ja' : 'Nein',
          cd.b2run_gruppe || '',
          '', // Strasse - nicht abgefragt
          '', // PLZ - nicht abgefragt
          '', // Stadt - nicht abgefragt
          cd.b2run_mobilnummer || '',
          cd.b2run_infoservice ? 'Ja' : 'Nein',
          cd.b2run_altersklasse || '',
        ];
      });
    } else {
      // Deloitte View: alle internen Felder
      headers = [
        'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'Email',
        'Department', 'Location', 'JobTitle', 'Phone',
        'Status', 'RegistrationDate',
      ];
      // Dynamisch alle Custom Field Labels aus dem Event sammeln
      const customLabels: Array<{ id: string; label: string }> = (selectedEvent.eventSpecificFields || []).map(f => ({ id: f.id, label: f.label }));
      headers = headers.concat(customLabels.map(cf => cf.label));

      rows = activeRegsForExport.map(r => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const anyReg = r as any;
        const cd = parseCustom(r.CustomData || '{}');
        const base = [
          String(r.TeilnehmerID || ''),
          r.Anrede || '',
          r.Vorname || '',
          r.Nachname || '',
          r.ParticipantEmail || '',
          anyReg.Department || '',
          anyReg.Location || '',
          anyReg.JobTitle || '',
          anyReg.Phone || '',
          r.Status || '',
          r.RegistrationDate ? new Date(r.RegistrationDate).toLocaleString('de-DE') : '',
        ];
        const customValues = customLabels.map(cf => {
          const v = cd[cf.id];
          if (v === undefined || v === null) return '';
          if (typeof v === 'boolean') return v ? 'Ja' : 'Nein';
          return String(v);
        });
        return base.concat(customValues);
      });
    }

    // CSV zusammenbauen (UTF-8 BOM fuer Excel-Umlaute + Semikolon-Separator fuer deutsche Excel)
    const csvBody = [headers, ...rows].map(row => row.map(esc).join(';')).join('\r\n');
    const csv = '\ufeff' + csvBody;

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    a.download = `${mode === 'b2run' ? 'B2Run' : 'Deloitte'}_${safeName}_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  };

  const handleSelectEvent = async (event: DeloitteEvent): Promise<void> => {
    setSelectedEvent(event);
    setIsLoadingRegs(true);
    setRegLoadError('');
    try {
      const regs = await getAllRegistrations(event.id);
      setRegistrations(regs);
    } catch {
      setRegistrations([]);
      setRegLoadError('Teilnehmerliste konnte nicht geladen werden.');
    }
    setIsLoadingRegs(false);
  };

  // Teilnehmerlisten-URL aus regListMap ableiten
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const getRegListUrl = (_event: DeloitteEvent): string => {
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
      <div className="page-container" role="main">
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
              {t('create.submit')}
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
                    <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)', margin: '2px 0 0' }}>
                      Organizer: {event.organizers.map(o => { const p = o.split(',').map(s => s.trim()); return p.length === 2 ? `${p[1]} ${p[0]}` : o; }).join(', ')}
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
  const query = searchQuery.toLowerCase().trim();
  const matchesSearch = (reg: SPRegistration): boolean => {
    if (!query) return true;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || '');
    return name.toLowerCase().includes(query)
      || (reg.ParticipantEmail || '').toLowerCase().includes(query)
      || String(reg.TeilnehmerID || '').includes(query);
  };

  const sortRegs = (a: SPRegistration, b: SPRegistration): number => {
    let cmp = 0;
    switch (sortColumn) {
      case 'id': cmp = (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0); break;
      case 'anrede': cmp = (a.Anrede || '').localeCompare(b.Anrede || '', 'de'); break;
      case 'name': {
        const na = (a.Vorname && a.Nachname) ? `${a.Nachname} ${a.Vorname}` : (a.ParticipantName || '');
        const nb = (b.Vorname && b.Nachname) ? `${b.Nachname} ${b.Vorname}` : (b.ParticipantName || '');
        cmp = na.localeCompare(nb, 'de');
        break;
      }
      case 'email': cmp = (a.ParticipantEmail || '').localeCompare(b.ParticipantEmail || ''); break;
      case 'status': cmp = (a.Status || '').localeCompare(b.Status || ''); break;
      case 'date': cmp = new Date(a.RegistrationDate || 0).getTime() - new Date(b.RegistrationDate || 0).getTime(); break;
    }
    return sortAsc ? cmp : -cmp;
  };

  const handleSort = (col: 'id' | 'anrede' | 'name' | 'email' | 'status' | 'date'): void => {
    if (sortColumn === col) { setSortAsc(!sortAsc); }
    else { setSortColumn(col); setSortAsc(true); }
  };

  const sortIcon = (col: string): string => col === sortColumn ? (sortAsc ? ' \u25B2' : ' \u25BC') : '';

  const activeRegs = registrations
    .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
    .filter(matchesSearch)
    .sort(sortRegs);
  const waitlistRegs = registrations.filter(r => r.Status === 'Warteliste').filter(matchesSearch)
    .sort((a, b) => new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
  const cancelledRegs = registrations.filter(r => r.Status === 'Abgemeldet').filter(matchesSearch);

  return (
    <div className="page-container" role="main">
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

      {/* Event-Bild - eigene kleine Kachel, Original-Aufloesung */}
      {selectedEvent.imageUrl && (
        <div
          className="card"
          style={{
            padding: 12,
            marginBottom: 24,
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'center',
            maxWidth: 360,
            background: 'var(--dex-gray-50, #fafafa)',
          }}
        >
          <img
            src={selectedEvent.imageUrl}
            alt={selectedEvent.title}
            style={{
              display: 'block',
              maxWidth: '100%',
              maxHeight: 220,
              borderRadius: 'var(--dex-radius, 12px)',
              objectFit: 'contain',
            }}
          />
        </div>
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
              <span className="settings-info__label">Organizer</span>
              <span>{selectedEvent.organizers.map(o => {
                const parts = o.split(',').map(s => s.trim());
                return parts.length === 2 ? `${parts[1]} ${parts[0]}` : o;
              }).join(', ')}</span>
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
            <button
              className="btn btn-secondary btn-block"
              onClick={() => {
                setEmailSubject(selectedEvent ? `${selectedEvent.title} - Info` : '');
                setEmailHeading(selectedEvent ? selectedEvent.title : '');
                setEmailBody('');
                setEmailShowPreview(false);
                setShowEmailModal(true);
              }}
            >
              <Mail size={16} /> {t('admin.emailall')}
            </button>
            {/* Deloitte-Export: alle Teilnehmer mit internen Feldern */}
            <button
              className="btn btn-secondary btn-block"
              onClick={() => exportCsv('deloitte')}
              title="CSV mit allen Teilnehmerdaten (Abteilung, Position, Location, ...)"
            >
              <FileText size={16} /> Deloitte View (CSV)
            </button>
            {/* B2Run-Export: nur bei B2Run Events */}
            {selectedEvent && selectedEvent.type === 'B2Run' && (
              <button
                className="btn btn-secondary btn-block"
                onClick={() => exportCsv('b2run')}
                title="CSV im B2Run-Format fuer die Anmeldung bei b2run.com"
              >
                <FileText size={16} /> B2Run View (CSV)
              </button>
            )}
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
        <div className="flex-between mb-16" style={{ flexWrap: 'wrap', gap: 12 }}>
          <h3 style={{ margin: 0 }}>
            <Users size={18} /> Teilnehmer ({activeRegs.length})
          </h3>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <input
              type="text"
              className="form-input"
              placeholder="Teilnehmer suchen..."
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              style={{ maxWidth: 280, padding: '6px 12px', fontSize: '0.85rem' }}
            />
            {isAdmin && (
              <button
                className="btn btn-secondary"
                style={{ fontSize: '0.75rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                disabled={isReorderingIDs || !selectedEvent?.subsiteUrl}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!confirm('TeilnehmerIDs neu vergeben (1, 2, 3, ...)? Sortierung nach Erstellungsreihenfolge.')) return;
                  setIsReorderingIDs(true);
                  setReorderResult(null);
                  try {
                    const result = await eventServiceRef.reorderParticipantIDs(selectedEvent.subsiteUrl);
                    setReorderResult(`${result.success} aktualisiert, ${result.errors} Fehler`);
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  } catch {
                    setReorderResult('Fehler beim Neuvergeben der IDs');
                  }
                  setIsReorderingIDs(false);
                }}
              >
                {isReorderingIDs ? 'IDs werden vergeben...' : 'IDs neu vergeben'}
              </button>
            )}
            {isAdmin && (
              <button
                className="btn btn-secondary"
                style={{ fontSize: '0.75rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                disabled={isFixingColumns || !selectedEvent?.subsiteUrl}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  setIsFixingColumns(true);
                  setFixColumnsResult(null);
                  try {
                    const result = await eventServiceRef.fixRegistrationListColumns(selectedEvent.subsiteUrl);
                    const msgs: string[] = [];
                    if (result.added.length > 0) msgs.push(`Spalten hinzugefuegt: ${result.added.join(', ')}`);
                    if (result.viewFixed) msgs.push('View-Reihenfolge korrigiert');
                    setFixColumnsResult(msgs.length > 0 ? msgs.join(' | ') : 'Alles OK, keine Aenderungen noetig');
                  } catch {
                    setFixColumnsResult('Fehler beim Fixen der Spalten');
                  }
                  setIsFixingColumns(false);
                }}
              >
                {isFixingColumns ? 'Spalten werden gefixt...' : 'Spalten fixen'}
              </button>
            )}
            {(reorderResult || fixColumnsResult) && (
              <span style={{ fontSize: '0.75rem', color: (reorderResult || fixColumnsResult || '').includes('Fehler') ? 'var(--dex-red)' : 'var(--dex-green)' }}>
                {reorderResult || fixColumnsResult}
              </span>
            )}
          </div>
        </div>

        {regLoadError ? (
          <p style={{ color: 'var(--dex-red)', fontStyle: 'italic' }}>{regLoadError}</p>
        ) : isLoadingRegs ? (
          <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>Lade Teilnehmer...</p>
        ) : activeRegs.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-400)' }}>Noch keine Teilnehmer registriert.</p>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                  {([['id', '#'], ['anrede', 'Anrede'], ['name', 'Name'], ['email', 'Email'], ['status', 'Status'], ['date', 'Registriert am']] as const).map(([col, label]) => (
                    <th
                      key={col}
                      style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', whiteSpace: 'nowrap' }}
                      onClick={() => handleSort(col)}
                    >
                      {label}{sortIcon(col)}
                    </th>
                  ))}
                  <th style={{ textAlign: 'left', padding: 8 }}>Aktion</th>
                </tr>
              </thead>
              <tbody>
                {activeRegs.map((reg, i) => (
                  <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                    <td style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{reg.TeilnehmerID || (i + 1)}</td>
                    <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{reg.Anrede || '-'}</td>
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
                          // Abmelde-Email und Outlook-Ausladen in Queue eintragen (falls nicht deaktiviert)
                          if (reg.ParticipantEmail) {
                            if (!selectedEvent.disableEmails) {
                              const emailData = cancellationEmail(name, selectedEvent.title);
                              eventServiceRef.queueEmail(
                                emailData.subject, reg.ParticipantEmail, name, emailData.body,
                                'Abmeldung', selectedEvent.title, selectedEvent.id
                              ).catch(err => console.warn('[DEX]', err));
                            }
                            if (!selectedEvent.disableOutlook) {
                              eventServiceRef.queueOutlookEvent(
                                reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
                              ).catch(err => console.warn('[DEX]', err));
                            }
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
                            if (reg.ParticipantEmail && !selectedEvent.disableEmails) {
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

      {/* ===== QUIZ-STATISTIK ===== */}
      {selectedEvent && selectedEvent.quiz && selectedEvent.quiz.length > 0 && (() => {
        const regsWithQuiz = registrations.filter(r => typeof r.QuizCompletedAt === 'string' && r.QuizCompletedAt && typeof r.QuizScore === 'number');
        const totalQuizzes = regsWithQuiz.length;

        // Pro Frage: wie oft richtig beantwortet?
        const perQuestion = selectedEvent.quiz.map((q, qIdx) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const correct = (q as any).correctIndices || [(q as any).correctIndex || 0];
          let correctCount = 0;
          for (const reg of regsWithQuiz) {
            try {
              const answers = JSON.parse(reg.QuizAnswers || '[]');
              const given: number[] = Array.isArray(answers[qIdx]) ? answers[qIdx] : [];
              const isRight = correct.length === given.length && correct.every((c: number) => given.indexOf(c) >= 0);
              if (isRight) correctCount++;
            } catch { /* skip */ }
          }
          return { question: q.question, correctCount, total: totalQuizzes };
        });

        // Top 10 nach Score (bei Gleichstand: frueher abgeschlossen)
        const top10 = regsWithQuiz.slice().sort((a, b) => {
          const sa = a.QuizScore || 0;
          const sb = b.QuizScore || 0;
          if (sb !== sa) return sb - sa;
          const ta = new Date(a.QuizCompletedAt || 0).getTime();
          const tb = new Date(b.QuizCompletedAt || 0).getTime();
          return ta - tb;
        }).slice(0, 10);

        return (
          <div className="card" style={{ padding: 24, marginTop: 24 }}>
            <h3 style={{ margin: 0, marginBottom: 16 }}>
              <FileText size={18} /> Quiz-Statistik
            </h3>

            {totalQuizzes === 0 ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: 0 }}>
                Noch kein Teilnehmer hat das Quiz abgeschlossen.
              </p>
            ) : (
              <>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12, marginBottom: 24 }}>
                  <div style={{ padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 12, textAlign: 'center' }}>
                    <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>{totalQuizzes}</div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Durchgefuehrt</div>
                  </div>
                  <div style={{ padding: 16, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, textAlign: 'center' }}>
                    <div style={{ fontSize: '1.6rem', fontWeight: 700 }}>{selectedEvent.quiz.length}</div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Fragen</div>
                  </div>
                  <div style={{ padding: 16, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, textAlign: 'center' }}>
                    <div style={{ fontSize: '1.6rem', fontWeight: 700 }}>
                      {totalQuizzes > 0
                        ? (regsWithQuiz.reduce((sum, r) => sum + (r.QuizScore || 0), 0) / totalQuizzes).toFixed(1)
                        : '0'}
                    </div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Durchschnitt</div>
                  </div>
                </div>

                {/* Pro Frage */}
                <h4 style={{ marginTop: 0, marginBottom: 12 }}>Pro Frage</h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 24 }}>
                  {perQuestion.map((pq, idx) => {
                    const pct = pq.total > 0 ? Math.round((pq.correctCount / pq.total) * 100) : 0;
                    return (
                      <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
                          <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>
                            {idx + 1}. {pq.question}
                          </span>
                          <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                            {pq.correctCount} / {pq.total} ({pct}%)
                          </span>
                        </div>
                        <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                          <div style={{
                            height: '100%',
                            width: `${pct}%`,
                            background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)',
                            transition: 'width 0.3s',
                          }} />
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Top 10 */}
                <h4 style={{ marginTop: 0, marginBottom: 12 }}>Top 10 Teilnehmer</h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                        <th style={{ textAlign: 'left', padding: 8, width: 40 }}>#</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Name</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Email</th>
                        <th style={{ textAlign: 'left', padding: 8, width: 80 }}>Score</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Abgeschlossen</th>
                      </tr>
                    </thead>
                    <tbody>
                      {top10.map((reg, i) => {
                        const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                        const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i + 1}.`;
                        return (
                          <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                            <td style={{ padding: 8, fontWeight: 700 }}>{medal}</td>
                            <td style={{ padding: 8, fontWeight: 500 }}>{name}</td>
                            <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                            <td style={{ padding: 8, fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>
                              {reg.QuizScore} / {selectedEvent.quiz.length}
                            </td>
                            <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.QuizCompletedAt || '')}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        );
      })()}

      {/* ===== EMAIL COMPOSE MODAL ===== */}
      {showEmailModal && selectedEvent && (
        <div
          style={{
            position: 'fixed', inset: 0, zIndex: 9999, display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: 'rgba(0,0,0,0.5)', padding: 20,
          }}
          onClick={(e) => { if (e.target === e.currentTarget && !emailSending) setShowEmailModal(false); }}
        >
          <div style={{
            background: '#fff', borderRadius: 'var(--dex-radius-lg, 16px)', width: '100%', maxWidth: 800, maxHeight: '90vh',
            overflow: 'auto', boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
          }}>
            {/* Header */}
            <div style={{ padding: '20px 24px', borderBottom: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0 }}>
                <Mail size={20} /> E-Mail an alle Teilnehmer
              </h3>
              <button onClick={() => !emailSending && setShowEmailModal(false)} style={{ background: 'none', border: 'none', fontSize: 20, cursor: 'pointer', color: 'var(--dex-gray-400)' }}>&times;</button>
            </div>

            {emailShowPreview ? (
              /* ===== PREVIEW ===== */
              <div style={{ padding: 24 }}>
                <div style={{ marginBottom: 16, display: 'flex', gap: 8 }}>
                  <button className="btn btn-secondary" onClick={() => setEmailShowPreview(false)}>
                    Bearbeiten
                  </button>
                  <button
                    className="btn btn-primary"
                    disabled={emailSending}
                    onClick={async () => {
                      if (!eventServiceRef || !selectedEvent) return;
                      const recipients = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
                      if (recipients.length === 0) return;
                      if (!confirm(`E-Mail an ${recipients.length} Teilnehmer senden?`)) return;
                      setEmailSending(true);
                      const fullBody = wrapTemplate('#86bc25', emailHeading, `Event ${selectedEvent.title}`, emailBody);
                      const allEmails = recipients.map(r => r.ParticipantEmail).join(';');
                      try {
                        await eventServiceRef.queueEmail(
                          emailSubject, allEmails, 'Alle Teilnehmer', fullBody,
                          'Massenmail', selectedEvent.title, selectedEvent.id
                        );
                        setEmailSending(false);
                        alert(`E-Mail an ${recipients.length} Teilnehmer in die Warteschlange eingetragen.`);
                        setShowEmailModal(false);
                      } catch {
                        setEmailSending(false);
                        alert('Fehler beim Eintragen der E-Mail.');
                      }
                    }}
                  >
                    <Send size={16} /> {emailSending ? 'Wird eingetragen...' : `An ${registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length} Teilnehmer senden`}
                  </button>
                </div>
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 8 }}>
                  Betreff: <strong>{emailSubject}</strong>
                </div>
                <div style={{
                  border: '1px solid var(--dex-gray-200)', borderRadius: 8, overflow: 'hidden', maxHeight: 500, overflowY: 'auto',
                }}>
                  <iframe
                    srcDoc={wrapTemplate('var(--dex-green, #86bc25)', emailHeading, `Event ${selectedEvent.title}`, emailBody)}
                    style={{ width: '100%', height: 500, border: 'none' }}
                    title="E-Mail Vorschau"
                  />
                </div>
              </div>
            ) : (
              /* ===== EDITOR ===== */
              <div style={{ padding: 24 }}>
                {/* Betreff */}
                <div style={{ marginBottom: 16 }}>
                  <label style={{ display: 'block', fontWeight: 600, marginBottom: 4, fontSize: '0.85rem' }}>Betreff</label>
                  <input
                    className="form-input"
                    value={emailSubject}
                    onChange={e => setEmailSubject(e.target.value)}
                    placeholder="E-Mail Betreff..."
                    style={{ width: '100%' }}
                  />
                </div>

                {/* Heading */}
                <div style={{ marginBottom: 16 }}>
                  <label style={{ display: 'block', fontWeight: 600, marginBottom: 4, fontSize: '0.85rem' }}>Titel (im E-Mail-Header)</label>
                  <input
                    className="form-input"
                    value={emailHeading}
                    onChange={e => setEmailHeading(e.target.value)}
                    placeholder="z.B. Wichtige Information"
                    style={{ width: '100%' }}
                  />
                </div>

                {/* Body - RichText */}
                <div style={{ marginBottom: 16 }}>
                  <label style={{ display: 'block', fontWeight: 600, marginBottom: 4, fontSize: '0.85rem' }}>Nachricht</label>
                  <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 'var(--dex-radius, 12px)', minHeight: 200, padding: '0 4px' }}>
                    <RichText
                      value={emailBody}
                      onChange={(text: string) => { setEmailBody(text); return text; }}
                    />
                  </div>
                </div>

                {/* Info */}
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 16 }}>
                  Die E-Mail wird im Deloitte-Design versendet (Logo, grüner Header, Footer) über das Gruppenpostfach <strong>no_reply.events@deloitte.de</strong>.
                  Empfänger: <strong>{registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length}</strong> aktive Teilnehmer.
                </div>

                {/* Actions */}
                <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                  <button className="btn btn-secondary" onClick={() => setShowEmailModal(false)}>
                    Abbrechen
                  </button>
                  <button
                    className="btn btn-primary"
                    disabled={!emailSubject.trim() || !emailBody.trim()}
                    onClick={() => setEmailShowPreview(true)}
                  >
                    Vorschau anzeigen
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
