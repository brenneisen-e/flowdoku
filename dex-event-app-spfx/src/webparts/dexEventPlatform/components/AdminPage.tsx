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
import { Plus, Users, FileText, Trash2, Copy, Mail, Send, Download } from './Icons';
import * as XLSX from 'xlsx';
import { EventService } from '../services/EventService';
import { qrCodeEmail, cancellationEmail, promotionEmail, wrapTemplate, replacePlaceholders, buildEmailFromTemplate } from '../services/EmailTemplates';
import { applyEventTemplateOverride, formatOrganizerList } from '../context/EventContext';
import { HtmlEditorModal } from './HtmlEditorModal';
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
  const { topLevelEvents: events, childEventsOf, isEventsLoading, getAllRegistrations, deleteEvent, updateEvent } = useEvents();
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
  const [isFixingFields, setIsFixingFields] = React.useState(false);
  const [fixFieldsResult, setFixFieldsResult] = React.useState<string | null>(null);
  const [isRefreshingProfiles, setIsRefreshingProfiles] = React.useState(false);
  const [refreshProfilesResult, setRefreshProfilesResult] = React.useState<string | null>(null);
  // Email Compose Modal
  const [showEmailModal, setShowEmailModal] = React.useState(false);
  const [emailSubject, setEmailSubject] = React.useState('');
  const [emailHeading, setEmailHeading] = React.useState('');
  const [emailBody, setEmailBody] = React.useState('');
  const [emailSending, setEmailSending] = React.useState(false);
  const [showExportMenu, setShowExportMenu] = React.useState(false);
  // Outlook-Decline-Check (Admin only): zeigt Teilnehmer, die in Outlook
  // abgesagt haben, aber in der Teilnehmerliste noch aktiv gelistet sind.
  const [isCheckingDeclines, setIsCheckingDeclines] = React.useState(false);
  const [declineResult, setDeclineResult] = React.useState<{
    declinedAndRegistered: Array<{ email: string; name: string; reg: SPRegistration }>;
    declinedTotal: number;
    error: string | null;
  } | null>(null);
  const [showDeclineModal, setShowDeclineModal] = React.useState(false);
  const [declineCopied, setDeclineCopied] = React.useState(false);

  // Admin-Toast fuer Abmelde-/Nachrueck-Feedback (seit v6.8):
  //  - 'cancelling': waehrend die Abmeldung + Nachrueck-Suche laeuft (orange, Spinner)
  //  - 'promoted'  : erfolgreicher Nachruecker mit Namen + Typ (gruen)
  //  - 'no-promote': Abmeldung ok, aber keiner auf der Warteliste (grau)
  type AdminToast =
    | { kind: 'cancelling'; name: string }
    | { kind: 'promoted'; name: string; email: string; type?: string }
    | { kind: 'no-promote'; name: string };
  const [adminToast, setAdminToast] = React.useState<AdminToast | null>(null);

  // v6.17: Spaltenkonfiguration der Teilnehmertabelle (pro Event, lokal gespeichert).
  //  - columnOrder = geordnete Liste sichtbarer Spalten-IDs
  //  - hiddenColumns = ausgeblendete Spalten-IDs (können übers "+ Spalte"-Popover wieder zugeschaltet werden)
  // Die Spezialspalten 'id' / 'name' / 'action' sind alwaysVisible und können nicht ausgeblendet werden.
  const [columnOrder, setColumnOrder] = React.useState<string[]>([]);
  const [hiddenColumns, setHiddenColumns] = React.useState<string[]>([]);
  const [showColumnPicker, setShowColumnPicker] = React.useState(false);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const spfxContext = (window as any).__dexSpfxContext;
  const eventServiceRef = React.useMemo(() => spfxContext ? new EventService(spfxContext) : null, []);

  // SuperAdmin sieht alle Events, EventAdmin nur seine + QR-Scanner-Events.
  // Zugriff wird strikt per E-Mail geprüft — NICHT per Namens-Substring
  // (hatte mehrere Jahre einen Match-per-Surname-Bug, der bei häufigen Nachnamen
  // zu False-Positives führte: z.B. eine Assistentin "Frau Müller" konnte Events
  // sehen, deren Organizer auf "Max Müller" hieß — weil "müller" in "max müller"
  // vorkommt. Seit v6.20 nur noch exakt per currentUser.email gegen
  // event.organizerEmails bzw. event.qrScannerEmails.)
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isQRScannerFor = (ev: DeloitteEvent): boolean =>
    !!currentEmailLc && !!ev.qrScannerEmails && ev.qrScannerEmails.some(e => e.toLowerCase() === currentEmailLc);
  const isOrganizerFor = (ev: DeloitteEvent): boolean =>
    !!currentEmailLc && !!ev.organizerEmails && ev.organizerEmails.some(e => e.toLowerCase() === currentEmailLc);
  const adminEvents = isAdmin
    ? events
    : events.filter(e => isOrganizerFor(e) || isQRScannerFor(e));
  // Wenn der User NUR QR-Scanner ist (nicht Organizer + nicht Admin), dann läuft die
  // Admin-Page im eingeschränkten Modus für das ausgewählte Event: nur KPI-Kacheln
  // + QR-Code-Scanner-Button sichtbar.
  const isQRScannerOnlyForSelected = !!selectedEvent && !isAdmin && !isOrganizerFor(selectedEvent) && isQRScannerFor(selectedEvent);

  // Fuer Admins: vergangene Events in eine einklappbare Sektion auslagern
  // (Organizer sehen nur ihre eigenen Events — dort bleiben auch abgelaufene
  // sichtbar, weil der Organizer sie fuer den Abschluss / CSV-Export etc.
  // evtl. direkt griffbereit braucht).
  const now = Date.now();
  const isPastEvent = (e: DeloitteEvent): boolean =>
    !!e.endDate && new Date(e.endDate).getTime() < now;
  const currentEvents = isAdmin ? adminEvents.filter(e => !isPastEvent(e)) : adminEvents;
  const pastEvents = isAdmin ? adminEvents.filter(isPastEvent) : [];
  const [showPastEvents, setShowPastEvents] = React.useState(false);

  // v6.17: Verfügbare Spalten der Teilnehmer-Tabelle aufbauen. MUSS vor dem
  // early return `if (!selectedEvent) return ...` stehen — sonst verletzen
  // die Hooks die Rules-of-Hooks (unterschiedliche Hook-Anzahl pro Render =
  // React Error #310).
  const availableColumns = React.useMemo(() => {
    const isSplit = !!selectedEvent
      && typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
    const userIds = (selectedEvent?.eventSpecificFields || [])
      .filter(f => f.type === 'user')
      .map(f => f.id);
    const cols: Array<{ id: string; label: string; alwaysVisible?: boolean }> = [
      { id: 'id', label: '#', alwaysVisible: true },
      { id: 'anrede', label: 'Anrede' },
      { id: 'name', label: 'Name', alwaysVisible: true },
      { id: 'email', label: 'Email' },
      { id: 'jobTitle', label: 'Job Title' },
      { id: 'location', label: 'Standort' },
    ];
    if (isSplit) cols.push({ id: 'starterType', label: 'Starter-Typ' });
    cols.push({ id: 'status', label: 'Status' });
    cols.push({ id: 'date', label: 'Registriert am' });
    cols.push({ id: 'registeredBy', label: 'Registriert von' });
    if (userIds.length > 0) cols.push({ id: 'roommate', label: 'Zimmerpartner' });
    for (const f of (selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim())) {
      cols.push({ id: `cf-${f.id}`, label: f.label });
    }
    cols.push({ id: 'action', label: 'Aktion', alwaysVisible: true });
    return cols;
  }, [
    selectedEvent?.id,
    selectedEvent?.durchstarterCapacity,
    selectedEvent?.funstarterCapacity,
    (selectedEvent?.eventSpecificFields || []).map(f => `${f.id}:${f.type}:${f.label}`).join(','),
  ]);

  const columnStorageKey = selectedEvent ? `dex_admin_columns_${selectedEvent.id}` : '';
  // localStorage-Load beim Event-Wechsel.
  React.useEffect(() => {
    if (!selectedEvent) { setColumnOrder([]); setHiddenColumns([]); return; }
    const allIds = availableColumns.map(c => c.id);
    try {
      const raw = localStorage.getItem(columnStorageKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.order) && Array.isArray(parsed.hidden)) {
          const knownOrder = parsed.order.filter((id: string) => allIds.indexOf(id) >= 0);
          const missing = allIds.filter(id => knownOrder.indexOf(id) < 0);
          setColumnOrder([...knownOrder, ...missing]);
          setHiddenColumns(parsed.hidden.filter((id: string) => allIds.indexOf(id) >= 0));
          return;
        }
      }
    } catch { /* kaputte Config ignorieren */ }
    setColumnOrder(allIds);
    setHiddenColumns([]);
  }, [columnStorageKey, availableColumns.map(c => c.id).join(',')]);

  // Persistieren bei Änderungen.
  React.useEffect(() => {
    if (!columnStorageKey || columnOrder.length === 0) return;
    try {
      localStorage.setItem(columnStorageKey, JSON.stringify({ order: columnOrder, hidden: hiddenColumns }));
    } catch { /* quota exceeded oder private mode → ignorieren */ }
  }, [columnStorageKey, columnOrder.join(','), hiddenColumns.join(',')]);

  // Helper: Spalte ausblenden / wieder einblenden / verschieben.
  const hideColumn = (id: string): void => {
    const col = availableColumns.find(c => c.id === id);
    if (!col || col.alwaysVisible) return;
    if (hiddenColumns.indexOf(id) < 0) setHiddenColumns([...hiddenColumns, id]);
  };
  const showColumn = (id: string): void => {
    setHiddenColumns(hiddenColumns.filter(h => h !== id));
    if (columnOrder.indexOf(id) < 0) {
      const actionIdx = columnOrder.indexOf('action');
      const next = [...columnOrder];
      if (actionIdx >= 0) next.splice(actionIdx, 0, id); else next.push(id);
      setColumnOrder(next);
    }
  };
  const moveColumn = (id: string, direction: -1 | 1): void => {
    const idx = columnOrder.indexOf(id);
    if (idx < 0) return;
    const target = idx + direction;
    if (target < 0 || target >= columnOrder.length) return;
    if (columnOrder[target] === 'action') return;
    const next = [...columnOrder];
    [next[idx], next[target]] = [next[target], next[idx]];
    setColumnOrder(next);
  };

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

    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    // XLSX Export — natives Excel-Format, automatische Spalten-Breiten, keine
    // CSV-Escaping-Quirks. Gilt fuer beide Modi (Teilnehmerliste + B2Run).
    const aoa: (string | number)[][] = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const colWidths = headers.map((h, ci) => {
      const maxLen = Math.max(h.length, ...rows.map(r => String(r[ci] || '').length));
      return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (ws as any)['!cols'] = colWidths;
    const wb = XLSX.utils.book_new();
    const sheetName = mode === 'b2run' ? 'B2Run' : 'Teilnehmer';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const filePrefix = mode === 'b2run' ? 'B2Run' : 'Teilnehmer';
    XLSX.writeFile(wb, `${filePrefix}_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
    // esc wird nicht mehr gebraucht; Hinweis an eslint
    void esc;
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

  // v6.20: Access-Gate — wer weder Admin noch Organizer eines Events noch QR-Scanner
  // eines Events ist, darf die Admin-Seite gar nicht erst sehen. Zeigt eine klare
  // "Kein Zugriff"-Meldung statt einer leeren Event-Liste.
  if (!selectedEvent && !isAdmin && adminEvents.length === 0) {
    return (
      <div className="page-container" role="main">
        <h2 className="mb-16">{t('admin.title')}</h2>
        <div className="card" style={{ padding: 48, textAlign: 'center' }}>
          <p style={{ color: 'var(--dex-gray-700)', marginBottom: 8, fontWeight: 600 }}>
            {t('admin.noaccess.title') || 'Kein Zugriff'}
          </p>
          <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.88rem', maxWidth: 520, margin: '0 auto' }}>
            {t('admin.noaccess.msg') || 'Du bist weder Organizer noch QR-Scanner eines Events. Nur Admins und Event-Organizer/Scanner haben Zugriff auf diesen Bereich. Wende dich bei Bedarf an einen Admin.'}
          </p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('landing')}>
            {t('reg.backtoevents') || 'Zurück'}
          </button>
        </div>
      </div>
    );
  }

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
          <>
          {(() => {
            const renderEventCard = (event: DeloitteEvent, opts?: { muted?: boolean }): React.ReactElement => (
              <div
                key={event.id}
                className="card card-clickable"
                style={{ padding: '20px 24px', cursor: 'pointer', opacity: opts?.muted ? 0.85 : 1 }}
              >
                <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
                  <div onClick={() => handleSelectEvent(event)} style={{ flex: '1 1 200px', display: 'flex', alignItems: 'center', gap: 16, cursor: 'pointer' }}>
                    {event.imageUrl && (
                      <div style={{
                        width: 60, height: 40, borderRadius: 'var(--dex-radius)', flexShrink: 0,
                        background: `url(${event.imageUrl}) center/cover no-repeat`,
                        filter: opts?.muted ? 'grayscale(0.4)' : 'none',
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
            );
            return (
              <>
                <div className="my-events-list">
                  {currentEvents.map(ev => renderEventCard(ev))}
                </div>
                {isAdmin && pastEvents.length > 0 && (
                  <div style={{ marginTop: 24 }}>
                    <button
                      type="button"
                      onClick={() => setShowPastEvents(!showPastEvents)}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        width: '100%', padding: '12px 20px',
                        background: 'var(--dex-gray-50, #f8f9fa)',
                        border: '1px dashed var(--dex-gray-300)',
                        borderRadius: 'var(--dex-radius, 12px)',
                        cursor: 'pointer',
                        fontSize: '0.9rem', fontWeight: 600,
                        color: 'var(--dex-gray-700)',
                        textAlign: 'left',
                      }}
                    >
                      <span style={{ fontSize: '1rem' }}>{showPastEvents ? '▾' : '▸'}</span>
                      <span style={{ flex: 1 }}>
                        Vergangene Events ({pastEvents.length})
                      </span>
                      <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                        Klicken zum {showPastEvents ? 'Einklappen' : 'Ausklappen'}
                      </span>
                    </button>
                    {showPastEvents && (
                      <div className="my-events-list" style={{ marginTop: 12 }}>
                        {pastEvents.map(ev => renderEventCard(ev, { muted: true }))}
                      </div>
                    )}
                  </div>
                )}
              </>
            );
          })()}
          </>
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
  // Seit v6.5: getrennte Wartelisten bei B2Run-Split-Kapazitäten (Durchstarter/Funstarter).
  // Die Split-Aktivierung erkennen wir daran, dass beide Kapazitäts-Felder gesetzt und > 0 sind.
  const isSplitCapacity = !!selectedEvent
    && typeof selectedEvent.durchstarterCapacity === 'number'
    && typeof selectedEvent.funstarterCapacity === 'number'
    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
  const waitlistDurch = isSplitCapacity
    ? waitlistRegs.filter(r => r.PreferredStarterType === 'Durchstarter')
    : [];
  const waitlistFun = isSplitCapacity
    ? waitlistRegs.filter(r => r.PreferredStarterType === 'Funstarter')
    : [];
  const waitlistUnassigned = isSplitCapacity
    ? waitlistRegs.filter(r => !r.PreferredStarterType || (r.PreferredStarterType !== 'Durchstarter' && r.PreferredStarterType !== 'Funstarter'))
    : [];

  // Roommate-Matching: durchsucht CustomData nach user-Type Feldern, extrahiert
  // Email aus "Name <email>"-Format, baut Map Email -> Partner-Email. Match-Badge,
  // wenn beide sich gegenseitig ausgewaehlt haben.
  const userFieldIds = (selectedEvent?.eventSpecificFields || [])
    .filter(f => f.type === 'user')
    .map(f => f.id);

  // Render-Funktionen pro Spalte — als eine Map, damit der Header + die Body-Zeilen
  // die gleiche stabile ID benutzen. Wird bei jedem Registration-Render neu aufgebaut,
  // weil die renderCell-Lambdas auf aktuelle Closures (handleSort, sortIcon, …) angewiesen sind.
  // Das ist günstig, weil die Funktion nur Pointer speichert.
  const roommateChoice: Record<string, { partnerEmail: string; partnerName: string }> = {};
  if (userFieldIds.length > 0) {
    const allActiveAndWaitlist = registrations.filter(r =>
      r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt' || r.Status === 'Warteliste'
    );
    for (const r of allActiveAndWaitlist) {
      const email = (r.ParticipantEmail || '').toLowerCase();
      if (!email) continue;
      let cd: Record<string, string> = {};
      try { cd = JSON.parse(r.CustomData || '{}'); } catch { /* */ }
      for (const fid of userFieldIds) {
        const v = cd[fid];
        if (!v) continue;
        const m = v.match(/<([^>]+@[^>]+)>/);
        if (!m) continue;
        const pEmail = m[1].trim().toLowerCase();
        const pName = v.replace(/<[^>]*>/, '').trim();
        if (pEmail && pEmail !== email) {
          roommateChoice[email] = { partnerEmail: pEmail, partnerName: pName };
          break; // nur erstes user-Feld
        }
      }
    }
  }
  const getRoommateInfo = (reg: { ParticipantEmail?: string }): { partnerName: string; mutual: boolean } | null => {
    const email = (reg.ParticipantEmail || '').toLowerCase();
    if (!email) return null;
    const choice = roommateChoice[email];
    if (!choice) return null;
    const reverse = roommateChoice[choice.partnerEmail];
    const mutual = !!reverse && reverse.partnerEmail === email;
    return { partnerName: choice.partnerName || choice.partnerEmail, mutual };
  };

  return (
    <div className="page-container" role="main">
      {/* Admin-Toast: drei Phasen beim Abmelden (seit v6.8).
          1. cancelling — orange, Spinner, laeuft waehrend der Abmeldung+Promote-Suche
          2. promoted   — gruen, zeigt den Nachruecker
          3. no-promote — grau, Abmeldung ok, keiner auf der Warteliste */}
      {adminToast && (() => {
        const accent = adminToast.kind === 'cancelling'
          ? 'var(--dex-orange, #ed8b00)'
          : adminToast.kind === 'promoted'
            ? 'var(--dex-green, #86bc25)'
            : 'var(--dex-gray-400)';
        const accentDark = adminToast.kind === 'cancelling'
          ? 'var(--dex-orange, #ed8b00)'
          : adminToast.kind === 'promoted'
            ? 'var(--dex-green-dark, #6b9a1e)'
            : 'var(--dex-gray-600)';
        const closable = adminToast.kind !== 'cancelling';
        return (
          <div style={{
            position: 'fixed', top: 80, right: 20, zIndex: 1000, maxWidth: 460,
            padding: '14px 18px', borderRadius: 'var(--dex-radius, 12px)',
            background: '#fff',
            border: `1px solid ${accent}`,
            borderLeft: `4px solid ${accent}`,
            boxShadow: '0 6px 20px rgba(0,0,0,0.15)',
            display: 'flex', alignItems: 'flex-start', gap: 12,
          }}>
            {adminToast.kind === 'cancelling' && (
              <div style={{
                width: 20, height: 20, marginTop: 2, flexShrink: 0,
                border: `3px solid var(--dex-gray-200)`,
                borderTopColor: accent,
                borderRadius: '50%',
                animation: 'dex-spin 0.8s linear infinite',
              }} />
            )}
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, color: accentDark, marginBottom: 4 }}>
                {adminToast.kind === 'cancelling' && `Abmeldung von ${adminToast.name} wird verarbeitet…`}
                {adminToast.kind === 'promoted' && `Nachgerückt: ${adminToast.name}`}
                {adminToast.kind === 'no-promote' && `Abmeldung von ${adminToast.name} verarbeitet`}
              </div>
              <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                {adminToast.kind === 'cancelling' && 'Teilnehmer wird abgemeldet, Warteliste wird geprüft und ggf. ein Nachrücker informiert.'}
                {adminToast.kind === 'promoted' && (
                  <>
                    <strong>{adminToast.email}</strong> wurde automatisch aus der Warteliste{adminToast.type ? ` (${adminToast.type})` : ''} nachgerückt. Nachrück-Mail + Outlook-Einladung wurden versendet.
                  </>
                )}
                {adminToast.kind === 'no-promote' && 'Aktuell ist niemand auf der Warteliste (bzw. kein passender Starter-Typ). Der Platz bleibt frei.'}
              </div>
            </div>
            {closable && (
              <button
                onClick={() => setAdminToast(null)}
                aria-label="Schließen"
                style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.2rem', color: 'var(--dex-gray-500)', lineHeight: 1, padding: 0 }}
              >×</button>
            )}
          </div>
        );
      })()}
      {/* Keyframes für Spinner */}
      <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
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

        {/* Aktionen-Card (Edit-Event, Mails versenden, Export, Decline-Check etc.)
            nur für Admins/Organizer sichtbar. QR-Scanner sehen stattdessen nur die
            Event-Info + KPIs + "QR-Code scannen"-Button (siehe unten). */}
        {!isQRScannerOnlyForSelected && (
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
                setShowEmailModal(true);
              }}
            >
              <Mail size={16} /> {t('admin.emailall')}
            </button>
            {isAdmin && (
              <button
                className="btn btn-secondary btn-block"
                disabled={isCheckingDeclines || !selectedEvent}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent) return;
                  setIsCheckingDeclines(true);
                  setDeclineResult(null);
                  setDeclineCopied(false);
                  try {
                    const result = await eventServiceRef.getDeclinedAttendees(selectedEvent.id);
                    if (result.ok) {
                      // Aktive Teilnehmer: nur Status ∈ {Angemeldet, QR versendet, Eingecheckt}
                      // werden gegen die Outlook-Decliner gematched. Wer bereits abgemeldet
                      // oder auf der Warteliste ist, ist hier nicht interessant.
                      const activeByEmail = new Map<string, SPRegistration>();
                      for (const r of registrations) {
                        if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') {
                          activeByEmail.set(String(r.ParticipantEmail || '').toLowerCase(), r);
                        }
                      }
                      const hits: Array<{ email: string; name: string; reg: SPRegistration }> = [];
                      for (const d of result.attendees) {
                        const reg = activeByEmail.get(d.email);
                        if (reg) hits.push({ email: d.email, name: d.name, reg });
                      }
                      setDeclineResult({
                        declinedAndRegistered: hits,
                        declinedTotal: result.attendees.length,
                        error: null,
                      });
                      setShowDeclineModal(true);
                    } else {
                      // result.message aus EventService hat Prioritaet (enthaelt Diagnose-Infos wie EventId),
                      // sonst Fallback auf generische Texte pro Reason.
                      let msg = result.message || 'Unbekannter Fehler beim Lesen des Outlook-Termins.';
                      if (!result.message) {
                        if (result.reason === 'no-pointer') {
                          msg = 'Für dieses Event ist kein Outlook-Termin verknüpft (OutlookEventId / CalendarLink fehlen).';
                        } else if (result.reason === 'not-found') {
                          msg = 'Outlook-Termin wurde im Postfach no_reply.events@deloitte.de nicht gefunden.';
                        } else if (result.reason === 'forbidden') {
                          msg = 'Graph-API-Zugriff abgelehnt (HTTP 403). Zwei Dinge müssen erfüllt sein:\n\n'
                            + '1) Tenant-Admin muss die Permission "Calendars.Read.Shared" im SharePoint Admin Center genehmigen:\n'
                            + '   SharePoint Admin Center → Advanced → API access → Pending requests → "Calendars.Read.Shared" → Approve\n\n'
                            + '2) Dein User muss Exchange-seitig Lese-Zugriff auf den Kalender von no_reply.events@deloitte.de haben. Der Exchange-Admin führt dafür aus:\n'
                            + '   Add-MailboxFolderPermission -Identity "no_reply.events@deloitte.de:\\Calendar" -User "<deine-email>" -AccessRights Reviewer\n\n'
                            + 'Beide Schritte sind einmalig nötig.';
                        }
                      }
                      setDeclineResult({ declinedAndRegistered: [], declinedTotal: 0, error: msg });
                      setShowDeclineModal(true);
                    }
                  } catch (err) {
                    setDeclineResult({
                      declinedAndRegistered: [],
                      declinedTotal: 0,
                      error: err instanceof Error ? err.message : String(err),
                    });
                    setShowDeclineModal(true);
                  }
                  setIsCheckingDeclines(false);
                }}
              >
                <Users size={16} /> {isCheckingDeclines ? 'Outlook wird geprüft...' : 'Outlook-Absagen prüfen'}
              </button>
            )}
            {/* Excel-Export-Dropdown: Deloitte-View oder B2Run-View */}
            <div style={{ position: 'relative' }}>
              <button
                className="btn btn-block"
                onClick={() => setShowExportMenu(!showExportMenu)}
                title="Teilnehmerliste als Excel herunterladen"
                style={{
                  background: 'var(--dex-green-dark, #4a7c1f)',
                  color: '#fff',
                  fontWeight: 600,
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                }}
              >
                <Download size={16} /> Download als Excel
                <span style={{ fontSize: '0.7rem', marginLeft: 4 }}>{showExportMenu ? '▴' : '▾'}</span>
              </button>
              {showExportMenu && (
                <div style={{
                  position: 'absolute', top: '100%', left: 0, right: 0,
                  background: '#fff', border: '1px solid var(--dex-gray-200)',
                  borderRadius: 'var(--dex-radius, 8px)',
                  boxShadow: '0 8px 24px rgba(0,0,0,0.15)',
                  marginTop: 4, padding: 6, zIndex: 100,
                }}>
                  <button
                    type="button"
                    onClick={() => { setShowExportMenu(false); exportCsv('deloitte'); }}
                    style={{
                      display: 'block', width: '100%', textAlign: 'left',
                      padding: '10px 12px', border: 'none', background: 'transparent',
                      cursor: 'pointer', borderRadius: 6,
                    }}
                    onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--dex-gray-50)'; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'transparent'; }}
                  >
                    <div style={{ fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>Deloitte Felder</div>
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', lineHeight: 1.4, marginTop: 2 }}>
                      Alle internen Felder: Name, E-Mail, Department, Location, Position, Status, Registrierungsdatum + alle Custom-Fields des Events.
                    </div>
                  </button>
                  {selectedEvent && selectedEvent.type === 'B2Run' && (
                    <button
                      type="button"
                      onClick={() => { setShowExportMenu(false); exportCsv('b2run'); }}
                      style={{
                        display: 'block', width: '100%', textAlign: 'left',
                        padding: '10px 12px', border: 'none', background: 'transparent',
                        cursor: 'pointer', borderRadius: 6,
                        borderTop: '1px solid var(--dex-gray-100)',
                      }}
                      onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--dex-gray-50)'; }}
                      onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'transparent'; }}
                    >
                      <div style={{ fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>B2Run View</div>
                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', lineHeight: 1.4, marginTop: 2 }}>
                        Spaltenformat exakt wie das B2Run-Excel-Template (Nr, Anrede, Name, E-Mail, Startblock, AGB, Gruppe, Mobilnummer, Altersklasse, …) — direkt importierbar in b2run.com.
                      </div>
                    </button>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
        )}
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

      {/* Split-Kapazitäts-Übersicht (seit v6.5): getrennte Belegung pro Starter-Typ. */}
      {isSplitCapacity && (() => {
        const active = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        const durchActive = active.filter(r => r.StarterType === 'Durchstarter').length;
        const funActive = active.filter(r => r.StarterType === 'Funstarter').length;
        const durchCap = selectedEvent?.durchstarterCapacity || 0;
        const funCap = selectedEvent?.funstarterCapacity || 0;
        const durchWait = waitlistDurch.length;
        const funWait = waitlistFun.length;
        return (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 24 }}>
            <div className="card" style={{ padding: 16, borderLeft: '3px solid var(--dex-green-dark, #6b9a1e)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <strong style={{ color: 'var(--dex-green-dark, #6b9a1e)' }}>Durchstarter</strong>
                <span style={{ fontSize: '1.2rem', fontWeight: 700 }}>
                  {durchActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{durchCap}</span>
                </span>
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                Warteliste: <strong style={{ color: 'var(--dex-orange)' }}>{durchWait}</strong>
              </div>
            </div>
            <div className="card" style={{ padding: 16, borderLeft: '3px solid var(--dex-orange, #ff8c00)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <strong style={{ color: 'var(--dex-orange, #ff8c00)' }}>Funstarter</strong>
                <span style={{ fontSize: '1.2rem', fontWeight: 700 }}>
                  {funActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{funCap}</span>
                </span>
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                Warteliste: <strong style={{ color: 'var(--dex-orange)' }}>{funWait}</strong>
              </div>
            </div>
          </div>
        );
      })()}

      {/* v6.19: QR-Code-Scanner-Modus — User ist nur als Scanner eingetragen und
          sieht deshalb NUR die Event-Info + KPIs + "QR-Code scannen"-Button.
          Alle Organizer-Aktionen (Teilnehmerliste, Mails, Edit, Export etc.)
          sind für Scanner ausgeblendet. */}
      {isQRScannerOnlyForSelected && (
        <div className="admin-actions" style={{ display: 'flex', marginBottom: 24 }}>
          <button
            className="btn btn-primary"
            onClick={() => navigate('check-in', selectedEvent.id)}
            style={{ flex: 1 }}
          >
            {t('admin.checkin') || 'QR-Code scannen'}
          </button>
        </div>
      )}

      {!isQRScannerOnlyForSelected && (<>
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
              // Vorname fuer die Anrede — fallback auf erstes Token aus ParticipantName
              const firstName = reg.Vorname || (reg.ParticipantName || '').trim().split(/\s+/)[0] || name;
              // QR-Code als Base64-Bild generieren
              let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
              try {
                const qrDataUrl = await QRCode.toDataURL(qrData, { width: 300, margin: 2 });
                qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
              } catch { /* Fallback: Text */ }
              // QR-Code E-Mail im Deloitte-Template queuen (Sprache folgt Event.EmailLanguage,
              // Anrede nutzt nur den Vornamen statt vollem Namen)
              const emailData = qrCodeEmail(firstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN');
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

      {/* ===== QUIZ-STATISTIK (collapsible, oberhalb Teilnehmerliste) ===== */}
      {selectedEvent && selectedEvent.quiz && selectedEvent.quiz.length > 0 && (() => {
        // Teilnehmer mit mindestens einer beantworteten Frage (nicht nur "komplett durchgefuehrt").
        // Dadurch erscheinen auch Teilnehmer, die mittendrin aufgehoert haben.
        const regsWithQuiz = registrations.filter(r => {
          if (!r.QuizAnswers) return false;
          try {
            const parsed = JSON.parse(r.QuizAnswers);
            return Array.isArray(parsed) && parsed.some((a: number[]) => Array.isArray(a) && a.length > 0);
          } catch { return false; }
        });
        const regsCompleted = regsWithQuiz.filter(r => typeof r.QuizCompletedAt === 'string' && r.QuizCompletedAt);
        const totalQuizzes = regsWithQuiz.length;
        const totalCompleted = regsCompleted.length;

        // Pro Frage: wie viele haben sie überhaupt beantwortet, wie viele richtig
        const perQuestion = selectedEvent.quiz.map((q, qIdx) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const correct = (q as any).correctIndices || [(q as any).correctIndex || 0];
          let correctCount = 0;
          let answeredCount = 0;
          for (const reg of regsWithQuiz) {
            try {
              const answers = JSON.parse(reg.QuizAnswers || '[]');
              const given: number[] = Array.isArray(answers[qIdx]) ? answers[qIdx] : [];
              if (given.length === 0) continue;
              answeredCount++;
              const isRight = correct.length === given.length && correct.every((c: number) => given.indexOf(c) >= 0);
              if (isRight) correctCount++;
            } catch { /* skip */ }
          }
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const imageBase64 = (q as any).imageBase64 as string | undefined;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const section = (q as any).section as string | undefined;
          return { question: q.question, imageBase64, section, correctCount, answeredCount, total: totalQuizzes };
        });

        // Top 10 nach Score, bei Gleichstand: abgeschlossene vor nicht-abgeschlossenen, dann Zeitpunkt
        const top10 = regsWithQuiz.slice().sort((a, b) => {
          const sa = a.QuizScore || 0;
          const sb = b.QuizScore || 0;
          if (sb !== sa) return sb - sa;
          const aDone = !!a.QuizCompletedAt;
          const bDone = !!b.QuizCompletedAt;
          if (aDone !== bDone) return aDone ? -1 : 1;
          const ta = new Date(a.QuizCompletedAt || 0).getTime();
          const tb = new Date(b.QuizCompletedAt || 0).getTime();
          return ta - tb;
        }).slice(0, 10);

        return (
          <details className="card" style={{ padding: 0, marginBottom: 16 }}>
            <summary style={{
              padding: '16px 24px', cursor: 'pointer', listStyle: 'none',
              display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
              fontSize: '1rem', fontWeight: 600,
            }}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <FileText size={18} /> Quiz-Statistik
              </span>
              <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                {totalQuizzes === 0
                  ? 'Keine Daten'
                  : `${totalCompleted} abgeschlossen, ${totalQuizzes - totalCompleted} teilweise (Klick zum Ausklappen)`}
              </span>
            </summary>
            <div style={{ padding: '0 24px 24px 24px' }}>
              {totalQuizzes === 0 ? (
                <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: 0 }}>
                  Noch kein Teilnehmer hat das Quiz gestartet.
                </p>
              ) : (
                <>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 24 }}>
                    <div style={{ padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>{totalCompleted}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Abgeschlossen</div>
                    </div>
                    <div style={{ padding: 16, background: 'var(--dex-orange-light, #fff7ed)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-orange, #ed8b00)' }}>{totalQuizzes - totalCompleted}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Teilweise</div>
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
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Ø Score</div>
                    </div>
                  </div>

                  {/* Pro Frage - gruppiert nach Bereich falls vorhanden */}
                  <h4 style={{ marginTop: 0, marginBottom: 12 }}>Pro Frage</h4>
                  {(() => {
                    const hasSections = perQuestion.some(pq => !!pq.section);
                    if (!hasSections) return null;
                    // Gruppen in Reihenfolge der ersten Erwaehnung
                    const sectionsInOrder: string[] = [];
                    for (const pq of perQuestion) {
                      if (pq.section && sectionsInOrder.indexOf(pq.section) < 0) sectionsInOrder.push(pq.section);
                    }
                    return (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 16, marginBottom: 24 }}>
                        {sectionsInOrder.map(sec => (
                          <div key={`stat-sec-${sec}`}>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.92rem' }}>
                              Bereich: {sec}
                            </h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => pq.section === sec ? (() => {
                                const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                                return (
                                  <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                                        {pq.imageBase64 && (
                                          <img src={pq.imageBase64} alt="" style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }} />
                                        )}
                                        <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>{idx + 1}. {pq.question}</span>
                                      </div>
                                      <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                                        {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                                      </span>
                                    </div>
                                    <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                                      <div style={{ height: '100%', width: `${pct}%`, background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)', transition: 'width 0.3s' }} />
                                    </div>
                                  </div>
                                );
                              })() : null)}
                            </div>
                          </div>
                        ))}
                        {/* Fragen ohne Bereich */}
                        {perQuestion.some(pq => !pq.section) && (
                          <div>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-gray-600)', fontSize: '0.92rem' }}>Ohne Bereich</h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => !pq.section ? (() => {
                                const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                                return (
                                  <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                                        {pq.imageBase64 && (
                                          <img src={pq.imageBase64} alt="" style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }} />
                                        )}
                                        <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>{idx + 1}. {pq.question}</span>
                                      </div>
                                      <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                                        {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                                      </span>
                                    </div>
                                    <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                                      <div style={{ height: '100%', width: `${pct}%`, background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)', transition: 'width 0.3s' }} />
                                    </div>
                                  </div>
                                );
                              })() : null)}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                  {!perQuestion.some(pq => !!pq.section) && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 24 }}>
                    {perQuestion.map((pq, idx) => {
                      const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                      return (
                        <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                              {pq.imageBase64 && (
                                <img
                                  src={pq.imageBase64}
                                  alt=""
                                  style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }}
                                />
                              )}
                              <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>
                                {idx + 1}. {pq.question}
                              </span>
                            </div>
                            <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                              {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
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
                  )}

                  {/* Top 10 */}
                  <h4 style={{ marginTop: 0, marginBottom: 12 }}>Top 10 Teilnehmer</h4>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                          <th style={{ textAlign: 'left', padding: 8, width: 40 }}>#</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>Name</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>E-Mail</th>
                          <th style={{ textAlign: 'left', padding: 8, width: 80 }}>Score</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {top10.map((reg, i) => {
                          const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                          const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i + 1}.`;
                          const done = !!reg.QuizCompletedAt;
                          // Beantwortete Fragen zaehlen (fuer Partial)
                          let answeredN = 0;
                          try {
                            const parsed = JSON.parse(reg.QuizAnswers || '[]');
                            if (Array.isArray(parsed)) answeredN = parsed.filter((a: number[]) => Array.isArray(a) && a.length > 0).length;
                          } catch { /* */ }
                          return (
                            <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                              <td style={{ padding: 8, fontWeight: 700 }}>{medal}</td>
                              <td style={{ padding: 8, fontWeight: 500 }}>{name}</td>
                              <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                              <td style={{ padding: 8, fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>
                                {reg.QuizScore ?? 0} / {selectedEvent.quiz.length}
                              </td>
                              <td style={{ padding: 8, color: done ? 'var(--dex-gray-500)' : 'var(--dex-orange, #ed8b00)' }}>
                                {done
                                  ? `Abgeschlossen ${formatDate(reg.QuizCompletedAt || '')}`
                                  : `Teilweise (${answeredN}/${selectedEvent.quiz.length})`}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          </details>
        );
      })()}

      {/* Teilnehmerliste */}
      <div className="card" style={{ padding: 24 }}>
        <div className="flex-between mb-16" style={{ flexWrap: 'wrap', gap: 12 }}>
          <h3 style={{ margin: 0 }}>
            <Users size={18} /> Teilnehmer ({activeRegs.length})
          </h3>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            {/* Seit v6.4: Sub-Events sind eigene DEX_Events-Items. Wenn der selektierte
                Event Child-Events hat, zeigen wir einen Dropdown zum schnellen
                Umschalten in die Admin-Ansicht der Child-Events (oder zurück zum Parent). */}
            {selectedEvent && (() => {
              const isChild = !!selectedEvent.parentEventId;
              const siblings = isChild
                ? childEventsOf(selectedEvent.parentEventId || '')
                : childEventsOf(selectedEvent.id);
              if (!isChild && siblings.length === 0) return null;
              const parent = isChild ? events.find(e => e.id === selectedEvent.parentEventId) : selectedEvent;
              return (
                <select
                  className="form-input"
                  value={selectedEvent.id}
                  onChange={e => {
                    const target = [parent, ...siblings].find(x => x && x.id === e.target.value);
                    // Nicht nur setSelectedEvent — sonst bleibt die alte
                    // Teilnehmerliste stehen (Parent-Teilnehmer tauchen dann in
                    // der Session-Ansicht auf). handleSelectEvent lädt die
                    // Registrations aus der richtigen Subsite neu.
                    if (target) handleSelectEvent(target).catch(() => { /* fehler wird intern gesetzt */ });
                  }}
                  style={{ maxWidth: 340, padding: '6px 12px', fontSize: '0.85rem' }}
                  aria-label="Event wechseln"
                >
                  {parent && (
                    <option value={parent.id}>Hauptevent: {parent.title}</option>
                  )}
                  {siblings.map(c => (
                    <option key={c.id} value={c.id}>
                      Session: {c.title || 'ohne Titel'}
                    </option>
                  ))}
                </select>
              );
            })()}
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
                    // Feature-Flags fuer das konkrete Event: steuert welche Spalten
                    // noetig sind und welche aktiv entfernt werden (StarterType auf
                    // Nicht-B2Run-Event, Quiz-Spalten auf Event ohne Quiz).
                    const isB2Run = !!(selectedEvent.durchstarterCapacity || selectedEvent.funstarterCapacity);
                    const hasQuiz = !!(selectedEvent.quiz && selectedEvent.quiz.length > 0);
                    const customFields = (selectedEvent.eventSpecificFields || []).map(f => ({
                      id: f.id, label: f.label, type: f.type, required: f.required, options: f.options,
                      // EventSpecificField hat kein `visible` — default true fuer den Fix
                      visible: true,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      spInternalName: (f as any).spInternalName || '',
                    }));
                    const result = await eventServiceRef.fixRegistrationListColumns(
                      selectedEvent.subsiteUrl,
                      { isB2Run, hasQuiz, customFields }
                    );
                    const msgs: string[] = [];
                    if (result.added.length > 0) msgs.push(`Spalten hinzugefuegt: ${result.added.join(', ')}`);
                    if (result.removed.length > 0) msgs.push(`Spalten entfernt: ${result.removed.join(', ')}`);
                    if (result.viewFixed) msgs.push('View-Reihenfolge korrigiert');

                    // Wenn neue spInternalNames fuer Custom-Fields angelegt wurden,
                    // ins DEX_Events-Item zurueckschreiben, damit upsertParticipant
                    // die Werte in die richtigen SP-Spalten schreiben kann.
                    if (result.customFieldMap && Object.keys(result.customFieldMap).length > 0) {
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const updatedCf = (customFields as any[]).map(f => {
                        const sp = result.customFieldMap![f.id];
                        return sp ? { ...f, spInternalName: sp } : f;
                      });
                      try {
                        await updateEvent(selectedEvent.id, { 'CustomFields': JSON.stringify(updatedCf) });
                        msgs.push(`Custom-Field-Zuordnung aktualisiert (${Object.keys(result.customFieldMap).length})`);
                      } catch {
                        msgs.push('WARN: Custom-Field-Mapping konnte nicht am Event gespeichert werden');
                      }
                    }

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
            {isAdmin && (
              <button
                className="btn btn-secondary"
                style={{ fontSize: '0.75rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                disabled={isFixingFields || !selectedEvent}
                onClick={async () => {
                  if (!selectedEvent) return;
                  setIsFixingFields(true);
                  setFixFieldsResult(null);
                  try {
                    // Normalisiert die Custom-Fields-Typen (Zustimmung/AGB -> checkbox,
                    // T-Shirt-Feld -> 'Kein T-Shirt benoetigt'-Option), und entfernt das
                    // redundante '(Pflicht)'-Suffix aus Labels, da das rote Sternchen den
                    // Pflicht-Status schon visualisiert.
                    const changes: string[] = [];
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const raw: any[] = (selectedEvent.eventSpecificFields || []).map((f: any) => ({ ...f }));
                    // B2Run: fehlende Infoservice/Anonym-Checkbox-Felder ergaenzen
                    const hasField = (id: string): boolean => raw.some(f => f.id === id);
                    const isB2Run = raw.some(f => String(f.id || '').indexOf('b2run_') === 0);
                    if (isB2Run) {
                      if (!hasField('b2run_infoservice')) {
                        raw.push({ id: 'b2run_infoservice', label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergaenzt: 'Infoservice'");
                      }
                      if (!hasField('b2run_anonym')) {
                        raw.push({ id: 'b2run_anonym', label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergaenzt: 'Anonym teilnehmen'");
                      }
                      // Laufshirt: ggf. als b2run_laufshirt anlegen, wenn weder das Feld
                      // noch ein gleichnamiges existiert
                      const hasLaufshirt = raw.some(f => f.id === 'b2run_laufshirt' || /laufshirt/i.test(String(f.label || '')));
                      if (!hasLaufshirt) {
                        raw.push({ id: 'b2run_laufshirt', label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true });
                        changes.push("Feld ergaenzt: 'Deloitte-Laufshirt' (Pflicht)");
                      }
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const fixed = raw.map((f: any) => {
                      const nf = { ...f };
                      const label = String(nf.label || '');
                      const lowLabel = label.toLowerCase();
                      // Zustimmung / AGB / Datenschutz -> checkbox
                      const isConsent = lowLabel.indexOf('zustimmung') >= 0
                        || lowLabel.indexOf('agb') >= 0
                        || lowLabel.indexOf('datenschutz') >= 0;
                      const isB2RunCheckbox = ['b2run_infoservice', 'b2run_anonym', 'b2run_datenschutz'].indexOf(nf.id) >= 0;
                      if ((isConsent || isB2RunCheckbox) && nf.type !== 'checkbox') {
                        nf.type = 'checkbox';
                        nf.options = [];
                        changes.push(`${label} -> Checkbox`);
                      }
                      // T-Shirt: 'Kein T-Shirt benoetigt' als Option hinzu, required=false
                      const isShirt = lowLabel.indexOf('t-shirt') >= 0 || lowLabel.indexOf('tshirt') >= 0 || lowLabel.indexOf('shirt') >= 0;
                      if (isShirt && nf.type === 'select') {
                        const opts: string[] = Array.isArray(nf.options) ? nf.options.slice() : [];
                        const hasNo = opts.some((o: string) => o.toLowerCase().indexOf('kein') >= 0);
                        if (!hasNo) {
                          opts.unshift('Ohne T-Shirt');
                          nf.options = opts;
                          changes.push(`${label} -> 'Ohne T-Shirt'-Option`);
                        }
                        if (nf.required) {
                          nf.required = false;
                          changes.push(`${label} -> optional`);
                        }
                      }
                      // '(Pflicht)'-Suffix aus Labels entfernen
                      const stripped = label.replace(/\s*\((?:pflicht|mandatory|required)\)\s*$/i, '').trim();
                      if (stripped && stripped !== label) {
                        nf.label = stripped;
                        changes.push(`Label "${label}" -> "${stripped}"`);
                      }
                      // B2Run-spezifische Korrekturen
                      if (nf.id === 'b2run_mobilnummer') {
                        if (nf.required) { nf.required = false; changes.push('Mobilnummer -> optional'); }
                        if (nf.label === 'Mobilnummer') {
                          nf.label = 'Mobilnummer (nur bei aktiviertem Infoservice)';
                          changes.push("Mobilnummer-Label praezisiert");
                        }
                      }
                      if (nf.id === 'b2run_infoservice' && nf.label && nf.label.indexOf('benoetigt') >= 0) {
                        nf.label = 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)';
                        changes.push('Infoservice-Label modernisiert');
                      }
                      if (nf.id === 'b2run_datenschutz') {
                        const needLinks = !Array.isArray(nf.externalLinks) || nf.externalLinks.length === 0;
                        if (needLinks) {
                          nf.externalLinks = [
                            { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
                            { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
                          ];
                          changes.push('B2Run-Datenschutz: AGB + Datenschutz Links ergaenzt');
                        }
                      }
                      // Deloitte-Laufshirt / Laufshirt: immer Pflicht + 'Kein Laufshirt benoetigt'
                      // als Option, damit man explizit keines waehlen kann ohne das
                      // Pflichtfeld leer zu lassen.
                      if (nf.id === 'b2run_laufshirt' || /laufshirt/i.test(label)) {
                        if (!nf.required) {
                          nf.required = true;
                          changes.push(`${label || nf.id}: als Pflichtfeld markiert`);
                        }
                        if (nf.type === 'select') {
                          const opts: string[] = Array.isArray(nf.options) ? nf.options.slice() : [];
                          const hasNo = opts.some((o: string) => o.toLowerCase().indexOf('kein') >= 0);
                          if (!hasNo) {
                            opts.unshift('Habe bereits ein Laufshirt');
                            nf.options = opts;
                            changes.push(`${label || nf.id}: 'Habe bereits ein Laufshirt'-Option hinzugefuegt`);
                          }
                        }
                      }
                      return nf;
                    });
                    // B2Run-Datenschutz-Checkbox ans Ende schieben
                    const dsIdx = fixed.findIndex((f: { id: string }) => f.id === 'b2run_datenschutz');
                    if (dsIdx >= 0 && dsIdx !== fixed.length - 1) {
                      const [ds] = fixed.splice(dsIdx, 1);
                      fixed.push(ds);
                      changes.push('Zustimmung-Checkbox ans Ende verschoben');
                    }
                    // Speichert zurueck in DEX_Events. Die SP-Spalte heisst 'CustomFields'
                    // (PascalCase) und enthaelt die Felder als JSON-String.
                    const ok = await updateEvent(selectedEvent.id, { CustomFields: JSON.stringify(fixed) });
                    if (ok) {
                      setFixFieldsResult(changes.length > 0
                        ? `Geaendert: ${changes.join(' | ')}`
                        : 'Keine Aenderungen noetig.');
                    } else {
                      setFixFieldsResult('Update fehlgeschlagen.');
                    }
                  } catch (err) {
                    setFixFieldsResult('Fehler: ' + (err instanceof Error ? err.message : String(err)));
                  }
                  setIsFixingFields(false);
                }}
              >
                {isFixingFields ? 'Felder werden repariert...' : 'Felder reparieren'}
              </button>
            )}
            {fixFieldsResult && (
              <span style={{ fontSize: '0.75rem', color: fixFieldsResult.startsWith('Fehler') || fixFieldsResult.startsWith('Update fehl') ? 'var(--dex-red)' : 'var(--dex-gray-500)' }}>
                {fixFieldsResult}
              </span>
            )}
            {isAdmin && (
              <button
                className="btn btn-secondary"
                style={{ fontSize: '0.75rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                disabled={isRefreshingProfiles || !selectedEvent?.subsiteUrl}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  // Frage wie viele der letzten Teilnehmer aufgefrischt werden sollen
                  const ans = prompt('Wie viele der letzten Teilnehmer sollen aus dem Benutzerprofil neu geladen werden? (JobTitle, Standort, Department, Phone)', '20');
                  if (!ans) return;
                  const n = parseInt(ans, 10);
                  if (isNaN(n) || n <= 0) { alert('Bitte eine positive Zahl eingeben.'); return; }
                  setIsRefreshingProfiles(true);
                  setRefreshProfilesResult(null);
                  try {
                    const result = await eventServiceRef.fixEventParticipantsProfileData(selectedEvent.subsiteUrl, n);
                    setRefreshProfilesResult(`${result.scanned} geprueft, ${result.updated} aktualisiert, ${result.failedLookups} Profil-Lookups fehlgeschlagen`);
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  } catch {
                    setRefreshProfilesResult('Fehler beim Auffrischen der Profile');
                  }
                  setIsRefreshingProfiles(false);
                }}
              >
                {isRefreshingProfiles ? 'Profile werden geladen...' : 'Profile neu laden'}
              </button>
            )}
            {(reorderResult || fixColumnsResult || refreshProfilesResult) && (
              <span style={{ fontSize: '0.75rem', color: (reorderResult || fixColumnsResult || refreshProfilesResult || '').includes('Fehler') ? 'var(--dex-red)' : 'var(--dex-green)' }}>
                {reorderResult || fixColumnsResult || refreshProfilesResult}
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
            {(() => {
              // v6.17: Spaltenkonfiguration — Header und Body-Zellen werden dynamisch
              // anhand `columnOrder` (+ `hiddenColumns`) gerendert. So kann der User
              // Spalten ein-/ausblenden und per Pfeilen umsortieren. Die Render-Logik
              // selbst (Sort-Buttons, Badges, Custom-Field-Anzeige etc.) bleibt gleich,
              // nur die Iteration ist umgebaut.
              const visibleColumnIds = columnOrder.filter(id => hiddenColumns.indexOf(id) < 0);

              const sortableCols: Record<string, 'id' | 'anrede' | 'name' | 'email' | 'status' | 'date'> = {
                id: 'id', anrede: 'anrede', name: 'name', email: 'email', status: 'status', date: 'date',
              };

              const hideButton = (id: string): React.ReactNode => {
                const col = availableColumns.find(c => c.id === id);
                if (!col || col.alwaysVisible) return null;
                return (
                  <button
                    type="button"
                    onClick={(e) => { e.stopPropagation(); hideColumn(id); }}
                    aria-label={`Spalte ${col.label} ausblenden`}
                    title="Spalte ausblenden"
                    style={{
                      marginLeft: 6, padding: 0, width: 16, height: 16, lineHeight: '14px',
                      border: 'none', background: 'transparent', cursor: 'pointer',
                      color: 'var(--dex-gray-400)', fontSize: '0.8rem', borderRadius: 3,
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.color = 'var(--dex-red, #c00)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--dex-gray-400)'; }}
                  >
                    ✕
                  </button>
                );
              };

              const renderHeader = (id: string): React.ReactNode => {
                const baseStyle: React.CSSProperties = { textAlign: 'left', padding: 8, whiteSpace: 'nowrap' };
                const sortable = sortableCols[id];
                if (sortable) {
                  return (
                    <th
                      key={id}
                      style={{ ...baseStyle, cursor: 'pointer', userSelect: 'none' }}
                      onClick={() => handleSort(sortable)}
                    >
                      {id === 'id' ? '#' : id === 'anrede' ? 'Anrede' : id === 'name' ? 'Name' : id === 'email' ? 'Email' : id === 'status' ? 'Status' : 'Registriert am'}
                      {sortIcon(sortable)}
                      {hideButton(id)}
                    </th>
                  );
                }
                if (id === 'jobTitle') return <th key={id} style={baseStyle}>Job Title{hideButton(id)}</th>;
                if (id === 'location') return <th key={id} style={baseStyle}>Standort{hideButton(id)}</th>;
                if (id === 'starterType') {
                  return (
                    <th key={id} style={baseStyle} title="Starter-Typ: Durchstarter oder Funstarter. Wird bei der Anmeldung gewählt und steuert die Split-Kapazität + Warteliste. Der eigentliche Startblock steht in der Custom-Field-Spalte 'Start block'.">
                      Starter-Typ{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'registeredBy') {
                  return (
                    <th key={id} style={baseStyle} title="Selbst = der Teilnehmer hat sich selbst registriert. Ansonsten Name des Users, der die Registrierung durchgefuehrt hat.">
                      Registriert von{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'roommate') {
                  return (
                    <th key={id} style={baseStyle} title="Ausgewaehlter Zimmerpartner. Match = beide haben sich gegenseitig ausgewaehlt.">
                      Zimmerpartner{hideButton(id)}
                    </th>
                  );
                }
                if (id === 'action') {
                  return <th key={id} style={{ textAlign: 'left', padding: 8 }}>Aktion</th>;
                }
                if (id.indexOf('cf-') === 0) {
                  const cfId = id.substring(3);
                  const field = (selectedEvent?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  const label = field.label || '';
                  return (
                    <th key={id} style={{ ...baseStyle, fontSize: '0.78rem' }} title={label}>
                      {label.length > 22 ? label.substring(0, 20) + '…' : label}
                      {hideButton(id)}
                    </th>
                  );
                }
                return null;
              };

              const renderCell = (id: string, reg: SPRegistration, i: number): React.ReactNode => {
                if (id === 'id') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{reg.TeilnehmerID || (i + 1)}</td>;
                }
                if (id === 'anrede') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{reg.Anrede || '-'}</td>;
                }
                if (id === 'name') {
                  return <td key={id} style={{ padding: 8, fontWeight: 500 }}>{(reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName}</td>;
                }
                if (id === 'email') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>;
                }
                if (id === 'jobTitle') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{(reg as any).JobTitle || '-'}</td>;
                }
                if (id === 'location') {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{(reg as any).Location || '-'}</td>;
                }
                if (id === 'starterType') {
                  return (
                    <td key={id} style={{ padding: 8, fontSize: '0.8rem' }}>
                      {(() => {
                        // Tatsächlicher Startblock (StarterType) + Wunsch (PreferredStarterType).
                        // Wenn beide identisch: nur einen anzeigen. Wenn unterschiedlich (z.B. per
                        // Fallback-Dialog auf anderen Typ umgestiegen): Wunsch in Klammern daneben.
                        const actual = reg.StarterType || '';
                        const pref = reg.PreferredStarterType || '';
                        if (!actual && !pref) return <span style={{ color: 'var(--dex-gray-400)' }}>—</span>;
                        if (actual && pref && actual !== pref) {
                          return <span>{actual} <span style={{ color: 'var(--dex-gray-500)' }}>(Wunsch: {pref})</span></span>;
                        }
                        return <span>{actual || `Wunsch: ${pref}`}</span>;
                      })()}
                    </td>
                  );
                }
                if (id === 'status') {
                  return (
                    <td key={id} style={{ padding: 8 }}>
                      <span className={`badge ${reg.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>{reg.Status}</span>
                    </td>
                  );
                }
                if (id === 'date') {
                  return <td key={id} style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>;
                }
                if (id === 'registeredBy') {
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>
                      {(() => {
                        const actorEmail = (reg.RegisteredByEmail || '').toLowerCase();
                        const participantEmail = (reg.ParticipantEmail || '').toLowerCase();
                        if (!actorEmail) return <span style={{ color: 'var(--dex-gray-400)' }}>-</span>;
                        if (actorEmail === participantEmail) {
                          return <span style={{ color: 'var(--dex-green-dark)' }}>Selbst</span>;
                        }
                        return (
                          <span title={reg.RegisteredByEmail || ''} style={{ color: 'var(--dex-orange)' }}>
                            {reg.RegisteredByName || reg.RegisteredByEmail}
                          </span>
                        );
                      })()}
                    </td>
                  );
                }
                if (id === 'roommate') {
                  return (
                    <td key={id} style={{ padding: 8, fontSize: '0.8rem' }}>
                      {(() => {
                        const info = getRoommateInfo(reg);
                        if (!info) return <span style={{ color: 'var(--dex-gray-300)' }}>-</span>;
                        return (
                          <span>
                            {info.partnerName}
                            {info.mutual && (
                              <span
                                className="badge"
                                style={{ marginLeft: 6, background: 'var(--dex-green)', color: '#fff', padding: '1px 6px', borderRadius: 4, fontSize: '0.7rem' }}
                                title="Beide haben sich gegenseitig als Zimmerpartner ausgewaehlt"
                              >
                                Match
                              </span>
                            )}
                          </span>
                        );
                      })()}
                    </td>
                  );
                }
                if (id.indexOf('cf-') === 0) {
                  const cfId = id.substring(3);
                  const field = (selectedEvent?.eventSpecificFields || []).find(f => f.id === cfId);
                  if (!field) return null;
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const spName = (field as any).spInternalName || '';
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  let val: any = spName ? (reg as any)[spName] : undefined;
                  if ((val === undefined || val === null || val === '') && reg.CustomData) {
                    try {
                      const cd = JSON.parse(reg.CustomData);
                      val = cd[field.id];
                    } catch { /* no-op */ }
                  }
                  let display: React.ReactNode = '-';
                  if (val !== undefined && val !== null && val !== '') {
                    if (field.type === 'checkbox') {
                      const truthy = val === true || val === 'true' || val === 1 || val === '1';
                      display = <span style={{ color: truthy ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>{truthy ? '✓' : '–'}</span>;
                    } else {
                      display = String(val);
                    }
                  }
                  return (
                    <td key={id} style={{ padding: 8, color: 'var(--dex-gray-700)', fontSize: '0.8rem', whiteSpace: 'nowrap', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis' }} title={String(val || '')}>
                      {display}
                    </td>
                  );
                }
                if (id === 'action') {
                  return (
                    <td key={id} style={{ padding: 8, display: 'flex', gap: 4 }}>
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
                          // Lade-Toast anzeigen
                          setAdminToast({ kind: 'cancelling', name });
                          // Typ des Abgemeldeten merken — für typ-bewusstes Nachrücken bei B2Run-Split.
                          const cancelledStarterType = reg.StarterType || '';
                          await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id, `${currentUser.firstName} ${currentUser.surname}`.trim(), currentUser.email);
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
                          // Seit v6.8: Bei Admin/Organizer-Cancel direkt client-seitig
                          // den Nachrücker promoten — so sieht der Admin/Organizer sofort
                          // wer nachgerückt ist. Der Flow später sieht Count_Active =
                          // MaxParticipants und überspringt seinen eigenen Promote-Zweig.
                          // Typ-bewusst: bei B2Run-Split (DurchstarterCapacity > 0 UND
                          // FunstarterCapacity > 0) wird nur ein Warteliste-Teilnehmer
                          // mit passendem PreferredStarterType nachgerückt.
                          const isB2RunSplit = typeof selectedEvent.durchstarterCapacity === 'number'
                            && typeof selectedEvent.funstarterCapacity === 'number'
                            && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
                          try {
                            const promoted = await eventServiceRef.promoteFirstWaitlistItem(
                              selectedEvent.subsiteUrl,
                              cancelledStarterType || undefined,
                              selectedEvent.maxParticipants,
                              (isB2RunSplit && cancelledStarterType) ? cancelledStarterType : undefined
                            );
                            if (promoted && promoted.success && promoted.email) {
                              // Erfolgs-Toast anzeigen
                              setAdminToast({
                                kind: 'promoted',
                                name: promoted.name || promoted.email,
                                email: promoted.email,
                                type: cancelledStarterType || undefined,
                              });
                              // Nachrück-Mail + Outlook-Einladung queuen
                              if (!selectedEvent.disableEmails) {
                                try {
                                  const lang = selectedEvent.emailLanguage || 'EN';
                                  const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
                                  const promoteVars = {
                                    Name: promotedFirstName,
                                    EventTitle: selectedEvent.title,
                                    Organizer: formatOrganizerList(selectedEvent.organizers, lang),
                                    AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                                    WaitlistPosition: '',
                                  };
                                  let emailData: { subject: string; body: string };
                                  const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
                                  const spTpl = applyEventTemplateOverride(spTplRaw, selectedEvent.emailTemplateOverrides, 'Nachruecken');
                                  if (spTpl) {
                                    emailData = buildEmailFromTemplate(spTpl, promoteVars);
                                  } else {
                                    emailData = promotionEmail(promotedFirstName, selectedEvent.title);
                                  }
                                  await eventServiceRef.queueEmail(
                                    emailData.subject, promoted.email, promoted.name || '', emailData.body,
                                    'Nachruecken', selectedEvent.title, selectedEvent.id
                                  );
                                } catch (err) { console.warn('[DEX] promote-email failed:', err); }
                              }
                              if (!selectedEvent.disableOutlook) {
                                try {
                                  await eventServiceRef.queueOutlookEvent(
                                    promoted.email, selectedEvent.id, selectedEvent.title, 'Einladen'
                                  );
                                } catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
                              }
                            } else {
                              // Kein passender Warteliste-Teilnehmer — grauer Info-Toast
                              setAdminToast({ kind: 'no-promote', name });
                            }
                          } catch (err) {
                            console.warn('[DEX] promoteFirstWaitlistItem failed:', err);
                            // Bei Fehler: "kein Nachrücker" anzeigen (Abmeldung selbst war erfolgreich)
                            setAdminToast({ kind: 'no-promote', name });
                          }
                          // IDReorder in Queue eintragen — der Flow macht nur noch Reorder
                          // (Promote ist oben schon passiert, Flow erkennt Count_Active = MaxParticipants).
                          if (selectedEvent.subsiteUrl) {
                            try {
                              const ok = await eventServiceRef.queueIDReorder(
                                selectedEvent.id, selectedEvent.eventNumber || 0,
                                selectedEvent.subsiteUrl, selectedEvent.title
                              );
                              if (!ok) {
                                console.warn('[DEX] queueIDReorder returned false');
                                alert('Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.');
                              }
                            } catch (err) {
                              console.warn('[DEX] queueIDReorder threw:', err);
                              alert('Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.');
                            }
                          }
                          const regs = await getAllRegistrations(selectedEvent.id);
                          setRegistrations(regs);
                        }}
                      >
                        Abmelden
                      </button>
                    </td>
                  );
                }
                return null;
              };

              return (
                <>
                  {/* v6.17: Kontrollzeile mit Column-Picker-Button. Der Popover
                      zeigt alle verfügbaren Spalten inkl. Checkbox zum Ein-/
                      Ausblenden und Pfeilen zum Umsortieren. Die Config wird
                      pro Event in localStorage persistiert (s. useEffect oben). */}
                  <div style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center', marginBottom: 8, gap: 8, position: 'relative' }}>
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                      onClick={() => setShowColumnPicker(!showColumnPicker)}
                    >
                      Spalten anpassen
                    </button>
                    {showColumnPicker && (
                      <div
                        style={{
                          position: 'absolute', right: 0, top: '100%', marginTop: 4,
                          background: '#fff', border: '1px solid var(--dex-gray-200)',
                          borderRadius: 8, padding: 12,
                          boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                          width: 280, zIndex: 100, maxHeight: 400, overflowY: 'auto',
                        }}
                      >
                        <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                          Spalten verwalten
                        </div>
                        {columnOrder.map((id, idx) => {
                          const col = availableColumns.find(c => c.id === id);
                          if (!col) return null;
                          const isHidden = hiddenColumns.indexOf(id) >= 0;
                          const isVisible = !isHidden;
                          const canMoveUp = isVisible && idx > 0 && columnOrder[idx - 1] !== undefined;
                          // "action" bleibt immer letzte → niemand darf unter "action" wandern
                          // und "action" selbst darf nicht verschoben werden.
                          const nextId = columnOrder[idx + 1];
                          const canMoveDown = isVisible && idx < columnOrder.length - 1 && id !== 'action' && nextId !== 'action';
                          return (
                            <div
                              key={id}
                              style={{
                                display: 'flex', alignItems: 'center', gap: 6,
                                padding: '4px 2px', fontSize: '0.82rem',
                                opacity: isVisible ? 1 : 0.55,
                              }}
                            >
                              <input
                                type="checkbox"
                                checked={isVisible}
                                disabled={!!col.alwaysVisible}
                                onChange={() => {
                                  if (col.alwaysVisible) return;
                                  if (isHidden) showColumn(id); else hideColumn(id);
                                }}
                                style={{ cursor: col.alwaysVisible ? 'not-allowed' : 'pointer' }}
                                title={col.alwaysVisible ? 'Pflicht-Spalte — kann nicht ausgeblendet werden' : (isHidden ? 'Einblenden' : 'Ausblenden')}
                              />
                              <span style={{ flex: 1, color: 'var(--dex-gray-700)' }}>{col.label}</span>
                              <button
                                type="button"
                                onClick={() => moveColumn(id, -1)}
                                disabled={!canMoveUp}
                                aria-label="Spalte nach oben"
                                title="Nach oben"
                                style={{
                                  border: 'none', background: 'transparent',
                                  cursor: canMoveUp ? 'pointer' : 'not-allowed',
                                  color: canMoveUp ? 'var(--dex-gray-600)' : 'var(--dex-gray-300)',
                                  fontSize: '0.9rem', padding: '0 4px',
                                }}
                              >
                                ↑
                              </button>
                              <button
                                type="button"
                                onClick={() => moveColumn(id, 1)}
                                disabled={!canMoveDown}
                                aria-label="Spalte nach unten"
                                title="Nach unten"
                                style={{
                                  border: 'none', background: 'transparent',
                                  cursor: canMoveDown ? 'pointer' : 'not-allowed',
                                  color: canMoveDown ? 'var(--dex-gray-600)' : 'var(--dex-gray-300)',
                                  fontSize: '0.9rem', padding: '0 4px',
                                }}
                              >
                                ↓
                              </button>
                            </div>
                          );
                        })}
                        <div style={{ marginTop: 8, textAlign: 'right' }}>
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.72rem', padding: '3px 8px' }}
                            onClick={() => setShowColumnPicker(false)}
                          >
                            Schließen
                          </button>
                        </div>
                      </div>
                    )}
                  </div>

                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                        {visibleColumnIds.map(id => renderHeader(id))}
                      </tr>
                    </thead>
                    <tbody>
                      {activeRegs.map((reg, i) => (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          {visibleColumnIds.map(id => renderCell(id, reg, i))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </>
              );
            })()}
          </div>
        )}

        {(() => {
          // Seit v6.5: bei B2Run-Split-Kapazitäten getrennte Wartelisten-Tabellen pro
          // PreferredStarterType. Ohne Split: eine einzige Warteliste wie bisher.
          const renderWaitlistTable = (title: string, regs: SPRegistration[], accentColor: string): React.ReactElement | null => {
            if (regs.length === 0) return null;
            return (
              <React.Fragment key={title}>
                <h4 style={{ marginTop: 24, color: accentColor }}>{title} ({regs.length})</h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                    <thead>
                      <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                        <th style={{ textAlign: 'left', padding: 8 }}>Platz</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Name</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Email</th>
                        {isSplitCapacity && <th style={{ textAlign: 'left', padding: 8 }}>Wunsch-Typ</th>}
                        <th style={{ textAlign: 'left', padding: 8 }}>Registriert am</th>
                        <th style={{ textAlign: 'left', padding: 8 }}>Aktion</th>
                      </tr>
                    </thead>
                    <tbody>
                      {regs.map((reg, i) => (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 8, fontWeight: 600, color: accentColor }}>{i + 1}</td>
                          <td style={{ padding: 8, fontWeight: 500 }}>{(reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName}</td>
                          <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                          {isSplitCapacity && (
                            <td style={{ padding: 8, color: 'var(--dex-gray-700)' }}>
                              {reg.PreferredStarterType || '—'}
                            </td>
                          )}
                          <td style={{ padding: 8, color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td style={{ padding: 8 }}>
                            <button
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
                              onClick={async () => {
                                if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                                const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                                if (!confirm(`${name} von der Warteliste entfernen?`)) return;
                                await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id, `${currentUser.firstName} ${currentUser.surname}`.trim(), currentUser.email);
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
                                if (selectedEvent.subsiteUrl) {
                                  try {
                                    const ok = await eventServiceRef.queueIDReorder(
                                      selectedEvent.id, selectedEvent.eventNumber || 0,
                                      selectedEvent.subsiteUrl, selectedEvent.title
                                    );
                                    if (!ok) {
                                      alert('Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.');
                                    }
                                  } catch {
                                    alert('Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.');
                                  }
                                }
                                const allRegs = await getAllRegistrations(selectedEvent.id);
                                setRegistrations(allRegs);
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
              </React.Fragment>
            );
          };

          if (isSplitCapacity) {
            return (
              <>
                {renderWaitlistTable('Warteliste Durchstarter', waitlistDurch, 'var(--dex-green-dark, #6b9a1e)')}
                {renderWaitlistTable('Warteliste Funstarter', waitlistFun, 'var(--dex-orange, #ff8c00)')}
                {renderWaitlistTable('Warteliste ohne Typ', waitlistUnassigned, 'var(--dex-gray-500)')}
              </>
            );
          }
          return renderWaitlistTable('Warteliste', waitlistRegs, 'var(--dex-orange)');
        })()}

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

      {/* ===== MASSENMAIL MODAL (HtmlEditorModal mit Toolbar, Variablen, Live-Preview) ===== */}
      {showEmailModal && selectedEvent && (() => {
        const recipients = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        const orgNames = (selectedEvent.organizers || []).join(', ');
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Organizer: orgNames,
        };
        const customLogo = (() => {
          try {
            const o = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const sendAction = async (): Promise<void> => {
          if (!eventServiceRef || !selectedEvent) return;
          if (recipients.length === 0) { alert('Keine aktiven Teilnehmer.'); return; }
          if (!confirm(`E-Mail an ${recipients.length} Teilnehmer senden?`)) return;
          setEmailSending(true);
          // Variablen einmalig aufloesen (Massenmail geht an alle zusammen)
          const resolvedSubject = replacePlaceholders(emailSubject, previewVars);
          const resolvedHeading = replacePlaceholders(emailHeading, previewVars);
          const resolvedBody = replacePlaceholders(emailBody, previewVars);
          const fullBody = wrapTemplate('#86bc25', resolvedHeading, `Event ${selectedEvent.title}`, resolvedBody);
          const allEmails = recipients.map(r => r.ParticipantEmail).join(';');
          try {
            await eventServiceRef.queueEmail(
              resolvedSubject, allEmails, 'Alle Teilnehmer', fullBody,
              'Massenmail', selectedEvent.title, selectedEvent.id
            );
            setEmailSending(false);
            alert(`E-Mail an ${recipients.length} Teilnehmer in die Warteschlange eingetragen.`);
            setShowEmailModal(false);
          } catch {
            setEmailSending(false);
            alert('Fehler beim Eintragen der E-Mail.');
          }
        };
        return (
          <HtmlEditorModal
            open={showEmailModal}
            onClose={() => !emailSending && setShowEmailModal(false)}
            title={`Massenmail an ${recipients.length} Teilnehmer: ${selectedEvent.title}`}
            value={emailBody}
            onChange={setEmailBody}
            previewMode="email"
            emailSubject={emailSubject}
            onEmailSubjectChange={setEmailSubject}
            emailHeading={emailHeading}
            onEmailHeadingChange={setEmailHeading}
            emailHeadingColor="#86bc25"
            previewVars={previewVars}
            insertableVars={[
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
            ]}
            imageBase64={customLogo}
            extraAction={{
              label: emailSending ? 'Wird eingetragen…' : `An ${recipients.length} Teilnehmer senden`,
              onClick: sendAction,
              disabled: emailSending || !emailSubject.trim() || !emailBody.trim() || recipients.length === 0,
              icon: <Send size={16} />,
            }}
          />
        );
      })()}

      {/* ===== OUTLOOK-DECLINE-CHECK MODAL (Admin only) ===== */}
      {showDeclineModal && declineResult && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 2000,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
          onClick={() => setShowDeclineModal(false)}
        >
          <div
            className="card"
            style={{ background: '#fff', maxWidth: 720, width: '100%', maxHeight: '80vh', overflow: 'auto', padding: 24 }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>Outlook-Absagen vs. Anmeldungen</h3>
              <button className="btn btn-secondary" style={{ padding: '4px 10px' }} onClick={() => setShowDeclineModal(false)}>
                Schließen
              </button>
            </div>
            {declineResult.error ? (
              <p style={{ color: 'var(--dex-red)', whiteSpace: 'pre-line' }}>{declineResult.error}</p>
            ) : declineResult.declinedAndRegistered.length === 0 ? (
              <p style={{ color: 'var(--dex-gray-600)' }}>
                Keine Diskrepanzen gefunden. {declineResult.declinedTotal > 0
                  ? `Es gibt ${declineResult.declinedTotal} Outlook-Absage(n), aber keiner davon ist in der Teilnehmerliste noch aktiv.`
                  : 'Niemand hat den Outlook-Termin abgelehnt.'}
              </p>
            ) : (
              <>
                <p style={{ color: 'var(--dex-gray-700)' }}>
                  <strong>{declineResult.declinedAndRegistered.length}</strong> Teilnehmer haben den Outlook-Termin abgelehnt,
                  stehen aber in der Teilnehmerliste noch als aktiv:
                </p>
                <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.8rem' }}
                    onClick={() => {
                      const emails = declineResult.declinedAndRegistered.map(d => d.email).join('; ');
                      navigator.clipboard.writeText(emails).then(() => {
                        setDeclineCopied(true);
                        setTimeout(() => setDeclineCopied(false), 2000);
                      }).catch(() => window.prompt('E-Mail-Adressen kopieren:', emails));
                    }}
                  >
                    <Copy size={14} /> {declineCopied ? 'Kopiert!' : 'E-Mails kopieren'}
                  </button>
                </div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ background: 'var(--dex-gray-50)' }}>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>ID</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>Name</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>E-Mail</th>
                      <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid var(--dex-gray-200)' }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {declineResult.declinedAndRegistered.map(d => {
                      const displayName = (d.reg.Vorname && d.reg.Nachname)
                        ? `${d.reg.Vorname} ${d.reg.Nachname}`
                        : (d.reg.ParticipantName || d.name);
                      return (
                        <tr key={d.email}>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.reg.TeilnehmerID ?? '-'}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{displayName}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.email}</td>
                          <td style={{ padding: '8px', borderBottom: '1px solid var(--dex-gray-100)' }}>{d.reg.Status}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 12 }}>
                  Insgesamt {declineResult.declinedTotal} Outlook-Absage(n) erfasst.
                </p>
              </>
            )}
          </div>
        </div>
      )}
      </>)}
    </div>
  );
}
