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
import { Plus, Users, FileText, Trash2, Copy, Mail, Send, Download, Pencil, ExternalLink, AlertCircle, Hash, Columns, Wrench, RefreshCw, X, Check, Link2 } from './Icons';
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

// v9.20: EventStatus-Labels lokalisieren (DE).
function localizeStatus(status: string): string {
  switch (status) {
    case 'Active': return 'Aktiv';
    case 'Under Construction': return 'Entwurf';
    case 'Completed': return 'Abgeschlossen';
    case 'Cancelled': return 'Abgesagt';
    default: return status;
  }
}

// Status-Werte sind in SP als deutsche Strings gespeichert ('Angemeldet',
// 'QR versendet', 'Warteliste', 'Eingecheckt', 'Abgemeldet'). Die App
// rendert sie hier in der UI-Sprache des Users, ohne den Datenbankwert
// zu aendern.
function translateStatus(status: string, isDe: boolean): string {
  if (isDe || !status) return status;
  switch (status) {
    case 'Angemeldet': return 'Registered';
    case 'QR versendet': return 'QR sent';
    case 'Warteliste': return 'Waitlist';
    case 'Eingecheckt': return 'Checked in';
    case 'Abgemeldet': return 'Cancelled';
    default: return status;
  }
}

// v11.14: Migriert hardcoded B2Run-Sonderbehandlungen aus dem Render-
// Code von RegistrationPage.tsx in echte Custom-Field-Properties:
//
// - b2run_mobilnummer ist nur sichtbar wenn b2run_infoservice='true'
//   → wird durch eine showIf-Bedingung auf dem Mobilnummer-Feld ersetzt.
//   Der Pflicht-Status bleibt dynamisch (true wenn sichtbar via showIf).
// - b2run_datenschutz hat im Render hardcoded externalLinks-Fallbacks
//   (B2Run-AGB + Datenschutz-URL) wenn die Field-Properties leer sind
//   → wird in das Field selbst persistiert.
// - b2run_laufshirt wird im Render auf required=true gezwungen
//   → wird in der Field-Property persistiert.
//
// Wird nur einmalig bei der Migration aufgerufen — die in-Memory-Field-
// Liste wird mutiert; Caller speichert das Ergebnis als CustomFields-
// JSON. Wenn keine relevanten Felder existieren, no-op.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function migrateB2RunFieldExtras(fields: any[]): { changed: boolean } {
  let changed = false;
  for (const f of fields) {
    const id = String(f.id || '').toLowerCase();
    if (id === 'b2run_mobilnummer') {
      // v11.14: showIf-Constraint auf b2run_infoservice='true'.
      // Damit übernimmt die generische Render-Logik die Sichtbarkeit
      // statt der hardcoded Sonderprüfung.
      const sf = f.showIf;
      const alreadySet = sf && sf.fieldId === 'b2run_infoservice'
        && Array.isArray(sf.values) && sf.values.indexOf('true') >= 0;
      if (!alreadySet) {
        f.showIf = { fieldId: 'b2run_infoservice', values: ['true'] };
        // Wenn der User Infoservice aktiviert, ist die Mobilnummer
        // Pflicht — über showIf gerendert ist die Pflicht-Logik
        // jetzt deterministisch (Feld sichtbar ⇒ Feld Pflicht).
        f.required = true;
        changed = true;
      }
    } else if (id === 'b2run_datenschutz') {
      // v11.14: Hardcoded B2Run-AGB- und Datenschutz-Links als
      // externalLinks-Property persistieren, sodass der Render-
      // Fallback-Path obsolet wird.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const links: any[] = Array.isArray(f.externalLinks) ? f.externalLinks : [];
      if (links.length === 0) {
        f.externalLinks = [
          { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ];
        changed = true;
      }
    } else if (id === 'b2run_laufshirt' || /laufshirt/i.test(String(f.label || ''))) {
      // v11.14: Hardcoded required=true persistieren.
      if (!f.required) {
        f.required = true;
        changed = true;
      }
    }
  }
  return { changed };
}

// v7.6: Wiederverwendbare Action-Kachel fuer den Aktionen-Bereich.
// Default in Grau, beim Hover/Focus kippt Border + Icon + Hintergrund auf
// Deloitte-Gruen. Unterstuetzt Button (onClick), Link (href, oeffnet in neuem
// Tab) und passive Wrapper (children-Mode fuer Spezialfaelle wie das
// Excel-Dropdown). Badge zeigt zwingend "Organizer" (gruener Tint) oder
// "Nur Admin" (oranger Tint), damit auf einen Blick klar ist, fuer welche
// Rolle die Aktion gedacht ist.
interface ActionTileProps {
  icon: React.ReactNode;
  title: string;
  desc: string;
  badge: 'organizer' | 'admin';
  onClick?: () => void;
  href?: string;
  disabled?: boolean;
  busy?: boolean;
  result?: string | null;
  resultIsError?: boolean;
  // v9.19: filled-Variante fuer Highlight-Aktionen (z.B. Event aktivieren).
  // accent='green' = grün gefuellt, accent='red' = rot gefuellt.
  accent?: 'green' | 'red';
  // children: zusaetzlicher Inhalt, der unterhalb der Standard-Tile-Inhalte
  // gerendert wird (z.B. das Excel-Dropdown-Menue).
  children?: React.ReactNode;
}
function ActionTile(props: ActionTileProps): React.ReactElement {
  const [hover, setHover] = React.useState(false);
  const isInteractive = !props.disabled && !props.busy;
  const greenAccent = isInteractive && hover;
  // v9.19/v9.20: filled-Look — Tile dezent eingefaerbt fuer
  // Highlight-Aktionen. Pastell statt voll gesaettigt, damit nicht
  // alarmierend wirkt.
  const isFilled = !!props.accent;
  const filledBg = props.accent === 'green' ? '#e3f0c5' : props.accent === 'red' ? '#ffe5e5' : '';
  const filledBorder = props.accent === 'green' ? 'var(--dex-green, #86bc25)' : props.accent === 'red' ? 'var(--dex-red, #da291c)' : '';
  const borderColor = isFilled ? filledBorder : (greenAccent ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)');
  const bg = isFilled ? filledBg : (greenAccent ? 'rgba(134,188,37,0.06)' : '#fff');
  // v9.20: bei pastell-gefuellten Tiles Text/Icon dunkel halten — auf
  // hellem Pastell-Hintergrund gut lesbar (im Gegensatz zum vorherigen
  // weiss auf saturated-Color).
  const filledIconColor = props.accent === 'green' ? 'var(--dex-green-dark, #4a7c1f)' : props.accent === 'red' ? '#a01e15' : 'var(--dex-gray-500, #6b7280)';
  const iconColor = isFilled ? filledIconColor : (greenAccent ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500, #6b7280)');
  const filledTextColor = isFilled
    ? (props.accent === 'green' ? 'var(--dex-green-dark, #3f5f10)' : props.accent === 'red' ? '#a01e15' : 'var(--dex-gray-800, #1f2937)')
    : 'var(--dex-gray-800, #1f2937)';
  const badgeLabel = props.badge === 'admin' ? 'Nur Admin' : 'Organizer';
  const badgeColors = props.badge === 'admin'
    ? { bg: 'rgba(237,139,0,0.12)', fg: 'var(--dex-orange, #ed8b00)' }
    : { bg: 'rgba(134,188,37,0.12)', fg: 'var(--dex-green-dark, #4a7c1f)' };
  const sharedStyle: React.CSSProperties = {
    textAlign: 'left', textDecoration: 'none', color: 'inherit',
    background: bg, border: `1px solid ${borderColor}`,
    borderRadius: 12, padding: 14,
    cursor: isInteractive ? 'pointer' : 'not-allowed',
    opacity: isInteractive ? 1 : 0.55,
    display: 'flex', flexDirection: 'column', gap: 8,
    fontFamily: 'inherit', fontSize: 'inherit',
    transition: 'all 0.15s ease',
    boxShadow: greenAccent ? '0 4px 12px rgba(134,188,37,0.18)' : 'none',
    position: 'relative',
    // width:100% sorgt dafuer, dass die Kachel auch in einem flex-Wrapper
    // (z.B. Excel-Export hat einen <div display:flex>-Wrapper fuer das
    // Dropdown-Positioning) auf die volle Grid-Zellen-Breite gestreckt
    // wird — sonst sieht sie schmaler aus als die direkten Grid-Geschwister.
    width: '100%',
    boxSizing: 'border-box',
  };
  const inner = (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, color: iconColor, transition: 'color 0.15s ease' }}>
          {props.icon}
          <span style={{ fontWeight: 600, fontSize: '0.88rem', color: filledTextColor }}>{props.title}</span>
        </span>
        <span style={{
          fontSize: '0.65rem', padding: '2px 8px', borderRadius: 999,
          // v9.20: Badge auf pastell Tiles in normalem badge-Look (auf hellem
          // Hintergrund gut sichtbar, im Gegensatz zur vorherigen
          // semi-transparenten weissen Variante auf saturated bg).
          background: badgeColors.bg,
          color: badgeColors.fg, fontWeight: 600,
          whiteSpace: 'nowrap', flexShrink: 0, letterSpacing: '0.02em',
        }}>{badgeLabel}</span>
      </div>
      {props.result && (
        <p style={{
          margin: 0, fontSize: '0.72rem',
          color: props.resultIsError ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)',
          fontStyle: 'italic',
        }}>{props.result}</p>
      )}
      {props.children}
      {hover && props.desc && (
        <div
          role="tooltip"
          style={{
            position: 'absolute',
            top: 'calc(100% + 8px)',
            left: 0,
            right: 0,
            zIndex: 50,
            background: 'var(--dex-gray-900, #1f2937)',
            color: '#fff',
            padding: '10px 12px',
            borderRadius: 8,
            fontSize: '0.76rem',
            lineHeight: 1.45,
            fontWeight: 400,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
            pointerEvents: 'none',
            whiteSpace: 'normal',
          }}
        >
          {props.desc}
        </div>
      )}
    </>
  );
  if (props.href) {
    return (
      <a
        href={props.href}
        target="_blank"
        rel="noopener noreferrer"
        style={sharedStyle}
        onMouseEnter={() => setHover(true)}
        onMouseLeave={() => setHover(false)}
      >
        {inner}
      </a>
    );
  }
  return (
    <button
      type="button"
      disabled={!isInteractive}
      onClick={props.onClick}
      style={sharedStyle}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
    >
      {inner}
    </button>
  );
}

export default function AdminPage(): React.ReactElement {
  const { navigate, selectedEventId } = useNavigation();
  const { topLevelEvents: events, childEventsOf, isEventsLoading, getAllRegistrations, deleteEvent, updateEvent, refreshEvents } = useEvents();
  const [isRefreshing, setIsRefreshing] = React.useState(false);
  const handleRefresh = async (): Promise<void> => {
    if (isRefreshing) return;
    setIsRefreshing(true);
    try {
      await refreshEvents();
      // Wenn ein Event gerade selektiert ist, auch dessen Registrations neu laden
      if (selectedEvent) {
        try {
          const regs = await getAllRegistrations(selectedEvent.id);
          setRegistrations(regs);
        } catch { /* */ }
      }
    } finally { setIsRefreshing(false); }
  };
  const { currentUser } = useCurrentUser();
  const { isAdmin, siteUrl, currentUserRole, searchUser, searchUsers } = useRoles();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  const [selectedEvent, setSelectedEvent] = React.useState<DeloitteEvent | null>(null);
  const [registrations, setRegistrations] = React.useState<SPRegistration[]>([]);
  // v11.0: Bei Events mit Teilnehmer-Upload alle Attachment-Listen
  // einmalig laden, sobald sich registrations oder das ausgewählte
  // Event ändern. Damit zeigt der „Anhang"-Button in der Action-Spalte
  // sofort die korrekte Anzahl.
  React.useEffect(() => {
    if (!selectedEvent || !selectedEvent.allowAttendeeUpload || !eventServiceRef || !selectedEvent.subsiteUrl) {
      setAttachmentsByReg({});
      return;
    }
    const subsiteUrl = selectedEvent.subsiteUrl;
    const ids = registrations.map(r => r.Id).filter(Boolean);
    let cancelled = false;
    (async () => {
      const map: Record<number, Array<{ fileName: string; serverRelativeUrl: string }>> = {};
      for (const id of ids) {
        try {
          const list = await eventServiceRef!.listRegistrationAttachments(subsiteUrl, id);
          if (list.length > 0) map[id] = list;
        } catch { /* */ }
      }
      if (!cancelled) setAttachmentsByReg(map);
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.allowAttendeeUpload, registrations.length]);
  // v9.29: Header-Refresh-Button triggert ein globales Event — wir hooken uns ein.
  React.useEffect(() => {
    const onRefresh = (): void => { void handleRefresh(); };
    window.addEventListener('dex-refresh-page', onRefresh);
    return () => window.removeEventListener('dex-refresh-page', onRefresh);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent]);
  const [isLoadingRegs, setIsLoadingRegs] = React.useState(false);
  const [regLoadError, setRegLoadError] = React.useState('');
  const [deletingId, setDeletingId] = React.useState<string | null>(null);
  // v9.0: Danger-Zone-Modal — User muss den Event-Titel exakt (lowercase)
  // eintippen bevor der Loesch-Button aktiv wird. Schutz gegen versehentliche
  // Loeschungen (frueher: Click-to-Confirm-Pattern, war zu schwach).
  const [confirmDeleteEvent, setConfirmDeleteEvent] = React.useState<DeloitteEvent | null>(null);
  const [confirmDeleteText, setConfirmDeleteText] = React.useState('');
  // v9.0: ChangeLog-Modal — Admin/Organizer sehen den Audit-Log aller
  // Event- und Teilnehmer-Aenderungen (DEX_ChangeLog).
  const [showChangeLogModal, setShowChangeLogModal] = React.useState(false);
  const [changeLogEntries, setChangeLogEntries] = React.useState<Array<{
    Id: number; Created: string; Action: string; TargetType: string;
    TargetId: string; TargetName: string; EventId: string; EventTitle: string;
    ActorName: string; ActorEmail: string; Details: string;
  }>>([]);
  const [changeLogLoading, setChangeLogLoading] = React.useState(false);
  const [changeLogFilterAction, setChangeLogFilterAction] = React.useState('');
  const [changeLogFilterEvent, setChangeLogFilterEvent] = React.useState('');
  const [changeLogFilterActor, setChangeLogFilterActor] = React.useState('');
  // Self-Actions (User registriert/storniert sich selbst) sind Datenrauschen
  // im Audit-Log — Admins/Organizer wollen normalerweise nur Aktionen sehen,
  // die jemand AUF einen anderen User angewendet hat. Marker im Details-JSON:
  // `"asActor":"self"` bei selbst durchgeführten Registrierungen/Stornos.
  const [changeLogHideSelf, setChangeLogHideSelf] = React.useState(true);
  const openChangeLog = async (): Promise<void> => {
    if (!eventServiceRef) return;
    setShowChangeLogModal(true);
    setChangeLogLoading(true);
    try {
      const entries = await eventServiceRef.readChangeLog({ top: 500 });
      setChangeLogEntries(entries);
    } finally {
      setChangeLogLoading(false);
    }
  };
  const [isDeleting, setIsDeleting] = React.useState(false);
  const [copiedEmails, setCopiedEmails] = React.useState(false);
  const [copiedDeepLink, setCopiedDeepLink] = React.useState(false);
  const [isSendingQR, setIsSendingQR] = React.useState(false);
  const [qrSentCount, setQrSentCount] = React.useState(0);
  // v9.15: QR-Code-Versand-Modal mit Test-/Volldurchlauf + Auto-Send-Toggle
  const [qrSendModalOpen, setQrSendModalOpen] = React.useState(false);
  const [qrAutoSendToggle, setQrAutoSendToggle] = React.useState(false);
  const [qrSendResult, setQrSendResult] = React.useState<string | null>(null);
  // v9.37: Vorschau der QR-Code-Mail (analog zur Live-Vorschau im Event-Wizard
  // unter „Kommunikation"). Der Organizer sieht damit vorab genau die Mail, die
  // beim Versand rausgeht — inklusive echtem QR-Code für ihn selbst als Empfänger.
  const [qrPreviewOpen, setQrPreviewOpen] = React.useState(false);
  const [qrPreviewHtml, setQrPreviewHtml] = React.useState('');
  const [qrPreviewSubject, setQrPreviewSubject] = React.useState('');
  const [qrPreviewLoading, setQrPreviewLoading] = React.useState(false);
  const [searchQuery, setSearchQuery] = React.useState('');
  const [sortColumn, setSortColumn] = React.useState<'id' | 'anrede' | 'vorname' | 'nachname' | 'email' | 'status' | 'date'>('id');
  const [sortAsc, setSortAsc] = React.useState(true);
  const [isReorderingIDs, setIsReorderingIDs] = React.useState(false);
  const [reorderResult, setReorderResult] = React.useState<string | null>(null);
  const [isResettingCounter, setIsResettingCounter] = React.useState(false);
  const [resetCounterResult, setResetCounterResult] = React.useState<string | null>(null);
  const [isFixingColumns, setIsFixingColumns] = React.useState(false);
  const [fixColumnsResult, setFixColumnsResult] = React.useState<string | null>(null);
  const [isFixingFields, setIsFixingFields] = React.useState(false);
  const [fixFieldsResult, setFixFieldsResult] = React.useState<string | null>(null);
  const [isRefreshingProfiles, setIsRefreshingProfiles] = React.useState(false);
  const [refreshProfilesResult, setRefreshProfilesResult] = React.useState<string | null>(null);
  // Globale Reparatur: Organizer-Email-Mismatch über alle Events fixen
  const [isRepairingOrganizers, setIsRepairingOrganizers] = React.useState(false);
  const [repairOrganizersResult, setRepairOrganizersResult] = React.useState<string | null>(null);
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

  // v8.0: In-App-Edit-Modal fuer Teilnehmer (Admin/Organizer kann jeden
  // Teilnehmer-Eintrag direkt aus der Liste editieren — Anrede, Name, Email,
  // Phone, Department, Location, JobTitle, Status, plus alle Custom-Felder).
  // Beim Save wird eine Audit-Zeile in ChangeLog geschrieben (wer/wann/was)
  // und LastModifiedDate gesetzt — kein direkter SP-Edit mehr noetig, was
  // gleichzeitig das deutsche Datumsformat-Problem in SP umgeht.
  const [editingReg, setEditingReg] = React.useState<SPRegistration | null>(null);
  const [editForm, setEditForm] = React.useState<Record<string, string>>({});
  const [isSavingEdit, setIsSavingEdit] = React.useState(false);
  const [editError, setEditError] = React.useState('');
  // v11.0: Attachment-Modal pro Teilnehmer-Reihe — nur wenn das Event
  // den Teilnehmer-Upload erlaubt (event.allowAttendeeUpload). Modal
  // listet alle Item-Attachments der jeweiligen Reg-Zeile + bietet
  // Download-Link + Lösch-Button (Admin/Organizer kann auch fremde
  // Uploads löschen). Map(regId → attachments) wird beim Laden der
  // Teilnehmerliste einmalig befüllt, damit der Anhang-Button die
  // Anzahl direkt anzeigen kann.
  const [attachmentsByReg, setAttachmentsByReg] = React.useState<Record<number, Array<{ fileName: string; serverRelativeUrl: string }>>>({});
  const [attachmentsModalReg, setAttachmentsModalReg] = React.useState<SPRegistration | null>(null);
  const [attachmentsBusy, setAttachmentsBusy] = React.useState(false);
  const openEditModal = (reg: SPRegistration): void => {
    setEditError('');
    setEditingReg(reg);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const r = reg as any;
    const initial: Record<string, string> = {
      Anrede: r.Anrede || '',
      Vorname: r.Vorname || '',
      Nachname: r.Nachname || '',
      ParticipantEmail: r.ParticipantEmail || '',
      Phone: r.Phone || '',
      Department: r.Department || '',
      Location: r.Location || '',
      JobTitle: r.JobTitle || '',
      Status: r.Status || '',
      // v10.13+: B2Run-Felder ins Edit-Form aufnehmen damit das B2Run-Modul
      // im Edit-Modal die aktuellen Werte vorbefüllen kann. Strings (auch
      // wenn leer) — bei Nicht-B2Run-Events werden die Felder im Modal
      // sowieso nicht angezeigt.
      StarterType: r.StarterType || '',
      PreferredStarterType: r.PreferredStarterType || '',
    };
    // Custom-Field-Werte aus dem reg laden (sie sind als SP-Spalten gespeichert)
    if (selectedEvent?.eventSpecificFields) {
      for (const f of selectedEvent.eventSpecificFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        if (sp) initial[sp] = (r[sp] !== undefined && r[sp] !== null) ? String(r[sp]) : '';
      }
    }
    setEditForm(initial);
  };
  const closeEditModal = (): void => {
    setEditingReg(null);
    setEditForm({});
    setEditError('');
  };
  const saveEdit = async (): Promise<void> => {
    if (!editingReg || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsSavingEdit(true);
    setEditError('');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const r = editingReg as any;

      // Stamm-Daten (Vorname, Nachname, E-Mail) werden seit v9.7 ebenfalls
      // editierbar gemacht — z.B. um Tippfehler nach manueller Anlage zu
      // korrigieren. Validierung:
      //   1. E-Mail muss eine Deloitte-Deutschland-Adresse sein (@deloitte.de).
      //      Die Plattform ist nur fuer DEALL freigeschaltet — auch @deloitte.com
      //      (US/Global) zaehlt als extern. Sonst Abbruch mit Fehler.
      //   2. Person muss in M365 existieren (searchUserByEmail). Sonst
      //      Abbruch mit "Tippfehler"-Hinweis.
      // Die uebrigen Profil-Felder (Phone, Department, Location, JobTitle)
      // bleiben read-only — sie kommen aus dem M365-Profil.
      const oldVorname = String(r.Vorname || '');
      const oldNachname = String(r.Nachname || '');
      const oldEmail = String(r.ParticipantEmail || '');
      const newVorname = (editForm.Vorname || '').trim();
      const newNachname = (editForm.Nachname || '').trim();
      const newEmail = (editForm.ParticipantEmail || '').trim();
      const stammChanged = newVorname !== oldVorname || newNachname !== oldNachname || newEmail !== oldEmail;

      const profileFields: { Department?: string; Location?: string; JobTitle?: string } = {};
      if (stammChanged) {
        // Plausibilitaet: nicht-leer
        if (!newVorname || !newNachname || !newEmail) {
          setEditError(isDe
            ? 'Vorname, Nachname und E-Mail dürfen nicht leer sein.'
            : 'First name, last name and email must not be empty.');
          return;
        }
        // Domain-Check: nur Deloitte-Adressen zulassen
        const lower = newEmail.toLowerCase();
        const isDeloitte = /@(.*\.)?deloitte\.de$/.test(lower);
        if (!isDeloitte) {
          setEditError(isDe
            ? `Externe E-Mail-Adresse — nicht erlaubt. Die Plattform ist nur für Deloitte Deutschland (@deloitte.de) freigeschaltet.`
            : `External email address — not allowed. The platform is only available for Deloitte Germany (@deloitte.de).`);
          return;
        }
        // Existenz-Check via M365-Profile (UPN!=SMTP-aware). Wenn wir hier
        // nichts finden, ist es entweder ein Tippfehler oder ein Account
        // der gar nicht (mehr) im Tenant ist — beides nicht akzeptabel.
        if (newEmail.toLowerCase() !== oldEmail.toLowerCase()) {
          const profile = await searchUser(newEmail);
          if (!profile || !profile.displayName) {
            setEditError(isDe
              ? `Person mit E-Mail "${newEmail}" wurde im Deloitte-Tenant nicht gefunden. Bitte Adresse prüfen (Tippfehler?).`
              : `No person found in the Deloitte tenant for "${newEmail}". Please check the address (typo?).`);
            return;
          }
          // Profil-Daten gleich mit-uebernehmen, damit der Eintrag konsistent
          // bleibt (Department / Location / JobTitle passen zum neuen User).
          profileFields.Department = ''; // searchUser liefert displayName/location/jobTitle
          profileFields.Location = profile.location || '';
          profileFields.JobTitle = profile.jobTitle || '';
        }
      }

      const oldValues: Record<string, unknown> = {};
      const patch: Record<string, unknown> = {};
      const fieldLabelMap: Record<string, string> = {};

      if (stammChanged) {
        if (newVorname !== oldVorname) {
          oldValues.Vorname = oldVorname; patch.Vorname = newVorname; fieldLabelMap.Vorname = isDe ? 'Vorname' : 'First name';
        }
        if (newNachname !== oldNachname) {
          oldValues.Nachname = oldNachname; patch.Nachname = newNachname; fieldLabelMap.Nachname = isDe ? 'Nachname' : 'Last name';
        }
        if (newEmail !== oldEmail) {
          oldValues.ParticipantEmail = oldEmail; patch.ParticipantEmail = newEmail; fieldLabelMap.ParticipantEmail = 'E-Mail';
          // Profil-Daten mit aktualisieren (nur wenn ueberhaupt was zurueckkam)
          if (profileFields.Location) {
            oldValues.Location = String(r.Location || ''); patch.Location = profileFields.Location;
            fieldLabelMap.Location = isDe ? 'Standort' : 'Location';
          }
          if (profileFields.JobTitle) {
            oldValues.JobTitle = String(r.JobTitle || ''); patch.JobTitle = profileFields.JobTitle;
            fieldLabelMap.JobTitle = 'Job Title';
          }
        }
      }

      // Custom-Felder des Events. v10.15+: nur Felder ins Patch aufnehmen die
      // sich tatsächlich geändert haben — sonst sendet ein unverändertes
      // Choice-Feld ohne ausgewählten Wert einen leeren String an SP, der
      // mit HTTP 400 'Invalid choice' kippt und das ganze Update abbricht.
      if (selectedEvent?.eventSpecificFields) {
        for (const f of selectedEvent.eventSpecificFields) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const sp = (f as any).spInternalName || '';
          if (!sp) continue;
          const oldVal = r[sp] !== undefined && r[sp] !== null ? String(r[sp]) : '';
          const newVal = editForm[sp] || '';
          if (newVal === oldVal) continue;  // unverändert → skip
          fieldLabelMap[sp] = f.label;
          oldValues[sp] = oldVal;
          patch[sp] = newVal;
        }
      }

      // v10.13+: B2Run-Felder explizit ins Patch aufnehmen, wenn sich was
      // geändert hat. Sind keine regulären customFields, daher werden sie
      // oben in der eventSpecificFields-Loop nicht abgeholt. Nur bei
      // Split-Capacity-Events relevant.
      const isSplitEvent = !!selectedEvent
        && (selectedEvent.durchstarterCapacity || 0) > 0
        && (selectedEvent.funstarterCapacity || 0) > 0;
      if (isSplitEvent) {
        const oldStarter = String(r.StarterType || '');
        const newStarter = (editForm.StarterType || '').trim();
        if (newStarter !== oldStarter) {
          oldValues.StarterType = oldStarter;
          patch.StarterType = newStarter;
          fieldLabelMap.StarterType = isDe ? 'Starter-Typ' : 'Starter type';
        }
        const oldPref = String(r.PreferredStarterType || '');
        const newPref = (editForm.PreferredStarterType || '').trim();
        if (newPref !== oldPref) {
          oldValues.PreferredStarterType = oldPref;
          patch.PreferredStarterType = newPref;
          fieldLabelMap.PreferredStarterType = isDe ? 'Wunsch-Starter-Typ' : 'Preferred starter type';
        }
      }
      if (Object.keys(patch).length === 0) {
        // Keine Aenderung — nichts zu tun.
        closeEditModal();
        return;
      }
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      const ok = await eventServiceRef.adminUpdateRegistration(
        selectedEvent.subsiteUrl, editingReg.Id, patch, actor, oldValues, fieldLabelMap
      );
      if (!ok) {
        // Häufigste 400-Ursache: eine SP-Spalte aus dem Patch existiert nicht
        // auf dieser Teilnehmerliste (z.B. StarterType auf einem v9-Event ohne
        // B2Run-Schema, oder ein neu hinzugefügtes Custom-Field ohne 'Spalten
        // fixen'-Run). Hilfreicher Hinweis auf den Repair-Button.
        setEditError(isDe
          ? 'Speichern fehlgeschlagen — vermutlich fehlt eine SP-Spalte in der Teilnehmerliste. Klicke einmal „Spalten fixen" im Toolbox-Bereich des Events, dann erneut versuchen.'
          : 'Save failed — likely a missing SP column on the participant list. Click „Fix columns" in the event toolbox once, then retry.');
        return;
      }
      // v9.0: Audit-Log mit Diff der geaenderten Felder
      try {
        const changes: Record<string, { old: unknown; new: unknown }> = {};
        for (const k of Object.keys(patch)) {
          if (oldValues[k] !== patch[k]) changes[k] = { old: oldValues[k], new: patch[k] };
        }
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          targetId: ((editingReg as any).ParticipantEmail || '') + '#' + editingReg.Id,
          targetName: `${editingReg.Vorname || ''} ${editingReg.Nachname || ''}`.trim(),
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { changes },
        });
      } catch { /* */ }
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
      closeEditModal();
    } catch (err) {
      console.warn('[DEX] saveEdit error:', err);
      setEditError(isDe
        ? 'Unerwarteter Fehler beim Speichern.'
        : 'Unexpected error while saving.');
    } finally {
      setIsSavingEdit(false);
    }
  };

  // v6.17: Spaltenkonfiguration der Teilnehmertabelle (pro Event, lokal gespeichert).
  //  - columnOrder = geordnete Liste sichtbarer Spalten-IDs
  //  - hiddenColumns = ausgeblendete Spalten-IDs (können übers "+ Spalte"-Popover wieder zugeschaltet werden)
  // Die Spezialspalten 'id' / 'vorname' / 'nachname' / 'action' sind alwaysVisible und können nicht ausgeblendet werden.
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
  // v9.18: Co-Organizer haben pro Event die gleichen Rechte wie der Hauptorganizer.
  // isOrganizerFor returned true sowohl fuer event.organizerEmails als auch
  // fuer event.coOrganizerEmails (per-Event-Rolle).
  const isOrganizerFor = (ev: DeloitteEvent): boolean => {
    if (!currentEmailLc) return false;
    if (ev.organizerEmails && ev.organizerEmails.some(e => e.toLowerCase() === currentEmailLc)) return true;
    if (ev.coOrganizerEmails && ev.coOrganizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc)) return true;
    return false;
  };
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
      .filter(f => f.type === 'user' || f.type === 'roommate')
      .map(f => f.id);
    const cols: Array<{ id: string; label: string; alwaysVisible?: boolean }> = [
      { id: 'id', label: '#', alwaysVisible: true },
      { id: 'anrede', label: 'Anrede' },
      // v11.26: getrennte Vorname / Nachname Spalten statt der einen
      // kombinierten 'name'-Spalte. Alte localStorage-Eintraege mit 'name'
      // werden im useEffect-Loader unten in 'vorname','nachname' migriert.
      { id: 'vorname', label: 'Vorname', alwaysVisible: true },
      { id: 'nachname', label: 'Nachname', alwaysVisible: true },
      { id: 'email', label: 'Email' },
      { id: 'jobTitle', label: 'Job Title' },
      { id: 'location', label: 'Standort' },
    ];
    // v11.6: bei Split-Capacity die frei waehlbaren Gruppen-Labels nutzen
    // (Fallback auf 'Starter-Typ' wenn keine Labels gesetzt sind).
    if (isSplit) {
      const lblA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || '';
      const lblB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || '';
      const colLabel = (lblA && lblB) ? `${lblA} / ${lblB}` : (isDe ? 'Gruppe' : 'Group');
      cols.push({ id: 'starterType', label: colLabel });
    }
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
          // v11.26: Migration alter Spaltenkonfigurationen — die zentrale
          // 'name'-Spalte wurde in 'vorname' + 'nachname' aufgeteilt. Wenn
          // ein gespeichertes Layout noch 'name' enthaelt, an gleicher
          // Position durch ['vorname','nachname'] ersetzen, damit der
          // User seine gewuenschte Reihenfolge beibehaelt.
          const migratedOrder: string[] = [];
          for (const id of parsed.order as string[]) {
            if (id === 'name') {
              if (migratedOrder.indexOf('vorname') < 0) migratedOrder.push('vorname');
              if (migratedOrder.indexOf('nachname') < 0) migratedOrder.push('nachname');
            } else {
              migratedOrder.push(id);
            }
          }
          const knownOrder = migratedOrder.filter((id: string) => allIds.indexOf(id) >= 0);
          const missing = allIds.filter(id => knownOrder.indexOf(id) < 0);
          setColumnOrder([...knownOrder, ...missing]);
          // 'name' aus hidden auch herausfiltern (wenn jemals manuell hidden gesetzt wurde,
          // unwahrscheinlich da alwaysVisible — aber defensiv).
          setHiddenColumns(parsed.hidden.filter((id: string) => id !== 'name' && allIds.indexOf(id) >= 0));
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
    const fileName = `${filePrefix}_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // v8.4: Manueller Blob-Download statt XLSX.writeFile. Im SPFx-Iframe-
    // Context ist saveAs/createObjectURL haeufig blockiert (CORS / Sandbox-
    // Policies), wodurch der Download stillschweigend nicht startet. Mit
    // anchor.click() laeuft das in jeder Browser-Umgebung zuverlaessig.
    try {
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    } catch (err) {
      console.warn('[DEX] Excel-Export fehlgeschlagen:', err);
      alert(isDe
        ? 'Excel-Export fehlgeschlagen. Bitte Browser-Console pruefen.'
        : 'Excel export failed. Please check the browser console.');
    }
    // esc wird nicht mehr gebraucht; Hinweis an eslint
    void esc;
  };

  const handleSelectEvent = async (event: DeloitteEvent): Promise<void> => {
    setSelectedEvent(event);
    // v10.19: NavigationContext.selectedEventId mitziehen, damit Header die
    // Page-ID granular ableiten kann (admin-center vs. admin-event) und der
    // Deep-Link-Kopier-Button immer die echte Item-ID des aktuell offenen
    // Events kennt. Skip falls bereits synchron — sonst doppelter History-
    // Eintrag beim Auto-Select via Deep-Link.
    if (selectedEventId !== event.id) {
      navigate('admin', event.id);
    }
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

  // v6.31: wenn navigation.selectedEventId gesetzt ist beim Mount (z.B. vom
  // Handbuch-Preview oder einem Deep-Link), direkt in die Detail-Ansicht
  // springen statt auf die Event-Auswahl-Liste.
  const didAutoSelectRef = React.useRef(false);
  React.useEffect(() => {
    if (didAutoSelectRef.current) return;
    if (!selectedEventId || selectedEvent) return;
    const match = adminEvents.find(e => e.id === selectedEventId);
    if (match) {
      didAutoSelectRef.current = true;
      handleSelectEvent(match).catch(() => { /* Fehler wird intern gesetzt */ });
    }
  }, [selectedEventId, adminEvents, selectedEvent]);

  // Soft-Refresh-Sync: wenn `events` durch refreshEvents() aktualisiert wurde
  // (z.B. nach Event-deaktivieren, Event-aktivieren, Edit-save), den lokalen
  // `selectedEvent`-State aus der frischen Liste neu derivieren. Sonst bleibt
  // der Status-Badge (z.B. „Aktiv" vs „Entwurf") nach Toggle stale, weil
  // `selectedEvent` ein eigener useState ist und nicht aus `events` derived.
  React.useEffect(() => {
    if (!selectedEvent) return;
    const fresh = adminEvents.find(e => e.id === selectedEvent.id);
    if (fresh && fresh !== selectedEvent) {
      setSelectedEvent(fresh);
    }
  }, [adminEvents]);

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

  // Danger-Zone-Modal als gemeinsames Element — wird in BEIDEN Render-Branches
  // (Event-Liste und Event-Detail) eingehängt, sonst läuft der Löschen-Klick auf
  // der Event-Liste ins Leere (Bug v9.x: Modal war nur im Detail-Branch gerendert).
  const dangerZoneModal: React.ReactElement | null = confirmDeleteEvent ? (() => {
    const expected = (confirmDeleteEvent.title || '').trim().toLowerCase();
    const typed = confirmDeleteText.trim().toLowerCase();
    const matches = !!expected && expected === typed;
    const close = (): void => { setConfirmDeleteEvent(null); setConfirmDeleteText(''); };
    return (
      <div
        style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
          display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
        }}
        onClick={() => { if (!isDeleting) close(); }}
      >
        <div
          className="card"
          style={{ width: '100%', maxWidth: 560, padding: 24, borderRadius: 16, background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.3)', borderTop: '4px solid var(--dex-red, #c00)' }}
          onClick={e => e.stopPropagation()}
        >
          <div className="flex-between mb-16">
            <h3 style={{ margin: 0, color: 'var(--dex-red, #c00)', display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Trash2 size={20} /> {isDe ? 'Danger Zone — Event löschen' : 'Danger Zone — Delete event'}
            </h3>
            <button
              onClick={close}
              disabled={isDeleting}
              style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: isDeleting ? 'not-allowed' : 'pointer', color: 'var(--dex-gray-500)' }}
            ><X size={20} /></button>
          </div>
          <p style={{ margin: '0 0 12px', fontSize: '0.88rem', lineHeight: 1.55 }}>
            {isDe
              ? <>Du bist dabei das Event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong> zu löschen.</>
              : <>You are about to delete the event <strong>&bdquo;{confirmDeleteEvent.title}&ldquo;</strong>.</>}
          </p>
          <ul style={{ margin: '0 0 16px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.55, paddingLeft: 18 }}>
            <li>{isDe ? 'Subsite (inkl. Teilnehmerliste) und Event-Item wandern in den SharePoint-Papierkorb.' : 'Subsite (incl. attendee list) and event item move to the SharePoint recycle bin.'}</li>
            <li>{isDe ? 'Wiederherstellung durch einen Admin innerhalb von 93 Tagen möglich (zweistufig).' : 'A site collection admin can restore within 93 days (two-stage).'}</li>
            <li>{isDe ? 'Outlook-Termin wird über den Power-Automate-Flow gelöscht.' : 'Outlook calendar event will be deleted via the Power Automate flow.'}</li>
            <li>{isDe ? 'Diese Aktion wird im DEX_ChangeLog mit deinem Namen + Datum protokolliert.' : 'This action is logged in DEX_ChangeLog with your name + date.'}</li>
          </ul>
          <div style={{ background: 'rgba(218,41,28,0.06)', border: '1px solid var(--dex-red, #c00)', padding: 12, borderRadius: 8, marginBottom: 16 }}>
            <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, marginBottom: 6 }}>
              {isDe
                ? <>Tippe zur Bestätigung den Event-Titel <strong>kleingeschrieben</strong> ein:</>
                : <>Type the event title <strong>in lowercase</strong> to confirm:</>}
            </label>
            <code style={{ display: 'inline-block', padding: '4px 8px', background: '#fff', borderRadius: 4, fontSize: '0.85rem', marginBottom: 8, wordBreak: 'break-all' }}>{expected}</code>
            <input
              className="form-input"
              value={confirmDeleteText}
              onChange={e => setConfirmDeleteText(e.target.value)}
              placeholder={isDe ? 'Event-Titel kleingeschrieben…' : 'Event title in lowercase…'}
              disabled={isDeleting}
              autoFocus
              style={{ width: '100%' }}
            />
          </div>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
            <button
              type="button"
              className="btn btn-outline"
              onClick={close}
              disabled={isDeleting}
            >{isDe ? 'Abbrechen' : 'Cancel'}</button>
            <button
              type="button"
              className="btn btn-danger"
              disabled={!matches || isDeleting}
              style={{
                background: matches && !isDeleting ? 'var(--dex-red, #c00)' : 'var(--dex-gray-300)',
                color: '#fff',
                border: 'none',
                cursor: matches && !isDeleting ? 'pointer' : 'not-allowed',
                padding: '8px 16px',
              }}
              onClick={async () => {
                if (!matches || !confirmDeleteEvent) return;
                setIsDeleting(true);
                setDeletingId(confirmDeleteEvent.id);
                try {
                  await deleteEvent(confirmDeleteEvent.id);
                } finally {
                  setIsDeleting(false);
                  setDeletingId(null);
                  close();
                }
              }}
            >
              <Trash2 size={14} /> {isDeleting ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Endgültig löschen' : 'Delete')}
            </button>
          </div>
        </div>
      </div>
    );
  })() : null;

  // ChangeLog-/Audit-Log-Modal als gemeinsames Element — wie das Danger-Zone-
  // Modal muss auch dieses in BEIDEN Render-Branches verfügbar sein, sonst
  // öffnet sich der "Audit log"-Button auf der Event-Liste ins Leere.
  const changeLogModal: React.ReactElement | null = showChangeLogModal ? (() => {
    const fa = changeLogFilterAction.toLowerCase().trim();
    const fe = changeLogFilterEvent.toLowerCase().trim();
    const fac = changeLogFilterActor.toLowerCase().trim();
    // Self-Action-Erkennung: Marker im Details-JSON ODER (als Fallback)
    // Actor-E-Mail == Target-E-Mail (User hat sich selbst registriert/abgemeldet).
    const isSelfAction = (e: typeof changeLogEntries[number]): boolean => {
      const d = (e.Details || '').toLowerCase();
      if (d.indexOf('"asactor":"self"') >= 0) return true;
      // Fallback: bei Participant-Aktionen ohne expliziten Marker prüfen wir,
      // ob Actor und Ziel dieselbe Person sind (Target trägt den Namen des
      // Participants, ActorName ist "Nachname, Vorname").
      const action = (e.Action || '').toLowerCase();
      if (action.indexOf('participant') < 0) return false;
      const tgt = (e.TargetName || '').toLowerCase().trim();
      const actorName = (e.ActorName || '').toLowerCase().trim();
      if (!tgt || !actorName) return false;
      // ActorName-Format "Nachname, Vorname" → in "Vorname Nachname" umdrehen
      const parts = actorName.split(',').map(s => s.trim());
      const flipped = parts.length === 2 ? `${parts[1]} ${parts[0]}` : actorName;
      return tgt === flipped || tgt === actorName;
    };
    const filtered = changeLogEntries.filter(e =>
      (!fa || (e.Action || '').toLowerCase().indexOf(fa) >= 0) &&
      (!fe || ((e.EventTitle || '').toLowerCase().indexOf(fe) >= 0 || (e.TargetName || '').toLowerCase().indexOf(fe) >= 0)) &&
      (!fac || (e.ActorName || '').toLowerCase().indexOf(fac) >= 0 || (e.ActorEmail || '').toLowerCase().indexOf(fac) >= 0) &&
      (!changeLogHideSelf || !isSelfAction(e))
    );
    const fmtDate = (iso: string): string => {
      if (!iso) return '';
      try { return new Date(iso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }
      catch { return iso; }
    };
    const actionColor = (a: string): string => {
      if (a.indexOf('Deleted') >= 0) return 'var(--dex-red, #c00)';
      if (a.indexOf('Created') >= 0) return 'var(--dex-green-dark)';
      if (a.indexOf('Cancelled') >= 0) return 'var(--dex-orange)';
      return 'var(--dex-gray-700)';
    };
    return (
      <div
        style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}
        onClick={() => setShowChangeLogModal(false)}
      >
        <div
          className="card"
          style={{ width: '100%', maxWidth: 1200, maxHeight: '90vh', overflow: 'auto', padding: 24, borderRadius: 16, background: '#fff' }}
          onClick={e => e.stopPropagation()}
        >
          <div className="flex-between mb-16">
            <h3 style={{ margin: 0 }}>
              <FileText size={18} /> {isDe ? 'Audit-Log (DEX_ChangeLog)' : 'Audit log (DEX_ChangeLog)'}
            </h3>
            <button onClick={() => setShowChangeLogModal(false)} style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}>
              <X size={20} />
            </button>
          </div>
          <p style={{ margin: '0 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
            {isDe
              ? <>Letzte <strong>{changeLogEntries.length}</strong> Einträge ({filtered.length} sichtbar). Schreibrechte: alle authentifizierten User; Leserechte: Organizer + Admin.</>
              : <>Last <strong>{changeLogEntries.length}</strong> entries ({filtered.length} visible). Write access: all authenticated users; read access: organizer + admin.</>}
          </p>
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.82rem', color: 'var(--dex-gray-700)', marginBottom: 12, cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={changeLogHideSelf}
              onChange={e => setChangeLogHideSelf(e.target.checked)}
            />
            {isDe
              ? 'Eigenaktionen der User ausblenden (nur Aktionen von Organizer/Admin anzeigen)'
              : 'Hide user self-actions (show only actions performed by organizer/admin)'}
          </label>
          {changeLogLoading && (
            <p style={{ textAlign: 'center', padding: 16, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Lade Einträge…' : 'Loading entries…'}
            </p>
          )}
          {!changeLogLoading && (
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                <thead style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50)' }}>
                  <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Datum' : 'Date'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Aktion' : 'Action'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Ziel' : 'Target'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Event' : 'Event'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Wer' : 'Actor'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Details' : 'Details'}</th>
                  </tr>
                  <tr style={{ borderBottom: '1px solid var(--dex-gray-200)', background: '#fff' }}>
                    <th style={{ padding: 4 }} />
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterAction} onChange={e => setChangeLogFilterAction(e.target.value)} placeholder={isDe ? 'z.B. Deleted' : 'e.g. Deleted'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }} />
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterEvent} onChange={e => setChangeLogFilterEvent(e.target.value)} placeholder={isDe ? 'Event-/Ziel-Name' : 'event/target'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterActor} onChange={e => setChangeLogFilterActor(e.target.value)} placeholder={isDe ? 'Name/E-Mail' : 'name/email'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }} />
                  </tr>
                </thead>
                <tbody>
                  {filtered.map(e => (
                    <tr key={e.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                      <td style={{ padding: 6, color: 'var(--dex-gray-600)', whiteSpace: 'nowrap' }}>{fmtDate(e.Created)}</td>
                      <td style={{ padding: 6, color: actionColor(e.Action), fontWeight: 600 }}>{e.Action}</td>
                      <td style={{ padding: 6 }}>{e.TargetName || e.TargetId || '-'}</td>
                      <td style={{ padding: 6, color: 'var(--dex-gray-700)' }}>{e.EventTitle || '-'}</td>
                      <td style={{ padding: 6 }}>
                        {e.ActorName || e.ActorEmail || '-'}
                        {e.ActorEmail && e.ActorName && (
                          <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{e.ActorEmail}</div>
                        )}
                      </td>
                      <td style={{ padding: 6, color: 'var(--dex-gray-600)', fontSize: '0.75rem', fontFamily: 'monospace', maxWidth: 320, wordBreak: 'break-word' }}>{e.Details}</td>
                    </tr>
                  ))}
                  {filtered.length === 0 && (
                    <tr><td colSpan={6} style={{ padding: 16, textAlign: 'center', color: 'var(--dex-gray-500)' }}>
                      {isDe ? 'Keine Einträge passen zum Filter.' : 'No entries match the filter.'}
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    );
  })() : null;

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
            {t('admin.noaccess.msg')}
          </p>
          <div style={{ marginTop: 24, display: 'flex', gap: 12, justifyContent: 'center', flexWrap: 'wrap' }}>
            {/* v7.2: Organizer ohne eigenes Event sehen hier einen direkten
                Shortcut zum Event-Erstellen — sonst sitzen sie in dieser
                Sackgasse ohne sichtbaren nächsten Schritt. Admins sehen den
                Button nicht, weil sie bereits alle Events in adminEvents haben. */}
            {currentUserRole !== 'User' && (
              <button className="btn btn-primary" onClick={() => navigate('create-event')}>
                + {t('admin.newevent')}
              </button>
            )}
            <button className="btn btn-secondary" onClick={() => navigate('landing')}>
              {t('reg.backtoevents') || 'Zurück'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (!selectedEvent) {
    // Event-Auswahl
    return (
      <div className="page-container" role="main" style={{ maxWidth: 1200, marginLeft: 'auto', marginRight: 'auto' }}>
        <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
        <h2 style={{ margin: '0 0 16px' }}>{t('admin.title')}</h2>
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
          {isAdmin && (
            <button className="btn btn-secondary" onClick={openChangeLog} style={{ fontSize: '0.85rem' }}>
              <FileText size={16} /> {isDe ? 'Audit-Log' : 'Audit log'}
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
                    {/* v9.11: Thumbnail-Container immer rendern (auch wenn kein Bild
                        gesetzt ist) — sonst rutscht der Text nach links und die
                        Reihen wirken inkonsistent neben Reihen mit Bild. */}
                    <div style={{
                      width: 60, height: 40, borderRadius: 'var(--dex-radius)', flexShrink: 0,
                      background: event.imageUrl
                        ? `url(${event.imageUrl}) center/cover no-repeat`
                        : 'linear-gradient(135deg, var(--dex-gray-200), var(--dex-gray-100))',
                      filter: opts?.muted ? 'grayscale(0.4)' : 'none',
                      display: 'flex', alignItems: 'center', justifyContent: 'center',
                      color: 'var(--dex-gray-400)', fontSize: '0.7rem',
                    }}>
                      {!event.imageUrl && '—'}
                    </div>
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
                      {/* v9.10: B2Run-Events haben maxParticipants=0 weil die Kapazitaet
                          auf durchstarter+funstarter aufgeteilt ist — Summe als
                          effektive Kapazitaet anzeigen statt "∞". */}
                      {(() => {
                        const split = (event.durchstarterCapacity || 0) + (event.funstarterCapacity || 0);
                        const eff = event.maxParticipants && event.maxParticipants > 0 ? event.maxParticipants : split;
                        return `${event.currentParticipants}/${eff || '∞'} Teilnehmer`;
                      })()}
                    </span>
                    {/* v9.20: Status-Badge mit Entwurfs-Override.
                        Wenn das Event als Entwurf markiert ist, wird "ENTWURF"
                        statt des EventStatus angezeigt — fuer den Organizer
                        ist dieser Hinweis wichtiger als der technische Status. */}
                    <span className="badge" style={{
                      background: event.isFictive ? 'rgba(237,139,0,0.15)' : getStatusColor(event.status) + '22',
                      color: event.isFictive ? 'var(--dex-orange-dark, #b35a00)' : getStatusColor(event.status),
                      fontWeight: 600,
                    }}>
                      {event.isFictive ? 'ENTWURF' : (isDe ? localizeStatus(event.status) : event.status)}
                    </span>
                    {/* v10.20 / v11.9: Migrations-Button fuer Legacy-B2Run-Events.
                        Erkennt das Event als 'altes B2Run' wenn entweder
                        type === 'B2Run' (alte EventType-Spalte) ODER mind.
                        ein b2run_*-Custom-Field in den eventSpecificFields
                        steht. Damit erscheint der Knopf auch wenn die alte
                        EventType-Spalte aus DEX_Events bereits geloescht
                        wurde — entscheidend ist die b2run_*-Spur in der
                        Felder-Konfiguration. Klick: entfernt b2run_*-Fields
                        aus customFields, persistiert 'Durchstarter' /
                        'Funstarter' als Gruppen-Labels, setzt EventType
                        best-effort auf 'Other'. Bestehende Anmeldungen,
                        Wartelisten und Sub-Events bleiben unveraendert. */}
                    {isAdmin && (event.type === 'B2Run' || (event.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_'))) && (
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 12px', color: 'var(--dex-green-dark, #4a7c1f)' }}
                        title={isDe
                          ? 'Auf neues Standard-Event-Schema migrieren (Type entfernen, Labels Durchstarter/Funstarter explizit speichern). Bestehende Anmeldungen bleiben unveraendert.'
                          : 'Migrate to the new standard event schema (drop type, persist Durchstarter/Funstarter labels). Existing registrations remain unchanged.'}
                        onClick={async (e) => {
                          e.stopPropagation();
                          if (!eventServiceRef) return;
                          // v11.9: Migration nimmt jetzt auch Legacy-B2Run-
                          // Sub-Events mit. Wir scannen alle Child-Events
                          // (childEventsOf) und migrieren die mit, die
                          // entweder type='B2Run' oder mind. ein b2run_*-
                          // Custom-Field haben.
                          const kids = childEventsOf(event.id);
                          const kidsToMigrate = kids.filter(k => k.type === 'B2Run' || (k.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_')));
                          const kidsHint = kidsToMigrate.length > 0
                            ? (isDe
                                ? `\n\nEs werden zusätzlich ${kidsToMigrate.length} Sub-Event(s) mitmigriert: ${kidsToMigrate.map(k => '„' + (k.title || '?') + '"').join(', ')}.`
                                : `\n\nAdditionally ${kidsToMigrate.length} sub-event(s) will be migrated: ${kidsToMigrate.map(k => '"' + (k.title || '?') + '"').join(', ')}.`)
                            : '';
                          const msg = isDe
                            ? `Event "${event.title}" auf Standard-Schema migrieren?\n\n• Type "B2Run" wird entfernt — Event sieht aus wie ein normales Deloitte-Event.\n• Bezeichnungen "Durchstarter" / "Funstarter" werden als Gruppen-Labels gespeichert (kannst du im Wizard frei ändern).\n• Falls Leistungsnachweis-Pflicht aktiv war: wird in ein reguläres Custom-Field „Leistungsnachweis vorhanden" (Checkbox, Pflicht, nur für Gruppe A) umgewandelt — bleibt also als richtige Frage erhalten.\n• Hardcoded Startblock-Mapping pro Gruppe wird ersatzlos entfernt. Bei Bedarf als Custom-Field mit Gruppen-Bindung wieder anlegen.\n• b2run_*-Custom-Fields (Altersgruppe, T-Shirt-Größe, Mobilnummer etc.) BLEIBEN als generische Custom-Fields erhalten — du kannst sie danach im Wizard umbenennen oder löschen, wenn nicht mehr gebraucht.\n• Anmeldungen, Wartelisten und Sub-Events bleiben inhaltlich unverändert.${kidsHint}`
                            : `Migrate event "${event.title}" to the standard schema?\n\n• Type "B2Run" is removed — the event will look like a standard Deloitte event.\n• Labels "Durchstarter" / "Funstarter" are persisted as group labels (editable later in the wizard).\n• If performance-proof requirement was active: it is converted into a regular custom field „Leistungsnachweis vorhanden" (checkbox, required, only for group A) — stays as a proper prompt.\n• Hardcoded per-group start-block mapping is removed. If needed, add it as a custom field bound to a group.\n• b2run_* custom fields (age group, t-shirt size, mobile etc.) are KEPT as generic custom fields — you can rename or remove them later in the wizard if no longer needed.\n• Registrations, waitlists and sub-events stay unchanged content-wise.${kidsHint}`;
                          if (!window.confirm(msg)) return;
                          const errors: string[] = [];
                          const migrateOne = async (ev: DeloitteEvent): Promise<void> => {
                            try {
                              // v11.11: KEINE Custom-Fields mehr löschen.
                              // Die b2run_*-Felder bleiben als generische
                              // Custom-Fields erhalten — der Organizer kann
                              // sie im Wizard danach selbst umbenennen oder
                              // entfernen. Vorher (v11.9) hat die Migration
                              // sie aggressiv aus customFields entfernt, was
                              // zu Datenverlust geführt hat (Altersgruppe,
                              // T-Shirt-Größe etc. waren weg, obwohl nur
                              // die Type-Spalte und Labels umgestellt
                              // werden sollten).
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              const keptFields: any[] = (ev.eventSpecificFields || []).map(f => ({ ...f }));
                              const splitActive = (ev.durchstarterCapacity || 0) > 0 && (ev.funstarterCapacity || 0) > 0;
                              const baseUpdates: Record<string, unknown> = {
                                'SplitLabelA': (ev.splitLabelA || 'Durchstarter'),
                                'SplitLabelB': (ev.splitLabelB || 'Funstarter'),
                              };
                              // v11.13: B2Run-Extras aus
                              // EmailTemplateOverrides._b2run nicht mehr nur
                              // löschen, sondern in echte Custom-Fields mit
                              // onlyForGroup-Bindung übersetzen:
                              // - durchstarterRequiresProof → Custom-Field
                              //   „Leistungsnachweis vorhanden" (Checkbox,
                              //   Pflicht, onlyForGroup='A').
                              // - durchstarterStartblock / funstarterStart-
                              //   block (Auto-Mapping) waren reine UI-
                              //   Convenience und werden ersatzlos entfernt.
                              //   Wenn der Organizer pro Gruppe einen
                              //   Startblock vorgeben will, lege er das
                              //   manuell als Custom-Field mit
                              //   onlyForGroup A bzw. B an.
                              try {
                                const overridesRaw = (ev.emailTemplateOverrides || '').toString();
                                if (overridesRaw.trim()) {
                                  const parsed = JSON.parse(overridesRaw);
                                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                  const b2 = parsed && typeof parsed === 'object' ? (parsed as any)._b2run : null;
                                  if (b2 && typeof b2 === 'object') {
                                    if (b2.durchstarterRequiresProof) {
                                      const PROOF_ID = 'b2run_leistungsnachweis';
                                      const existing = keptFields.find(f => String(f.id || '').toLowerCase() === PROOF_ID);
                                      if (existing) {
                                        existing.onlyForGroup = 'A';
                                        existing.required = true;
                                        if (!existing.label) existing.label = 'Leistungsnachweis vorhanden';
                                        if (!existing.type) existing.type = 'checkbox';
                                        if (!existing.helpText) existing.helpText = 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.';
                                      } else {
                                        keptFields.push({
                                          id: PROOF_ID,
                                          label: 'Leistungsnachweis vorhanden',
                                          type: 'checkbox',
                                          required: true,
                                          options: [],
                                          visible: true,
                                          onlyForGroup: 'A',
                                          helpText: 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.',
                                        });
                                      }
                                      baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                                    }
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    delete (parsed as any)._b2run;
                                    baseUpdates['EmailTemplateOverrides'] = JSON.stringify(parsed);
                                  }
                                }
                              } catch { /* invalid JSON → einfach ignorieren */ }
                              // v11.14: hardcoded B2Run-Field-Specials in
                              // echte Field-Properties migrieren (showIf
                              // für Mobilnummer, externalLinks für
                              // Datenschutz, required für Laufshirt).
                              const fieldExtras = migrateB2RunFieldExtras(keptFields);
                              if (fieldExtras.changed) {
                                baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                              }
                              const ok = await updateEvent(ev.id, baseUpdates);
                              try { await updateEvent(ev.id, { 'EventType': 'Other' }); } catch { /* SP-Spalte evtl. nicht vorhanden — ignoriert */ }
                              if (!ok) { errors.push(`„${ev.title}"`); return; }
                              // v11.11: Subsite-Spalten syncen — fehlende
                              // Spalten werden angelegt. Die b2run_*-Spalten
                              // bleiben drin, weil sie auch in customFields
                              // bleiben.
                              if (ev.subsiteUrl && eventServiceRef) {
                                try {
                                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                  const cfForFix: any[] = keptFields.map(f => ({
                                    id: f.id,
                                    label: f.label,
                                    type: f.type,
                                    required: !!f.required,
                                    visible: true,
                                    options: f.options || [],
                                    /* eslint-disable @typescript-eslint/no-explicit-any */
                                    spInternalName: (f as any).spInternalName || '',
                                    ...((f as any).helpText ? { helpText: (f as any).helpText } : {}),
                                    ...((f as any).multi ? { multi: true } : {}),
                                    ...((f as any).showIf ? { showIf: (f as any).showIf } : {}),
                                    /* eslint-enable @typescript-eslint/no-explicit-any */
                                  }));
                                  await eventServiceRef.fixRegistrationListColumns(ev.subsiteUrl, {
                                    isB2Run: splitActive,
                                    hasQuiz: (ev.quiz || []).length > 0,
                                    customFields: cfForFix,
                                  });
                                } catch (err) { console.warn('[DEX] fixRegistrationListColumns failed for', ev.id, err); }
                              }
                            } catch (err) {
                              console.warn('[DEX] migrate event failed:', ev.id, err);
                              errors.push(`„${ev.title}"`);
                            }
                          };
                          try {
                            await migrateOne(event);
                            for (const k of kidsToMigrate) {
                              await migrateOne(k);
                            }
                            await refreshEvents();
                            const total = 1 + kidsToMigrate.length;
                            if (errors.length === 0) {
                              window.alert(isDe
                                ? `Migration abgeschlossen — ${total} Event(s) auf das Standard-Schema umgestellt.`
                                : `Migration completed — ${total} event(s) migrated to the standard schema.`);
                            } else {
                              window.alert(isDe
                                ? `Migration teilweise fehlgeschlagen bei: ${errors.join(', ')}. Siehe Browser-Console.`
                                : `Migration partially failed for: ${errors.join(', ')}. See browser console.`);
                            }
                          } catch (err) {
                            console.warn('[DEX] migrate B2Run event failed:', err);
                            window.alert(isDe ? 'Migration fehlgeschlagen — siehe Browser-Console.' : 'Migration failed — see browser console.');
                          }
                        }}
                      >
                        {isDe ? 'B2Run migrieren' : 'Migrate B2Run'}
                      </button>
                    )}
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.8rem', padding: '6px 12px', color: 'var(--dex-red, #c00)' }}
                      onClick={(e) => {
                        e.stopPropagation();
                        setConfirmDeleteEvent(event);
                        setConfirmDeleteText('');
                      }}
                      disabled={isDeleting}
                    >
                      <Trash2 size={14} /> {isDeleting && deletingId === event.id ? 'Wird gelöscht...' : 'Löschen'}
                    </button>
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
        {dangerZoneModal}
        {changeLogModal}
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
      // v11.26: getrennte Vorname / Nachname Sortierung.
      case 'vorname': {
        const na = (a.Vorname || (a.ParticipantName || '').split(' ')[0] || '');
        const nb = (b.Vorname || (b.ParticipantName || '').split(' ')[0] || '');
        cmp = na.localeCompare(nb, 'de');
        break;
      }
      case 'nachname': {
        // Fallback fuer Alt-Daten ohne separates Vorname/Nachname:
        // Letztes Wort aus ParticipantName als Nachname.
        const lastWord = (s: string): string => {
          const parts = s.trim().split(/\s+/);
          return parts.length > 0 ? parts[parts.length - 1] : '';
        };
        const na = a.Nachname || lastWord(a.ParticipantName || '');
        const nb = b.Nachname || lastWord(b.ParticipantName || '');
        cmp = na.localeCompare(nb, 'de');
        break;
      }
      case 'email': cmp = (a.ParticipantEmail || '').localeCompare(b.ParticipantEmail || ''); break;
      case 'status': cmp = (a.Status || '').localeCompare(b.Status || ''); break;
      case 'date': cmp = new Date(a.RegistrationDate || 0).getTime() - new Date(b.RegistrationDate || 0).getTime(); break;
    }
    return sortAsc ? cmp : -cmp;
  };

  const handleSort = (col: 'id' | 'anrede' | 'vorname' | 'nachname' | 'email' | 'status' | 'date'): void => {
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
      {/* v9.29: Inline Zurück + Aktualisieren entfernt — beides liegt jetzt im Header.
          Eventauswahl-Reset („zurück zur Event-Liste") triggern wir über den Header-Back —
          siehe Listener weiter oben, der bei navigate-Wechsel selectedEvent zurücksetzt. */}

      {/* Event-Info + Aktionen
          Bei Hochformat-Bildern: Bild links neben den Detail-Rows.
          Bei Querformat-Bildern (oder solange die Orientierung noch
          unbekannt ist): Bild als Banner ueber den Detail-Rows. */}
      <div className="admin-event-info-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24, marginBottom: 24 }}>
        <div className="card" style={{ padding: 24 }}>
          {/* Header: Event-Titel + Status-Badge */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 16 }}>
            <h2 style={{ margin: 0, fontSize: '1.2rem', lineHeight: 1.2 }}>{selectedEvent.title}</h2>
            <span className="badge" style={{
              background: getStatusColor(selectedEvent.status) + '22',
              color: getStatusColor(selectedEvent.status),
            }}>
              {selectedEvent.isFictive ? 'ENTWURF' : (isDe ? localizeStatus(selectedEvent.status) : selectedEvent.status)}
            </span>
          </div>
          {/* Foto immer als Kreis links, Detail-Rows rechts. Layout
              unabhaengig vom Bildformat (cover-Crop sorgt fuer den Kreis). */}
          <div style={{ display: 'flex', gap: 20, alignItems: 'flex-start' }}>
            {selectedEvent.imageUrl && (
              <div
                style={{
                  flex: '0 0 auto',
                  width: 110,
                  height: 110,
                  borderRadius: '50%',
                  overflow: 'hidden',
                  background: 'var(--dex-gray-50, #fafafa)',
                  border: '1px solid var(--dex-gray-200, #e5e7eb)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                }}
              >
                <img
                  src={selectedEvent.imageUrl}
                  alt={selectedEvent.title}
                  style={{
                    display: 'block',
                    width: '100%',
                    height: '100%',
                    objectFit: 'cover',
                  }}
                />
              </div>
            )}
            <div style={{ flex: 1, minWidth: 0 }}>
              <h3 className="mb-16">{isDe ? 'Event-Details' : 'Event details'}</h3>
                {/* v11.28: Bookmark-Tabs statt Dropdown fuer schnelles Umschalten
                    zwischen Hauptevent und Sub-Events. Pro Tab wird die aktuelle
                    Teilnehmerzahl (currentParticipants aus EventContext) als
                    kleiner Badge angezeigt. */}
                {selectedEvent && (() => {
                  const isChild = !!selectedEvent.parentEventId;
                  const siblings = isChild
                    ? childEventsOf(selectedEvent.parentEventId || '')
                    : childEventsOf(selectedEvent.id);
                  if (!isChild && siblings.length === 0) return null;
                  const parent = isChild ? events.find(e => e.id === selectedEvent.parentEventId) : selectedEvent;
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const tabs: Array<{ id: string; label: string; count: number; isParent: boolean; ev: any }> = [];
                  if (parent) {
                    tabs.push({ id: parent.id, label: parent.title || (isDe ? 'Hauptevent' : 'Main event'), count: parent.currentParticipants || 0, isParent: true, ev: parent });
                  }
                  for (const c of siblings) {
                    tabs.push({ id: c.id, label: c.title || (isDe ? 'ohne Titel' : 'untitled'), count: c.currentParticipants || 0, isParent: false, ev: c });
                  }
                  return (
                    <div
                      role="tablist"
                      aria-label={isDe ? 'Event wechseln' : 'Switch event'}
                      style={{
                        display: 'flex', flexWrap: 'wrap', gap: 6,
                        marginBottom: 16,
                        borderBottom: '1px solid var(--dex-gray-200)',
                        paddingBottom: 0,
                      }}
                    >
                      {tabs.map(t => {
                        const active = t.id === selectedEvent.id;
                        return (
                          <button
                            key={t.id}
                            type="button"
                            role="tab"
                            aria-selected={active}
                            onClick={() => handleSelectEvent(t.ev).catch(() => { /* */ })}
                            style={{
                              display: 'inline-flex', alignItems: 'center', gap: 8,
                              padding: '8px 14px',
                              border: '1px solid var(--dex-gray-200)',
                              borderBottom: active ? '2px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-200)',
                              borderRadius: '8px 8px 0 0',
                              background: active ? '#fff' : 'var(--dex-gray-50, #fafafa)',
                              color: active ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)',
                              fontWeight: active ? 700 : 500,
                              fontSize: '0.85rem',
                              cursor: 'pointer',
                              marginBottom: -1,
                              whiteSpace: 'nowrap',
                              maxWidth: 280,
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              transition: 'background 0.15s, color 0.15s, border-color 0.15s',
                            }}
                            title={t.label}
                          >
                            {t.isParent && (
                              <span style={{ fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: 0.4, color: active ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>
                                {isDe ? 'Haupt' : 'Main'}
                              </span>
                            )}
                            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{t.label}</span>
                            <span
                              style={{
                                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                minWidth: 24, height: 20, padding: '0 6px',
                                borderRadius: 999,
                                background: active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)',
                                color: active ? '#fff' : 'var(--dex-gray-700)',
                                fontSize: '0.72rem', fontWeight: 700,
                              }}
                            >
                              {t.count}
                            </span>
                          </button>
                        );
                      })}
                    </div>
                  );
                })()}
              {/* Eigenes Row-Layout (zwei Spalten: Label fett, Wert links-
                  buendig). Das globale .settings-info SCSS macht stattdessen
                  space-between (also Wert rechts-buendig) — hier wollen wir
                  beide Spalten links ausgerichtet. */}
              {(() => {
                const rowStyle: React.CSSProperties = {
                  display: 'grid',
                  gridTemplateColumns: '160px 1fr',
                  gap: 12,
                  padding: '10px 0',
                  borderBottom: '1px solid var(--dex-gray-200)',
                  fontSize: '0.9rem',
                };
                const labelStyle: React.CSSProperties = { fontWeight: 700, color: 'var(--dex-gray-700)' };
                const valueStyle: React.CSSProperties = { fontWeight: 400, color: 'var(--dex-gray-800)' };
                return (
                  <>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Zeitraum' : 'Time period'}</span>
                      <span style={valueStyle}>{formatDate(selectedEvent.startDate)} - {formatDate(selectedEvent.endDate)}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Organizer' : 'Organizer'}</span>
                      <span style={valueStyle}>{selectedEvent.organizers.map(o => {
                        const parts = o.split(',').map(s => s.trim());
                        return parts.length === 2 ? `${parts[1]} ${parts[0]}` : o;
                      }).join(', ')}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Ort' : 'Location'}</span>
                      <span style={valueStyle}>{selectedEvent.location || '-'}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Max. Teilnehmer' : 'Max. attendees'}</span>
                      <span style={valueStyle}>{(() => {
                        // v9.11: B2Run-Events nutzen Split-Kapazitaet statt maxParticipants —
                        // hier die Summe anzeigen statt "Unbegrenzt".
                        const split = (selectedEvent.durchstarterCapacity || 0) + (selectedEvent.funstarterCapacity || 0);
                        const eff = selectedEvent.maxParticipants && selectedEvent.maxParticipants > 0
                          ? selectedEvent.maxParticipants
                          : split;
                        return eff || (isDe ? 'Unbegrenzt' : 'Unlimited');
                      })()}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Aktuell registriert' : 'Currently registered'}</span>
                      <span style={valueStyle}>{activeRegs.length}</span>
                    </div>
                    {waitlistRegs.length > 0 && (
                      <div style={rowStyle}>
                        <span style={labelStyle}>{isDe ? 'Warteliste' : 'Waitlist'}</span>
                        <span style={valueStyle}>{waitlistRegs.length}</span>
                      </div>
                    )}
                    {selectedEvent.eventSpecificFields && selectedEvent.eventSpecificFields.length > 0 && (
                      <div style={{ ...rowStyle, alignItems: 'flex-start', borderBottom: 'none' }}>
                        <span style={labelStyle}>{isDe ? 'Abgefragte Felder' : 'Requested fields'}</span>
                        <span style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                          {selectedEvent.eventSpecificFields.map(f => (
                            <span
                              key={f.id}
                              title={f.helpText || f.label}
                              style={{
                                display: 'inline-flex', alignItems: 'center', gap: 4,
                                padding: '2px 8px', borderRadius: 999,
                                background: 'rgba(134,188,37,0.10)',
                                border: '1px solid rgba(134,188,37,0.35)',
                                color: 'var(--dex-green-dark, #4a7c1f)',
                                fontSize: '0.72rem', fontWeight: 600, whiteSpace: 'nowrap',
                              }}
                            >
                              {f.label}
                              {f.required && <span style={{ color: 'var(--dex-red, #c00)' }}>*</span>}
                            </span>
                          ))}
                        </span>
                      </div>
                    )}
                  </>
                );
              })()}
            </div>
          </div>
        </div>

        {/* v7.6: Aktionen-Bereich als Kachel-Grid (auto-fit ab 220px, max 4
            pro Zeile auf Desktop). Default Grau, beim Hover Deloitte-Gruen mit
            leichtem Schatten. Jede Kachel zeigt SVG-Icon + Titel + ausfuehrliche
            Beschreibung + Rollen-Badge ("Organizer" oder "Nur Admin"). Die
            ehemals in der TN-Toolbar versteckten Wartungs-Aktionen (IDs neu
            vergeben, Spalten fixen, Felder reparieren, Profile neu laden) sind
            seit v7.6 hier integriert — der Organizer/Admin findet alle Event-
            relevanten Aktionen an einem Ort. QR-Scanner sehen den ganzen Block
            nicht. */}
        {!isQRScannerOnlyForSelected && (
        <div className="card" style={{ padding: 24 }}>
          <h3 className="mb-16">Aktionen</h3>
          <div className="admin-actions-grid" style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
            gap: 12,
          }}>
            {/* v9.20: Check-In starten — prominent als erster Tile.
                Sowohl Organizer als auch Check-In-Team-Mitglieder duerfen
                diese Aktion ausloesen (siehe Header.canCheckIn-Logik). */}
            <ActionTile
              icon={<Hash size={18} />}
              title={t('admin.checkin')}
              desc="Öffnet das Check-In-Tool: QR-Codes scannen, manuell ein-/auschecken, Live-KPIs (wie viele angemeldet / eingecheckt / ausstehend) sehen. Am Eventtag das wichtigste Werkzeug."
              badge="organizer"
              onClick={() => navigate('check-in', selectedEvent.id)}
            />

            {/* v9.20: QR-Codes versenden als ActionTile (Modal-Trigger). */}
            <ActionTile
              icon={<Send size={18} />}
              title={isSendingQR ? `QR-Codes werden versendet... (${qrSentCount})` : `QR-Codes versenden`}
              desc="Öffnet ein Modal mit drei Optionen: Test (nur an dich), Volldurchlauf an alle Angemeldeten, oder Auto-Send aktivieren (jede neue Anmeldung kriegt automatisch ihren QR-Code)."
              badge="organizer"
              busy={isSendingQR}
              onClick={() => {
                setQrAutoSendToggle(!!selectedEvent?.autoSendQRCode);
                setQrSendResult(null);
                setQrSendModalOpen(true);
              }}
            />

            {/* v9.19: Event aktivieren/deaktivieren — prominent gefuellt.
                Aktiv → User koennen sich anmelden. Deaktiv (Under Construction)
                → Event nur fuer Organizer/Admin/Test-Team sichtbar, Anmeldung
                blockiert. Der Toggle ist die schnellste Moeglichkeit, ein
                Event live zu schalten oder kurzfristig "auf Pause" zu setzen
                (z.B. waehrend kurzfristige Aenderungen). */}
            {(() => {
              const isActive = selectedEvent.status === 'Active';
              return (
                <ActionTile
                  icon={isActive ? <X size={18} /> : <Check size={18} />}
                  title={isActive ? 'Event deaktivieren' : 'Event aktivieren'}
                  desc={isActive
                    ? 'Setzt das Event auf "Entwurf" zurück — reguläre User können sich nicht mehr anmelden, sehen das Event nicht mehr in der Eventliste. Bestehende Anmeldungen bleiben erhalten. Du kannst jederzeit wieder auf "aktiv" stellen.'
                    : 'Schaltet das Event auf "Active" — ab jetzt sehen alle Berechtigten das Event in der Liste und können sich anmelden. Mails + Outlook-Termine laufen wie konfiguriert.'}
                  badge="organizer"
                  // v9.19/v9.20: nur "Aktivieren" wird gruen highlighted —
                  // "Deaktivieren" bleibt unauffaellig (Standard-Tile-Look),
                  // damit der Button nicht alarmierend wirkt.
                  accent={isActive ? undefined : 'green'}
                  onClick={async () => {
                    if (!eventServiceRef) return;
                    const newStatus = isActive ? 'Under Construction' : 'Active';
                    const confirmMsg = isActive
                      ? 'Event auf "Entwurf" zurücksetzen? Reguläre User sehen das Event danach nicht mehr.'
                      : 'Event auf "Active" schalten? Alle Berechtigten können sich danach anmelden.';
                    if (!window.confirm(confirmMsg)) return;
                    await updateEvent(selectedEvent.id, { 'EventStatus': newStatus });
                    await refreshEvents();
                  }}
                />
              );
            })()}

            {/* 1. Event bearbeiten */}
            <ActionTile
              icon={<Pencil size={18} />}
              title={t('admin.editbutton') || 'Event bearbeiten'}
              desc="Öffnet das Event im 7-Schritte-Wizard. Titel, Datum, Ort, Kapazität, Custom-Fields, E-Mail-Templates und Quiz nachträglich anpassen."
              badge="organizer"
              onClick={() => navigate('edit-event', selectedEvent.id)}
            />

            {/* 2. Teilnehmerliste in SharePoint öffnen */}
            <ActionTile
              icon={<ExternalLink size={18} />}
              title={t('admin.opensp') || 'In SharePoint öffnen'}
              desc="Öffnet die SharePoint-Teilnehmerliste der Subsite in einem neuen Tab — für tiefere Bearbeitung jenseits dieser App (z.B. Massen-Edit per Spreadsheet-View)."
              badge="organizer"
              href={selectedEvent.subsiteUrl ? `${selectedEvent.subsiteUrl}/Lists/Teilnehmer/AllItems.aspx` : `${siteUrl}/Lists`}
            />

            {/* v10.19: Deep-Link kopieren — Organizer/Admin können den Link
                des aktuell offenen Events in die Zwischenablage legen und z.B.
                an Co-Organizer / Helfer weitergeben. Zielseite ist exakt
                dieses Admin-Center-Detail (?action=admin&event=<SP-ID>). Beim
                Aufruf landet der Empfänger nach Login direkt auf der gleichen
                Detail-Ansicht statt in der Event-Auswahl-Liste. */}
            <ActionTile
              icon={<Link2 size={18} />}
              title={copiedDeepLink ? (t('admin.copied') || 'Kopiert') : 'Deep-Link kopieren'}
              desc="Legt den direkten Link auf dieses Event-Admin in die Zwischenablage. Per Mail / Teams an Co-Organizer schicken — sie landen nach Login direkt hier, ohne sich erst durch die Event-Liste klicken zu müssen."
              badge="organizer"
              onClick={() => {
                const base = (typeof window !== 'undefined' && window.location)
                  ? `${window.location.origin}${window.location.pathname}`
                  : `${siteUrl}/SitePages/DEX.aspx`;
                const url = `${base}?env=WebView&action=admin&event=${selectedEvent.id}`;
                if (typeof navigator !== 'undefined' && navigator.clipboard) {
                  navigator.clipboard.writeText(url).then(() => {
                    setCopiedDeepLink(true);
                    setTimeout(() => setCopiedDeepLink(false), 2000);
                  }).catch(() => { window.prompt('Deep-Link kopieren:', url); });
                } else {
                  window.prompt('Deep-Link kopieren:', url);
                }
              }}
            />

            {/* 3. E-Mail-Adressen kopieren */}
            <ActionTile
              icon={<Copy size={18} />}
              title={copiedEmails ? (t('admin.copied') || 'Kopiert') : (t('admin.copyemails') || 'E-Mails kopieren')}
              desc="Legt alle aktiven Teilnehmer-Mails (Semikolon-getrennt) in die Zwischenablage. Direkt in Outlook-Empfänger oder externe Tools einfügbar."
              badge="organizer"
              onClick={() => {
                const emails = registrations
                  .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
                  .map(r => r.ParticipantEmail)
                  .join('; ');
                if (emails) {
                  navigator.clipboard.writeText(emails).then(() => {
                    setCopiedEmails(true);
                    setTimeout(() => setCopiedEmails(false), 2000);
                  }).catch(() => { window.prompt('E-Mail-Adressen kopieren:', emails); });
                }
              }}
            />

            {/* 4. Massenmail an alle aktiven Teilnehmer */}
            <ActionTile
              icon={<Mail size={18} />}
              title={t('admin.emailall') || 'Massenmail senden'}
              desc="Öffnet einen RichText-Editor mit Deloitte-Mail-Template. Geht an alle aktiven Teilnehmer (nicht Wartelistler / Abgemeldete)."
              badge="organizer"
              onClick={() => {
                setEmailSubject(selectedEvent ? `${selectedEvent.title} - Info` : '');
                setEmailHeading(selectedEvent ? selectedEvent.title : '');
                setEmailBody('');
                setShowEmailModal(true);
              }}
            />

            {/* 5. Excel-Download (mit Dropdown Deloitte/B2Run-View)
                Wrapper braucht display:flex, damit der innere Button auf die
                volle Grid-Zellen-Hoehe gestreckt wird — sonst sieht die Kachel
                niedriger aus als ihre Nachbarn, die zwei Zeilen Titel haben. */}
            <div style={{ position: 'relative', display: 'flex' }}>
              <ActionTile
                icon={<Download size={18} />}
                title="Excel-Export"
                desc={selectedEvent && selectedEvent.type === 'B2Run'
                  ? "Lädt die Teilnehmerliste als Excel. Wahl zwischen 'Deloitte Felder' (alle internen Spalten + Custom-Fields) oder 'B2Run View' (importierbar in b2run.com)."
                  : "Lädt die Teilnehmerliste als Excel mit allen internen Spalten + Custom-Fields des Events."}
                badge="organizer"
                onClick={() => {
                  // Bei B2Run-Events: Dropdown mit Wahl zwischen Deloitte-
                  // und B2Run-View. Bei normalen Events gibt es nur den
                  // Deloitte-Export — direkt ausloesen, kein Dropdown noetig.
                  if (selectedEvent && selectedEvent.type === 'B2Run') {
                    setShowExportMenu(!showExportMenu);
                  } else {
                    exportCsv('deloitte');
                  }
                }}
              />
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
                      Alle internen Felder: Name, E-Mail, Department, Standort, Position, Status, Registrierungsdatum + alle Custom-Fields des Events.
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

            {/* 6. Outlook-Absagen prüfen — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<AlertCircle size={18} />}
                title={isCheckingDeclines ? 'Outlook wird geprüft…' : 'Outlook-Absagen prüfen'}
                desc="Liest die Outlook-Absagen aus dem no_reply.events-Postfach und matched sie gegen aktive Teilnehmer. Zeigt, wer den Termin abgelehnt hat, aber noch in der Liste steht."
                badge="admin"
                busy={isCheckingDeclines}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent) return;
                  setIsCheckingDeclines(true);
                  setDeclineResult(null);
                  setDeclineCopied(false);
                  try {
                    const result = await eventServiceRef.getDeclinedAttendees(selectedEvent.id);
                    if (result.ok) {
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
                      let msg = result.message || 'Unbekannter Fehler beim Lesen des Outlook-Termins.';
                      if (!result.message) {
                        if (result.reason === 'no-pointer') {
                          msg = 'Für dieses Event ist kein Outlook-Termin verknüpft (OutlookEventId / CalendarLink fehlen).';
                        } else if (result.reason === 'not-found') {
                          msg = 'Outlook-Termin wurde im Postfach no_reply.events@deloitte.de nicht gefunden.';
                        } else if (result.reason === 'forbidden') {
                          msg = 'Graph-API-Zugriff abgelehnt (HTTP 403). Tenant-Admin muss "Calendars.Read.Shared" genehmigen, und der User braucht Reviewer-Rechte auf dem Postfach-Kalender.';
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
              />
            )}

            {/* 7. TeilnehmerIDs neu vergeben — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<Hash size={18} />}
                title={isReorderingIDs ? 'IDs werden vergeben…' : 'IDs neu vergeben'}
                desc="Vergibt die TeilnehmerIDs sequentiell (1, 2, 3, …) nach Erstellungsreihenfolge. Schließt Lücken nach Stornos und sortiert die Liste sauber durch."
                badge="admin"
                busy={isReorderingIDs}
                disabled={!selectedEvent?.subsiteUrl}
                result={reorderResult}
                resultIsError={!!reorderResult && reorderResult.indexOf('Fehler') >= 0}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!confirm('TeilnehmerIDs neu vergeben (1, 2, 3, …)? Sortierung nach Erstellungsreihenfolge.')) return;
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
              />
            )}

            {/* 7b. Counter zurücksetzen — Admin only (v9.13 → v11.27).
                Recovery-Button um den DEX_TeilnehmerCounter EXAKT auf
                max(TeilnehmerID) der Subsite zu setzen. Bidirektional:
                Counter wird hochgezogen wenn er drunter steht (gegen
                Doppel-IDs), oder runtergesetzt wenn er drueber steht
                (z.B. nach vielen Abmeldungen, die TIDs gefressen
                haben). Vorher (vor v11.27) lief es nur monotonic-up,
                weshalb ein zu hoher Counter (Counter=11, Max-TID=4)
                nicht zurueckgesetzt wurde — Klick auf den Button
                hatte dann keinen sichtbaren Effekt. */}
            {isAdmin && (
              <ActionTile
                icon={<Hash size={18} />}
                title={isResettingCounter ? 'Counter wird zurückgesetzt…' : 'Counter zurücksetzen'}
                desc="Setzt den TeilnehmerID-Counter exakt auf den aktuellen Max-TID der Teilnehmerliste. Hilft, wenn neue Anmeldungen mit zu hohen IDs starten (Lücken durch frühere Abmeldungen) oder wenn sie versehentlich bei zu niedrigen IDs (z.B. wieder bei 1) starten würden. Bidirektional — egal ob der Counter zu hoch oder zu niedrig steht."
                badge="admin"
                busy={isResettingCounter}
                disabled={!selectedEvent?.subsiteUrl}
                result={resetCounterResult}
                resultIsError={!!resetCounterResult && resetCounterResult.indexOf('Fehler') >= 0}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!confirm('Counter auf aktuellen Max-Wert zurücksetzen?')) return;
                  setIsResettingCounter(true);
                  setResetCounterResult(null);
                  try {
                    const result = await eventServiceRef.resetCounterToMax(selectedEvent.subsiteUrl);
                    setResetCounterResult(`Counter steht jetzt auf ${result.counter} (Max-TID: ${result.max})`);
                  } catch {
                    setResetCounterResult('Fehler beim Zurücksetzen des Counters');
                  }
                  setIsResettingCounter(false);
                }}
              />
            )}

            {/* 8. Spalten fixen — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<Columns size={18} />}
                title={isFixingColumns ? 'Spalten werden gefixt…' : 'Spalten fixen'}
                desc="Legt fehlende SP-Spalten in der Teilnehmerliste an, entfernt überflüssige (z.B. StarterType bei Nicht-B2Run-Events) und korrigiert die Default-View-Reihenfolge."
                badge="admin"
                busy={isFixingColumns}
                disabled={!selectedEvent?.subsiteUrl}
                result={fixColumnsResult}
                resultIsError={!!fixColumnsResult && fixColumnsResult.indexOf('Fehler') >= 0}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  setIsFixingColumns(true);
                  setFixColumnsResult(null);
                  try {
                    const isB2Run = !!(selectedEvent.durchstarterCapacity || selectedEvent.funstarterCapacity);
                    const hasQuiz = !!(selectedEvent.quiz && selectedEvent.quiz.length > 0);
                    const customFields = (selectedEvent.eventSpecificFields || []).map(f => ({
                      id: f.id, label: f.label, type: f.type, required: f.required, options: f.options,
                      visible: true,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      spInternalName: (f as any).spInternalName || '',
                    }));
                    const result = await eventServiceRef.fixRegistrationListColumns(
                      selectedEvent.subsiteUrl,
                      { isB2Run, hasQuiz, customFields }
                    );
                    const msgs: string[] = [];
                    if (result.added.length > 0) msgs.push(`Spalten hinzugefügt: ${result.added.join(', ')}`);
                    if (result.removed.length > 0) msgs.push(`Spalten entfernt: ${result.removed.join(', ')}`);
                    if (result.viewFixed) msgs.push('View-Reihenfolge korrigiert');
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
                    setFixColumnsResult(msgs.length > 0 ? msgs.join(' | ') : 'Alles OK, keine Änderungen nötig');
                  } catch {
                    setFixColumnsResult('Fehler beim Fixen der Spalten');
                  }
                  setIsFixingColumns(false);
                }}
              />
            )}

            {/* 8b. Organizer-Mails reparieren (alle Events) — Admin only.
                Findet Events mit Längen-Mismatch zwischen organizers (Names) und
                organizerEmails — typisch nach Legacy-Korruption aus v10.0–v10.2-
                Closure-Bug. Versucht via Graph-Search die fehlenden Emails per
                Lastname-Match nachzufüllen. Persistiert das gefixte Pair-Mapping
                via updateEvent. Bricht NICHT bei einzelnen Fehlern ab — ein Event
                mit unauflösbarem Namen wird übersprungen, der Rest läuft weiter.
                Operiert über ALLE adminEvents (nicht nur das gerade ausgewählte). */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                title={isRepairingOrganizers ? 'Reparatur läuft…' : 'Organizer-Mails reparieren (alle Events)'}
                desc="Scannt alle Events nach Mismatches zwischen Organizer-Namen und Organizer-Emails (Legacy-Korruption aus früheren App-Versionen). Sucht fehlende Emails per Tenant-Suche über den Nachnamen und persistiert die gefixten Paare. Manuell nicht auflösbare Personen bleiben mit leerem Email-Slot — User muss diese im Wizard nachziehen."
                badge="admin"
                busy={isRepairingOrganizers}
                result={repairOrganizersResult}
                resultIsError={!!repairOrganizersResult && repairOrganizersResult.indexOf('Fehler') >= 0}
                onClick={async () => {
                  if (!eventServiceRef) return;
                  if (!confirm(`Organizer-Mails über ALLE ${adminEvents.length} Events reparieren? Dauert je nach Anzahl ca. 1–2 Minuten und schreibt direkt in DEX_Events zurück.`)) return;
                  setIsRepairingOrganizers(true);
                  setRepairOrganizersResult(null);
                  let scanned = 0;
                  let mismatched = 0;
                  let eventsUpdated = 0;
                  let orgsRecovered = 0;
                  let orgsUnresolved = 0;
                  const unresolvedNames: string[] = [];
                  try {
                    for (const ev of adminEvents) {
                      scanned++;
                      const names = (ev.organizers || []).slice();
                      const emails = (ev.organizerEmails || []).slice();
                      // Pad das kürzere Array auf max() — nichts geht verloren
                      const max = Math.max(names.length, emails.length);
                      while (names.length < max) names.push('');
                      while (emails.length < max) emails.push('');
                      // Mismatch erkannt? Mindestens ein Name ohne Email oder eine Email ohne Name
                      const hasMismatch = names.some((n, i) => (n || '').trim() && !((emails[i] || '').trim()))
                        || emails.some((e, i) => (e || '').trim() && !((names[i] || '').trim()));
                      if (!hasMismatch) continue;
                      mismatched++;
                      // Für jeden Slot mit Name aber ohne Email: Graph-Search nach Lastname
                      // pro Slot, EINS-zu-EINS-Match wenn Local-Part den Lastname enthält.
                      let recoveredHere = 0;
                      const fixedNames = names.slice();
                      const fixedEmails = emails.slice();
                      for (let i = 0; i < max; i++) {
                        const name = (fixedNames[i] || '').trim();
                        const email = (fixedEmails[i] || '').trim();
                        if (!name || email) continue;
                        // Lastname extrahieren — egal ob "Lastname, Firstname" oder
                        // "Firstname Lastname", als Suchquery für Graph nehmen wir
                        // den ganzen Namen (Graph ist tolerant).
                        try {
                          // Lastname als Suchterm — Graph-Search ist tolerant für
                          // 'Lastname' als Query und liefert eindeutigere Ergebnisse
                          // als die kombinierte Form 'Lastname, Firstname'.
                          const queryRaw = name.indexOf(',') >= 0 ? name.split(',')[0].trim() : name;
                          const hits = await searchUsers(queryRaw);
                          // Lastname-Substring-Match: filtere die Hits auf Personen,
                          // deren Email-Local-Part den Lastname enthält. Damit greifen
                          // wir den richtigen Eintrag auch bei Häufigkeitsnamen.
                          const lastname = queryRaw.toLowerCase().split(/\s+/).filter(t => t.length >= 3).pop() || '';
                          const matched = lastname
                            ? hits.filter(h => ((h.email || '').toLowerCase().split('@')[0]).indexOf(lastname) >= 0)
                            : hits;
                          if (matched.length === 1 && matched[0].email) {
                            fixedEmails[i] = matched[0].email;
                            recoveredHere++;
                            orgsRecovered++;
                          } else if (matched.length === 0 && hits.length === 1 && hits[0].email) {
                            // Kein Lastname-Match aber genau 1 Treffer überhaupt — übernehmen
                            fixedEmails[i] = hits[0].email;
                            recoveredHere++;
                            orgsRecovered++;
                          } else {
                            // Mehrdeutig oder nichts gefunden — leer lassen
                            unresolvedNames.push(`${name} (Event ${ev.eventNumber})`);
                            orgsUnresolved++;
                          }
                        } catch {
                          unresolvedNames.push(`${name} (Event ${ev.eventNumber})`);
                          orgsUnresolved++;
                        }
                      }
                      // Nichts wiederhergestellt? Skip Update — Storage ist eh schon
                      // im aktuellen Zustand, Pad allein bringt keinen Mehrwert
                      // (bei Save aus Wizard heilt sich das ohnehin).
                      if (recoveredHere === 0) continue;
                      // Alle vollständig leeren Slots aussortieren bevor wir schreiben
                      const finalPairs = fixedNames.map((n, i) => ({ n: (n || '').trim(), e: (fixedEmails[i] || '').trim() }))
                        .filter(p => p.n || p.e);
                      const finalNames = finalPairs.map(p => p.n).join('; ');
                      const finalEmails = finalPairs.map(p => p.e).join(';');
                      try {
                        const ok = await updateEvent(ev.id, { 'Organizer': finalNames, 'OrganizerEmail': finalEmails });
                        if (ok) eventsUpdated++;
                      } catch {
                        // Update fehlgeschlagen — counts trotzdem belassen, einfach
                        // skip dieses Event.
                      }
                    }
                    const lines = [`Gescannt: ${scanned}`, `Mit Mismatch: ${mismatched}`, `Aktualisiert: ${eventsUpdated}`, `Emails wiederhergestellt: ${orgsRecovered}`];
                    if (orgsUnresolved > 0) {
                      lines.push(`Manuell nachziehen (${orgsUnresolved}): ${unresolvedNames.slice(0, 5).join(', ')}${unresolvedNames.length > 5 ? '…' : ''}`);
                    }
                    setRepairOrganizersResult(lines.join(' · '));
                  } catch (err) {
                    setRepairOrganizersResult(`Fehler: ${err instanceof Error ? err.message : String(err)}`);
                  }
                  setIsRepairingOrganizers(false);
                }}
              />
            )}

            {/* v11.9: B2Run-Migration als Action-Tile im Admin-Event-Detail.
                Erkennt Legacy-Events (type='B2Run' oder b2run_*-Custom-
                Fields vorhanden) und bietet die gleiche Migration an wie
                der „B2Run migrieren"-Button in der Event-Liste. Damit
                findet der Admin den Knopf auch wenn er das Event bereits
                ausgewählt hat. */}
            {isAdmin && selectedEvent && (selectedEvent.type === 'B2Run' || (selectedEvent.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_'))) && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                title="Legacy-B2Run migrieren"
                desc="Entfernt den B2Run-Type und persistiert 'Durchstarter' / 'Funstarter' als reguläre Gruppen-Labels (kannst du danach im Wizard frei umbenennen). b2run_*-Custom-Fields (Altersgruppe, T-Shirt-Größe etc.) BLEIBEN als generische Custom-Fields erhalten. Anmeldungen, Wartelisten und Sub-Events bleiben unverändert."
                badge="admin"
                onClick={async () => {
                  if (!eventServiceRef) return;
                  const kids = childEventsOf(selectedEvent.id);
                  const kidsToMigrate = kids.filter(k => k.type === 'B2Run' || (k.eventSpecificFields || []).some(f => (f.id || '').toLowerCase().startsWith('b2run_')));
                  const msg = `Event "${selectedEvent.title}" auf Standard-Schema migrieren?\n\n` +
                    `• B2Run-Type wird entfernt — Event sieht aus wie ein normales Deloitte-Event.\n` +
                    `• Bezeichnungen "Durchstarter" / "Funstarter" werden als Gruppen-Labels gespeichert (frei umbenennbar im Wizard).\n` +
                    `• Falls Leistungsnachweis-Pflicht aktiv war: wird in ein reguläres Custom-Field „Leistungsnachweis vorhanden" (Checkbox, Pflicht, nur für Gruppe A) umgewandelt.\n` +
                    `• Hardcoded Startblock-Mapping pro Gruppe wird ersatzlos entfernt.\n` +
                    `• b2run_*-Custom-Fields (Altersgruppe, T-Shirt-Größe, Mobilnummer etc.) BLEIBEN als generische Custom-Fields erhalten.\n` +
                    `• Anmeldungen, Wartelisten und Sub-Events bleiben inhaltlich unverändert.\n\n` +
                    (kidsToMigrate.length > 0
                      ? `Es werden zusätzlich ${kidsToMigrate.length} Sub-Event(s) mitmigriert: ${kidsToMigrate.map(k => '„' + (k.title || '?') + '"').join(', ')}.`
                      : `Keine Sub-Events mit Legacy-B2Run-Spuren gefunden — nur das Hauptevent wird migriert.`);
                  if (!window.confirm(msg)) return;
                  const errors: string[] = [];
                  const migrateOne = async (ev: DeloitteEvent): Promise<void> => {
                    try {
                      // v11.11: Custom-Fields werden NICHT mehr gelöscht.
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const keptFields: any[] = (ev.eventSpecificFields || []).map(f => ({ ...f }));
                      const baseUpdates: Record<string, unknown> = {
                        'SplitLabelA': (ev.splitLabelA || 'Durchstarter'),
                        'SplitLabelB': (ev.splitLabelB || 'Funstarter'),
                      };
                      // v11.13: B2Run-Extras aus EmailTemplateOverrides._b2run
                      // in echte Custom-Fields mit onlyForGroup übersetzen
                      // (siehe ausführlicher Kommentar im Card-Button-Pfad).
                      try {
                        const overridesRaw = (ev.emailTemplateOverrides || '').toString();
                        if (overridesRaw.trim()) {
                          const parsed = JSON.parse(overridesRaw);
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const b2 = parsed && typeof parsed === 'object' ? (parsed as any)._b2run : null;
                          if (b2 && typeof b2 === 'object') {
                            if (b2.durchstarterRequiresProof) {
                              const PROOF_ID = 'b2run_leistungsnachweis';
                              const existing = keptFields.find(f => String(f.id || '').toLowerCase() === PROOF_ID);
                              if (existing) {
                                existing.onlyForGroup = 'A';
                                existing.required = true;
                                if (!existing.label) existing.label = 'Leistungsnachweis vorhanden';
                                if (!existing.type) existing.type = 'checkbox';
                                if (!existing.helpText) existing.helpText = 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.';
                              } else {
                                keptFields.push({
                                  id: PROOF_ID,
                                  label: 'Leistungsnachweis vorhanden',
                                  type: 'checkbox',
                                  required: true,
                                  options: [],
                                  visible: true,
                                  onlyForGroup: 'A',
                                  helpText: 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.',
                                });
                              }
                              baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                            }
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            delete (parsed as any)._b2run;
                            baseUpdates['EmailTemplateOverrides'] = JSON.stringify(parsed);
                          }
                        }
                      } catch { /* invalid JSON → einfach ignorieren */ }
                      // v11.14: hardcoded B2Run-Field-Specials in echte
                      // Field-Properties migrieren.
                      const fieldExtras = migrateB2RunFieldExtras(keptFields);
                      if (fieldExtras.changed) {
                        baseUpdates['CustomFields'] = JSON.stringify(keptFields);
                      }
                      const ok = await updateEvent(ev.id, baseUpdates);
                      try { await updateEvent(ev.id, { 'EventType': 'Other' }); } catch { /* SP-Spalte evtl. nicht vorhanden — ignoriert */ }
                      if (!ok) errors.push(`„${ev.title}"`);
                    } catch (err) {
                      console.warn('[DEX] migrate event failed:', ev.id, err);
                      errors.push(`„${ev.title}"`);
                    }
                  };
                  try {
                    // Hauptevent zuerst, dann alle b2run-Sub-Events.
                    await migrateOne(selectedEvent);
                    for (const k of kidsToMigrate) {
                      await migrateOne(k);
                    }
                    await refreshEvents();
                    if (errors.length === 0) {
                      const total = 1 + kidsToMigrate.length;
                      window.alert(`Migration abgeschlossen — ${total} Event(s) auf das Standard-Schema umgestellt.`);
                    } else {
                      window.alert(`Migration teilweise fehlgeschlagen bei: ${errors.join(', ')}. Siehe Browser-Console für Details.`);
                    }
                  } catch (err) {
                    console.warn('[DEX] migrate B2Run event failed:', err);
                    window.alert('Migration fehlgeschlagen — siehe Browser-Console.');
                  }
                }}
              />
            )}

            {/* v11.11: Custom-Fields aus Versionsverlauf zurückholen.
                Hilft den Admins, denen die v11.9-Migration die b2run_*-
                Felder (Altersgruppe, T-Shirt-Größe etc.) versehentlich
                aus customFields entfernt hat. Liest die SP-Versionen des
                Event-Items, sucht die jüngste Version mit b2run_*-
                Feldern und mergt diese zurück in das aktuelle
                CustomFields-Array. Bestehende Felder bleiben unverändert
                — es werden NUR fehlende b2run_*-Felder ergänzt. */}

            {isAdmin && selectedEvent && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                title="Custom-Fields aus Versionsverlauf zurückholen"
                desc="Liest den SharePoint-Versionsverlauf des Events und holt verloren gegangene b2run_*-Custom-Fields (Altersgruppe, T-Shirt-Größe, Startblock, Mobilnummer etc.) zurück. Nützlich nach der v11.9-Migration, die diese Felder versehentlich gelöscht hat. Bestehende Felder werden NICHT überschrieben — es werden nur fehlende Felder ergänzt."
                badge="admin"
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent) return;
                  try {
                    const history = await eventServiceRef.getEventCustomFieldsHistory(parseInt(selectedEvent.id, 10));
                    if (history.length === 0) {
                      window.alert('Kein Versionsverlauf gefunden — entweder hat das Event keine Versionen oder der Zugriff wurde verweigert.');
                      return;
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const currentFields: any[] = (selectedEvent.eventSpecificFields || []).map(f => ({ ...f }));
                    const currentIds = new Set(currentFields.map(f => String(f.id || '').toLowerCase()));
                    // Jüngste Version mit b2run_*-Feldern finden, die noch
                    // NICHT in currentFields stecken.
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    let foundFields: any[] = [];
                    let foundVersion = '';
                    let foundModified = '';
                    for (const v of history) {
                      const missingB2run = v.customFields.filter(f => {
                        const id = String(f.id || '').toLowerCase();
                        return id.indexOf('b2run_') === 0 && !currentIds.has(id);
                      });
                      if (missingB2run.length > 0) {
                        foundFields = missingB2run;
                        foundVersion = v.versionLabel;
                        foundModified = v.modified;
                        break;
                      }
                    }
                    if (foundFields.length === 0) {
                      window.alert('Keine fehlenden b2run_*-Felder im Versionsverlauf gefunden — entweder sind alle Felder schon vorhanden oder es gab nie welche.');
                      return;
                    }
                    const fieldList = foundFields.map(f => `• ${f.label || f.id}`).join('\n');
                    const modifiedDate = foundModified ? new Date(foundModified).toLocaleString('de-DE') : '?';
                    if (!window.confirm(`Folgende ${foundFields.length} Custom-Field(s) aus Version ${foundVersion} (${modifiedDate}) zurückholen?\n\n${fieldList}\n\nDie Felder werden ans Ende deiner aktuellen Felder-Liste angehängt. Du kannst sie danach im Wizard frei umbenennen, neu sortieren oder löschen.`)) {
                      return;
                    }
                    const merged = [...currentFields, ...foundFields];
                    const ok = await updateEvent(selectedEvent.id, { 'CustomFields': JSON.stringify(merged) });
                    if (!ok) {
                      window.alert('Update fehlgeschlagen — siehe Browser-Console.');
                      return;
                    }
                    // Subsite-Spalten gleich mit-syncen, damit die b2run_*-
                    // Spalten in der Teilnehmerliste wieder existieren.
                    if (selectedEvent.subsiteUrl) {
                      try {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const cfForFix: any[] = merged.map((f: any) => ({
                          id: f.id,
                          label: f.label,
                          type: f.type,
                          required: !!f.required,
                          visible: true,
                          options: f.options || [],
                          spInternalName: f.spInternalName || '',
                          ...(f.helpText ? { helpText: f.helpText } : {}),
                          ...(f.multi ? { multi: true } : {}),
                          ...(f.showIf ? { showIf: f.showIf } : {}),
                        }));
                        const splitActive = (selectedEvent.durchstarterCapacity || 0) > 0 && (selectedEvent.funstarterCapacity || 0) > 0;
                        await eventServiceRef.fixRegistrationListColumns(selectedEvent.subsiteUrl, {
                          isB2Run: splitActive,
                          hasQuiz: (selectedEvent.quiz || []).length > 0,
                          customFields: cfForFix,
                        });
                      } catch (err) { console.warn('[DEX] fixRegistrationListColumns nach Restore fehlgeschlagen:', err); }
                    }
                    await refreshEvents();
                    window.alert(`${foundFields.length} Custom-Field(s) erfolgreich aus Version ${foundVersion} zurückgeholt.`);
                  } catch (err) {
                    console.warn('[DEX] restore custom fields from history failed:', err);
                    window.alert('Zurückholen fehlgeschlagen — siehe Browser-Console.');
                  }
                }}
              />
            )}

            {/* 9. Felder reparieren — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                title={isFixingFields ? 'Felder werden repariert…' : 'Felder reparieren'}
                desc="Normalisiert Custom-Fields: AGB/Datenschutz → Checkbox, T-Shirt → 'Kein T-Shirt'-Option, B2Run-Spezialfelder ergänzen, redundante '(Pflicht)'-Suffixe entfernen."
                badge="admin"
                busy={isFixingFields}
                disabled={!selectedEvent}
                result={fixFieldsResult}
                resultIsError={!!fixFieldsResult && (fixFieldsResult.startsWith('Fehler') || fixFieldsResult.startsWith('Update fehl'))}
                onClick={async () => {
                  if (!selectedEvent) return;
                  setIsFixingFields(true);
                  setFixFieldsResult(null);
                  try {
                    const changes: string[] = [];
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const raw: any[] = (selectedEvent.eventSpecificFields || []).map((f: any) => ({ ...f }));
                    const hasField = (id: string): boolean => raw.some(f => f.id === id);
                    const isB2Run = raw.some(f => String(f.id || '').indexOf('b2run_') === 0);
                    if (isB2Run) {
                      if (!hasField('b2run_infoservice')) {
                        raw.push({ id: 'b2run_infoservice', label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergänzt: 'Infoservice'");
                      }
                      if (!hasField('b2run_anonym')) {
                        raw.push({ id: 'b2run_anonym', label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergänzt: 'Anonym teilnehmen'");
                      }
                      const hasLaufshirt = raw.some(f => f.id === 'b2run_laufshirt' || /laufshirt/i.test(String(f.label || '')));
                      if (!hasLaufshirt) {
                        raw.push({ id: 'b2run_laufshirt', label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true });
                        changes.push("Feld ergänzt: 'Deloitte-Laufshirt' (Pflicht)");
                      }
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const fixed = raw.map((f: any) => {
                      const nf = { ...f };
                      const label = String(nf.label || '');
                      const lowLabel = label.toLowerCase();
                      const isConsent = lowLabel.indexOf('zustimmung') >= 0
                        || lowLabel.indexOf('agb') >= 0
                        || lowLabel.indexOf('datenschutz') >= 0;
                      const isB2RunCheckbox = ['b2run_infoservice', 'b2run_anonym', 'b2run_datenschutz'].indexOf(nf.id) >= 0;
                      if ((isConsent || isB2RunCheckbox) && nf.type !== 'checkbox') {
                        nf.type = 'checkbox';
                        nf.options = [];
                        changes.push(`${label} -> Checkbox`);
                      }
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
                      const stripped = label.replace(/\s*\((?:pflicht|mandatory|required)\)\s*$/i, '').trim();
                      if (stripped && stripped !== label) {
                        nf.label = stripped;
                        changes.push(`Label "${label}" -> "${stripped}"`);
                      }
                      if (nf.id === 'b2run_mobilnummer') {
                        if (nf.required) { nf.required = false; changes.push('Mobilnummer -> optional'); }
                        if (nf.label === 'Mobilnummer') {
                          nf.label = 'Mobilnummer (nur bei aktiviertem Infoservice)';
                          changes.push("Mobilnummer-Label präzisiert");
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
                          changes.push('B2Run-Datenschutz: AGB + Datenschutz Links ergänzt');
                        }
                      }
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
                            changes.push(`${label || nf.id}: 'Habe bereits ein Laufshirt'-Option hinzugefügt`);
                          }
                        }
                      }
                      return nf;
                    });
                    const dsIdx = fixed.findIndex((f: { id: string }) => f.id === 'b2run_datenschutz');
                    if (dsIdx >= 0 && dsIdx !== fixed.length - 1) {
                      const [ds] = fixed.splice(dsIdx, 1);
                      fixed.push(ds);
                      changes.push('Zustimmung-Checkbox ans Ende verschoben');
                    }
                    const ok = await updateEvent(selectedEvent.id, { CustomFields: JSON.stringify(fixed) });
                    if (ok) {
                      setFixFieldsResult(changes.length > 0
                        ? `Geändert: ${changes.join(' | ')}`
                        : 'Keine Änderungen nötig.');
                    } else {
                      setFixFieldsResult('Update fehlgeschlagen.');
                    }
                  } catch (err) {
                    setFixFieldsResult('Fehler: ' + (err instanceof Error ? err.message : String(err)));
                  }
                  setIsFixingFields(false);
                }}
              />
            )}

            {/* 10. Profile neu laden — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                title={isRefreshingProfiles ? 'Profile werden geladen…' : 'Profile neu laden'}
                desc="Frischt JobTitle, Standort, Department und Telefonnummer der letzten N Teilnehmer aus dem Microsoft-365-Benutzerprofil auf — wenn z.B. nach einem Org-Wechsel die Teilnehmerdaten veraltet sind."
                badge="admin"
                busy={isRefreshingProfiles}
                disabled={!selectedEvent?.subsiteUrl}
                result={refreshProfilesResult}
                resultIsError={!!refreshProfilesResult && refreshProfilesResult.indexOf('Fehler') >= 0}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  const ans = prompt('Wie viele der letzten Teilnehmer sollen aus dem Benutzerprofil neu geladen werden? (JobTitle, Standort, Department, Phone)', '20');
                  if (!ans) return;
                  const n = parseInt(ans, 10);
                  if (isNaN(n) || n <= 0) { alert('Bitte eine positive Zahl eingeben.'); return; }
                  setIsRefreshingProfiles(true);
                  setRefreshProfilesResult(null);
                  try {
                    const result = await eventServiceRef.fixEventParticipantsProfileData(selectedEvent.subsiteUrl, n);
                    setRefreshProfilesResult(`${result.scanned} geprüft, ${result.updated} aktualisiert, ${result.failedLookups} Profil-Lookups fehlgeschlagen`);
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  } catch {
                    setRefreshProfilesResult('Fehler beim Auffrischen der Profile');
                  }
                  setIsRefreshingProfiles(false);
                }}
              />
            )}
          </div>
        </div>
        )}
      </div>

      {/* Zähler + QR/Check-in Aktionen.
          v9.14: Warteliste-KPI wird nur gerendert wenn Event eine Warteliste hat.
          Sonst Grid auf 4 Spalten. */}
      <div className="admin-counters" style={{ display: 'grid', gridTemplateColumns: `repeat(${(selectedEvent?.waitlistEnabled && (selectedEvent?.maxParticipants || 0) > 0) ? 5 : 4}, 1fr)`, gap: 12, marginBottom: 24 }}>
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
        {(selectedEvent?.waitlistEnabled && (selectedEvent?.maxParticipants || 0) > 0) && (
          <div className="card" style={{ padding: 16, textAlign: 'center' }}>
            <div style={{ fontSize: '1.8rem', fontWeight: 700, color: 'var(--dex-orange)' }}>
              {registrations.filter(r => r.Status === 'Warteliste').length}
            </div>
            <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.waitlist')}</div>
          </div>
        )}
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
        // v11.6: frei waehlbare Gruppen-Labels statt der hardcodeten
        // 'Durchstarter'/'Funstarter'-Begriffe — fallback auf die alten
        // Labels wenn der Organizer keine eigenen gesetzt hat.
        const labelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
        const labelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
        // v11.29: Label + Anzahl beide linksbuendig — vorher
        // justify-content: space-between schob die Anzahl an den
        // rechten Rand der Karte (bei breiten Cards optisch
        // disconnected vom Label).
        const cardA = (
          <div className="card" style={{ padding: 16, borderLeft: '3px solid var(--dex-green-dark, #6b9a1e)' }}>
            <div style={{ display: 'flex', alignItems: 'baseline', gap: 12, flexWrap: 'wrap' }}>
              <strong style={{ color: 'var(--dex-green-dark, #6b9a1e)' }}>{labelA}</strong>
              <span style={{ fontSize: '1.2rem', fontWeight: 700 }}>
                {durchActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{durchCap}</span>
              </span>
            </div>
            <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
              Warteliste: <strong style={{ color: 'var(--dex-orange)' }}>{durchWait}</strong>
            </div>
          </div>
        );
        const cardB = (
          <div className="card" style={{ padding: 16, borderLeft: '3px solid var(--dex-orange, #ff8c00)' }}>
            <div style={{ display: 'flex', alignItems: 'baseline', gap: 12, flexWrap: 'wrap' }}>
              <strong style={{ color: 'var(--dex-orange, #ff8c00)' }}>{labelB}</strong>
              <span style={{ fontSize: '1.2rem', fontWeight: 700 }}>
                {funActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{funCap}</span>
              </span>
            </div>
            <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
              Warteliste: <strong style={{ color: 'var(--dex-orange)' }}>{funWait}</strong>
            </div>
          </div>
        );
        // v11.25: gleiche Display-Reihenfolge wie auf der Registrierungs-Seite.
        const reversed = !!selectedEvent?.splitDisplayOrderReversed;
        return (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 24 }}>
            {reversed ? <>{cardB}{cardA}</> : <>{cardA}{cardB}</>}
          </div>
        );
      })()}

      {/* v9.20: Check-In starten + QR-Codes versenden sind jetzt im Aktionen-Grid
          unten als ActionTile gerendert (nicht mehr als eigene Button-Reihe).
          Damit sind alle Quick-Actions an EINEM Ort zusammengefasst. Auch fuer
          Check-In-only-User (qrScanner-Mode) — die sehen weiterhin nur den
          Check-In-Tile, da das Aktionen-Grid fuer sie unten gefiltert ist. */}
      {!isQRScannerOnlyForSelected && (<>

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
        {/* v11.28: Suchfeld direkt neben dem „Teilnehmer (N)"-Header
            statt rechtsbuendig — fluessiger Lese-Flow von links nach
            rechts, kein Sprung ueber die ganze Card-Breite mehr. */}
        <div className="mb-16" style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
          <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            <Users size={18} /> Teilnehmer ({activeRegs.length})
          </h3>
          <input
            type="text"
            className="form-input"
            placeholder="Teilnehmer suchen..."
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            style={{ maxWidth: 280, padding: '6px 12px', fontSize: '0.85rem' }}
          />
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

              const sortableCols: Record<string, 'id' | 'anrede' | 'vorname' | 'nachname' | 'email' | 'status' | 'date'> = {
                id: 'id', anrede: 'anrede', vorname: 'vorname', nachname: 'nachname', email: 'email', status: 'status', date: 'date',
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
                      {id === 'id' ? '#' : id === 'anrede' ? 'Anrede' : id === 'vorname' ? 'Vorname' : id === 'nachname' ? 'Nachname' : id === 'email' ? 'Email' : id === 'status' ? 'Status' : 'Registriert am'}
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
                if (id === 'vorname') {
                  // Fallback fuer Alt-Daten: erstes Wort aus ParticipantName.
                  const v = reg.Vorname || ((reg.ParticipantName || '').split(' ')[0] || '');
                  return <td key={id} style={{ padding: 8, fontWeight: 500 }}>{v || '-'}</td>;
                }
                if (id === 'nachname') {
                  // Fallback fuer Alt-Daten: alles ausser dem ersten Wort als Nachname.
                  let n = reg.Nachname || '';
                  if (!n && reg.ParticipantName) {
                    const parts = reg.ParticipantName.trim().split(/\s+/);
                    if (parts.length > 1) n = parts.slice(1).join(' ');
                  }
                  return <td key={id} style={{ padding: 8, fontWeight: 500 }}>{n || '-'}</td>;
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
                      <span className={`badge ${reg.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>
                        {translateStatus(reg.Status, isDe)}
                      </span>
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
                    } else if (field.type === 'select' && field.multi) {
                      // v7.11: Mehrfachauswahl wird " | "-getrennt gespeichert.
                      // In der Admin-Tabelle als Komma-Liste anzeigen, damit
                      // der Spalten-Inhalt sauberer scanbar ist.
                      display = String(val).split(' | ').map(s => s.trim()).filter(Boolean).join(', ');
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
                  const att = attachmentsByReg[reg.Id] || [];
                  return (
                    <td key={id} style={{ padding: 8, display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                      <button
                        className="btn btn-secondary"
                        style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                        title={isDe ? 'Teilnehmer-Daten bearbeiten' : 'Edit attendee data'}
                        onClick={() => openEditModal(reg)}
                      >
                        <Pencil size={12} /> {isDe ? 'Bearbeiten' : 'Edit'}
                      </button>
                      {/* v11.0: Anhang-Button — nur wenn das Event den
                          Teilnehmer-Upload erlaubt hat. Zeigt Counter wenn
                          mind. eine Datei hochgeladen wurde. */}
                      {selectedEvent?.allowAttendeeUpload && (
                        <button
                          className="btn btn-secondary"
                          style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                          title={isDe ? 'Hochgeladene Dateien anzeigen' : 'Show uploaded files'}
                          onClick={() => setAttachmentsModalReg(reg)}
                        >
                          <FileText size={12} />
                          {att.length > 0 ? `${isDe ? 'Datei' : 'File'} (${att.length})` : (isDe ? 'Datei' : 'File')}
                        </button>
                      )}
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
                          // Typ-bewusst: bei aktiver Split-Capacity (DurchstarterCapacity > 0
                          // UND FunstarterCapacity > 0) wird per Default nur ein
                          // Warteliste-Teilnehmer mit passendem PreferredStarterType
                          // nachgerueckt — es sei denn, das Event ist auf
                          // splitSharedWaitlist=true gesetzt (v10.20). Dann faellt der
                          // Filter weg und der aelteste Wartelistler rueckt nach,
                          // unabhaengig vom Typ.
                          const isSplitEvent = typeof selectedEvent.durchstarterCapacity === 'number'
                            && typeof selectedEvent.funstarterCapacity === 'number'
                            && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
                          const useTypeFilter = isSplitEvent && !selectedEvent.splitSharedWaitlist;
                          try {
                            const promoted = await eventServiceRef.promoteFirstWaitlistItem(
                              selectedEvent.subsiteUrl,
                              cancelledStarterType || undefined,
                              selectedEvent.maxParticipants,
                              (useTypeFilter && cancelledStarterType) ? cancelledStarterType : undefined
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
            // v11.6: Wartelisten-Tabellen mit den frei wählbaren Gruppen-
            // Labels statt hartcodeten 'Durchstarter'/'Funstarter'.
            const wlLabelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
            const wlLabelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
            // v11.29: Reihenfolge respektiert splitDisplayOrderReversed
            // (gleicher Toggle wie auf Register-Page + Kapazitaets-Cards).
            const wlA = renderWaitlistTable(`Warteliste ${wlLabelA}`, waitlistDurch, 'var(--dex-green-dark, #6b9a1e)');
            const wlB = renderWaitlistTable(`Warteliste ${wlLabelB}`, waitlistFun, 'var(--dex-orange, #ff8c00)');
            const reversed = !!selectedEvent?.splitDisplayOrderReversed;
            return (
              <>
                {reversed ? <>{wlB}{wlA}</> : <>{wlA}{wlB}</>}
                {renderWaitlistTable('Warteliste ohne Gruppe', waitlistUnassigned, 'var(--dex-gray-500)')}
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

      {/* ===== TEILNEHMER-EDIT MODAL (v8.0) ===== */}
      {dangerZoneModal}

      {changeLogModal}

      {/* v9.15: QR-Code-Versand-Modal — Test (nur Organizer) / Volldurchlauf
          (alle Angemeldeten) / Auto-Send-Toggle fuer zukuenftige Anmeldungen. */}
      {qrSendModalOpen && selectedEvent && (
        <div
          role="dialog"
          aria-modal="true"
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 2000,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
          onClick={() => { if (!isSendingQR) setQrSendModalOpen(false); }}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 'var(--dex-radius, 8px)',
              maxWidth: 520, width: '100%', padding: 24,
              boxShadow: 'var(--dex-shadow-hover)',
            }}
          >
            <h3 style={{ margin: '0 0 8px', fontSize: '1.05rem' }}>QR-Codes versenden</h3>
            <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              Wähle, wie der Versand laufen soll. Der QR-Code geht im Deloitte-Layout an die Empfänger und enthält unter dem Code Name + Event als Klartext (für manuellen Check-in).
            </p>
            {/* v9.22: Hinweis bei externen Teilnehmern. */}
            {(() => {
              const externalCount = registrations
                .filter(r => r.Status === 'Angemeldet')
                .filter(r => r.ParticipantEmail && !/@(.*\.)?deloitte\.de$/i.test(r.ParticipantEmail))
                .length;
              if (externalCount === 0) return null;
              return (
                <p style={{
                  margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-orange-dark, #b35a00)',
                  background: '#fff3e0', border: '1px solid #ed8b00', borderRadius: 6,
                  padding: '8px 12px', lineHeight: 1.5,
                }}>
                  <strong>Hinweis:</strong> {externalCount} {externalCount === 1 ? 'externer Teilnehmer' : 'externe Teilnehmer'} in der Liste (keine @deloitte.de-Adresse). Diese bekommen <strong>keine</strong> QR-Code-Mail — stattdessen landet der jeweilige QR-Code bei dir als Organizer mit klar markiertem Subject. Drucke ihn aus oder leite ihn unter Beachtung der Deloitte-Datenschutzrichtlinien intern weiter.
                </p>
              );
            })()}
            {/* v9.19: Hinweis aufs Handbuch — User klickt auf den Link und
                landet direkt in der Check-In-Sektion. */}
            <p style={{ margin: '0 0 16px', fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
              Tipp: Wie der Check-In am Eventtag genau abläuft — vom QR-Code-Scan über manuelle Eincheck-Vorgänge bis zu Sonderfällen — steht im{' '}
              <a
                href="javascript:void(0)"
                onClick={(e) => { e.preventDefault(); navigate('manual'); window.location.hash = 'check-in'; }}
                style={{ color: 'var(--dex-green-dark)', fontWeight: 500 }}
              >
                Handbuch-Artikel &bdquo;Check-In am Event-Tag&ldquo;
              </a>.
            </p>

            {/* v9.29: Hinweis falls Organizer selbst NICHT für das Event angemeldet ist —
                "Nur Test (an mich)" verschickt zwar die QR-Mail, aber der Check-In-Scan
                wird die Person nicht in der Teilnehmerliste finden. */}
            {(() => {
              const orgEmail = (currentUser.email || '').toLowerCase();
              const isOrgRegistered = !!orgEmail && registrations.some(r =>
                (r.ParticipantEmail || '').toLowerCase() === orgEmail
                && (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
              );
              if (isOrgRegistered) return null;
              return (
                <div style={{
                  display: 'flex', alignItems: 'flex-start', gap: 10,
                  padding: '10px 12px', marginBottom: 16,
                  background: '#fff8e1', border: '1px solid #f5b400',
                  borderRadius: 8, fontSize: '0.82rem', lineHeight: 1.5,
                  color: '#7a5a00',
                }}>
                  <span style={{ fontSize: '1.05rem', flexShrink: 0 }}>⚠</span>
                  <div>
                    <strong>Du bist selbst nicht für dieses Event angemeldet.</strong> Beim Klick auf <strong>{'„Nur Test (an mich)“'}</strong> bekommst du zwar die QR-Code-Mail, aber das anschließende Check-In wird <strong>nicht funktionieren</strong> — beim Scan wird die Teilnehmerliste durchsucht, dort fehlst du. Wenn du den kompletten Flow inklusive Check-In testen willst, melde dich vorher selbst zum Event an.
                  </div>
                </div>
              );
            })()}

            {/* Auto-Send-Toggle */}
            <label style={{
              display: 'flex', alignItems: 'flex-start', gap: 10, padding: '12px',
              border: '1px solid var(--dex-gray-200)', borderRadius: 8, marginBottom: 16,
              background: qrAutoSendToggle ? 'var(--dex-green-light, #f0f8e8)' : '#fff',
              cursor: 'pointer',
            }}>
              <input
                type="checkbox"
                checked={qrAutoSendToggle}
                onChange={e => setQrAutoSendToggle(e.target.checked)}
                disabled={isSendingQR}
                style={{ marginTop: 3 }}
              />
              <div style={{ fontSize: '0.85rem' }}>
                <div style={{ fontWeight: 600, marginBottom: 4 }}>Automatisch bei Anmeldung versenden</div>
                <div style={{ color: 'var(--dex-gray-600)' }}>
                  Wenn aktiv, bekommt jeder neue Teilnehmer direkt nach erfolgreicher Anmeldung seinen QR-Code per Mail. Diese Einstellung wird am Event gespeichert und wirkt für ALLE zukünftigen Anmeldungen.
                </div>
              </div>
            </label>

            {qrSendResult && (
              <div style={{
                padding: '8px 12px', marginBottom: 12, borderRadius: 6,
                background: qrSendResult.startsWith('Fehler') ? 'var(--dex-red-light, #ffe5e5)' : 'var(--dex-green-light, #f0f8e8)',
                fontSize: '0.85rem', color: qrSendResult.startsWith('Fehler') ? 'var(--dex-red-dark, #b00)' : 'var(--dex-green-dark)',
              }}>{qrSendResult}</div>
            )}

            <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', flexWrap: 'wrap' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setQrSendModalOpen(false)}
                disabled={isSendingQR}
                style={{ fontSize: '0.85rem' }}
              >
                Abbrechen
              </button>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <button
                  className="btn btn-outline"
                  disabled={isSendingQR || qrPreviewLoading}
                  onClick={async () => {
                    if (!selectedEvent) return;
                    setQrPreviewLoading(true);
                    try {
                      const orgEmail = currentUser.email;
                      const orgFullName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || orgEmail;
                      const orgFirstName = currentUser.firstName || orgFullName.split(/\s+/)[0] || orgFullName;
                      const qrData = `DEX|${selectedEvent.eventNumber}|${orgEmail}`;
                      let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
                      try {
                        const qrDataUrl = await QRCode.toDataURL(qrData, { width: 300, margin: 2 });
                        qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
                      } catch { /* */ }
                      const emailData = qrCodeEmail(orgFirstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', orgFullName);
                      setQrPreviewSubject(emailData.subject);
                      setQrPreviewHtml(emailData.body);
                      setQrPreviewOpen(true);
                    } finally { setQrPreviewLoading(false); }
                  }}
                  style={{ fontSize: '0.85rem' }}
                  title="So sieht die Mail aus, die beim Versand rausgeht — inklusive echtem QR-Code für dich als Empfänger."
                >
                  {qrPreviewLoading ? 'Lade Vorschau…' : '👁 Vorschau Mail'}
                </button>
                <button
                  className="btn btn-secondary"
                  onClick={async () => {
                    if (!eventServiceRef || !selectedEvent) return;
                    // Toggle persistieren (einmal speichern reicht — auch wenn user nur testet)
                    if (qrAutoSendToggle !== !!selectedEvent.autoSendQRCode) {
                      try { await eventServiceRef.updateEvent(parseInt(selectedEvent.id, 10), { AutoSendQRCode: qrAutoSendToggle }); } catch { /* */ }
                    }
                    // Test-Modus: an Organizer-Mail des aktuellen Users versenden
                    setIsSendingQR(true);
                    setQrSendResult(null);
                    setQrSentCount(0);
                    try {
                      const orgEmail = currentUser.email;
                      const orgFullName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || orgEmail;
                      const orgFirstName = currentUser.firstName || orgFullName.split(/\s+/)[0] || orgFullName;
                      const qrData = `DEX|${selectedEvent.eventNumber}|${orgEmail}`;
                      let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
                      try {
                        const qrDataUrl = await QRCode.toDataURL(qrData, { width: 300, margin: 2 });
                        qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
                      } catch { /* */ }
                      const emailData = qrCodeEmail(orgFirstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', orgFullName);
                      await eventServiceRef.queueEmail(
                        emailData.subject, orgEmail, orgFullName, emailData.body,
                        'QRCode', selectedEvent.title, selectedEvent.id
                      );
                      setQrSendResult(`Test-Mail an ${orgEmail} verschickt — bitte in deinem Postfach prüfen.`);
                    } catch (err) {
                      setQrSendResult('Fehler beim Test-Versand: ' + (err instanceof Error ? err.message : String(err)));
                    }
                    setIsSendingQR(false);
                  }}
                  disabled={isSendingQR}
                  style={{ fontSize: '0.85rem' }}
                >
                  Nur Test (an mich)
                </button>
                <button
                  className="btn btn-primary"
                  onClick={async () => {
                    if (!eventServiceRef || !selectedEvent) return;
                    const eligible = registrations.filter(r => r.Status === 'Angemeldet');
                    if (eligible.length === 0) {
                      setQrSendResult('Fehler: Keine angemeldeten Teilnehmer.');
                      return;
                    }
                    if (!window.confirm(`QR-Codes an ${eligible.length} angemeldete Teilnehmer versenden?`)) return;

                    // Auto-Send-Toggle persistieren
                    if (qrAutoSendToggle !== !!selectedEvent.autoSendQRCode) {
                      try { await eventServiceRef.updateEvent(parseInt(selectedEvent.id, 10), { AutoSendQRCode: qrAutoSendToggle }); } catch { /* */ }
                    }

                    setIsSendingQR(true);
                    setQrSendResult(null);
                    setQrSentCount(0);
                    let sent = 0;
                    let extCount = 0;
                    for (const reg of eligible) {
                      const qrData = `DEX|${selectedEvent.eventNumber}|${reg.ParticipantEmail}`;
                      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                      const firstName = reg.Vorname || (reg.ParticipantName || '').trim().split(/\s+/)[0] || name;
                      let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
                      try {
                        const qrDataUrl = await QRCode.toDataURL(qrData, { width: 300, margin: 2 });
                        qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
                      } catch { /* */ }
                      const emailData = qrCodeEmail(firstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', name);
                      // v9.22: Externe Mail-Adresse → QR an Organizer umleiten.
                      const isExternal = !!reg.ParticipantEmail && !/@(.*\.)?deloitte\.de$/i.test(reg.ParticipantEmail);
                      if (isExternal) {
                        const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
                        const orgRecipient = orgEmails.length > 0 ? orgEmails.join(';') : currentUser.email;
                        const orgSubject = `[Externer Teilnehmer] QR-Code für ${name} — ${selectedEvent.title}`;
                        const qrExternalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
                          + `<strong>QR-Code für externen Teilnehmer.</strong><br>`
                          + `Eigentlich für <strong>${reg.ParticipantEmail}</strong> (${name}). Da externe Adressen keinen Mail-Versand bekommen, landet der QR-Code bei dir — drucke ihn aus oder leite die Mail intern an den Empfänger weiter (Datenschutzrichtlinien Deloitte Deutschland beachten).`
                          + `</div>`;
                        const qrBody = emailData.body.replace(/<body([^>]*)>/i, `<body$1>${qrExternalHint}`);
                        await eventServiceRef.queueEmail(
                          orgSubject, orgRecipient, 'Organizer', qrBody,
                          'QRCode', selectedEvent.title, selectedEvent.id
                        );
                        extCount++;
                      } else {
                        await eventServiceRef.queueEmail(
                          emailData.subject, reg.ParticipantEmail, name, emailData.body,
                          'QRCode', selectedEvent.title, selectedEvent.id
                        );
                      }
                      if (selectedEvent.subsiteUrl) {
                        await eventServiceRef.setQRSentStatus(selectedEvent.subsiteUrl, reg.Id);
                      }
                      sent++;
                      setQrSentCount(sent);
                    }
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                    setIsSendingQR(false);
                    setQrSendResult(extCount > 0
                      ? `${sent} QR-Codes verschickt (davon ${extCount} an dich/Organizer umgeleitet — externe Adressen).`
                      : `${sent} QR-Codes verschickt.`);
                  }}
                  disabled={isSendingQR}
                  style={{ fontSize: '0.85rem' }}
                >
                  {isSendingQR ? `Versende... (${qrSentCount})` : 'An alle Angemeldeten'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {editingReg && selectedEvent && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => { if (!isSavingEdit) closeEditModal(); }}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 920, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                <Pencil size={18} />{' '}
                {isDe ? 'Teilnehmer bearbeiten' : 'Edit attendee'}
                {' — '}
                <span style={{ color: 'var(--dex-green-dark)' }}>
                  {editForm.Vorname} {editForm.Nachname}
                </span>
              </h3>
              <button
                onClick={closeEditModal}
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}
                aria-label={isDe ? 'Schließen' : 'Close'}
                disabled={isSavingEdit}
              ><X size={20} /></button>
            </div>
            <p style={{ margin: '0 0 16px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              {isDe
                ? 'Hier kannst du Vorname, Nachname und E-Mail-Adresse korrigieren (z.B. nach einem Tippfehler bei der manuellen Anlage) sowie die event-spezifischen Felder anpassen. Beim Ändern der E-Mail wird geprüft, ob die Adresse zum Deloitte-Tenant gehört und die Person dort existiert — externe Adressen sind nicht erlaubt. Phone, Department, Standort und Job Title kommen aus dem M365-Profil und sind read-only — sie werden bei einem Mail-Wechsel automatisch nachgezogen. Den Status änderst du über die Aktions-Buttons in der Liste. Jede Änderung wird im Audit-Log und im ChangeLog des Teilnehmers mit Datum und deinem Namen protokolliert.'
                : 'You can fix first name, last name and email address here (e.g. after a typo during manual creation) and adjust event-specific fields. When changing the email, the app verifies that the address belongs to the Deloitte tenant and that the person exists there — external addresses are not allowed. Phone, Department, Location and Job Title come from the M365 profile and are read-only — they are refreshed automatically when the email changes. The status is changed via the action buttons in the list. Every change is logged in the audit log and in the attendee\'s ChangeLog with date and your name.'}
            </p>

            {(() => {
              // Vorname / Nachname / E-Mail sind seit v9.7 editierbar (mit
              // Deloitte-Domain- und Tenant-Existenz-Check beim Speichern).
              // Die uebrigen Profil-Felder bleiben read-only — sie kommen
              // aus dem M365-Profil und werden bei einer Mail-Aenderung
              // mit den Profil-Daten der neuen Person ueberschrieben.
              const editableStammFields: Array<{ key: string; label: string; type?: string }> = [
                { key: 'Vorname', label: isDe ? 'Vorname' : 'First name' },
                { key: 'Nachname', label: isDe ? 'Nachname' : 'Last name' },
                { key: 'ParticipantEmail', label: 'E-Mail', type: 'email' },
              ];
              const readOnlyFields: Array<{ key: string; label: string }> = [
                { key: 'Anrede', label: isDe ? 'Anrede' : 'Salutation' },
                { key: 'Phone', label: isDe ? 'Telefon' : 'Phone' },
                { key: 'Department', label: 'Department' },
                { key: 'Location', label: isDe ? 'Standort' : 'Location' },
                { key: 'JobTitle', label: 'Job Title' },
                { key: 'Status', label: 'Status' },
              ];
              const renderReadOnly = (label: string, value: string): React.ReactNode => (
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                    {label}
                  </label>
                  <div style={{
                    width: '100%', padding: '8px 12px',
                    background: 'var(--dex-gray-50, #fafafa)',
                    border: '1px solid var(--dex-gray-200, #e5e7eb)',
                    borderRadius: 6, fontSize: '0.88rem',
                    color: value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                    minHeight: 38, lineHeight: 1.5,
                  }}>
                    {value || (isDe ? '— nicht gesetzt —' : '— not set —')}
                  </div>
                </div>
              );
              const renderEditable = (key: string, label: string, type?: string): React.ReactNode => (
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                    {label}
                  </label>
                  <input
                    className="form-input"
                    type={type || 'text'}
                    value={editForm[key] || ''}
                    onChange={e => setEditForm(prev => ({ ...prev, [key]: e.target.value }))}
                    style={{ fontSize: '0.88rem' }}
                  />
                </div>
              );
              return (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                  {/* Anrede zuerst (read-only) */}
                  <div>{renderReadOnly(isDe ? 'Anrede' : 'Salutation', editForm.Anrede || '')}</div>
                  {/* Vorname + Nachname editierbar */}
                  {editableStammFields.filter(f => f.key !== 'ParticipantEmail').map(f => (
                    <div key={f.key}>
                      {renderEditable(f.key, f.label, f.type)}
                    </div>
                  ))}
                  {/* E-Mail editierbar (volle Breite) */}
                  <div style={{ gridColumn: '1 / -1' }}>
                    {renderEditable('ParticipantEmail', 'E-Mail', 'email')}
                    <p style={{ margin: '4px 0 0', fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                      {isDe
                        ? 'Nur Deloitte-Adressen (@deloitte.de / @deloitte.com). Beim Speichern wird die Person im Tenant verifiziert.'
                        : 'Only Deloitte addresses (@deloitte.de / @deloitte.com). The person is verified in the tenant on save.'}
                    </p>
                  </div>
                  {/* Restliche Profil-Felder read-only */}
                  {readOnlyFields.filter(f => f.key !== 'Anrede').map(f => (
                    <div key={f.key}>
                      {renderReadOnly(f.label, editForm[f.key] || '')}
                    </div>
                  ))}

                  {/* B2Run-Starter-Typ (Funstarter / Durchstarter). Hardcoded
                      SP-Spalte auf der Teilnehmerliste (kein regulärer
                      Custom-Field-Eintrag), daher explizit hier gerendert.
                      Updates BEIDE intern getrackten Felder zugleich
                      (StarterType + PreferredStarterType) — die getrennte
                      Speicherung von „aktuell vs. Wunsch" ist Implementierungs-
                      Detail für die Warteliste-Nachrück-Logik und braucht im
                      Edit-Modal keine UI-Komplexität. v10.15+ */}
                  {selectedEvent.durchstarterCapacity !== undefined
                    && selectedEvent.funstarterCapacity !== undefined
                    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0) && (
                    <div style={{ gridColumn: '1 / -1', marginTop: 12, paddingTop: 16, borderTop: '1px solid var(--dex-gray-200)' }}>
                      <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                        {isDe ? 'B2Run-Starter-Typ' : 'B2Run starter type'}
                      </label>
                      <select
                        value={editForm.StarterType || ''}
                        onChange={e => {
                          const v = e.target.value;
                          // Beide Felder synchron halten — der Aktuelle wechselt
                          // mit, der Wunsch ebenfalls (User-Erwartung: „ich ändere
                          // den Starter-Typ" = beides ändert sich).
                          setEditForm(prev => ({ ...prev, StarterType: v, PreferredStarterType: v }));
                        }}
                        className="form-input"
                        style={{ maxWidth: 320 }}
                      >
                        <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
                        <option value="Durchstarter">Durchstarter</option>
                        <option value="Funstarter">Funstarter</option>
                      </select>
                    </div>
                  )}

                  {/* Custom Fields des Events — DAS ist der editierbare Teil.
                      Renderer abhaengig vom Field-Type (text/number/select/
                      checkbox). Multi-Select speichert Werte als " | "-
                      getrennten String, identisch zum Registrierungs-Pfad. */}
                  {selectedEvent.eventSpecificFields && selectedEvent.eventSpecificFields.length > 0 && (
                    <div style={{ gridColumn: '1 / -1', marginTop: 12, paddingTop: 16, borderTop: '1px solid var(--dex-gray-200)' }}>
                      <h4 style={{ margin: '0 0 4px', fontSize: '0.92rem', color: 'var(--dex-gray-800)' }}>
                        {isDe ? 'Event-spezifische Felder (editierbar)' : 'Event-specific fields (editable)'}
                      </h4>
                      <p style={{ margin: '0 0 12px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                        {isDe
                          ? 'Nur diese Felder werden gespeichert.'
                          : 'Only these fields will be saved.'}
                      </p>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                        {selectedEvent.eventSpecificFields.map(cf => {
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const sp = (cf as any).spInternalName || '';
                          if (!sp) return null;
                          const value = editForm[sp] || '';
                          const setVal = (v: string): void => setEditForm(prev => ({ ...prev, [sp]: v }));
                          const labelEl = (
                            <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                              {cf.label}{cf.required && <span style={{ color: 'var(--dex-red, #c00)' }}> *</span>}
                            </label>
                          );

                          // Single-Select-Dropdown
                          if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <select
                                  className="form-select"
                                  value={value}
                                  onChange={e => setVal(e.target.value)}
                                  style={{ width: '100%' }}
                                >
                                  <option value="">{isDe ? '— bitte wählen —' : '— please choose —'}</option>
                                  {cf.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                                </select>
                              </div>
                            );
                          }

                          // Multi-Select (Checkbox-Liste, Werte mit ' | ' joinen)
                          if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                            const selectedSet = new Set(value.split(' | ').map(s => s.trim()).filter(Boolean));
                            const toggle = (opt: string): void => {
                              if (selectedSet.has(opt)) selectedSet.delete(opt);
                              else selectedSet.add(opt);
                              setVal(Array.from(selectedSet).join(' | '));
                            };
                            return (
                              <div key={cf.id} style={{ gridColumn: '1 / -1' }}>
                                {labelEl}
                                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, padding: 8, border: '1px solid var(--dex-gray-200)', borderRadius: 6, background: '#fff' }}>
                                  {cf.options.map(opt => (
                                    <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', cursor: 'pointer', padding: '4px 8px', background: selectedSet.has(opt) ? 'rgba(134,188,37,0.12)' : 'var(--dex-gray-50)', borderRadius: 6 }}>
                                      <input
                                        type="checkbox"
                                        checked={selectedSet.has(opt)}
                                        onChange={() => toggle(opt)}
                                        style={{ accentColor: 'var(--dex-green)' }}
                                      />
                                      {opt}
                                    </label>
                                  ))}
                                </div>
                              </div>
                            );
                          }

                          // Checkbox (true/false)
                          if (cf.type === 'checkbox') {
                            const isChecked = value === 'true' || value === '1';
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '6px 10px', background: isChecked ? 'rgba(134,188,37,0.12)' : 'var(--dex-gray-50)', borderRadius: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                                  <input
                                    type="checkbox"
                                    checked={isChecked}
                                    onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                                    style={{ accentColor: 'var(--dex-green)' }}
                                  />
                                  {isChecked ? (isDe ? 'Ja' : 'Yes') : (isDe ? 'Nein' : 'No')}
                                </label>
                              </div>
                            );
                          }

                          // Number
                          if (cf.type === 'number') {
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <input
                                  className="form-input"
                                  type="number"
                                  value={value}
                                  onChange={e => setVal(e.target.value)}
                                  style={{ width: '100%' }}
                                />
                              </div>
                            );
                          }

                          // Default: text-Input (auch fuer 'text', 'user', 'roommate')
                          return (
                            <div key={cf.id}>
                              {labelEl}
                              <input
                                className="form-input"
                                value={value}
                                onChange={e => setVal(e.target.value)}
                                style={{ width: '100%' }}
                              />
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>
              );
            })()}

            {editError && (
              <div style={{ marginTop: 12, padding: '8px 12px', background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', borderRadius: 6, fontSize: '0.85rem', color: 'var(--dex-red, #c00)' }}>
                {editError}
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20 }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={closeEditModal}
                disabled={isSavingEdit}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={saveEdit}
                disabled={isSavingEdit}
                style={{ opacity: isSavingEdit ? 0.6 : 1 }}
              >
                {isSavingEdit ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* v9.37: Vorschau-Modal für die QR-Code-Mail. Rendert das wirklich
          versendete Mail-HTML in einem sandboxed iframe — analog zur Live-
          Preview im Event-Wizard unter Kommunikation. Editieren ist hier
          NICHT vorgesehen, der Body wird zentral aus der QR-Code-Vorlage
          gebaut. */}
      {qrPreviewOpen && (
        <div
          role="dialog"
          aria-modal="true"
          onClick={() => setQrPreviewOpen(false)}
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 2100,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 12,
              width: '100%', maxWidth: 720, maxHeight: '90vh',
              display: 'flex', flexDirection: 'column',
              boxShadow: '0 16px 48px rgba(0,0,0,0.28)',
            }}
          >
            <div style={{
              padding: '14px 18px', borderBottom: '1px solid var(--dex-gray-200)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
            }}>
              <div style={{ minWidth: 0, flex: 1 }}>
                <h3 style={{ margin: 0, fontSize: '1.05rem' }}>Vorschau: QR-Code-Mail</h3>
                <p style={{ margin: '2px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  So sieht die Mail aus, die jeder angemeldete Teilnehmer bekommt — der QR-Code in der Vorschau ist auf dich ausgestellt.
                </p>
                <p style={{ margin: '6px 0 0', fontSize: '0.8rem', color: 'var(--dex-gray-700)' }}>
                  <strong>Betreff:</strong> <span style={{ color: 'var(--dex-gray-600)' }}>{qrPreviewSubject}</span>
                </p>
              </div>
              <button
                type="button"
                onClick={() => setQrPreviewOpen(false)}
                style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-500)', padding: 4 }}
                aria-label="Schließen"
              >
                <X size={22} />
              </button>
            </div>
            <div style={{ flex: 1, overflow: 'hidden', background: '#f5f5f5', padding: 12 }}>
              <iframe
                title="QR-Code-Mail-Vorschau"
                srcDoc={qrPreviewHtml}
                sandbox=""
                style={{ width: '100%', height: '100%', minHeight: 480, border: 'none', borderRadius: 6, background: '#fff' }}
              />
            </div>
            <div style={{ padding: '12px 18px', borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'flex-end' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setQrPreviewOpen(false)}
                style={{ fontSize: '0.85rem' }}
              >
                Schließen
              </button>
            </div>
          </div>
        </div>
      )}

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

      {/* v11.0: Modal für Teilnehmer-Attachments. Liste der hochgeladenen
          Dateien mit Download-Link + Lösch-Button (Admin/Organizer kann
          fremde Uploads löschen). Plus optional eigener Upload-Button für
          den Admin (z.B. Bestätigungsbescheinigung im Namen des
          Teilnehmers anhängen). */}
      {attachmentsModalReg && (() => {
        const reg = attachmentsModalReg;
        const list = attachmentsByReg[reg.Id] || [];
        const close = (): void => setAttachmentsModalReg(null);
        const refreshOne = async (regId: number): Promise<void> => {
          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
          try {
            const fresh = await eventServiceRef.listRegistrationAttachments(selectedEvent.subsiteUrl, regId);
            setAttachmentsByReg(prev => ({ ...prev, [regId]: fresh }));
          } catch { /* */ }
        };
        const onDelete = async (fileName: string): Promise<void> => {
          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
          if (!window.confirm(isDe ? `Datei „${fileName}" wirklich löschen?` : `Really delete file „${fileName}"?`)) return;
          setAttachmentsBusy(true);
          try {
            await eventServiceRef.deleteRegistrationAttachment(selectedEvent.subsiteUrl, reg.Id, fileName);
            await refreshOne(reg.Id);
          } finally { setAttachmentsBusy(false); }
        };
        const onAdd = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
          const f = e.target.files && e.target.files[0];
          e.target.value = '';
          if (!f || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
          if (f.size > 10 * 1024 * 1024) {
            window.alert(isDe ? 'Datei ist größer als 10 MB.' : 'File is larger than 10 MB.');
            return;
          }
          setAttachmentsBusy(true);
          try {
            await eventServiceRef.addRegistrationAttachment(selectedEvent.subsiteUrl, reg.Id, f);
            await refreshOne(reg.Id);
          } finally { setAttachmentsBusy(false); }
        };
        const fullName = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantEmail || '–';
        return (
          <div
            role="dialog"
            aria-modal="true"
            style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 2000,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
            }}
            onClick={() => { if (!attachmentsBusy) close(); }}
          >
            <div
              onClick={e => e.stopPropagation()}
              className="card"
              style={{ width: '100%', maxWidth: 560, maxHeight: '85vh', overflow: 'auto', padding: 24, background: '#fff', borderRadius: 8 }}
            >
              <h3 style={{ margin: '0 0 4px', fontSize: '1.1rem' }}>
                {isDe ? 'Hochgeladene Dateien' : 'Uploaded files'}
              </h3>
              <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                {fullName}{reg.ParticipantEmail ? ` · ${reg.ParticipantEmail}` : ''}
              </p>
              {list.length === 0 ? (
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', fontStyle: 'italic', margin: '12px 0' }}>
                  {isDe ? 'Noch keine Dateien hochgeladen.' : 'No files uploaded yet.'}
                </p>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 16 }}>
                  {list.map(f => (
                    <div key={f.fileName} style={{
                      display: 'flex', alignItems: 'center', gap: 10,
                      padding: '8px 12px', borderRadius: 6,
                      background: 'rgba(134,188,37,0.08)',
                      border: '1px solid rgba(134,188,37,0.30)',
                      fontSize: '0.85rem',
                    }}>
                      <FileText size={16} />
                      <a
                        href={f.serverRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ flex: 1, color: 'var(--dex-gray-800)', textDecoration: 'none', wordBreak: 'break-all' }}
                      >
                        {f.fileName}
                      </a>
                      <a
                        href={f.serverRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.72rem', padding: '2px 10px', textDecoration: 'none' }}
                      >
                        <Download size={12} /> {isDe ? 'Download' : 'Download'}
                      </a>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.72rem', padding: '2px 10px', color: 'var(--dex-red, #c00)' }}
                        disabled={attachmentsBusy}
                        onClick={() => onDelete(f.fileName)}
                        title={isDe ? 'Löschen' : 'Delete'}
                      >
                        <X size={12} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
              <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
                <label className="btn btn-outline" style={{ fontSize: '0.82rem', padding: '6px 14px', cursor: attachmentsBusy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <Plus size={14} /> {attachmentsBusy ? (isDe ? 'Wird übertragen…' : 'Uploading…') : (isDe ? 'Datei hinzufügen' : 'Add file')}
                  <input
                    type="file"
                    accept="application/pdf,image/*,.doc,.docx"
                    style={{ display: 'none' }}
                    onChange={onAdd}
                    disabled={attachmentsBusy}
                  />
                </label>
                <button className="btn btn-primary" onClick={close} disabled={attachmentsBusy}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}
