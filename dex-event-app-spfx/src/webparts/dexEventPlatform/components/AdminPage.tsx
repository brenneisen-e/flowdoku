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
import { Plus, Users, FileText, Trash2, Copy, Mail, Send, Download, Pencil, ExternalLink, AlertCircle, Hash, Columns, Wrench, RefreshCw, X, Check, Link2, ChevronUp, ChevronDown } from './Icons';
import * as XLSX from 'xlsx';
import { EventService } from '../services/EventService';
import { qrCodeEmail, cancellationEmail, promotionEmail, wrapTemplate, replacePlaceholders, buildEmailFromTemplate, getCachedLogoBase64, getCachedOrbBase64 } from '../services/EmailTemplates';
import { applyEventTemplateOverride, formatOrganizerList } from '../context/EventContext';
import { HtmlEditorModal } from './HtmlEditorModal';
import { InfoTooltip } from './InfoTooltip';
import { MultiSelectDropdown } from './MultiSelectDropdown';
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

// v11.41: Einladungsmail-Empfaenger-Blocker. Die Einladungsmail darf NIE an
// komplette Standort-Verteiler ('de.duesseldorf@...', 'duesseldorf@...' etc.)
// oder an pauschale 'all'-Listen ('deall@...', 'all@...', 'alldeloitte@...')
// gehen. Hintergrund: solche Aussendungen sind ohne CMC-/Marketing-Freigabe
// nicht erlaubt — kleinere explizite Verteilergruppen (Team-Mailboxen,
// Funktions-Accounts) bleiben aber zulaessig.
const DEX_LOCATION_TOKENS: string[] = [
  'berlin', 'dresden', 'duesseldorf', 'dusseldorf', 'düsseldorf',
  'frankfurt', 'goerlitz', 'görlitz', 'halle', 'hamburg', 'hannover',
  'koeln', 'köln', 'cologne', 'leipzig', 'magdeburg', 'mannheim',
  'muenchen', 'münchen', 'munich', 'nuernberg', 'nürnberg', 'nuremberg',
  'stuttgart', 'walldorf',
];

/** Wenn die Adresse als unerlaubter Massen-Verteiler erkannt wird, gibt den
 *  Block-Grund zurueck — sonst null. Heuristik bewusst konservativ: matched
 *  nur, wenn der Local-Part (bzw. der gesamte Token, falls kein '@' vorhanden)
 *  eindeutig ein Standort-/All-Verteiler ist. Team-Mailboxen wie
 *  'frankfurt-event-team@' bleiben erlaubt.
 *
 *  v11.44: Auch reine Tokens ohne '@' werden geprueft — der Mailverteiler
 *  kann Eintraege wie 'All' oder 'Duesseldorf' enthalten, die direkt aus dem
 *  Standort-/Location-Picker stammen. Vorher wurden die durchgelassen, weil
 *  der Parser an `at <= 0` zurueckkehrte. */
function getBlockedInviteReason(email: string): string | null {
  const lc = (email || '').trim().toLowerCase();
  if (!lc) return null;
  const at = lc.indexOf('@');
  // Mit '@': Local-Part vor dem '@' pruefen. Ohne '@': gesamten Token pruefen
  // (z.B. wenn die Sichtbarkeit per Location-Picker auf 'All' gesetzt war
  // und 'All' so im audienceFilter landet).
  const local = at > 0 ? lc.slice(0, at) : lc;
  // Token-Split: Local-Part nach .-_ tokenisieren.
  const tokens = local.split(/[._-]/).filter(Boolean);
  // (1) deall / de.all / de-all / alldeloitte etc. — globaler DE-Verteiler.
  if (local === 'deall' || local === 'alldeloitte' || tokens.includes('deall') || tokens.includes('alldeloitte')) {
    return 'globaler Deloitte-DE-Verteiler';
  }
  // (2) 'all' als eigenstaendiger Token oder Local-Part — pauschale Liste.
  if (local === 'all' || tokens.includes('all')) {
    return 'globaler "all"-Verteiler';
  }
  // (3) Standort-Verteiler: Local-Part ist exakt eine Stadt ODER beginnt /
  //     endet mit 'de.' / 'de-' und enthaelt eine Stadt als Token.
  for (const loc of DEX_LOCATION_TOKENS) {
    if (local === loc) return `Standort-Verteiler (${loc})`;
    // 'de.<loc>' / 'de-<loc>' / '<loc>.de' / '<loc>-de'
    if (tokens.length === 2 && tokens.includes('de') && tokens.includes(loc)) {
      return `Standort-Verteiler (${loc})`;
    }
  }
  return null;
}

/** Liefert pro Empfaenger die Block-Begruendung — leeres Array = alles OK. */
function getBlockedInviteRecipients(emails: string[]): Array<{ email: string; reason: string }> {
  const out: Array<{ email: string; reason: string }> = [];
  for (const e of emails) {
    const reason = getBlockedInviteReason(e);
    if (reason) out.push({ email: e, reason });
  }
  return out;
}

// v9.20: EventStatus-Labels lokalisieren (DE).
// v11.89: 'Under Construction' wird transparent als 'Entwurf' angezeigt,
// solange noch Legacy-Daten existieren — neue Events nutzen IsFictive.
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

// v11.98: Pill-Toggle für die Aktiv-Teilnehmer-Tabelle bei Split-Kapazität.
// Default 'split' = getrennte Tabellen pro Gruppe. 'merged' = einzelne
// Tabelle (alter Look).
function SplitMergeToggle(props: {
  view: 'split' | 'merged';
  setView: (v: 'split' | 'merged') => void;
  isDe: boolean;
}): React.ReactElement {
  const pill = (active: boolean): React.CSSProperties => ({
    padding: '5px 12px',
    borderRadius: 999,
    fontSize: '0.78rem',
    fontWeight: 600,
    cursor: 'pointer',
    border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
    background: active ? 'rgba(134,188,37,0.10)' : '#fff',
    color: active ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
    transition: 'all 0.12s ease',
  });
  return (
    <div style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
      <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginRight: 6 }}>
        {props.isDe ? 'Ansicht:' : 'View:'}
      </span>
      <button type="button" onClick={() => props.setView('split')} style={pill(props.view === 'split')}>
        {props.isDe ? 'Getrennt' : 'Split'}
      </button>
      <button type="button" onClick={() => props.setView('merged')} style={pill(props.view === 'merged')}>
        {props.isDe ? 'Zusammen' : 'Merged'}
      </button>
    </div>
  );
}

export default function AdminPage(): React.ReactElement {
  const { navigate, selectedEventId } = useNavigation();
  const { topLevelEvents: events, childEventsOf, isEventsLoading, getAllRegistrations, deleteEvent, updateEvent, refreshEvents, addTeamMember, transferTeamLead } = useEvents();
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
  // v11.97/v11.98: bei Events mit Split-Kapazität (zwei Gruppen) wird die
  // Aktiv-Teilnehmer-Tabelle standardmäßig nach Gruppe getrennt angezeigt
  // (kleinere Gruppe zuerst). Per Toggle umschaltbar auf zusammengeführte
  // Sicht. Default: 'split'. Bei Events ohne Split-Kapazität ohne Wirkung.
  const [splitParticipantsView, setSplitParticipantsView] = React.useState<'split' | 'merged'>('split');
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
  // v11.36: Überbuchungs-Bereinigung
  const [isDetectingOverbook, setIsDetectingOverbook] = React.useState(false);
  const [detectOverbookResult, setDetectOverbookResult] = React.useState<string | null>(null);
  // Modal-State für „Bestätigen" (einzeln oder Sammel) bzw. „Platz behalten".
  // mode: 'confirm' = auf Warteliste; 'keep' = Platz behalten.
  // targets: betroffene Registrierungen (1 = einzeln, n = Sammel „Alle").
  const [overbookModal, setOverbookModal] = React.useState<{
    mode: 'confirm' | 'keep';
    targets: SPRegistration[];
  } | null>(null);
  const [obWithMail, setObWithMail] = React.useState(true);
  const [obMailSubject, setObMailSubject] = React.useState('');
  const [obMailBody, setObMailBody] = React.useState('');
  const [obMailLang, setObMailLang] = React.useState<'DE' | 'EN'>('DE');
  const [obRemoveCalendar, setObRemoveCalendar] = React.useState(true);
  const [obKeepVariant, setObKeepVariant] = React.useState<'active' | 'firstWaitlist'>('firstWaitlist');
  const [obBusy, setObBusy] = React.useState(false);
  // v11.36: Fortschritts-Overlay für die ID-Neuvergabe (0..100 %, null = aus).
  const [reorderProgress, setReorderProgress] = React.useState<number | null>(null);
  const [reorderProgressLabel, setReorderProgressLabel] = React.useState('');
  const [isFixingFields, setIsFixingFields] = React.useState(false);
  const [fixFieldsResult, setFixFieldsResult] = React.useState<string | null>(null);
  // v11.84: Teams-Section (Admin Center Team-Management).
  const [teamsCollapsed, setTeamsCollapsed] = React.useState<boolean>(false);
  // Add-Member-Modal pro Team — gleiche Mechanik wie MyEventsPage.
  const [adminAddMemberDialog, setAdminAddMemberDialog] = React.useState<{
    teamId: string;
    teamName: string;
    freeSlots: number;
  } | null>(null);
  const [adminAddMemberPick, setAdminAddMemberPick] = React.useState<{ email: string; displayName: string } | null>(null);
  const [adminAddMemberQuery, setAdminAddMemberQuery] = React.useState('');
  const [adminAddMemberResults, setAdminAddMemberResults] = React.useState<Array<{ email: string; displayName: string }>>([]);
  const [adminAddMemberSearching, setAdminAddMemberSearching] = React.useState(false);
  const [adminAddMemberConsent, setAdminAddMemberConsent] = React.useState(false);
  const [adminAddMemberBusy, setAdminAddMemberBusy] = React.useState(false);
  const [adminAddMemberError, setAdminAddMemberError] = React.useState('');
  const adminAddMemberQueryTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  // Lead-Transfer-Dropdown: pro Team ein offener Dropdown-Index (TeamId-Key).
  const [leadTransferOpenFor, setLeadTransferOpenFor] = React.useState<string | null>(null);
  const [leadTransferBusy, setLeadTransferBusy] = React.useState(false);
  // Toast nach erfolgreicher Aktion in der Teams-Section.
  const [teamsToast, setTeamsToast] = React.useState<string>('');
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
  // v11.40: Einladungsmail-Modal — Mail mit Anmelde-Link an Organizer (zum
  // Weiterleiten) oder direkt an den hinterlegten Mailverteiler des Events.
  const [showInviteModal, setShowInviteModal] = React.useState(false);
  const [inviteSubject, setInviteSubject] = React.useState('');
  const [inviteHeading, setInviteHeading] = React.useState('');
  const [inviteBody, setInviteBody] = React.useState('');
  const [inviteTarget, setInviteTarget] = React.useState<'organizer' | 'audience'>('organizer');
  const [inviteSending, setInviteSending] = React.useState(false);
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

  // v11.36: Fairer Wartelisten-Rang einer Person in ihrer Gruppe — gleiche
  // Logik wie die Review-Box. Genutzt für die "neue Warteliste-Position" im
  // Mailtext (Vorschlag + Sammel-Versand).
  const getFairWaitlistRank = (reg: SPRegistration): number => {
    if (!selectedEvent) return 0;
    const ACT = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const isSplit = isSplitCapacity;
    const keyOf = (r: SPRegistration): string => isSplit ? (r.StarterType || r.PreferredStarterType || '?') : 'all';
    const capOf = (k: string): number => !isSplit
      ? (selectedEvent.maxParticipants || 0)
      : (k === 'Durchstarter' ? (selectedEvent.durchstarterCapacity || 0) : k === 'Funstarter' ? (selectedEvent.funstarterCapacity || 0) : 0);
    const k = keyOf(reg);
    const activeSorted = registrations
      .filter(r => ACT.indexOf(r.Status) >= 0 && keyOf(r) === k)
      .slice().sort((a, b) => a.Id - b.Id);
    const cap = capOf(k);
    const overCap = cap > 0 ? activeSorted.slice(cap) : [];
    const existingWl = registrations.filter(r => r.Status === 'Warteliste' && keyOf(r) === k);
    const fairWl = [...overCap, ...existingWl].sort((a, b) =>
      new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
    const idx = fairWl.findIndex(x => x.Id === reg.Id);
    return idx >= 0 ? idx + 1 : 0;
  };

  // v11.36: Beim Öffnen des „Bestätigen"-Dialogs die Mail-Sprache aus dem
  // Event vorbelegen (Default DE wenn nicht explizit EN) — umschaltbar.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm') {
      setObMailLang((selectedEvent?.emailLanguage || '').toUpperCase() === 'EN' ? 'EN' : 'DE');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal]);

  // v11.36: Mailtext vorbefüllen — reagiert auf Dialog-Öffnen UND Sprachwahl.
  // Enthält die neue Wartelisten-Position der ersten Zielperson.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm' && eventServiceRef && selectedEvent) {
      const t = overbookModal.targets[0];
      const nm = t ? ((t.Vorname && t.Nachname) ? `${t.Vorname} ${t.Nachname}` : t.ParticipantName) : '';
      const pos = t ? getFairWaitlistRank(t) : 0;
      const m = eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, pos);
      setObMailSubject(m.subject);
      setObMailBody(m.body);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal, obMailLang]);

  // v11.36: Überbuchungs-Entscheidung ausführen (einzeln oder Sammel) und
  // danach IDs neu vergeben + Counter/Seat-Sync + Liste neu laden.
  const runOverbookResolution = async (): Promise<void> => {
    if (!overbookModal || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setObBusy(true);
    const sub = selectedEvent.subsiteUrl;
    const isBulk = overbookModal.targets.length > 1;
    try {
      for (const reg of overbookModal.targets) {
        const grp = reg.StarterType || reg.PreferredStarterType || '';
        const nm = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
        if (overbookModal.mode === 'confirm') {
          await eventServiceRef.resolveOverbookToWaitlist(sub, reg.Id, grp);
          if (obWithMail && reg.ParticipantEmail && !selectedEvent.disableEmails) {
            // Einzeln: ggf. vom Admin editierter Text. Sammel: pro Person
            // frisch personalisiert aus dem Standard-Template.
            const mail = isBulk
              ? eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, getFairWaitlistRank(reg))
              : { subject: obMailSubject, body: obMailBody };
            try {
              await eventServiceRef.queueEmail(
                mail.subject, reg.ParticipantEmail, nm, mail.body,
                'Info', selectedEvent.title, selectedEvent.id
              );
            } catch { /* Mail-Fehler darf Korrektur nicht blockieren */ }
          }
          if (obRemoveCalendar && reg.ParticipantEmail && !selectedEvent.disableOutlook) {
            try {
              await eventServiceRef.queueOutlookEvent(
                reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
              );
            } catch { /* Kalender-Abmeldung best-effort */ }
          }
        } else {
          // Platz behalten
          if (obKeepVariant === 'active') {
            await eventServiceRef.resolveOverbookKeepActive(sub, reg.Id);
          } else {
            await eventServiceRef.resolveOverbookKeepAsFirstWaitlist(sub, reg.Id, grp);
          }
        }
      }
      // IDs neu vergeben (Aktive 1..N, Warteliste N+1..) + Counter + Seat-Sync.
      // Mit Fortschritts-Overlay, damit man bei großen Listen sieht wie weit.
      setReorderProgressLabel('IDs werden neu vergeben…');
      setReorderProgress(0);
      try { await eventServiceRef.reorderParticipantIDs(sub, pct => setReorderProgress(pct)); } catch { /* */ }
      try { await eventServiceRef.syncSeatsToActiveCount(sub, { isSplit: isSplitCapacity }); } catch { /* */ }
      setReorderProgress(null);
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch { /* einzelne Fehler werden geschluckt; Liste wird trotzdem neu geladen */ }
    setObBusy(false);
    setOverbookModal(null);
  };

  // v11.36: TeilnehmerIDs neu vergeben — gemeinsam von der Toolbox-Kachel
  // UND dem Hinweis-Modal genutzt (mit %-Fortschritts-Overlay).
  const runIdReorder = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsReorderingIDs(true);
    setReorderResult(null);
    setReorderProgressLabel('IDs werden neu vergeben…');
    setReorderProgress(0);
    try {
      const result = await eventServiceRef.reorderParticipantIDs(
        selectedEvent.subsiteUrl,
        pct => setReorderProgress(pct)
      );
      setReorderResult(`${result.success} aktualisiert, ${result.errors} Fehler`);
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch {
      setReorderResult('Fehler beim Neuvergeben der IDs');
    }
    setReorderProgress(null);
    setIsReorderingIDs(false);
  };

  // v11.70 / v11.71: Hinweis-Box „IDs sind ggf. nicht korrekt" wird jetzt
  // an die tatsaechliche TeilnehmerID-Sequenz gekoppelt — nicht mehr an
  // eine 10-Minuten-Zeit-Heuristik nach der letzten Abmeldung.
  //
  // Erwartet: alle nicht-abgemeldeten Eintraege (Status in
  // Angemeldet/QR versendet/Eingecheckt/Warteliste) haben TeilnehmerIDs,
  // die nach Sortierung lueckenlos 1..N durchlaufen. Sobald
  //   - eine ID fehlt (Luecke),
  //   - eine ID doppelt vorkommt,
  //   - ein nicht-abgemeldeter Eintrag keine (oder ≤0) ID hat,
  // ist der Zustand „IDs evtl. nicht korrekt". Typischer Trigger: gerade
  // erfolgte Abmeldung, der DEX_IDReorder-Flow ist noch nicht fertig.
  // Das gibt einen ehrlichen Status — die Box verschwindet automatisch,
  // sobald der Flow durch ist (statt nach willkuerlichen 10 Minuten).
  const recentCancellation = (regs: SPRegistration[]): { recent: boolean; whenIso: string } => {
    const active = regs.filter(r => r.Status !== 'Abgemeldet');
    if (active.length === 0) return { recent: false, whenIso: '' };
    const ids: number[] = [];
    for (const r of active) {
      const id = Number(r.TeilnehmerID);
      if (!isFinite(id) || id <= 0) {
        // Eintrag ohne gueltige ID — IDs sind kaputt, Letzten-Cancel
        // mitgeben (oder leer wenn keiner).
        return { recent: true, whenIso: latestCancelIso(regs) };
      }
      ids.push(id);
    }
    ids.sort((a, b) => a - b);
    // Lueckenlos 1..N pruefen
    for (let i = 0; i < ids.length; i++) {
      if (ids[i] !== i + 1) {
        return { recent: true, whenIso: latestCancelIso(regs) };
      }
    }
    return { recent: false, whenIso: '' };
  };
  // Hilfsfunktion: jüngste CancellationDate aus der Liste (fuer den
  // optionalen Zeit-Hinweis in der Box).
  const latestCancelIso = (regs: SPRegistration[]): string => {
    let latest = 0;
    for (const r of regs) {
      if (r.Status !== 'Abgemeldet') continue;
      const t = new Date(r.CancellationDate || '').getTime();
      if (!isNaN(t) && t > latest) latest = t;
    }
    return latest > 0 ? new Date(latest).toISOString() : '';
  };
  const idFixCheckedForRef = React.useRef<string | null>(null);

  // v11.70: kein Modal mehr beim Event-Oeffnen — der Hinweis steht ab
  // jetzt direkt als Box oben in der Teilnehmerliste, solange die
  // Bedingung erfuellt ist (siehe Render-Block unten). Der Ref bleibt
  // erhalten, um in Zukunft ein erneutes „Mount-Trigger"-Verhalten
  // einbauen zu koennen, ohne den Save-Pfad zu touchen.
  React.useEffect(() => {
    if (!selectedEvent || isLoadingRegs || registrations.length === 0) return;
    if (idFixCheckedForRef.current === selectedEvent.id) return;
    idFixCheckedForRef.current = selectedEvent.id;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations, isLoadingRegs]);

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
    if (userIds.length > 0) {
      // v11.56: Label aus dem ersten roommate-/user-Feld ableiten, statt hart
      // „Zimmerpartner" zu nennen. Wenn ein roommate-Feld existiert, nimm dessen
      // Label (User-Picker-Pairs); sonst das erste user-Feld; Fallback bleibt
      // der deutsche Default.
      const fields = selectedEvent?.eventSpecificFields || [];
      const firstRoommate = fields.filter(f => f.type === 'roommate' && f.label && f.label.trim())[0];
      const firstUser = fields.filter(f => f.type === 'user' && f.label && f.label.trim())[0];
      const roommateLabel = (firstRoommate?.label || firstUser?.label || 'Zimmerpartner').trim();
      cols.push({ id: 'roommate', label: roommateLabel });
    }
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

  // Roommate-Matching: durchsucht CustomData nach roommate-Type Feldern, extrahiert
  // Email aus "Name <email>"-Format, baut Map Email -> Partner-Email. Match-Badge,
  // wenn beide sich gegenseitig ausgewaehlt haben.
  // v11.65: ausschliesslich `roommate`-Felder, nicht mehr `user`. Bei Assistant-
  // /generischen User-Pickern macht ein „Match"-Badge semantisch keinen Sinn —
  // der wurde faelschlich auch dort gezeigt, wenn Person A und B sich
  // gegenseitig als Assistant eingetragen haben.
  const userFieldIds = (selectedEvent?.eventSpecificFields || [])
    .filter(f => f.type === 'roommate')
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
  const getRoommateInfo = (reg: { ParticipantEmail?: string }): { partnerName: string; partnerEmail: string; mutual: boolean } | null => {
    const email = (reg.ParticipantEmail || '').toLowerCase();
    if (!email) return null;
    const choice = roommateChoice[email];
    if (!choice) return null;
    const reverse = roommateChoice[choice.partnerEmail];
    const mutual = !!reverse && reverse.partnerEmail === email;
    return { partnerName: choice.partnerName || choice.partnerEmail, partnerEmail: choice.partnerEmail, mutual };
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
                    {/* v12.2: 'Abgefragte Felder'-Zeile entfernt — die
                        Custom-Field-Pills hier waren redundant; sie tauchen
                        ohnehin als Spalten in der Teilnehmer-Tabelle auf. */}
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

            {/* v11.89: Event Live/Entwurf-Toggle — flippt IsFictive
                (vorher EventStatus). Live → alle Berechtigten sehen das
                Event und können sich anmelden. Entwurf → nur Organizer,
                Admin und Test-Team sehen es. Bestehende Anmeldungen
                bleiben in beiden Modi erhalten. */}
            {(() => {
              const isDraft = !!selectedEvent.isFictive;
              return (
                <ActionTile
                  icon={isDraft ? <Check size={18} /> : <X size={18} />}
                  // v12.4: Aktueller Status immer im Titel sichtbar, damit
                  // der Admin/Organizer auf einen Blick weiß, in welchem
                  // Zustand das Event gerade ist — und was der Klick bewirkt.
                  title={isDraft
                    ? 'Aktuell: Entwurf — Event live schalten'
                    : 'Aktuell: Live — Auf Entwurf setzen'}
                  desc={isDraft
                    ? 'Schaltet das Event live — ab jetzt sehen alle Berechtigten das Event in der Liste und können sich anmelden. Mails + Outlook-Termine laufen wie konfiguriert.'
                    : 'Setzt das Event auf "Entwurf" zurück — reguläre User können sich nicht mehr anmelden, sehen das Event nicht mehr in der Eventliste. Bestehende Anmeldungen bleiben erhalten. Du kannst jederzeit wieder live schalten.'}
                  badge="organizer"
                  accent={isDraft ? 'green' : undefined}
                  onClick={async () => {
                    if (!eventServiceRef) return;
                    const nextIsFictive = !isDraft;
                    const confirmMsg = nextIsFictive
                      ? 'Event auf "Entwurf" zurücksetzen? Reguläre User sehen das Event danach nicht mehr.'
                      : 'Event live schalten? Alle Berechtigten können sich danach anmelden.';
                    if (!window.confirm(confirmMsg)) return;
                    // Legacy-Cleanup: Falls das Event noch EventStatus='Under Construction'
                    // hatte, beim Live-Schalten direkt mit auf 'Active' setzen.
                    const patch: Record<string, unknown> = { 'IsFictive': nextIsFictive };
                    if (!nextIsFictive) patch['EventStatus'] = 'Active';
                    await updateEvent(selectedEvent.id, patch);
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

            {/* v11.40: 4b. Einladungsmail — Mail mit Anmelde-Link an dich
                (zum Weiterleiten an Kollegen / Teams / externe Adressen)
                oder direkt an den auf dem Event hinterlegten Mailverteiler.
                Default-Text + Link werden vorbefuellt, sind aber im RichText-
                Editor frei editierbar. */}
            <ActionTile
              icon={<Send size={18} />}
              title={isDe ? 'Einladungsmail' : 'Invitation email'}
              desc={isDe
                ? 'Versendet eine Einladungs-Mail mit Anmelde-Link — an dich zum Weiterleiten oder direkt an den hinterlegten Mailverteiler des Events.'
                : 'Sends an invitation email with the registration link — to yourself for forwarding or directly to the configured mail distribution list of the event.'}
              badge="organizer"
              onClick={() => {
                if (!selectedEvent) return;
                const appUrl = `${siteUrl}/SitePages/DEX.aspx?env=WebView`;
                const linkHtml = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">${appUrl}</a>`;
                // v11.44: Signatur in zwei Stufen:
                //   1. "Das <Event-Titel> Orga Team" als kollektive Absender-
                //      bezeichnung — gibt der Mail einen erkennbaren Anker.
                //   2. Darunter die einzelnen Organizer mit vollem Namen
                //      (Vorname Nachname), in Reihenfolge wie im Wizard
                //      eingetragen — Haupt-Organizer zuerst.
                const orgList = (selectedEvent.organizers || [])
                  .map(s => (s || '').trim())
                  .filter(Boolean);
                const teamLine = isDe
                  ? `Das ${selectedEvent.title} Orga Team`
                  : `The ${selectedEvent.title} Organizer Team`;
                const signatureNames = orgList.length > 0
                  ? `${teamLine}<br />${orgList.join('<br />')}`
                  : teamLine;
                const defaultBody = isDe
                  ? `<p>Hallo,</p>
<p>wir laden dich herzlich zum Event <strong>${selectedEvent.title}</strong> ein.</p>
<p>Du kannst dich ab sofort über unsere Event-Plattform anmelden:</p>
<p>${linkHtml}</p>
<p>Falls du dich im Nachgang doch nicht beteiligen kannst, ist eine <strong>Abmeldung jederzeit über dieselbe Plattform</strong> möglich — bitte gib uns rechtzeitig Bescheid, damit Wartelisten-Plätze nachrücken können.</p>
<p>Bei Rückfragen meld dich gern bei uns.</p>
<p>Viele Grüße<br />${signatureNames}</p>`
                  : `<p>Hello,</p>
<p>we would like to invite you to the event <strong>${selectedEvent.title}</strong>.</p>
<p>You can register via our event platform:</p>
<p>${linkHtml}</p>
<p>If you change your mind, you can <strong>cancel anytime via the same platform</strong> — please let us know early so people on the waitlist can move up.</p>
<p>Feel free to reach out if you have any questions.</p>
<p>Best regards<br />${signatureNames}</p>`;
                setInviteSubject(isDe
                  ? `Einladung: ${selectedEvent.title}`
                  : `Invitation: ${selectedEvent.title}`);
                setInviteHeading(isDe
                  ? `Einladung zu ${selectedEvent.title}`
                  : `Invitation to ${selectedEvent.title}`);
                setInviteBody(defaultBody);
                setInviteTarget('organizer');
                setShowInviteModal(true);
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

            {/* 7. TeilnehmerIDs neu vergeben — Admin ODER Organizer des Events (v11.36) */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Hash size={18} />}
                title={isReorderingIDs ? 'IDs werden vergeben…' : 'IDs neu vergeben'}
                desc="Vergibt die TeilnehmerIDs sequentiell (1, 2, 3, …) nach Erstellungsreihenfolge. Schließt Lücken nach Stornos und sortiert die Liste sauber durch. Hinweis: nicht ausführen während gerade viele Anmeldungen laufen — erst wenn die Anmeldewelle vorbei ist."
                badge="organizer"
                busy={isReorderingIDs}
                disabled={!selectedEvent?.subsiteUrl}
                result={reorderResult}
                resultIsError={!!reorderResult && reorderResult.indexOf('Fehler') >= 0}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!confirm('TeilnehmerIDs neu vergeben (1, 2, 3, …)? Sortierung nach Erstellungsreihenfolge.\n\nNICHT ausführen, während gerade viele Anmeldungen laufen — bitte erst wenn die Anmeldewelle vorbei ist.')) return;
                  await runIdReorder();
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

            {/* 7c. Überbuchung prüfen — Admin ODER Organizer des Events (v11.36).
                Markiert pro Gruppe (bzw. gesamt) die zuletzt über Kapazität
                Angemeldeten mit OverbookReview='Pending'. Ändert KEINEN
                Status — Admin/Organizer entscheidet danach pro Person über
                die Buttons in der „Überbuchung – zu prüfen"-Box oben in der
                Teilnehmerliste. Organizer dürfen das für eigene Events, weil
                es Teilnehmerverwaltung ist (analog Abmelden/QR/Massenmail). */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Users size={18} />}
                title={isDetectingOverbook ? 'Wird geprüft…' : 'Überbuchung prüfen'}
                desc="Findet pro Gruppe (Durchstarter/Funstarter, bzw. gesamt) die zuletzt angemeldeten Personen ÜBER der Kapazität und markiert sie zur Prüfung. Es wird nichts automatisch geändert — danach entscheidest du pro Person (auf Warteliste / Platz behalten) über die Buttons oben in der Teilnehmerliste."
                badge="organizer"
                busy={isDetectingOverbook}
                disabled={!selectedEvent?.subsiteUrl}
                result={detectOverbookResult}
                resultIsError={!!detectOverbookResult && detectOverbookResult.indexOf('Fehler') >= 0}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!confirm('Überbuchung prüfen und betroffene Personen markieren? (ändert keinen Status)')) return;
                  setIsDetectingOverbook(true);
                  setDetectOverbookResult(null);
                  try {
                    const res = await eventServiceRef.detectOverbooking(selectedEvent.subsiteUrl, {
                      isSplit: isSplitCapacity,
                      maxParticipants: selectedEvent.maxParticipants || 0,
                      durchstarterCapacity: selectedEvent.durchstarterCapacity || 0,
                      funstarterCapacity: selectedEvent.funstarterCapacity || 0,
                    });
                    // Counter mit echtem Bestand abgleichen (best-effort).
                    try { await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit: isSplitCapacity }); } catch { /* */ }
                    const parts = res.groups
                      .map(g => `${g.group}: ${g.activeBefore}/${g.cap || '∞'} → ${g.marked} markiert`)
                      .join(' · ');
                    setDetectOverbookResult(res.total > 0
                      ? `${res.total} markiert (${parts})${res.errors ? ` — ${res.errors} Fehler` : ''}`
                      : `Keine Überbuchung gefunden (${parts})`);
                    const regs = await getAllRegistrations(selectedEvent.id);
                    setRegistrations(regs);
                  } catch {
                    setDetectOverbookResult('Fehler beim Prüfen der Überbuchung');
                  }
                  setIsDetectingOverbook(false);
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
                      { isB2Run, hasQuiz, customFields },
                      (count, titles) => {
                        const preview = titles.slice(0, 8).map(t => `„${t}"`).join(', ');
                        const more = titles.length > 8 ? ` …und ${titles.length - 8} weitere` : '';
                        return window.confirm(
                          `${count} überflüssige (leere) Duplikat-Spalten in der Teilnehmerliste gefunden ` +
                          `(${titles.length} Titel betroffen: ${preview}${more}).\n\n` +
                          `Diese werden jetzt gelöscht (irreversibel). Spalten mit Daten bleiben erhalten ` +
                          `und werden zur manuellen Prüfung gemeldet.\n\nFortfahren?`
                        );
                      }
                    );
                    const msgs: string[] = [];
                    if (result.added.length > 0) msgs.push(`Spalten hinzugefügt: ${result.added.join(', ')}`);
                    if (result.removed.length > 0) msgs.push(`Spalten entfernt: ${result.removed.join(', ')}`);
                    if (result.duplicatesRemoved && result.duplicatesRemoved.length > 0) {
                      msgs.push(`${result.duplicatesRemoved.length} leere Duplikate gelöscht`);
                    }
                    if (result.duplicatesWithData && result.duplicatesWithData.length > 0) {
                      const list = result.duplicatesWithData.map(t => `„${t}"`).join(', ');
                      msgs.push(`${result.duplicatesWithData.length} Duplikate mit Daten — bitte manuell prüfen: ${list}`);
                    }
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
          Sonst Grid auf 4 Spalten.
          v11.32: Bei Split-Capacity wird die separate Kapazitaets-Karten-Reihe
          unten in die „Angemeldet"-Kachel hochgezogen. Die Kachel bekommt
          dann doppelte Breite (2fr) damit Group-A/B-Breakdown sauber drin
          Platz hat — keine zwei breiten Vollbreite-Karten mehr. */}
      {(() => {
        const hasWaitlistKPI = !!(selectedEvent?.waitlistEnabled && (selectedEvent?.maxParticipants || 0) > 0);
        // Fraktionen pro Spalte — Angemeldet bekommt 2fr wenn Split aktiv ist.
        const angeFr = isSplitCapacity ? '2fr' : '1fr';
        const tail = `1fr 1fr${hasWaitlistKPI ? ' 1fr' : ''} 1fr`; // QR / Eingecheckt / [Warteliste] / Abgemeldet
        const gridCols = `${angeFr} ${tail}`;
        const active = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        const totalActive = active.length;
        const durchActive = active.filter(r => r.StarterType === 'Durchstarter').length;
        const funActive = active.filter(r => r.StarterType === 'Funstarter').length;
        const durchCap = selectedEvent?.durchstarterCapacity || 0;
        const funCap = selectedEvent?.funstarterCapacity || 0;
        const labelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
        const labelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
        const reversed = !!selectedEvent?.splitDisplayOrderReversed;
        const grpA = (
          <div key="grpA" style={{ display: 'flex', alignItems: 'baseline', gap: 8 }}>
            <span style={{ color: 'var(--dex-green-dark, #6b9a1e)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={labelA}>● {labelA}</span>
            <strong style={{ whiteSpace: 'nowrap' }}>{durchActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{durchCap}</span></strong>
          </div>
        );
        const grpB = (
          <div key="grpB" style={{ display: 'flex', alignItems: 'baseline', gap: 8 }}>
            <span style={{ color: 'var(--dex-orange, #ff8c00)', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={labelB}>● {labelB}</span>
            <strong style={{ whiteSpace: 'nowrap' }}>{funActive}<span style={{ color: 'var(--dex-gray-400)' }}>/{funCap}</span></strong>
          </div>
        );
        return (
          <div className="admin-counters" style={{ display: 'grid', gridTemplateColumns: gridCols, gap: 12, marginBottom: 24 }}>
            <div className="card" style={{ padding: 16, textAlign: 'center' }}>
              <div style={{ fontSize: '1.8rem', fontWeight: 700, color: '#1565c0' }}>{totalActive}</div>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>{t('status.registered')}</div>
              {isSplitCapacity && (
                <div style={{
                  marginTop: 10, paddingTop: 10,
                  borderTop: '1px solid var(--dex-gray-200)',
                  fontSize: '0.82rem', textAlign: 'left',
                  display: 'flex', flexDirection: 'column', gap: 4,
                }}>
                  {reversed ? <>{grpB}{grpA}</> : <>{grpA}{grpB}</>}
                </div>
              )}
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
            {hasWaitlistKPI && (
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

        {/* v11.70: Inline-Hinweisbox statt Modal — bei einer kürzlich
            erfolgten Abmeldung läuft die automatische Korrektur evtl. noch
            (Nachrücken + ID-Neuvergabe per Power-Automate-Batch). Solange
            sich die IDs evtl. noch verschieben, soll der Organizer nicht
            parallel manuell „IDs neu vergeben" anstoßen. */}
        {(() => {
          if (!selectedEvent) return null;
          const info = recentCancellation(registrations);
          if (!info.recent) return null;
          const whenStr = info.whenIso ? formatDate(info.whenIso) : '';
          return (
            <div style={{
              margin: '0 0 16px',
              padding: '14px 16px',
              borderRadius: 8,
              background: 'rgba(237,139,0,0.10)',
              border: '1px solid var(--dex-orange, #ed8b00)',
              color: 'var(--dex-orange-dark, #b35a00)',
            }}>
              <div style={{ fontWeight: 700, marginBottom: 6, fontSize: '0.9rem' }}>
                Achtung — TeilnehmerIDs sind aktuell ggf. nicht korrekt
              </div>
              <div style={{ fontSize: '0.82rem', lineHeight: 1.5 }}>
                Es gab gerade eine Abmeldung{whenStr ? <> (zuletzt: <strong>{whenStr}</strong>)</> : ''}. Die automatische Korrektur — <strong>Nachrücken von der Warteliste</strong> und <strong>TeilnehmerID-Neuvergabe</strong> — läuft im Hintergrund und ist evtl. noch nicht fertig. Bitte ein paar Minuten warten, bevor du manuell &bdquo;IDs neu vergeben&ldquo; nutzt — sonst läuft die manuelle Korrektur in die noch laufende automatische Batch-Korrektur hinein und es kann zu Doppel-Nachrücken / Inkonsistenzen kommen. Meist musst du gar nichts tun.
              </div>
            </div>
          );
        })()}

        {(() => {
          // v11.36: Überbuchungs-Review-Box. Zeigt alle per „Überbuchung
          // prüfen" markierten Personen (OverbookReview='Pending') mit
          // Fairness-Kontext + Aktions-Buttons. Erst durch eine Aktion
          // ändert sich der Status.
          const flagged = registrations.filter(r => r.OverbookReview === 'Pending');
          if (flagged.length === 0 || !selectedEvent) return null;
          const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
          const ACTIVE_ST = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
          // Gruppen-Key: bei Split die Gruppe, sonst ein gemeinsamer Topf.
          const keyOf = (r: SPRegistration): string => isSplitCapacity ? (groupOf(r) || '?') : 'all';
          const capOf = (key: string): number => {
            if (!isSplitCapacity) return selectedEvent.maxParticipants || 0;
            if (key === 'Durchstarter') return selectedEvent.durchstarterCapacity || 0;
            if (key === 'Funstarter') return selectedEvent.funstarterCapacity || 0;
            return 0;
          };
          // Pro Gruppe: aktive Anmeldungen in Anmeldereihenfolge (Id asc =
          // Reihenfolge der Registrierung — identisch zur Detect-Logik).
          const activeByGroup: Record<string, SPRegistration[]> = {};
          registrations
            .filter(r => ACTIVE_ST.indexOf(r.Status) >= 0)
            .slice()
            .sort((a, b) => a.Id - b.Id)
            .forEach(r => { const k = keyOf(r); (activeByGroup[k] = activeByGroup[k] || []).push(r); });
          // Faire Wartelisten-Reihenfolge je Gruppe: die über Kapazität
          // Aktiven + bereits vorhandene Warteliste, nach RegistrationDate.
          const fairWaitByGroup: Record<string, SPRegistration[]> = {};
          Object.keys(activeByGroup).forEach(k => {
            const cap = capOf(k);
            const overCap = cap > 0 ? activeByGroup[k].slice(cap) : [];
            const existingWl = registrations.filter(r => r.Status === 'Warteliste' && keyOf(r) === k);
            fairWaitByGroup[k] = [...overCap, ...existingWl].sort((a, b) =>
              new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
          });
          // Faire Aktiv-Gesamtzahl (bei sauberer Liste) für die faire ID.
          let totalFairActive = 0;
          Object.keys(activeByGroup).forEach(k => {
            const cap = capOf(k);
            totalFairActive += cap > 0 ? Math.min(activeByGroup[k].length, cap) : activeByGroup[k].length;
          });
          const fmtGap = (ms: number): string => {
            if (!isFinite(ms) || ms < 0) return '—';
            const s = Math.round(ms / 1000);
            if (s < 90) return `${s} Sek`;
            const m = Math.round(s / 60);
            if (m < 90) return `${m} Min`;
            const h = Math.round(m / 60);
            if (h < 48) return `${h} Std`;
            return `${Math.round(h / 24)} Tage`;
          };
          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 12 }}>
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.95rem' }}>
                  Überbuchung – zu prüfen ({flagged.length})
                </strong>
                <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                  Über Kapazität angemeldet. Pro Person entscheiden — danach werden IDs automatisch neu vergeben.
                </span>
                <div style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.78rem', padding: '5px 12px', color: 'var(--dex-red, #c00)' }}
                    onClick={() => { setOverbookModal({ mode: 'confirm', targets: flagged }); setObWithMail(true); setObRemoveCalendar(true); }}
                  >
                    Alle bestätigen ({flagged.length})
                  </button>
                  <InfoTooltip placement="left" text={
                    <>
                      <strong>Sammel-Aktion:</strong> setzt <strong>alle</strong> markierten Personen auf die <strong>Warteliste</strong> (gruppentreu).<br /><br />
                      Die Optionen <strong>mit/ohne Mail</strong>, <strong>Kalender-Abmeldung</strong> und <strong>Sprache</strong> gelten <strong>für alle gleich</strong> — eine gemeinsame Entscheidung.<br /><br />
                      Der Mailtext ist trotzdem <strong>pro Person personalisiert</strong> (Name + individuelle neue Warteliste-Position).<br /><br />
                      Sollen einzelne Personen <strong>anders</strong> behandelt werden (z.B. &bdquo;Platz behalten&ldquo;), nutze stattdessen die <strong>Einzel-Buttons</strong> pro Zeile.
                    </>
                  } />
                </div>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid rgba(237,139,0,0.4)', textAlign: 'left', color: 'var(--dex-gray-600)' }}>
                      <th style={{ padding: '4px 8px' }}>Aktuell</th>
                      <th style={{ padding: '4px 8px' }}>Name</th>
                      <th style={{ padding: '4px 8px' }}>Gruppe</th>
                      <th style={{ padding: '4px 8px' }}>Angemeldet</th>
                      <th style={{ padding: '4px 8px' }}>Über Kapazität</th>
                      <th style={{ padding: '4px 8px' }}>Abstand zum letzten fairen Platz</th>
                      <th style={{ padding: '4px 8px' }}>Fairer Platz</th>
                      <th style={{ padding: '4px 8px', textAlign: 'right' }}>Aktion</th>
                    </tr>
                  </thead>
                  <tbody>
                    {flagged.map(reg => {
                      const nm = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                      const k = keyOf(reg);
                      const grpLabel = isSplitCapacity ? (groupOf(reg) || '—') : '—';
                      const cap = capOf(k);
                      const bucket = activeByGroup[k] || [];
                      const idx = bucket.findIndex(x => x.Id === reg.Id); // 0-basiert
                      const position = idx >= 0 ? idx + 1 : null;
                      const overBy = (position !== null && cap > 0) ? position - cap : null;
                      const cutoff = (cap > 0 && cap - 1 < bucket.length) ? bucket[cap - 1] : null;
                      const cutoffNm = cutoff ? ((cutoff.Vorname && cutoff.Nachname) ? `${cutoff.Vorname} ${cutoff.Nachname}` : cutoff.ParticipantName) : '';
                      const gapMs = cutoff ? (new Date(reg.RegistrationDate).getTime() - new Date(cutoff.RegistrationDate).getTime()) : NaN;
                      const wl = fairWaitByGroup[k] || [];
                      const wlRank = wl.findIndex(x => x.Id === reg.Id) + 1; // 1-basiert; 0 = nicht gefunden
                      const fairId = totalFairActive + (wlRank > 0 ? wlRank : (overBy || 0));
                      return (
                        <tr key={reg.Id} style={{ borderBottom: '1px solid rgba(237,139,0,0.25)' }}>
                          <td style={{ padding: '6px 8px', fontWeight: 600 }}>#{reg.TeilnehmerID ?? '—'}</td>
                          <td style={{ padding: '6px 8px' }}>
                            {nm}
                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{reg.ParticipantEmail}</div>
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>{grpLabel}</td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-500)' }}>{formatDate(reg.RegistrationDate)}</td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {position !== null && cap > 0
                              ? <>Platz <strong>{position}</strong> bei Kap. {cap} <span style={{ color: 'var(--dex-red, #c00)' }}>(+{overBy})</span></>
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {cutoff
                              ? <><strong>+{fmtGap(gapMs)}</strong><div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>nach {cutoffNm} ({formatDate(cutoff.RegistrationDate)})</div></>
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                            {wlRank > 0
                              ? <>Warteliste-Platz <strong>{wlRank}</strong>{isSplitCapacity ? ` (${grpLabel})` : ''}<div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>= TeilnehmerID ~#{fairId} bei sauberer Liste</div></>
                              : '—'}
                          </td>
                          <td style={{ padding: '6px 8px', textAlign: 'right', whiteSpace: 'nowrap' }}>
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, marginRight: 10 }}>
                              <button
                                className="btn btn-secondary"
                                style={{ fontSize: '0.75rem', padding: '4px 10px', color: 'var(--dex-red, #c00)' }}
                                onClick={() => { setOverbookModal({ mode: 'confirm', targets: [reg] }); setObWithMail(true); setObRemoveCalendar(true); }}
                              >
                                Auf Warteliste
                              </button>
                              <InfoTooltip placement="left" text={
                                <>
                                  <strong>&bdquo;Auf Warteliste&ldquo;</strong> — die Person wird (gruppentreu) auf die <strong>Warteliste</strong> gesetzt; sie hatte fälschlich einen Platz.<br /><br />
                                  Im nächsten Dialog wählst du: <strong>mit oder ohne Entschuldigungs-Mail</strong> (Deloitte-Layout, geht in die Mail-Queue — nicht direkt versendet) und ob sie <strong>vom Kalendereintrag abgemeldet</strong> wird.<br /><br />
                                  Es wird ein <strong>Audit-Eintrag</strong> geschrieben (war fälschlich angemeldet, Original-Registrierung). Danach werden die <strong>TeilnehmerIDs automatisch neu vergeben</strong>.
                                </>
                              } />
                            </span>
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                              <button
                                className="btn btn-secondary"
                                style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                onClick={() => { setOverbookModal({ mode: 'keep', targets: [reg] }); setObKeepVariant('firstWaitlist'); }}
                              >
                                Platz behalten
                              </button>
                              <InfoTooltip placement="left" text={
                                <>
                                  <strong>&bdquo;Platz behalten&ldquo;</strong> — die Person verliert den Platz <strong>nicht</strong>. Im nächsten Dialog wählst du:<br /><br />
                                  <strong>(a) Erste(r) auf der Warteliste</strong> der Gruppe — rückt beim nächsten frei werdenden Platz garantiert als Erste(r) nach.<br /><br />
                                  <strong>(b) Bleibt angemeldet</strong> — die Gruppe ist dann <strong>+1</strong> über Kapazität; der nächste frei werdende Platz wird <strong>einmal nicht</strong> nachgerückt, bis die Überzahl absorbiert ist.<br /><br />
                                  Beide Varianten mit <strong>Audit-Eintrag</strong>, danach IDs neu.
                                </>
                              } />
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          );
        })()}

        {(() => {
          // v11.84: Teams-Section — Admin-Center-Team-Management.
          // Sichtbar nur fuer Events mit aktivierter Team-Anmeldung. Listet
          // alle Teams (gruppiert per TeamId, abgemeldete Mitglieder
          // ausgeblendet), mit Lead-Badge und Buttons fuer „+ Person
          // hinzufuegen" und „Lead-Rolle uebergeben". Reagiert live auf
          // `registrations` — kein zusaetzlicher Roundtrip.
          if (!selectedEvent || !selectedEvent.teamRegistrationEnabled) return null;
          if (isLoadingRegs) return null;

          // groupBy TeamId, abgemeldete Personen NICHT eingehen lassen.
          const teamsByid: Record<string, SPRegistration[]> = {};
          for (const r of registrations) {
            const tid = r.TeamId || '';
            if (!tid) continue;
            if (r.Status === 'Abgemeldet') continue;
            (teamsByid[tid] = teamsByid[tid] || []).push(r);
          }
          // Sortierung: aelteste Lead-RegistrationDate zuerst.
          const teamEntries = Object.entries(teamsByid)
            .map(([tid, members]) => {
              // Lead oben, dann TeilnehmerID aufsteigend.
              members.sort((a, b) => {
                if (!!a.TeamLead !== !!b.TeamLead) return a.TeamLead ? -1 : 1;
                const aT = (a.TeilnehmerID ?? 9_999_999) as number;
                const bT = (b.TeilnehmerID ?? 9_999_999) as number;
                return aT - bT;
              });
              const lead = members.find(m => !!m.TeamLead) || members[0];
              const leadDate = lead?.RegistrationDate ? new Date(lead.RegistrationDate).getTime() : Number.MAX_SAFE_INTEGER;
              return { tid, members, lead, leadDate };
            })
            .sort((a, b) => a.leadDate - b.leadDate);

          const teamSizeCfg = selectedEvent.teamSize || 0;
          const count = teamEntries.length;
          const canManage = isAdmin || isOrganizerFor(selectedEvent);

          const statusBadge = (st: string): React.ReactElement | null => {
            if (!st || st === 'Angemeldet') return null;
            const colorMap: Record<string, string> = {
              'Warteliste': '#b35a00',
              'QR versendet': '#3a7dbf',
              'Eingecheckt': '#4a7c1f',
            };
            const color = colorMap[st] || 'var(--dex-gray-500)';
            return (
              <span style={{
                display: 'inline-block', padding: '1px 8px', borderRadius: 10,
                background: `${color}15`, color, fontSize: '0.7rem', fontWeight: 600, marginLeft: 6,
              }}>{st}</span>
            );
          };

          return (
            <div style={{ marginBottom: 20, padding: 16, borderRadius: 12, border: '1px solid var(--dex-gray-200)', background: '#fff' }}>
              <div
                role="button"
                tabIndex={0}
                onClick={() => setTeamsCollapsed(v => !v)}
                onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setTeamsCollapsed(v => !v); } }}
                style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', userSelect: 'none' }}
              >
                <Users size={20} />
                <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1rem' }}>
                  Teams ({count})
                </strong>
                <span style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  {teamsCollapsed ? <ChevronDown size={18} /> : <ChevronUp size={18} />}
                </span>
              </div>
              {!teamsCollapsed && (
                <div style={{ marginTop: 14, display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {teamEntries.length === 0 && (
                    <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.88rem', fontStyle: 'italic' }}>
                      Keine Team-Anmeldungen bisher.
                    </div>
                  )}
                  {teamEntries.map(({ tid, members, lead }) => {
                    const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
                    const total = members.length;
                    const free = teamSizeCfg > 0 ? Math.max(0, teamSizeCfg - total) : 0;
                    const canAdd = canManage && (teamSizeCfg === 0 || total < teamSizeCfg);
                    const leadEmail = lead?.ParticipantEmail || '';
                    const otherMembers = members.filter(m => m.Id !== lead?.Id);
                    return (
                      <div
                        key={tid}
                        style={{
                          padding: 14,
                          border: '1px solid var(--dex-gray-200)',
                          borderRadius: 10,
                          background: 'var(--dex-gray-50, #f7f7f7)',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap', marginBottom: 10 }}>
                          <strong style={{ fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
                            {teamName ? `Team „${teamName}"` : 'Team (ohne Namen)'}
                          </strong>
                          <span style={{ color: 'var(--dex-gray-600)', fontSize: '0.85rem' }}>
                            {teamSizeCfg > 0 ? `${total}/${teamSizeCfg} belegt` : `${total} Mitglieder`}
                          </span>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {members.map(m => {
                            const name = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                            const isLead = !!m.TeamLead;
                            return (
                              <div
                                key={m.Id}
                                style={{
                                  display: 'flex', alignItems: 'center', gap: 10,
                                  padding: '4px 0',
                                }}
                              >
                                <div style={{ position: 'relative', width: 32, height: 32, flexShrink: 0 }}>
                                  <img
                                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                                    alt={name}
                                    onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                    style={{
                                      width: 32, height: 32, borderRadius: '50%', objectFit: 'cover',
                                      background: 'var(--dex-gray-100)',
                                      transition: 'transform 0.18s ease',
                                      transformOrigin: 'left center',
                                      cursor: 'pointer',
                                    }}
                                    onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)'; (e.currentTarget as HTMLImageElement).style.zIndex = '10'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 8px 24px rgba(0,0,0,0.25)'; }}
                                    onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.zIndex = ''; (e.currentTarget as HTMLImageElement).style.position = ''; (e.currentTarget as HTMLImageElement).style.boxShadow = ''; }}
                                  />
                                </div>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: '0.88rem', fontWeight: 500 }}>
                                    {name}
                                    {statusBadge(m.Status)}
                                  </div>
                                  <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                </div>
                                {isLead && (
                                  <span style={{
                                    display: 'inline-block', padding: '2px 10px', borderRadius: 12,
                                    background: 'var(--dex-green, #86bc25)', color: '#fff',
                                    fontSize: '0.72rem', fontWeight: 700,
                                  }}>Lead</span>
                                )}
                              </div>
                            );
                          })}
                        </div>
                        {canManage && (
                          <div style={{ marginTop: 12, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', position: 'relative' }}>
                            {canAdd && (
                              <button
                                type="button"
                                className="btn btn-secondary"
                                style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                onClick={() => {
                                  setAdminAddMemberDialog({ teamId: tid, teamName, freeSlots: free });
                                  setAdminAddMemberPick(null);
                                  setAdminAddMemberQuery('');
                                  setAdminAddMemberResults([]);
                                  setAdminAddMemberConsent(false);
                                  setAdminAddMemberError('');
                                }}
                              >
                                <Plus size={14} /> Person hinzufügen
                                {teamSizeCfg > 0 && ` (${free} Slot${free === 1 ? '' : 's'} frei)`}
                              </button>
                            )}
                            {otherMembers.length > 0 && (
                              <>
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  style={{ fontSize: '0.8rem', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                                  onClick={() => setLeadTransferOpenFor(leadTransferOpenFor === tid ? null : tid)}
                                >
                                  <RefreshCw size={14} /> Lead-Rolle übergeben
                                </button>
                                {leadTransferOpenFor === tid && (
                                  <div style={{
                                    position: 'absolute', top: '100%', left: 0, marginTop: 6,
                                    background: '#fff', border: '1px solid var(--dex-gray-300)',
                                    borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
                                    zIndex: 20, minWidth: 280, maxWidth: 360, padding: 6,
                                  }}>
                                    <div style={{ padding: '6px 10px', fontSize: '0.78rem', color: 'var(--dex-gray-600)', borderBottom: '1px solid var(--dex-gray-100)' }}>
                                      Neue Lead-Rolle übertragen an:
                                    </div>
                                    {otherMembers.map(m => {
                                      const nm = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
                                      return (
                                        <button
                                          key={m.Id}
                                          type="button"
                                          disabled={leadTransferBusy}
                                          onClick={async () => {
                                            if (leadTransferBusy) return;
                                            setLeadTransferBusy(true);
                                            try {
                                              const res = await transferTeamLead(selectedEvent.id, tid, m.ParticipantEmail);
                                              if (res.ok) {
                                                setTeamsToast(`Lead-Rolle wurde an ${nm} übergeben.`);
                                                const regs = await getAllRegistrations(selectedEvent.id);
                                                setRegistrations(regs);
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              } else {
                                                setTeamsToast(`Lead-Übergabe fehlgeschlagen: ${res.reason || 'Unbekannter Fehler'}.`);
                                                window.setTimeout(() => setTeamsToast(''), 4500);
                                              }
                                            } finally {
                                              setLeadTransferBusy(false);
                                              setLeadTransferOpenFor(null);
                                            }
                                          }}
                                          style={{
                                            display: 'flex', alignItems: 'center', gap: 8,
                                            width: '100%', padding: '8px 10px', border: 'none',
                                            background: 'transparent', cursor: 'pointer',
                                            textAlign: 'left', borderRadius: 6,
                                          }}
                                          onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = 'var(--dex-gray-100)'; }}
                                          onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = 'transparent'; }}
                                        >
                                          <img
                                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=S`}
                                            alt={nm}
                                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                            style={{ width: 24, height: 24, borderRadius: '50%' }}
                                          />
                                          <div style={{ minWidth: 0 }}>
                                            <div style={{ fontSize: '0.85rem', fontWeight: 500 }}>{nm}</div>
                                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{m.ParticipantEmail}</div>
                                          </div>
                                        </button>
                                      );
                                    })}
                                    <button
                                      type="button"
                                      onClick={() => setLeadTransferOpenFor(null)}
                                      style={{
                                        width: '100%', padding: '6px 10px',
                                        border: 'none', borderTop: '1px solid var(--dex-gray-100)',
                                        background: 'transparent', cursor: 'pointer',
                                        fontSize: '0.78rem', color: 'var(--dex-gray-500)',
                                      }}
                                    >Abbrechen</button>
                                  </div>
                                )}
                              </>
                            )}
                            {/* leadEmail nur als Referenz fuer den Lead-Lookup behalten — nicht fuer's TS-Linting wegwerfen. */}
                            <span style={{ display: 'none' }}>{leadEmail}</span>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })()}

        {teamsToast && (
          <div style={{
            marginBottom: 14, padding: '10px 14px', borderRadius: 8,
            background: 'rgba(134,188,37,0.12)', border: '1px solid var(--dex-green, #86bc25)',
            color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.88rem',
          }}>
            {teamsToast}
          </div>
        )}

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
                  // v11.56: Label dynamisch aus availableColumns nehmen (entstammt dem
                  // ersten roommate-/user-Feld der Custom-Field-Definition) statt
                  // hartcodiertem „Zimmerpartner".
                  const roommateCol = availableColumns.find(c => c.id === 'roommate');
                  const roommateLabel = roommateCol?.label || 'Zimmerpartner';
                  return (
                    <th key={id} style={baseStyle} title="Ausgewählter User-Picker-Wert aus diesem Feld. Match = beide haben sich gegenseitig ausgewählt.">
                      {roommateLabel}{hideButton(id)}
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
                        const photoEmail = (info.partnerEmail || '').trim();
                        return (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            {photoEmail && (
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(photoEmail)}&size=S`}
                                alt={info.partnerName}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                              />
                            )}
                            <span>{info.partnerName}</span>
                            {info.mutual && (
                              <span
                                className="badge"
                                style={{ marginLeft: 2, background: 'var(--dex-green)', color: '#fff', padding: '1px 6px', borderRadius: 4, fontSize: '0.7rem' }}
                                title="Beide haben sich gegenseitig als Zimmerpartner ausgewählt"
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

                  {/* v11.98: Split-/Merged-Toggle bei Split-Kapazität.
                      Default 'split' — getrennte Tabellen pro Gruppe,
                      kleinere zuerst. */}
                  {(() => {
                    const renderTable = (rows: SPRegistration[], indexOffset: number): React.ReactElement => (
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                        <thead>
                          <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                            {visibleColumnIds.map(id => renderHeader(id))}
                          </tr>
                        </thead>
                        <tbody>
                          {rows.map((reg, i) => {
                            const isOverbook = reg.OverbookReview === 'Pending';
                            return (
                              <tr
                                key={reg.Id}
                                title={isOverbook ? 'Über Kapazität angemeldet — siehe Box „Überbuchung – zu prüfen" oben' : undefined}
                                style={{
                                  borderBottom: '1px solid var(--dex-gray-100)',
                                  ...(isOverbook
                                    ? { background: 'rgba(237,139,0,0.13)', boxShadow: 'inset 3px 0 0 var(--dex-orange, #ed8b00)' }
                                    : {}),
                                }}
                              >
                                {visibleColumnIds.map(id => renderCell(id, reg, indexOffset + i))}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    );

                    if (!isSplitCapacity || splitParticipantsView === 'merged') {
                      return (
                        <>
                          {isSplitCapacity && (
                            <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                              <SplitMergeToggle view={splitParticipantsView} setView={setSplitParticipantsView} isDe={isDe} />
                            </div>
                          )}
                          {renderTable(activeRegs, 0)}
                        </>
                      );
                    }

                    // Split-View: nach Gruppe trennen (StarterType ||
                    // PreferredStarterType), kleinere Gruppe zuerst.
                    const lblA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
                    const lblB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
                    const groupA = activeRegs.filter(r => (r.StarterType || r.PreferredStarterType) === 'Durchstarter');
                    const groupB = activeRegs.filter(r => (r.StarterType || r.PreferredStarterType) === 'Funstarter');
                    const groupNone = activeRegs.filter(r => !(r.StarterType || r.PreferredStarterType));
                    const groups = [
                      { label: lblA, key: 'A', rows: groupA, cap: selectedEvent?.durchstarterCapacity || 0 },
                      { label: lblB, key: 'B', rows: groupB, cap: selectedEvent?.funstarterCapacity || 0 },
                    ].sort((x, y) => x.rows.length - y.rows.length);
                    let runningIdx = 0;
                    return (
                      <>
                        <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
                          <SplitMergeToggle view={splitParticipantsView} setView={setSplitParticipantsView} isDe={isDe} />
                        </div>
                        {groups.map(g => {
                          const offset = runningIdx;
                          runningIdx += g.rows.length;
                          return (
                            <div key={g.key} style={{ marginBottom: 20 }}>
                              <h4 style={{
                                margin: '0 0 8px', color: 'var(--dex-green-dark, #4a7c1f)',
                                fontSize: '0.95rem', fontWeight: 700, display: 'flex',
                                alignItems: 'baseline', gap: 8,
                              }}>
                                <span>{g.label}</span>
                                <span style={{ color: 'var(--dex-gray-500)', fontWeight: 500, fontSize: '0.85rem' }}>
                                  ({g.rows.length}{g.cap > 0 ? ` / ${g.cap}` : ''})
                                </span>
                              </h4>
                              {g.rows.length === 0 ? (
                                <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
                                  {isDe ? 'Keine Teilnehmer in dieser Gruppe.' : 'No participants in this group.'}
                                </p>
                              ) : renderTable(g.rows, offset)}
                            </div>
                          );
                        })}
                        {groupNone.length > 0 && (
                          <div style={{ marginBottom: 20 }}>
                            <h4 style={{ margin: '0 0 8px', color: 'var(--dex-gray-500)', fontSize: '0.95rem', fontWeight: 700 }}>
                              {isDe ? 'Ohne Gruppe' : 'No group'} <span style={{ color: 'var(--dex-gray-400)', fontWeight: 500, fontSize: '0.85rem' }}>({groupNone.length})</span>
                            </h4>
                            {renderTable(groupNone, runningIdx)}
                          </div>
                        )}
                      </>
                    );
                  })()}
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

                          // v11.89: Multi-Select-Dropdown (vorher Checkbox-Liste).
                          // Werte werden weiterhin ' | '-getrennt gespeichert.
                          if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                            const selected = value.split(' | ').map(s => s.trim()).filter(Boolean);
                            return (
                              <div key={cf.id} style={{ gridColumn: '1 / -1' }}>
                                {labelEl}
                                <MultiSelectDropdown
                                  options={cf.options}
                                  value={selected}
                                  onChange={next => setVal(next.join(' | '))}
                                  placeholder={isDe ? '— bitte wählen —' : '— please choose —'}
                                />
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

      {/* ===== EINLADUNGSMAIL MODAL (v11.40) ===== */}
      {showInviteModal && selectedEvent && (() => {
        const audienceEmails = (selectedEvent.audienceFilter || [])
          .map(s => (s || '').trim())
          .filter(Boolean);
        const myEmail = currentUser.email || '';
        const myDisplayName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || myEmail;
        const targetEmails = inviteTarget === 'organizer' ? [myEmail].filter(Boolean) : audienceEmails;
        // v11.43: Organizer-Mails als CC mitschicken — damit alle Organizer
        // sehen, dass die Einladung raus ist und ggf. auf Rueckfragen
        // antworten koennen. Duplikate gegenueber TO werden rausgefiltert
        // (z.B. wenn der Sender selbst Organizer ist und 'An mich' waehlt).
        const toLcSet = new Set(targetEmails.map(e => (e || '').toLowerCase()));
        const ccEmails = (selectedEvent.organizerEmails || [])
          .map(s => (s || '').trim())
          .filter(Boolean)
          .filter(e => !toLcSet.has(e.toLowerCase()));
        // v11.41: Blocked-Check fuer den aktuell gewaehlten Empfaenger-Modus.
        // 'organizer'-Modus blockt eigentlich nie — die eigene Mail ist immer
        // eine Person, kein Verteiler — aber wir laufen das defensiv mit.
        const blockedInTargets = getBlockedInviteRecipients(targetEmails);
        const blockedInAudience = getBlockedInviteRecipients(audienceEmails);
        const recipientLabel = inviteTarget === 'organizer'
          ? (isDe ? `An mich (${myEmail})` : `To me (${myEmail})`)
          : (isDe
            ? `An Mailverteiler (${audienceEmails.length === 0 ? 'leer' : audienceEmails.length + ' Empfänger'})`
            : `To mail distribution (${audienceEmails.length === 0 ? 'empty' : audienceEmails.length + ' recipients'})`);
        const orgNames = (selectedEvent.organizers || []).join(', ');
        const appUrl = `${siteUrl}/SitePages/DEX.aspx?env=WebView`;
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Organizer: orgNames,
          Link: appUrl,
        };
        const customLogo = (() => {
          try {
            const o = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const sendAction = async (): Promise<void> => {
          if (!eventServiceRef || !selectedEvent) return;
          if (targetEmails.length === 0) {
            alert(isDe
              ? (inviteTarget === 'audience'
                ? 'Es ist kein Mailverteiler auf dem Event hinterlegt. Bitte zuerst in Schritt 3 (Sichtbarkeit) Empfänger ergänzen.'
                : 'Keine eigene E-Mail-Adresse verfügbar.')
              : (inviteTarget === 'audience'
                ? 'No mail distribution list configured on the event. Please add recipients in step 3 (Visibility) first.'
                : 'No own email address available.'));
            return;
          }
          // v11.41: Hart blocken — Einladungsmail darf nie an pauschale
          // Standort-/All-Verteiler ('deall', 'all', 'de.<stadt>') gehen.
          if (blockedInTargets.length > 0) {
            const lines = blockedInTargets.map(b => `• ${b.email}  (${b.reason})`).join('\n');
            alert(isDe
              ? `Die Einladungs-Mail darf NICHT an pauschale Standort- oder All-Verteiler verschickt werden.\n\nFolgende Empfänger sind blockiert:\n\n${lines}\n\nBitte entferne diese Adressen aus dem Mailverteiler in Schritt 3 des Event-Edits oder nutze die Option „An mich (zum Weiterleiten)".`
              : `The invitation email must NOT be sent to entire location or all-distribution lists.\n\nThe following recipients are blocked:\n\n${lines}\n\nPlease remove these addresses from the mail distribution in step 3 of event edit, or use the option "To me (for forwarding)".`);
            return;
          }
          const confirmMsg = isDe
            ? (inviteTarget === 'organizer'
              ? `Einladungs-Mail an dich selbst (${myEmail}) senden? Du kannst sie anschließend aus Outlook an deinen Verteiler weiterleiten.`
              : `Einladungs-Mail an ${audienceEmails.length} Empfänger des Mailverteilers senden?\n\n${audienceEmails.join(', ')}`)
            : (inviteTarget === 'organizer'
              ? `Send invitation email to yourself (${myEmail})? You can then forward it from Outlook to your distribution list.`
              : `Send invitation email to ${audienceEmails.length} recipients of the mail distribution?\n\n${audienceEmails.join(', ')}`);
          if (!confirm(confirmMsg)) return;
          setInviteSending(true);
          const resolvedSubject = replacePlaceholders(inviteSubject, previewVars);
          const resolvedHeading = replacePlaceholders(inviteHeading, previewVars);
          const resolvedBody = replacePlaceholders(inviteBody, previewVars);
          const fullBody = wrapTemplate('#86bc25', resolvedHeading, `Event ${selectedEvent.title}`, resolvedBody);
          const allEmails = targetEmails.join(';');
          const ccString = ccEmails.join(';');
          const recipientName = inviteTarget === 'organizer' ? myDisplayName : (isDe ? 'Mailverteiler' : 'Mail distribution');
          try {
            await eventServiceRef.queueEmail(
              resolvedSubject, allEmails, recipientName, fullBody,
              'Einladung', selectedEvent.title, selectedEvent.id,
              ccString || undefined,
            );
            setInviteSending(false);
            alert(isDe
              ? `Einladungs-Mail an ${targetEmails.length} Empfänger in die Warteschlange eingetragen.`
              : `Invitation email queued for ${targetEmails.length} recipient(s).`);
            setShowInviteModal(false);
          } catch {
            setInviteSending(false);
            alert(isDe ? 'Fehler beim Eintragen der E-Mail.' : 'Error queueing the email.');
          }
        };
        const headerExtra = (
          <div style={{
            padding: 12,
            background: 'var(--dex-gray-50, #fafafa)',
            border: '1px solid var(--dex-gray-200)',
            borderRadius: 'var(--dex-radius)',
            marginBottom: 4,
          }}>
            <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 8 }}>
              {isDe ? 'Empfänger' : 'Recipient'}
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: 'pointer', fontSize: '0.82rem' }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'organizer'}
                onChange={() => setInviteTarget('organizer')}
                style={{ marginTop: 3 }}
              />
              <span>
                <strong>{isDe ? 'An mich — zum Weiterleiten' : 'To me — for forwarding'}</strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>
                  {myEmail}
                </span>
              </span>
            </label>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: audienceEmails.length === 0 ? 'not-allowed' : 'pointer', fontSize: '0.82rem', opacity: audienceEmails.length === 0 ? 0.55 : 1 }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'audience'}
                onChange={() => setInviteTarget('audience')}
                disabled={audienceEmails.length === 0}
                style={{ marginTop: 3 }}
              />
              <span style={{ flex: 1 }}>
                <strong>
                  {isDe
                    ? `An Mailverteiler des Events (${audienceEmails.length})`
                    : `To event mail distribution (${audienceEmails.length})`}
                </strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem', wordBreak: 'break-word' }}>
                  {audienceEmails.length === 0
                    ? (isDe
                      ? 'Kein Mailverteiler auf dem Event hinterlegt — in Schritt 3 (Sichtbarkeit) im Event-Edit ergänzen.'
                      : 'No mail distribution configured — add recipients in step 3 (Visibility) of event edit.')
                    : audienceEmails.join(', ')}
                </span>
                {blockedInAudience.length > 0 && (
                  <div style={{
                    marginTop: 6, padding: '6px 8px',
                    background: '#fef3f2', border: '1px solid #c9302c',
                    borderRadius: 6, color: '#7a1f1c',
                    fontSize: '0.75rem', lineHeight: 1.4,
                  }}>
                    <strong>
                      {isDe ? '⚠ Blockierte Empfänger im Mailverteiler:' : '⚠ Blocked recipients in the distribution list:'}
                    </strong>
                    <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                      {blockedInAudience.map(b => (
                        <li key={b.email}><code>{b.email}</code> — {b.reason}</li>
                      ))}
                    </ul>
                    <div style={{ marginTop: 4 }}>
                      {isDe
                        ? 'Pauschale Standort- oder All-Verteiler sind für Einladungs-Mails nicht zulässig. Bitte aus dem Mailverteiler entfernen (Event-Edit, Schritt 3) — sonst wird das Senden blockiert.'
                        : 'Entire location or all-distribution lists are not allowed for invitation emails. Please remove from the distribution list (event edit, step 3) — otherwise sending is blocked.'}
                    </div>
                  </div>
                )}
              </span>
            </label>
            {ccEmails.length > 0 && (
              <div style={{
                marginTop: 10, paddingTop: 8,
                borderTop: '1px dashed var(--dex-gray-200)',
                fontSize: '0.78rem', color: 'var(--dex-gray-600)',
              }}>
                <strong style={{ color: 'var(--dex-gray-700)' }}>{isDe ? 'CC' : 'CC'}: </strong>
                <span style={{ wordBreak: 'break-word' }}>{ccEmails.join(', ')}</span>
                <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                  {isDe
                    ? 'Alle Organizer dieses Events werden automatisch in CC gesetzt.'
                    : 'All organizers of this event are automatically added in CC.'}
                </div>
              </div>
            )}
          </div>
        );
        return (
          <HtmlEditorModal
            open={showInviteModal}
            onClose={() => !inviteSending && setShowInviteModal(false)}
            title={isDe
              ? `Einladungsmail: ${selectedEvent.title}`
              : `Invitation email: ${selectedEvent.title}`}
            value={inviteBody}
            onChange={setInviteBody}
            previewMode="email"
            emailSubject={inviteSubject}
            onEmailSubjectChange={setInviteSubject}
            emailHeading={inviteHeading}
            onEmailHeadingChange={setInviteHeading}
            emailHeadingColor="#86bc25"
            previewVars={previewVars}
            insertableVars={[
              { key: '{{EventTitle}}', label: isDe ? 'Event-Titel' : 'Event title' },
              { key: '{{Link}}', label: isDe ? 'Anmelde-Link' : 'Registration link' },
              { key: '{{Organizer}}', label: 'Organizer' },
            ]}
            imageBase64={customLogo}
            headerExtra={headerExtra}
            extraAction={{
              label: inviteSending
                ? (isDe ? 'Wird eingetragen…' : 'Queueing…')
                : (isDe ? `Senden — ${recipientLabel}` : `Send — ${recipientLabel}`),
              onClick: sendAction,
              disabled: inviteSending
                || !inviteSubject.trim()
                || !inviteBody.trim()
                || targetEmails.length === 0,
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

      {/* v11.36: Fortschritts-Overlay für die ID-Neuvergabe (mit %). */}
      {reorderProgress !== null && (
        <div
          role="dialog"
          aria-modal="true"
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 2100,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
        >
          <div className="card" style={{ width: '100%', maxWidth: 420, padding: 28, background: '#fff', borderRadius: 10, textAlign: 'center' }}>
            <div style={{ fontWeight: 600, marginBottom: 14 }}>{reorderProgressLabel || 'IDs werden neu vergeben…'}</div>
            <div style={{ height: 12, borderRadius: 6, background: 'var(--dex-gray-200, #e5e5e5)', overflow: 'hidden' }}>
              <div
                style={{
                  height: '100%', width: `${reorderProgress}%`,
                  background: 'var(--dex-green, #86bc25)', borderRadius: 6,
                  transition: 'width 0.25s ease',
                }}
              />
            </div>
            <div style={{ marginTop: 10, fontSize: '1.1rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {reorderProgress}%
            </div>
            <div style={{ marginTop: 6, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              Bitte warten — das Fenster nicht schließen.
            </div>
          </div>
        </div>
      )}

      {/* v11.70: kein Modal mehr — der Hinweis wird inline ueber der
          Teilnehmerliste angezeigt (siehe Render-Block oberhalb der
          Teilnehmer-Tabelle). */}

      {/* v11.36: Überbuchungs-Entscheidungs-Modal (Bestätigen / Platz behalten) */}
      {overbookModal && selectedEvent && (
        <div
          role="dialog"
          aria-modal="true"
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 2000,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
          onClick={() => { if (!obBusy) setOverbookModal(null); }}
        >
          <div
            onClick={e => e.stopPropagation()}
            className="card"
            style={{ width: '100%', maxWidth: 560, maxHeight: '88vh', overflow: 'auto', padding: 24, background: '#fff', borderRadius: 8 }}
          >
            {overbookModal.mode === 'confirm' ? (
              <>
                <h3 style={{ marginTop: 0 }}>
                  Auf Warteliste bestätigen ({overbookModal.targets.length})
                </h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  {overbookModal.targets.length === 1
                    ? `${(overbookModal.targets[0].Vorname && overbookModal.targets[0].Nachname) ? `${overbookModal.targets[0].Vorname} ${overbookModal.targets[0].Nachname}` : overbookModal.targets[0].ParticipantName} wird auf die Warteliste der Gruppe zurückgesetzt. Im Audit-Log wird vermerkt, dass die Person fälschlich angemeldet war.`
                    : `${overbookModal.targets.length} Personen werden auf die Warteliste zurückgesetzt. Im Audit-Log jeder Person wird der Vorgang vermerkt.`}
                </p>
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.88rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="checkbox" checked={obWithMail} onChange={e => setObWithMail(e.target.checked)} disabled={obBusy} />
                  Mit Entschuldigungs-Mail (Deloitte-Layout, in die Mail-Queue)
                </label>
                {obWithMail && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '0 0 10px', fontSize: '0.82rem' }}>
                    <span style={{ color: 'var(--dex-gray-600)' }}>Sprache:</span>
                    {(['DE', 'EN'] as const).map(lng => (
                      <button
                        key={lng}
                        type="button"
                        className="btn btn-secondary"
                        disabled={obBusy}
                        onClick={() => setObMailLang(lng)}
                        style={{
                          fontSize: '0.75rem', padding: '3px 12px',
                          ...(obMailLang === lng ? { background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 600 } : {}),
                        }}
                      >
                        {lng === 'DE' ? 'Deutsch' : 'English'}
                      </button>
                    ))}
                  </div>
                )}
                {obWithMail && overbookModal.targets.length === 1 && (
                  <div style={{ marginBottom: 10 }}>
                    <input
                      className="form-input"
                      value={obMailSubject}
                      onChange={e => setObMailSubject(e.target.value)}
                      disabled={obBusy}
                      style={{ width: '100%', marginBottom: 6, padding: '6px 10px', fontSize: '0.82rem' }}
                    />
                    <textarea
                      value={obMailBody}
                      onChange={e => setObMailBody(e.target.value)}
                      disabled={obBusy}
                      rows={5}
                      style={{ width: '100%', fontFamily: 'monospace', fontSize: '0.72rem', padding: 8 }}
                    />
                    <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', margin: '4px 0 8px' }}>
                      Vorschlagstext — editierbar. Wird in die Mail-Queue gelegt, nicht direkt versendet.
                    </p>
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', marginBottom: 4 }}>Vorschau (echte Deloitte-Mail):</div>
                    <div
                      style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 6, maxHeight: 280, overflow: 'auto', background: '#fff' }}
                      dangerouslySetInnerHTML={{
                        __html: obMailBody
                          .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
                          .replace(/\{\{ORB_URL\}\}/g, getCachedOrbBase64() || ''),
                      }}
                    />
                  </div>
                )}
                {obWithMail && overbookModal.targets.length > 1 && (() => {
                  const t0 = overbookModal.targets[0];
                  const nm0 = (t0.Vorname && t0.Nachname) ? `${t0.Vorname} ${t0.Nachname}` : t0.ParticipantName;
                  const prev = eventServiceRef
                    ? eventServiceRef.buildOverbookApologyEmail(nm0, selectedEvent.title, obMailLang, getFairWaitlistRank(t0)).body
                    : '';
                  return (
                    <div style={{ marginBottom: 10 }}>
                      <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', margin: '0 0 6px' }}>
                        Bei &bdquo;Alle&ldquo; wird der Standardtext je Person personalisiert versendet (eigene Wartelisten-Position). Vorschau am Beispiel der ersten Person:
                      </p>
                      <div
                        style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 6, maxHeight: 260, overflow: 'auto', background: '#fff' }}
                        dangerouslySetInnerHTML={{
                          __html: prev
                            .replace(/\{\{LOGO_URL\}\}/g, getCachedLogoBase64() || '')
                            .replace(/\{\{ORB_URL\}\}/g, getCachedOrbBase64() || ''),
                        }}
                      />
                    </div>
                  );
                })()}
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.88rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="checkbox" checked={obRemoveCalendar} onChange={e => setObRemoveCalendar(e.target.checked)} disabled={obBusy} />
                  Vom Kalendereintrag abmelden (falls vorhanden)
                </label>
              </>
            ) : (
              <>
                <h3 style={{ marginTop: 0 }}>Platz behalten ({overbookModal.targets.length})</h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  Wie soll die Person ihren Platz behalten?
                </p>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.85rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="radio" name="obkeep" checked={obKeepVariant === 'firstWaitlist'} onChange={() => setObKeepVariant('firstWaitlist')} disabled={obBusy} style={{ marginTop: 3 }} />
                  <span><strong>Erste(r) auf der Warteliste</strong> — rückt beim nächsten frei werdenden Platz der Gruppe garantiert als Erste(r) nach (risikoarm).</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.85rem', margin: '10px 0', cursor: 'pointer' }}>
                  <input type="radio" name="obkeep" checked={obKeepVariant === 'active'} onChange={() => setObKeepVariant('active')} disabled={obBusy} style={{ marginTop: 3 }} />
                  <span><strong>Bleibt angemeldet</strong> (als Letzte(r)) — Gruppe bleibt +1, der nächste frei werdende Platz wird einmal nicht nachgerückt, bis die Überzahl absorbiert ist.</span>
                </label>
              </>
            )}
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 18 }}>
              <button className="btn btn-secondary" onClick={() => setOverbookModal(null)} disabled={obBusy}>
                Abbrechen
              </button>
              <button className="btn btn-primary" onClick={() => { runOverbookResolution().catch(() => { /* */ }); }} disabled={obBusy}>
                {obBusy ? 'Wird ausgeführt…' : (overbookModal.mode === 'confirm' ? 'Bestätigen & IDs neu vergeben' : 'Übernehmen & IDs neu vergeben')}
              </button>
            </div>
          </div>
        </div>
      )}

      {adminAddMemberDialog && selectedEvent && (() => {
        const closeDlg = (): void => {
          setAdminAddMemberDialog(null);
          setAdminAddMemberPick(null);
          setAdminAddMemberQuery('');
          setAdminAddMemberResults([]);
          setAdminAddMemberConsent(false);
          setAdminAddMemberError('');
          setAdminAddMemberBusy(false);
        };
        const submit = async (): Promise<void> => {
          if (!adminAddMemberDialog || !adminAddMemberPick || !adminAddMemberConsent || adminAddMemberBusy) return;
          setAdminAddMemberBusy(true);
          setAdminAddMemberError('');
          try {
            const res = await addTeamMember(
              selectedEvent.id,
              adminAddMemberDialog.teamId,
              adminAddMemberDialog.teamName || undefined,
              adminAddMemberPick
            );
            if (!res.ok) {
              if (res.reason && res.reason.startsWith('already-registered')) {
                setAdminAddMemberError('Diese Person ist bereits beim Event angemeldet — bitte abmelden lassen, bevor du sie zum Team hinzufügst.');
              } else if (res.reason === 'team-full') {
                setAdminAddMemberError('Das Team ist bereits voll.');
              } else {
                setAdminAddMemberError('Hinzufügen fehlgeschlagen.');
              }
              setAdminAddMemberBusy(false);
              return;
            }
            const pickedName = adminAddMemberPick.displayName || adminAddMemberPick.email;
            setTeamsToast(`${pickedName} wurde zum Team hinzugefügt — Mail + Outlook werden versendet.`);
            window.setTimeout(() => setTeamsToast(''), 4500);
            const regs = await getAllRegistrations(selectedEvent.id);
            setRegistrations(regs);
            closeDlg();
          } catch {
            setAdminAddMemberError('Hinzufügen fehlgeschlagen.');
            setAdminAddMemberBusy(false);
          }
        };
        return (
          <div
            role="dialog"
            aria-modal="true"
            onClick={() => { if (!adminAddMemberBusy) closeDlg(); }}
            style={{
              position: 'fixed', inset: 0, zIndex: 2000,
              background: 'rgba(0,0,0,0.55)',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              padding: 16,
            }}
          >
            <div
              onClick={e => e.stopPropagation()}
              style={{
                background: '#fff', borderRadius: 12, padding: '28px 32px',
                maxWidth: 540, width: '100%',
                boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
                display: 'flex', flexDirection: 'column', gap: 14,
                maxHeight: '90vh', overflowY: 'auto',
              }}
            >
              <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
                {adminAddMemberDialog.teamName
                  ? `Person zum Team „${adminAddMemberDialog.teamName}" hinzufügen`
                  : 'Person zum Team hinzufügen'}
              </h3>
              <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                {(selectedEvent.teamSize || 0) > 0
                  ? `Team-Belegung: ${(selectedEvent.teamSize || 0) - adminAddMemberDialog.freeSlots}/${selectedEvent.teamSize}`
                  : 'Belegung wird nach dem Hinzufügen aktualisiert.'}
              </div>
              <div style={{
                padding: '14px 16px',
                background: 'rgba(237,139,0,0.10)',
                border: '2px solid var(--dex-orange, #ed8b00)',
                borderRadius: 8,
                color: '#7a4a00',
                fontSize: '0.88rem',
                lineHeight: 1.5,
              }}>
                <div style={{ fontWeight: 700, marginBottom: 6 }}>
                  Vorab die Zustimmung des Mitglieds einholen
                </div>
                <div>
                  {'Mit dem Hinzufügen meldest du diese Person an. Sie erhält automatisch '}
                  {'eine Anmeldebestätigung per Mail, einen Outlook-Termin und sieht das '}
                  {'Event in „Meine Events". Bitte stelle sicher, dass die Person ihrer '}
                  {'Anmeldung '}<strong>vorher zugestimmt</strong>{' hat.'}
                </div>
              </div>
              <div>
                <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'block', marginBottom: 4 }}>
                  <span style={{ color: 'var(--dex-red)' }}>*</span> Person auswählen
                </label>
                {adminAddMemberPick ? (
                  <div style={{
                    display: 'inline-flex', alignItems: 'center', gap: 10,
                    padding: '6px 10px 6px 6px',
                    border: '1px solid var(--dex-gray-200)',
                    borderRadius: 'var(--dex-radius)',
                    background: 'var(--dex-gray-50, #f7f7f7)',
                    maxWidth: '100%',
                  }}>
                    <img
                      src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(adminAddMemberPick.email)}&size=S`}
                      alt={adminAddMemberPick.displayName}
                      onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                      style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                    />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{adminAddMemberPick.displayName}</div>
                      <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{adminAddMemberPick.email}</div>
                    </div>
                    <button
                      type="button"
                      onClick={() => { setAdminAddMemberPick(null); setAdminAddMemberQuery(''); setAdminAddMemberResults([]); }}
                      title="Auswahl entfernen"
                      style={{
                        background: 'var(--dex-gray-200)', border: 'none', color: 'var(--dex-gray-700)',
                        width: 22, height: 22, borderRadius: '50%', cursor: 'pointer',
                        fontSize: '0.9rem', lineHeight: 1,
                      }}
                    >×</button>
                  </div>
                ) : (
                  <div style={{ position: 'relative' }}>
                    <input
                      className="form-input"
                      value={adminAddMemberQuery}
                      placeholder="Name oder E-Mail eingeben…"
                      onChange={e => {
                        const val = e.target.value;
                        setAdminAddMemberQuery(val);
                        if (adminAddMemberQueryTimer.current) clearTimeout(adminAddMemberQueryTimer.current);
                        if (val.length >= 2) {
                          adminAddMemberQueryTimer.current = setTimeout(async () => {
                            setAdminAddMemberSearching(true);
                            try {
                              const res = await searchUsers(val);
                              setAdminAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
                            } catch { setAdminAddMemberResults([]); }
                            setAdminAddMemberSearching(false);
                          }, 300);
                        } else {
                          setAdminAddMemberResults([]);
                        }
                      }}
                    />
                    {(adminAddMemberResults.length > 0 || adminAddMemberSearching) && (
                      <div style={{
                        position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 10,
                        background: '#fff', border: '1px solid var(--dex-gray-200)',
                        borderRadius: 6, marginTop: 4, maxHeight: 220, overflowY: 'auto',
                        boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
                      }}>
                        {adminAddMemberSearching && (
                          <div style={{ padding: 10, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                            Suche…
                          </div>
                        )}
                        {adminAddMemberResults.map(r => (
                          <button
                            key={r.email}
                            type="button"
                            onClick={() => { setAdminAddMemberPick(r); setAdminAddMemberResults([]); setAdminAddMemberQuery(''); }}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8,
                              width: '100%', padding: '6px 10px', border: 'none',
                              background: '#fff', cursor: 'pointer', textAlign: 'left',
                            }}
                          >
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(r.email)}&size=S`}
                              alt={r.displayName}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                              style={{ width: 28, height: 28, borderRadius: '50%' }}
                            />
                            <div>
                              <div style={{ fontSize: '0.85rem', fontWeight: 600 }}>{r.displayName}</div>
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{r.email}</div>
                            </div>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                <input
                  type="checkbox"
                  checked={adminAddMemberConsent}
                  onChange={e => setAdminAddMemberConsent(e.target.checked)}
                  style={{ marginTop: 3 }}
                />
                <span>
                  <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                  Ich bestätige, dass die Person ihrer Anmeldung zugestimmt hat.
                </span>
              </label>
              {adminAddMemberError && (
                <div style={{ padding: 10, borderRadius: 6, background: 'rgba(220,38,38,0.10)', color: '#b91c1c', fontSize: '0.85rem' }}>
                  {adminAddMemberError}
                </div>
              )}
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={closeDlg}
                  disabled={adminAddMemberBusy}
                >
                  Abbrechen
                </button>
                <button
                  type="button"
                  className="btn btn-primary"
                  onClick={() => { submit().catch(() => { /* */ }); }}
                  disabled={!adminAddMemberPick || !adminAddMemberConsent || adminAddMemberBusy}
                >
                  {adminAddMemberBusy ? 'Wird hinzugefügt…' : 'Hinzufügen'}
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}
