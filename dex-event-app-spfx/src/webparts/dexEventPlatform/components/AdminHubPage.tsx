/**
 * v23.41: Admin-Hub — zentrale Anlaufstelle für Admin-Themen.
 *
 * Erreichbar über die „Admin"-Kachel der Startseite (nur Admins, auch im
 * Demo-Modus via originalIsAdmin). Bündelt:
 *  - Werkzeuge: Prozessübersicht, Rollenverwaltung, Einstellungen/Templates, Handbuch
 *  - Erklärung aller SharePoint-Listen (was macht welche Liste)
 *  - Archiv & Löschung (archivieren + alte Archiv-Einträge löschen)
 */
import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents, FixColumnsDetail } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { useDialog } from '../context/DialogContext';
import { useIsMobile } from '../utils/useIsMobile';
import { Settings, Users, Mail, Book, FileText, Trash2, Columns, BarChart3, Wrench, GraduationCap, Search, ChevronDown } from './Icons';
// v31.3: Gemeinsame Klassen des Organizer Centers (Kacheln, Aktions-Gruppen,
// Tabellen). Inline-Styles können kein `:hover` — bis v31.2 hing der Hover der
// Hub-Kacheln deshalb an onMouseEnter/Leave, und die Wartungs-Karten hatten
// gar keinen, obwohl ihr Knopf etwas tut.
import { cx, ensureDexUiStyles } from './dexUi';
import { InfoTooltip } from './InfoTooltip';
import { useTutorial } from './tutorial/TutorialGuide';
import { splitReleaseNote, RELEASE_NOTES, RELEASE_BEREICHE } from '../data/releaseNotes';
import { EventService, PermCleanupReport, OrphanScanResult } from '../services/EventService';
import { healAllEvents } from './admin/logic/healAllEvents';
import Modal from './Modal';
import { setCachedLogoBase64, setCachedOrbBase64 } from '../services/EmailTemplates';

// Erklärung aller DEX_*-Listen in Klartext (was tut die Liste, warum gibt es sie).
const LIST_DOCS: Array<{ name: string; de: string }> = [
  { name: 'DEX_Events', de: 'Das Herzstück: ein Eintrag pro Event (und pro Sub-Event). Enthält Titel, Datum, Ort, Sichtbarkeit, Kapazität, Kommunikations-Einstellungen usw.' },
  { name: 'DEX_Roles', de: 'Wer welche Rolle hat: User, Organizer oder Admin. Steuert, wer Events anlegen/verwalten darf.' },
  { name: 'DEX_Participants', de: 'Übergreifendes Register aller Personen, die sich jemals für ein Event angemeldet haben (für Statistik/KPIs).' },
  { name: 'DEX_Emails', de: 'Warteschlange für alle ausgehenden Mails (Bestätigungen, Erinnerungen, Wochenbericht …). Ein Hintergrundprozess verschickt sie und setzt den Status auf „Sent".' },
  { name: 'DEX_EmailTemplates', de: 'Die anpassbaren Mail-Vorlagen plus zentrale Konfiguration (Logo, Standardbild, KPI-Zähler).' },
  { name: 'DEX_Outlook', de: 'Warteschlange für Outlook-Kalendereinladungen: Teilnehmer ein-/ausladen, Termin aktualisieren/löschen.' },
  { name: 'DEX_OutlookLocks', de: 'Technische Sperre, damit zwei Kalender-Vorgänge desselben Events sich nicht in die Quere kommen.' },
  { name: 'DEX_IDReorder', de: 'Aufträge zum Neu-Durchnummerieren der Teilnehmer-IDs und zum Nachrücken von der Warteliste.' },
  { name: 'DEX_ChangeLog', de: 'Änderungsprotokoll: wer hat wann was getan (Anmeldung, Abmeldung, Check-in, Bearbeitung …) — der Nachweis pro Event.' },
  { name: 'DEX_AccessFix', de: 'Aufträge, damit bei stellvertretenden Anmeldungen die richtige Person ihre Anmeldung sieht und sich selbst abmelden kann.' },
  { name: 'DEX_TeamJoinRequests', de: 'Beitritts-Anfragen für Team-Anmeldungen, die der Team-Kapitän bestätigen muss.' },
  { name: 'DEX_OrganizerRequests', de: 'Anträge „Organizer werden" — Admins sehen offene Anträge in der App und geben sie frei.' },
  { name: 'DEX_WeeklyReports', de: 'Protokoll des automatischen Wochenberichts: hält fest, wann zuletzt ein Bericht verschickt wurde (damit er nicht doppelt kommt).' },
  { name: 'DEX_Archive', de: 'End-Ablage: hierhin werden alte Zeilen aus den Arbeitslisten verschoben, damit diese schlank bleiben. Einträge werden rund 1 Monat nach Ablauf des Events zum Löschen vorgeschlagen.' },
];

export default function AdminHubPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isAdmin, originalIsAdmin, siteUrl, roles } = useRoles();
  // v30.25: Tutorial-Einstieg für Admins (Header-Pille entfällt für sie).
  const { openTutorial } = useTutorial();
  const listUrl = (name: string): string => `${siteUrl}/Lists/${name}`;
  const { events: allEvents, getArchivableCount, runArchiveExpired, getDeletableArchiveCount, runDeleteOldArchive, fixAllEventColumns, repairAllOrganizerPermissions, restoreCustomFieldDescriptions, reseedDefaultEmailTemplates, maybeSendWeeklyReport, recomputeEventKpiOnly, getAllRegistrations } = useEvents();
  const { locale } = useLanguage();
  const { confirmDialog, showAlert } = useDialog();
  const isDe = locale === 'de';
  const isMobile = useIsMobile();
  const adminLike = isAdmin || originalIsAdmin;
  // v31.3: Das gemeinsame Stylesheet direkt beim Rendern sicherstellen (nicht im
  // Effect) — der Hub ist die erste Admin-Seite und rendert oft, bevor irgendein
  // Modal es injiziert hätte; im Effect gäbe es einen Moment ohne Klassen.
  ensureDexUiStyles();

  const [archTotal, setArchTotal] = React.useState(0);
  const [delTotal, setDelTotal] = React.useState(0);
  // v31.3: Ob die beiden Zähler überhaupt gelesen werden konnten. Solange nicht
  // (noch am Laden oder Lesefehler), ist die Zahl UNBEKANNT — die alte Anzeige
  // schrieb dort „0 Zeilen stehen an" und behauptete damit etwas über Daten, die
  // niemand gesehen hat (CLAUDE.md: ein Lesefehler ist keine Null). Rein für die
  // Anzeige; die Sperre der Knöpfe hängt weiter an der Zahl selbst.
  const [countsRead, setCountsRead] = React.useState<{ arch: boolean; del: boolean }>({ arch: false, del: false });
  const [busy, setBusy] = React.useState<'' | 'arch' | 'del' | 'fixcols' | 'perms' | 'restoredesc' | 'reseed' | 'weekly' | 'kpi'>('');
  // v26.63: zuletzt neu berechnete Events-Zahl (für die Erfolgs-Anzeige).
  const [kpiResult, setKpiResult] = React.useState<number | null>(null);
  // v28.26: Teilnehmer-Register bereinigen (Dubletten zusammenführen).
  const [regCleanBusy, setRegCleanBusy] = React.useState(false);
  const [regCleanResult, setRegCleanResult] = React.useState<string | null>(null);
  const [regCleanIsError, setRegCleanIsError] = React.useState(false);
  // v28.29: Fortschritt der Register-Bereinigung. Das Register hat mehrere
  // tausend Zeilen — Lesen und Zusammenführen dauern spürbar, und die Kachel
  // sah bis zum Ergebnis aus, als würde nichts passieren. `total: 0` = Phase
  // ohne bekannte Gesamtzahl (Lesen) → unbestimmter Balken.
  const [regCleanProgress, setRegCleanProgress] = React.useState<{ done: number; total: number; label: string } | null>(null);
  // v24.33: Fortschritt für das globale „Spalten fixen".
  const [fixProgress, setFixProgress] = React.useState<{ done: number; total: number; label: string } | null>(null);
  // v30.70: „Nachrücken & IDs für ALLE Events nachholen" — Sammel-Heilung
  // nach einem Ausfall des Nachrück-Flows. Fortschritt getrennt vom Ergebnis,
  // damit die Kachel während des Laufs sagt, bei welchem Event sie gerade ist.
  const [healBusy, setHealBusy] = React.useState(false);
  const [healProgress, setHealProgress] = React.useState<string | null>(null);
  const [healResult, setHealResult] = React.useState<{ text: string; isError: boolean } | null>(null);
  // v30.58: Befund des letzten Spalten-Laufs, je Event. `null` = noch nicht gelaufen.
  const [fixReport, setFixReport] = React.useState<FixColumnsDetail[] | null>(null);
  // v30.39: Fortschritt der Berechtigungs-Reparatur über alle Event-Bäume.
  const [permProgress, setPermProgress] = React.useState<{ done: number; total: number; label: string } | null>(null);
  const [restoreProgress, setRestoreProgress] = React.useState<{ done: number; total: number; label: string } | null>(null);
  const [restorePreview, setRestorePreview] = React.useState<Array<{ eventId: string; eventTitle: string; fields: Array<{ label: string; props: string[] }> }> | null>(null);
  // v26.81: Berechtigungen aufräumen — Modal mit Prüf-/Korrektur-Ablauf.
  const [permCleanupOpen, setPermCleanupOpen] = React.useState(false);
  const [permCleanupBusy, setPermCleanupBusy] = React.useState(false);
  const [permCleanupProgress, setPermCleanupProgress] = React.useState<{ msg: string; done: number; total: number } | null>(null);
  const [permCleanupReport, setPermCleanupReport] = React.useState<PermCleanupReport | null>(null);
  // v26.81: Verwaiste Subsites prüfen — Modal mit Bericht + Einzel-Löschung.
  const [orphanOpen, setOrphanOpen] = React.useState(false);
  const [orphanBusy, setOrphanBusy] = React.useState(false);
  const [orphanProgress, setOrphanProgress] = React.useState<{ msg: string; done: number; total: number } | null>(null);
  const [orphanResult, setOrphanResult] = React.useState<OrphanScanResult | null>(null);
  const [orphanDeleting, setOrphanDeleting] = React.useState<Record<string, boolean>>({});
  const [orphanDeleted, setOrphanDeleted] = React.useState<Record<string, boolean>>({});
  // v26.51: Logo & Branding — zentrales Default-Logo (Mails) + Logo-Video.
  // AdminHubPage hat sonst keinen SPFx-Kontext; Instanz wie in AdminPage erzeugen.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const spfxContext = (window as any).__dexSpfxContext;
  const eventServiceRef = React.useMemo(() => spfxContext ? new EventService(spfxContext) : null, []);
  // v26.58: getrennte Assets — logoBase64 = Deloitte-Logo (E-Mail-Kopfzeile,
  // weißer Schriftzug!), orbBase64 = das eigentliche DEX-Logo (bunter Ring).
  const [branding, setBranding] = React.useState<{ logoBase64: string; orbBase64: string; videoUrl: string; videoFileName: string } | null>(null);
  const [brandingBusy, setBrandingBusy] = React.useState<'' | 'logo' | 'orb' | 'video'>('');
  // Cache-Buster fürs Video: fester Dateiname + Overwrite ⇒ ohne ?ver würde der
  // Browser nach einem Tausch weiter das alte Video aus dem Cache zeigen.
  const [videoVer, setVideoVer] = React.useState(0);
  const logoInputRef = React.useRef<HTMLInputElement>(null);
  const orbInputRef = React.useRef<HTMLInputElement>(null);
  const videoInputRef = React.useRef<HTMLInputElement>(null);
  // v31.3: Welcher Aufklapper „Was diese Aktionen genau tun" offen ist ('' = keiner).
  // Die ausführlichen Erklärungen der Wartungs-Aktionen stehen dort — auf der
  // Kachel selbst steht nur eine Zeile Folge, sonst liest niemand mehr, was ein
  // Klick auslöst.
  const [openInfo, setOpenInfo] = React.useState<string>('');
  // Release-Notes: Volltext-Suche + Bereichs-Filter + Art-Filter.
  const [rnSearch, setRnSearch] = React.useState('');
  const [rnBereich, setRnBereich] = React.useState<string>('');
  const [rnType, setRnType] = React.useState<string>('');
  const filteredNotes = React.useMemo(() => {
    const q = rnSearch.trim().toLowerCase();
    return RELEASE_NOTES.filter(n => {
      if (rnBereich && n.bereich !== rnBereich) return false;
      if (rnType && n.type !== rnType) return false;
      if (!q) return true;
      return (
        n.text.toLowerCase().indexOf(q) >= 0 ||
        n.bereich.toLowerCase().indexOf(q) >= 0 ||
        ('v' + n.version).toLowerCase().indexOf(q) >= 0
      );
    });
  }, [rnSearch, rnBereich, rnType]);
  const fmtDate = (iso: string): string => {
    const d = new Date(iso);
    return isFinite(d.getTime())
      ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })
      : iso;
  };

  React.useEffect(() => {
    if (!adminLike) { navigate('start'); return; }
    // v31.3 (Nachzug): Erst zählen, wenn die Event-Liste da ist. `getArchivableCount`
    // antwortet bei leerer Liste mit 0 (archiveAndPurge.ts: `allIds.size === 0`) —
    // das ist die Aussage „noch nichts gelesen", nicht „nichts zu archivieren".
    // Ohne diese Sperre stünde eine ehrliche 0 neben einer erfundenen, und der
    // Zähler-Merker (countsRead) würde die erfundene als gelesen ausweisen.
    if (!allEvents.length) return;
    let cancelled = false;
    getArchivableCount().then(r => { if (!cancelled) { setArchTotal(r.total); setCountsRead(p => ({ ...p, arch: true })); } }).catch(() => { /* */ });
    getDeletableArchiveCount().then(n => { if (!cancelled) { setDelTotal(n); setCountsRead(p => ({ ...p, del: true })); } }).catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [adminLike, allEvents.length]);

  // v26.51: aktuelles Branding (Default-Logo + Logo-Video) einmalig laden.
  React.useEffect(() => {
    if (!eventServiceRef || !adminLike) return;
    let cancelled = false;
    eventServiceRef.getBranding().then(b => { if (!cancelled) setBranding(b); }).catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (!adminLike) return <div className="page-container" />;

  // v26.63: Startseiten-Zähler (KPI „Events"/„Teilnehmer") sofort neu berechnen.
  // Hintergrund: Der angezeigte Wert ist ein persistenter Zähler in _Config, der
  // von normalen Usern nur per ±1 bewegt wird und über die Zeit vom echten Stand
  // abweichen kann. Die automatische Voll-Neuberechnung läuft sonst nur 1×/Session
  // beim Admin-Boot im Hintergrund — dieser Button erzwingt sie jederzeit.
  const doRecomputeKpi = async (): Promise<void> => {
    if (busy) return;
    setBusy('kpi');
    setKpiResult(null);
    try {
      // v26.63: NUR die Events-Zahl neu berechnen — allein aus DEX_Events, ohne
      // den teuren Subsite-Scan. Der Teilnehmer-Zählerwert bleibt unberührt.
      const events = await recomputeEventKpiOnly();
      if (events === null) {
        showAlert(isDe ? 'Neuberechnung fehlgeschlagen — bitte später erneut versuchen.' : 'Recompute failed — please try again later.', { variant: 'error' });
        return;
      }
      setKpiResult(events);
      try { sessionStorage.setItem('dex-kpi-cache-refreshed', '1'); } catch { /* */ }
      showAlert(
        isDe
          ? `Events-Zähler neu berechnet: ${events} Events. Gespeichert — die Startseite zeigt den Wert ab dem nächsten Laden.`
          : `Events counter recomputed: ${events} events. Saved — the landing page shows it on next load.`,
        { variant: 'success' });
    } catch {
      showAlert(isDe ? 'Neuberechnung fehlgeschlagen.' : 'Recompute failed.', { variant: 'error' });
    } finally { setBusy(''); }
  };

  // v26.81: Berechtigungen aufräumen — Prüf-/Korrektur-Lauf über die gesamte
  // Site-Collection. apply=false = Dry-Run (Bericht ohne Änderung), apply=true =
  // Über-Freigaben entfernen + Element-Sicherheit korrigieren. Baut den
  // Rollen-Kontext (Admins/Organizer aus DEX_Roles, Organizer je Subsite aus
  // den Events) und ruft den Service.
  const runPermCleanup = async (apply: boolean): Promise<void> => {
    if (!eventServiceRef) { showAlert(isDe ? 'Kein SharePoint-Kontext verfügbar.' : 'No SharePoint context available.', { variant: 'error' }); return; }
    // Sicherheits-Guard: Ohne geladene DEX_Roles wäre die „erlaubt"-Liste leer
    // → alle Admins/Organizer würden fälschlich als Über-Freigabe erscheinen
    // (und beim Korrigieren entfernt). Lieber abbrechen und neu laden lassen.
    if (!roles || roles.length === 0) {
      showAlert(isDe
        ? 'Die Rollenliste ist noch nicht geladen. Bitte kurz warten oder die Seite neu laden und erneut versuchen — ohne geladene Rollen kann die Aufräumung nicht sicher laufen.'
        : 'The role list is not loaded yet. Please wait a moment or reload the page and try again — the cleanup cannot run safely without the roles.', { variant: 'error' });
      return;
    }
    const adminEmails = roles.filter(r => r.role === 'Admin' || r.role === 'IT-Admin').map(r => r.userEmail).filter(Boolean);
    // Sicherheitsnetz: JEDE in DEX_Roles gepflegte Person gilt als sanktioniert
    // und wird NIE als Über-Freigabe entfernt (reguläre User stehen nicht drin).
    const organizerEmails = roles.map(r => r.userEmail).filter(Boolean);
    // Organizer je Subsite (Haupt- + Co-Organizer). Schlüssel: absoluter
    // Subsite-URL UND server-relativer Pfad (ohne Trailing-Slash).
    const subsiteOrganizers: Record<string, string> = {};
    for (const ev of allEvents) {
      const su = (ev.subsiteUrl || '').trim();
      if (!su) continue;
      const orgs = [...(ev.organizerEmails || []), ...(ev.coOrganizerEmails || [])].map(e => (e || '').trim()).filter(Boolean);
      if (orgs.length === 0) continue;
      const keys: string[] = [su.toLowerCase().replace(/\/+$/, '')];
      try { keys.push(new URL(su).pathname.toLowerCase().replace(/\/+$/, '')); } catch { /* kein gültiger URL */ }
      for (const k of keys) {
        const existing = subsiteOrganizers[k] ? subsiteOrganizers[k].split(';') : [];
        subsiteOrganizers[k] = Array.from(new Set([...existing, ...orgs])).join(';');
      }
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const selfEmail = (spfxContext as any)?.pageContext?.user?.email || '';
    setPermCleanupBusy(true);
    setPermCleanupProgress({ msg: isDe ? 'Wird gestartet…' : 'Starting…', done: 0, total: 1 });
    if (apply) setPermCleanupReport(null);
    try {
      const report = await eventServiceRef.auditOrCleanupPermissions(
        apply,
        { adminEmails, organizerEmails, subsiteOrganizers, selfEmail },
        (msg, done, total) => setPermCleanupProgress({ msg, done, total }),
      );
      setPermCleanupReport(report);
    } catch (err) {
      showAlert((isDe ? 'Fehler bei der Berechtigungs-Prüfung: ' : 'Error during the permission check: ') + (err instanceof Error ? err.message : String(err)), { variant: 'error' });
    } finally {
      setPermCleanupBusy(false);
      setPermCleanupProgress(null);
    }
  };

  // v26.81: Verwaiste Subsites suchen (Analyse, ändert nichts).
  const runOrphanScan = async (): Promise<void> => {
    if (!eventServiceRef) { showAlert(isDe ? 'Kein SharePoint-Kontext verfügbar.' : 'No SharePoint context available.', { variant: 'error' }); return; }
    setOrphanBusy(true);
    setOrphanDeleted({});
    setOrphanProgress({ msg: isDe ? 'Wird gestartet…' : 'Starting…', done: 0, total: 1 });
    try {
      const res = await eventServiceRef.findOrphanSubsites((msg, done, total) => setOrphanProgress({ msg, done, total }));
      setOrphanResult(res);
    } catch (err) {
      showAlert((isDe ? 'Fehler bei der Subsite-Prüfung: ' : 'Error during the subsite check: ') + (err instanceof Error ? err.message : String(err)), { variant: 'error' });
    } finally {
      setOrphanBusy(false);
      setOrphanProgress(null);
    }
  };

  const deleteOrphan = async (url: string, label: string, participantCount: number): Promise<void> => {
    if (!eventServiceRef) return;
    const ok = await confirmDialog(isDe
      ? `Subsite „${label}" endgültig löschen?\n\nDie komplette Subsite inkl. aller Listen${participantCount > 0 ? ` (mit ${participantCount} Teilnehmer-Zeilen)` : ''} wird unwiderruflich entfernt. Das lässt sich NICHT rückgängig machen.`
      : `Permanently delete subsite „${label}"?\n\nThe entire subsite including all lists${participantCount > 0 ? ` (with ${participantCount} participant rows)` : ''} will be removed irreversibly. This CANNOT be undone.`,
      { danger: true, confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' });
    if (!ok) return;
    setOrphanDeleting(prev => ({ ...prev, [url]: true }));
    try {
      const done = await eventServiceRef.deleteSubsiteWeb(url);
      if (done) {
        setOrphanDeleted(prev => ({ ...prev, [url]: true }));
        showAlert(isDe ? `Subsite „${label}" gelöscht.` : `Subsite „${label}" deleted.`, { variant: 'success' });
      } else {
        showAlert(isDe ? `Subsite „${label}" konnte nicht gelöscht werden (evtl. Unter-Webs vorhanden oder fehlende Rechte).` : `Could not delete subsite „${label}" (maybe it has subwebs or you lack permissions).`, { variant: 'error' });
      }
    } catch (err) {
      showAlert((isDe ? 'Löschen fehlgeschlagen: ' : 'Deletion failed: ') + (err instanceof Error ? err.message : String(err)), { variant: 'error' });
    } finally {
      setOrphanDeleting(prev => ({ ...prev, [url]: false }));
    }
  };

  const doArchive = async (): Promise<void> => {
    if (busy || archTotal === 0) return;
    if (!(await confirmDialog(isDe ? `${archTotal} Zeilen abgelaufener/gelöschter Events ins Archiv verschieben?` : `Move ${archTotal} rows of expired/deleted events to the archive?`, { confirmLabel: isDe ? 'Archivieren' : 'Archive' }))) return;
    setBusy('arch');
    try {
      const r = await runArchiveExpired();
      showAlert(isDe ? `${r.archived} Zeilen archiviert${r.failed ? `, ${r.failed} fehlgeschlagen` : ''}.` : `${r.archived} rows archived${r.failed ? `, ${r.failed} failed` : ''}.`, { variant: r.failed ? 'error' : 'success' });
      try { const a = await getArchivableCount(); setArchTotal(a.total); const d = await getDeletableArchiveCount(); setDelTotal(d); } catch { /* */ }
    } catch { showAlert(isDe ? 'Archivierung fehlgeschlagen.' : 'Archiving failed.', { variant: 'error' }); }
    finally { setBusy(''); }
  };

  const doDelete = async (): Promise<void> => {
    if (busy || delTotal === 0) return;
    if (!(await confirmDialog(isDe ? `${delTotal} Archiv-Einträge (älter als 1 Monat) endgültig löschen?` : `Permanently delete ${delTotal} archive entries (older than 1 month)?`, { danger: true, confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' }))) return;
    setBusy('del');
    try {
      const r = await runDeleteOldArchive();
      showAlert(isDe ? `${r.deleted} alte Archiv-Einträge gelöscht${r.failed ? `, ${r.failed} fehlgeschlagen` : ''}.` : `${r.deleted} old archive entries deleted${r.failed ? `, ${r.failed} failed` : ''}.`, { variant: r.failed ? 'error' : 'success' });
      try { const d = await getDeletableArchiveCount(); setDelTotal(d); } catch { /* */ }
    } catch { showAlert(isDe ? 'Löschen fehlgeschlagen.' : 'Deletion failed.', { variant: 'error' }); }
    finally { setBusy(''); }
  };

  // v24.33: Globales „Spalten fixen" über ALLE Events (inkl. Sub-Events) +
  // Company-Backfill bestehender Teilnehmer — mit Fortschrittsanzeige.
  const doFixAllColumns = async (): Promise<void> => {
    if (busy) return;
    if (!(await confirmDialog(
      isDe
        ? 'Die Teilnehmerlisten ALLER Events (inkl. Sub-Events) prüfen, fehlende Spalten anlegen und die Unternehmenszugehörigkeit für bestehende Teilnehmer nachtragen? Je nach Anzahl der Events kann das einen Moment dauern.'
        : 'Check the participant lists of ALL events (incl. sub-events), add missing columns and backfill the company affiliation for existing attendees? This may take a moment depending on the number of events.',
      { confirmLabel: isDe ? 'Jetzt prüfen' : 'Check now' }
    ))) return;
    setBusy('fixcols');
    setFixProgress({ done: 0, total: 0, label: '' });
    try {
      const r = await fixAllEventColumns((done, total, label) => setFixProgress({ done, total, label }));
      const msg = r.anyChange
        ? (isDe
            ? `Fertig: ${r.lists} Teilnehmerlisten geprüft, ${r.columnsAdded} Spalte(n) ergänzt, ${r.backfilled} Unternehmens-Angabe(n) nachgetragen${r.errors ? `, ${r.errors} mit Fehler` : ''}.`
            : `Done: ${r.lists} lists checked, ${r.columnsAdded} column(s) added, ${r.backfilled} company value(s) backfilled${r.errors ? `, ${r.errors} with errors` : ''}.`)
        : (isDe
            ? `Alles war schon korrekt — ${r.lists} Teilnehmerlisten geprüft, nichts zu tun${r.errors ? ` (${r.errors} mit Fehler)` : ''}.`
            : `Everything was already fine — ${r.lists} lists checked, nothing to do${r.errors ? ` (${r.errors} with errors)` : ''}.`);
      // v30.58: Der Befund JE EVENT — das ist der eigentliche Zweck des Laufs.
      // Eine Zahl („12 Spalten ergänzt") beantwortet nicht, warum bei drei
      // Personen die Klammer-Anmeldung scheitert. Was nach dem Fix immer noch
      // fehlt, bringt jeden Insert zu Fall, in dem die Spalte vorkommt — bei
      // einer Klammer-Liste also jede Anmeldung, bei der die Person das
      // betreffende Hauptevent-Feld ausgefüllt hat.
      setFixReport(r.details || []);
      showAlert(msg, { variant: r.errors ? 'error' : 'success' });
    } catch { showAlert(isDe ? 'Spalten-Prüfung fehlgeschlagen.' : 'Column check failed.', { variant: 'error' }); }
    finally { setBusy(''); setFixProgress(null); }
  };

  // v30.39: Berechtigungen aller Organizer über ALLE Events nachziehen.
  const doRepairPermissions = async (): Promise<void> => {
    if (busy) return;
    if (!(await confirmDialog(
      isDe
        ? 'Für ALLE Events prüfen, ob jeder Organizer und Co-Organizer Zugriff auf die Teilnehmerlisten hat — auf dem Haupt-Event UND auf jedem Sub-Event? Fehlende Rechte werden ergänzt. Es wird nichts entzogen und nichts gelöscht. Je nach Anzahl der Events kann das einige Minuten dauern.'
        : 'Check for ALL events whether every organizer and co-organizer has access to the participant lists — on the main event AND on every sub-event? Missing permissions are added. Nothing is revoked and nothing is deleted. This may take a few minutes depending on the number of events.',
      { confirmLabel: isDe ? 'Jetzt prüfen' : 'Check now' }
    ))) return;
    setBusy('perms');
    setPermProgress({ done: 0, total: 0, label: '' });
    try {
      const r = await repairAllOrganizerPermissions((done, total, label) => setPermProgress({ done, total, label }));
      // Die Zahl der Zuweisungen ist bewusst NICHT als „so viele waren kaputt"
      // formuliert: SharePoint meldet bei addroleassignment nicht, ob das Recht
      // neu ist. Deshalb steht dort, was getan wurde, nicht was gefehlt hat.
      const un = r.unresolved.length
        ? (isDe
            ? ` ${r.unresolved.length} Adresse(n) konnten nicht zugeordnet werden: ${r.unresolved.join(', ')} — meist ehemalige Kolleg:innen.`
            : ` ${r.unresolved.length} address(es) could not be resolved: ${r.unresolved.join(', ')} — usually former colleagues.`)
        : '';
      showAlert(isDe
        ? `Fertig: ${r.trees} Event(s) mit insgesamt ${r.sites} Teilnehmerliste(n) durchlaufen, ${r.grants} Zuweisung(en) gesetzt${r.errors ? `, ${r.errors} mit Fehler` : ''}.${un}`
        : `Done: ${r.trees} event(s) with ${r.sites} participant list(s) processed, ${r.grants} assignment(s) applied${r.errors ? `, ${r.errors} with errors` : ''}.${un}`,
        { variant: r.errors ? 'error' : 'success' });
    } catch {
      showAlert(isDe ? 'Berechtigungs-Prüfung fehlgeschlagen.' : 'Permission check failed.', { variant: 'error' });
    } finally { setBusy(''); setPermProgress(null); }
  };

  // v26.13: Wiederherstellung versehentlich gelöschter Custom-Field-
  // Eigenschaften (Beschreibungen, Bedingungen, Mehrfachauswahl, EN-Varianten …)
  // aus der SharePoint-Versionshistorie.
  // v26.13: Trockenlauf — ermittelt OHNE zu schreiben, was wiederhergestellt würde.
  const doPreviewDescriptions = async (): Promise<void> => {
    if (busy) return;
    setBusy('restoredesc');
    setRestoreProgress({ done: 0, total: 0, label: '' });
    setRestorePreview(null);
    try {
      const r = await restoreCustomFieldDescriptions((done, total, label) => setRestoreProgress({ done, total, label }), true);
      setRestorePreview(r.details || []);
      showAlert(
        r.fieldsRestored > 0
          ? (isDe ? `Vorschau: ${r.fieldsRestored} Feld-Eigenschaft(en) in ${r.eventsChanged} Event(s) könnten wiederhergestellt werden (${r.events} geprüft). Es wurde NICHTS verändert.` : `Preview: ${r.fieldsRestored} field propert(ies) in ${r.eventsChanged} event(s) could be restored (${r.events} checked). NOTHING was changed.`)
          : (isDe ? `Vorschau: nichts wiederherzustellen (${r.events} Event(s) geprüft). Prüfe die Browser-Konsole (Filter „[DEX restore]") — dort steht pro Event, ob in der Versionshistorie überhaupt Beschreibungen vorhanden sind.` : `Preview: nothing to restore (${r.events} event(s) checked). Check the browser console (filter „[DEX restore]").`),
        { variant: 'success' });
    } catch { showAlert(isDe ? 'Vorschau fehlgeschlagen.' : 'Preview failed.', { variant: 'error' }); }
    finally { setBusy(''); setRestoreProgress(null); }
  };

  const doRestoreDescriptions = async (): Promise<void> => {
    if (busy) return;
    if (!(await confirmDialog(
      isDe
        ? 'Aus der Versionshistorie die zuvor gespeicherten Feld-Eigenschaften (Beschreibungen, Anzeige-Bedingungen, Mehrfachauswahl, Englisch-Varianten u.a.) für ALLE Events wiederherstellen? Es werden nur FEHLENDE Werte aufgefüllt — aktuelle Eingaben bleiben unangetastet.'
        : 'Restore previously saved field properties (descriptions, display conditions, multi-select, English variants, etc.) for ALL events from the version history? Only MISSING values are filled in — current entries are left untouched.',
      { confirmLabel: isDe ? 'Jetzt wiederherstellen' : 'Restore now' }
    ))) return;
    setBusy('restoredesc');
    setRestoreProgress({ done: 0, total: 0, label: '' });
    try {
      const r = await restoreCustomFieldDescriptions((done, total, label) => setRestoreProgress({ done, total, label }));
      const msg = r.fieldsRestored > 0
        ? (isDe
            ? `Fertig: ${r.fieldsRestored} Feld-Eigenschaft(en) in ${r.eventsChanged} Event(s) wiederhergestellt (${r.events} geprüft)${r.errors ? `, ${r.errors} mit Fehler` : ''}.`
            : `Done: restored ${r.fieldsRestored} field propert(ies) in ${r.eventsChanged} event(s) (${r.events} checked)${r.errors ? `, ${r.errors} with errors` : ''}.`)
        : (isDe
            ? `Nichts wiederherzustellen — ${r.events} Event(s) geprüft, alle Eigenschaften aktuell vorhanden${r.errors ? ` (${r.errors} mit Fehler)` : ''}.`
            : `Nothing to restore — ${r.events} event(s) checked, all properties present${r.errors ? ` (${r.errors} with errors)` : ''}.`);
      showAlert(msg, { variant: r.errors ? 'error' : 'success' });
    } catch { showAlert(isDe ? 'Wiederherstellung fehlgeschlagen.' : 'Restore failed.', { variant: 'error' }); }
    finally { setBusy(''); setRestoreProgress(null); }
  };

  // v24.97: Globale Mail-Werkzeuge — aus dem per-Event-Aktionsmenü hierher
  // verschoben (gehören als globale Admin-Aktion in den Admin-Hub).
  const doReseed = async (): Promise<void> => {
    if (busy) return;
    if (!(await confirmDialog(
      isDe
        ? 'Alle Standard-Mail-Vorlagen mit den eingebauten Texten aus dem aktuellen Stand der App überschreiben? Eigene Anpassungen an den Standard-Vorlagen gehen dabei verloren.'
        : 'Overwrite all default mail templates with the built-in texts from the current app version? Customizations to the standard templates will be lost.',
      { danger: true, confirmLabel: isDe ? 'Überschreiben' : 'Overwrite' }))) return;
    setBusy('reseed');
    try {
      const r = await reseedDefaultEmailTemplates();
      showAlert(
        r.failed > 0
          ? (isDe ? `Mit Fehlern: ${r.failed} Vorlage(n) fehlgeschlagen.` : `With errors: ${r.failed} template(s) failed.`)
          : (isDe ? `Erledigt: ${r.created} neu angelegt, ${r.updated} aktualisiert, ${r.skipped} unverändert.` : `Done: ${r.created} created, ${r.updated} updated, ${r.skipped} unchanged.`),
        { variant: r.failed > 0 ? 'error' : 'success' });
    } catch { showAlert(isDe ? 'Zurücksetzen fehlgeschlagen.' : 'Reset failed.', { variant: 'error' }); }
    finally { setBusy(''); }
  };

  const doWeekly = async (): Promise<void> => {
    if (busy) return;
    if (!(await confirmDialog(
      isDe
        ? 'Den Wochenbericht JETZT (sofort, ohne 7-Tage-Sperre) an alle Admins versenden? Nutze das nur zum Testen — der nächste reguläre Bericht zählt dann ab jetzt.'
        : 'Send the weekly report NOW (immediately, bypassing the 7-day lock) to all admins? Use this only for testing.',
      { confirmLabel: isDe ? 'Jetzt senden' : 'Send now' }))) return;
    setBusy('weekly');
    try {
      const r = await maybeSendWeeklyReport({ force: true });
      showAlert(
        r.sent
          ? (isDe ? `In die Warteschlange gelegt — eine Mail an ${r.admins} Admin(s). Sie wird in Kürze versendet.` : `Queued — one mail to ${r.admins} admin(s). It will be sent shortly.`)
          : (isDe ? 'Versand nicht möglich — keine Admins gefunden? Bitte Rollen prüfen.' : 'Sending not possible — no admins found? Please check the roles.'),
        { variant: r.sent ? 'success' : 'error' });
    } catch { showAlert(isDe ? 'Versand fehlgeschlagen.' : 'Sending failed.', { variant: 'error' }); }
    finally { setBusy(''); }
  };

  // v31.3: Der Ablauf der Register-Bereinigung stand als 170-zeiliger
  // onClick-Ausdruck mitten im JSX — unlesbar und beim Umbau der Kacheln nicht
  // verschiebbar. Reiner Umzug, kein Schritt verändert.
  const doRegistryCleanup = (): void => {
    (async () => {
      if (!eventServiceRef) return;
      setRegCleanBusy(true);
      setRegCleanResult(null);
      setRegCleanIsError(false);
      setRegCleanProgress({ done: 0, total: 0, label: isDe ? 'Teilnehmer-Register wird gelesen…' : 'Reading participant registry…' });
      try {
        const validNumbers = allEvents
          .map(e => e.eventNumber)
          .filter((n): n is number => typeof n === 'number' && n > 0);
        const info = await eventServiceRef.analyzeParticipantRegistry(validNumbers, loaded => {
          setRegCleanProgress({ done: loaded, total: 0, label: isDe ? `${loaded} Einträge gelesen…` : `${loaded} records read…` });
        });
        const orphanNote = info.orphanNumbers > 0
          ? (isDe
            ? ` ${info.orphanNumbers} Verweis(e) zeigen auf gelöschte Events.`
            : ` ${info.orphanNumbers} reference(s) point to deleted events.`)
          : '';
        // v29.0: Zweite Stufe — das Register gegen die
        // TEILNEHMERLISTEN abgleichen. Die Dubletten-Prüfung oben
        // sieht nur mehrfache Einträge und Verweise auf gelöschte
        // Events; ein Verweis auf ein EXISTIERENDES Event ohne Zeile
        // in dessen Liste fiel bisher durch. Genau der lässt „Meine
        // Events" eine Anmeldung zeigen, die es nicht gibt (v28.99).
        const cmp = await eventServiceRef.analyzeRegistryAgainstLists(
          allEvents.map(e => ({ eventNumber: e.eventNumber, title: e.title, subsiteUrl: e.subsiteUrl })),
          (done, total, title) => setRegCleanProgress({
            done, total,
            label: isDe
              ? `Teilnehmerlisten werden verglichen… ${done}/${total}${title ? ` — ${title}` : ''}`
              : `Comparing attendee lists… ${done}/${total}${title ? ` — ${title}` : ''}`,
          }),
        );
        setRegCleanProgress(null);
        /**
         * v29.4: Verweise auf GELÖSCHTE Events mitnehmen. Bis v29.3
         * wurden sie nur gezählt („wirkungslos, aber harmlos") — das
         * stimmt technisch, aber es sind personenbezogene Reste
         * gelöschter Events, und genau die soll das Register nicht
         * behalten. Die Event-Nummern kommen dafür STRIKT aus
         * DEX_Events (nicht aus der geladenen Event-Liste, die bei
         * einem Mapping-Fehler still Events auslässt) — sonst würde
         * ein Lesefehler gültige Verweise als verwaist ausweisen.
         */
        let orphanPairs: Array<{ email: string; eventNumber: number }> = [];
        let orphanReadError = '';
        try {
          orphanPairs = await eventServiceRef.collectOrphanRegistryNumbers(loaded =>
            setRegCleanProgress({
              done: loaded, total: 0,
              label: isDe ? `Verweise auf gelöschte Events werden gesucht… ${loaded}` : `Looking for references to deleted events… ${loaded}`,
            }));
        } catch (e) {
          orphanReadError = (e instanceof Error ? e.message : String(e || '')).slice(0, 200);
        }
        const orphanErrNote = orphanReadError
          ? (isDe
            ? `\n\nVerweise auf gelöschte Events konnten NICHT geprüft werden (${orphanReadError}) — sie bleiben unangetastet.`
            : `\n\nReferences to deleted events could NOT be checked (${orphanReadError}) — they stay untouched.`)
          : '';
        const staleAll: Array<{ email: string; eventNumber: number; title: string }> = [
          ...cmp.stale,
          ...orphanPairs.map(o => ({
            email: o.email, eventNumber: o.eventNumber,
            title: isDe ? `gelöschtes Event #${o.eventNumber}` : `deleted event #${o.eventNumber}`,
          })),
        ];
        const orphanFoundNote = orphanPairs.length > 0
          ? (isDe
            ? `\n\nEnthalten sind ${orphanPairs.length} Verweis(e) auf Events, die es in der Event-Liste NICHT MEHR GIBT (gelöschte Events). Sie laufen ins Leere und werden mit entfernt.`
            : `\n\nIncluded are ${orphanPairs.length} reference(s) to events that NO LONGER EXIST in the event list (deleted events). They point nowhere and are removed as well.`)
          : '';
        const staleNote = cmp.stale.length > 0
          ? (isDe
            ? ` ${cmp.stale.length} Verweis(e) zeigen auf ein Event, in dessen Teilnehmerliste die Person NICHT steht.`
            : ` ${staleAll.length} reference(s) point to an event whose attendee list does not contain the person.`)
          : '';
        // v29.1: Events, bei denen (fast) ALLE Verweise ins Leere zeigen,
        // sind kein Aufräum-Fall, sondern ein Hinweis darauf, dass
        // Register und Liste dort nicht vergleichbar sind. Sie werden
        // benannt statt stillschweigend bereinigt.
        const suspNote = cmp.suspiciousEvents.length > 0
          ? (isDe
            ? `\n\nNICHT bereinigt werden ${cmp.suspiciousEvents.length} Event(s), bei denen nahezu ALLE Verweise ins Leere zeigen — dort stimmt eher die Zuordnung nicht als hunderte Abmeldungen:\n`
              + cmp.suspiciousEvents.slice(0, 8).map(e => `• ${e.title || e.eventNumber}: ${e.missing} von ${e.referenced} Verweisen ohne Zeile, Liste hat ${e.rows} aktive Zeile(n)`).join('\n')
              + (cmp.suspiciousEvents.length > 8 ? `\n… und ${cmp.suspiciousEvents.length - 8} weitere` : '')
            : `\n\nNOT cleaned: ${cmp.suspiciousEvents.length} event(s) where nearly ALL references point nowhere — there the mapping is more likely wrong than hundreds of cancellations:\n`
              + cmp.suspiciousEvents.slice(0, 8).map(e => `• ${e.title || e.eventNumber}: ${e.missing} of ${e.referenced} references without a row, list has ${e.rows} active row(s)`).join('\n')
              + (cmp.suspiciousEvents.length > 8 ? `\n… and ${cmp.suspiciousEvents.length - 8} more` : ''))
          : '';
        // v29.3: Events, deren Teilnehmerliste NICHT MEHR EXISTIERT
        // (HTTP 404). Das ist der Regelfall hinter den meisten
        // verwaisten Verweisen: Das 3-Monats-Löschkonzept recycelt
        // die Subsite und lässt das Event-Item stehen — die
        // Register-Verweise darauf sind genau der Rückstand, den
        // diese Löschung hätte mitnehmen sollen. Sie werden bereinigt
        // (das ist kein Datenverlust, sondern der fehlende Rest der
        // Löschung), aber vorher benannt.
        const goneNote = cmp.deletedListEvents.length > 0
          ? (isDe
            ? `\n\nDavon entfallen ${cmp.deletedListEvents.reduce((n, e) => n + e.referenced, 0)} Verweis(e) auf ${cmp.deletedListEvents.length} Event(s), deren Teilnehmerliste es NICHT MEHR GIBT — typischerweise nach dem 3-Monats-Löschkonzept (Liste gelöscht, Event bleibt bestehen). Hier ist das Entfernen der Rest der Löschung, kein Datenverlust:\n`
              + cmp.deletedListEvents.slice(0, 8).map(e => `• ${e.title || e.eventNumber}: ${e.referenced} Verweis(e)`).join('\n')
              + (cmp.deletedListEvents.length > 8 ? `\n… und ${cmp.deletedListEvents.length - 8} weitere` : '')
            : `\n\nOf these, ${cmp.deletedListEvents.reduce((n, e) => n + e.referenced, 0)} reference(s) belong to ${cmp.deletedListEvents.length} event(s) whose attendee list NO LONGER EXISTS — typically after the 3-month retention deletion (list deleted, event kept). Removing them completes that deletion, it does not lose data:\n`
              + cmp.deletedListEvents.slice(0, 8).map(e => `• ${e.title || e.eventNumber}: ${e.referenced} reference(s)`).join('\n')
              + (cmp.deletedListEvents.length > 8 ? `\n… and ${cmp.deletedListEvents.length - 8} more` : ''))
          : '';
        const skipNote = cmp.skippedEvents > 0
          ? (isDe
            ? ` ${cmp.skippedEvents} Event(s) konnten nicht gelesen werden (z.B. fehlende Rechte oder Drosselung) und wurden übersprungen — ihre Verweise bleiben unangetastet.`
            : ` ${cmp.skippedEvents} event(s) could not be read (e.g. missing permissions or throttling) and were skipped — their references stay untouched.`)
          : '';
        if (info.duplicateGroups === 0 && staleAll.length === 0 && cmp.suspiciousEvents.length > 0) {
          setRegCleanIsError(true);
          setRegCleanResult(isDe
            ? `Keine Dubletten und keine einzeln verwaisten Verweise — ABER bei ${cmp.suspiciousEvents.length} Event(s) zeigen nahezu alle Verweise ins Leere. Das sieht nach einem Zuordnungsproblem aus und wurde deshalb NICHT bereinigt: ${cmp.suspiciousEvents.slice(0, 5).map(e => `${e.title || e.eventNumber} (${e.missing}/${e.referenced}, Liste ${e.rows})`).join('; ')}.${skipNote}`
            : `No duplicates and no individually orphaned references — BUT for ${cmp.suspiciousEvents.length} event(s) nearly all references point nowhere. That looks like a mapping problem and was NOT cleaned: ${cmp.suspiciousEvents.slice(0, 5).map(e => `${e.title || e.eventNumber} (${e.missing}/${e.referenced}, list ${e.rows})`).join('; ')}.${skipNote}`);
          setRegCleanProgress(null);
          setRegCleanBusy(false);
          return;
        }
        if (info.duplicateGroups === 0 && staleAll.length === 0) {
          setRegCleanResult(isDe
            ? `Alles sauber: keine Dubletten, und alle Verweise haben eine Zeile in der Teilnehmerliste (${info.total} Einträge, ${cmp.checkedEvents} Event(s) verglichen).${orphanNote}${skipNote}${info.noEmail > 0 ? ` ${info.noEmail} Eintrag/Einträge ohne E-Mail-Adresse.` : ''}`
            : `All clean: no duplicates, and every reference has a row in the attendee list (${info.total} records, ${cmp.checkedEvents} event(s) compared).${orphanNote}${skipNote}${info.noEmail > 0 ? ` ${info.noEmail} record(s) without an email address.` : ''}`);
          setRegCleanProgress(null);
          setRegCleanBusy(false);
          return;
        }
        if (info.duplicateGroups === 0) {
          // Nur verwaiste Verweise — einzeln nachfragen und entfernen.
          const examples = staleAll.slice(0, 5)
            .map(x => `• ${x.email} → ${x.title || x.eventNumber}`).join('\n');
          const okStale = await confirmDialog(isDe
            ? `${staleAll.length} Verweis(e) im Register zeigen auf ein Event, in dessen Teilnehmerliste die Person nicht steht — typischerweise eine Abmeldung, bei der das Nachziehen scheiterte, oder eine von Hand gelöschte Zeile.\n\n${examples}${staleAll.length > 5 ? `\n… und ${staleAll.length - 5} weitere` : ''}\n\nDiese Verweise jetzt entfernen? Die Einträge selbst bleiben mit ihren übrigen Events bestehen. An den Teilnehmerlisten wird nichts geändert.${orphanFoundNote}${goneNote}${suspNote}${orphanErrNote}${skipNote ? `\n\nHinweis:${skipNote}` : ''}`
            : `${staleAll.length} reference(s) point to an event whose attendee list does not contain the person — typically a cancellation whose registry update failed, or a manually deleted row.\n\n${examples}${staleAll.length > 5 ? `\n… and ${staleAll.length - 5} more` : ''}\n\nRemove these references now? The records themselves stay with their remaining events. Attendee lists are not touched.${orphanFoundNote}${goneNote}${suspNote}${orphanErrNote}${skipNote ? `\n\nNote:${skipNote}` : ''}`,
            { confirmLabel: isDe ? 'Verweise entfernen' : 'Remove references' });
          if (!okStale) { setRegCleanProgress(null); setRegCleanBusy(false); return; }
          setRegCleanProgress({ done: 0, total: 0, label: isDe ? 'Verweise werden entfernt…' : 'Removing references…' });
          const pr = await eventServiceRef.pruneStaleRegistryNumbers(staleAll, (done, total) =>
            setRegCleanProgress({ done, total, label: isDe ? 'Verweise werden entfernt…' : 'Removing references…' }));
          setRegCleanIsError(pr.failed > 0);
          setRegCleanResult(isDe
            ? `${pr.removed} Verweis(e) bei ${pr.updated} Person(en) entfernt${pr.failed > 0 ? `, ${pr.failed} fehlgeschlagen` : ''}.${orphanNote}`
            : `${pr.removed} reference(s) removed for ${pr.updated} person(s)${pr.failed > 0 ? `, ${pr.failed} failed` : ''}.${orphanNote}`);
          setRegCleanProgress(null);
          setRegCleanBusy(false);
          return;
        }
        setRegCleanProgress(null);
        const ok = await confirmDialog(isDe
          ? `${info.duplicateGroups} Person(en) haben mehrere Einträge im Teilnehmer-Register (${info.surplusRecords} überzählige Zeile(n) von ${info.total} insgesamt).\n\nJetzt zusammenführen? Je Person bleibt der älteste Eintrag und erhält ALLE Event-Nummern der Dubletten; die überzähligen Zeilen werden gelöscht. Anmeldungen gehen dabei nicht verloren.${orphanNote || staleNote ? `\n\nHinweis:${orphanNote}${staleNote} Die Verweise räumst du auf, indem du die Aktion nach dem Zusammenführen noch einmal startest.` : ''}`
          : `${info.duplicateGroups} person(s) have multiple records in the participant registry (${info.surplusRecords} surplus row(s) out of ${info.total} total).\n\nMerge now? Per person the oldest record is kept and receives ALL event numbers; the surplus rows are deleted. No registrations are lost.${orphanNote ? `\n\nNote:${orphanNote}` : ''}`,
          { confirmLabel: isDe ? 'Zusammenführen' : 'Merge' });
        if (!ok) { setRegCleanProgress(null); setRegCleanBusy(false); return; }
        setRegCleanProgress({ done: 0, total: info.duplicateGroups, label: isDe ? 'Einträge werden zusammengeführt…' : 'Merging records…' });
        const r = await eventServiceRef.mergeDuplicateParticipants(
          (done, total) => setRegCleanProgress({
            done, total,
            label: isDe ? 'Einträge werden zusammengeführt…' : 'Merging records…',
          }),
          loaded => setRegCleanProgress({ done: loaded, total: 0, label: isDe ? `${loaded} Einträge gelesen…` : `${loaded} records read…` }),
        );
        setRegCleanIsError(r.failed > 0);
        setRegCleanResult(isDe
          ? `${r.groups} Person(en) zusammengeführt, ${r.deleted} überzählige Zeile(n) entfernt${r.failed > 0 ? `, ${r.failed} fehlgeschlagen` : ''}.${staleNote ? `${staleNote} Starte die Aktion noch einmal, um sie zu entfernen.` : ''}`
          : `${r.groups} person(s) merged, ${r.deleted} surplus row(s) removed${r.failed > 0 ? `, ${r.failed} failed` : ''}.`);
      } catch (err) {
        setRegCleanIsError(true);
        setRegCleanResult((isDe ? 'Fehler: ' : 'Error: ') + (err instanceof Error ? err.message : String(err || '')).slice(0, 300));
      }
      setRegCleanProgress(null);
      setRegCleanBusy(false);
    })().catch(() => { /* */ });
  };

  // v31.3: Ebenfalls aus dem JSX gezogen — Sammel-Heilung „Nachrücken & IDs".
  const doHealAll = (): void => {
    if (!eventServiceRef) return;
    setHealBusy(true);
    setHealResult(null);
    healAllEvents({
      svc: eventServiceRef, allEvents, isDe, getAllRegistrations, confirmDialog,
      onProgress: setHealProgress,
    }).then(r => {
      if (!r.cancelled) setHealResult({ text: r.text, isError: r.isError });
    }).catch(err => {
      setHealResult({ text: (isDe ? 'Fehler: ' : 'Error: ') + (err instanceof Error ? err.message : String(err || '')).slice(0, 300), isError: true });
    }).then(() => { setHealProgress(null); setHealBusy(false); });
  };

  // v26.51: Logo & Branding — Download des aktuellen Logos (Data-URI → Datei).
  const doDownloadLogo = (): void => {
    if (!branding || !branding.logoBase64) return;
    const a = document.createElement('a');
    a.href = branding.logoBase64;
    a.download = 'Deloitte_Logo.png';
    a.click();
  };

  // v26.58: Download des DEX-Logos (Orb).
  const doDownloadOrb = (): void => {
    if (!branding || !branding.orbBase64) return;
    const a = document.createElement('a');
    a.href = branding.orbBase64;
    a.download = 'DEX_Logo_Orb.png';
    a.click();
  };

  // v26.51: Neues Default-Logo (PNG) hochladen — gilt für alle NEU versendeten Mails.
  const onLogoFileChosen = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file || !eventServiceRef) return;
    // Klare Regel: > 1,5 MB ablehnen — das Logo wird in jede Mail eingebettet.
    if (file.size > 1.5 * 1024 * 1024) {
      showAlert(
        isDe
          ? 'Die Datei ist größer als 1,5 MB. Das Logo wird in jede E-Mail eingebettet — bitte das PNG vorher komprimieren (z. B. auf unter 1,5 MB) und erneut hochladen.'
          : 'The file is larger than 1.5 MB. The logo is embedded into every email — please compress the PNG first (e.g. to below 1.5 MB) and upload again.',
        { variant: 'error' });
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUri = typeof reader.result === 'string' ? reader.result : '';
      if (dataUri.indexOf('data:image/png') !== 0) {
        showAlert(isDe ? 'Bitte eine PNG-Datei auswählen.' : 'Please choose a PNG file.', { variant: 'error' });
        return;
      }
      setBrandingBusy('logo');
      eventServiceRef.saveBrandingLogo(dataUri)
        .then(ok => {
          if (ok) {
            setCachedLogoBase64(dataUri); // Mails derselben Sitzung sofort mit neuem Logo
            setBranding(prev => prev ? { ...prev, logoBase64: dataUri } : { logoBase64: dataUri, orbBase64: '', videoUrl: '', videoFileName: '' });
            showAlert(isDe ? 'Neues Logo gespeichert — alle neuen Mails nutzen es ab sofort.' : 'New logo saved — all new emails will use it from now on.', { variant: 'success' });
          } else {
            showAlert(isDe ? 'Logo konnte nicht gespeichert werden.' : 'The logo could not be saved.', { variant: 'error' });
          }
        })
        .catch(() => showAlert(isDe ? 'Logo konnte nicht gespeichert werden.' : 'The logo could not be saved.', { variant: 'error' }))
        .finally(() => setBrandingBusy(''));
    };
    reader.readAsDataURL(file);
  };

  // v26.58: Neues DEX-Logo (Orb, PNG) hochladen — Default-Mail-Bild für Events
  // ohne eigenes Bild ({{ORB_URL}}-Fallback) + zentrale Download-Quelle.
  const onOrbFileChosen = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file || !eventServiceRef) return;
    if (file.size > 1.5 * 1024 * 1024) {
      showAlert(
        isDe
          ? 'Die Datei ist größer als 1,5 MB. Das Bild wird in Mails eingebettet — bitte das PNG vorher komprimieren und erneut hochladen.'
          : 'The file is larger than 1.5 MB. The image is embedded into emails — please compress the PNG first and upload again.',
        { variant: 'error' });
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUri = typeof reader.result === 'string' ? reader.result : '';
      if (dataUri.indexOf('data:image/png') !== 0) {
        showAlert(isDe ? 'Bitte eine PNG-Datei auswählen.' : 'Please choose a PNG file.', { variant: 'error' });
        return;
      }
      setBrandingBusy('orb');
      eventServiceRef.saveBrandingOrb(dataUri)
        .then(ok => {
          if (ok) {
            setCachedOrbBase64(dataUri);
            setBranding(prev => prev ? { ...prev, orbBase64: dataUri } : { logoBase64: '', orbBase64: dataUri, videoUrl: '', videoFileName: '' });
            showAlert(isDe ? 'Neues DEX-Logo gespeichert.' : 'New DEX logo saved.', { variant: 'success' });
          } else {
            showAlert(isDe ? 'DEX-Logo konnte nicht gespeichert werden.' : 'The DEX logo could not be saved.', { variant: 'error' });
          }
        })
        .catch(() => showAlert(isDe ? 'DEX-Logo konnte nicht gespeichert werden.' : 'The DEX logo could not be saved.', { variant: 'error' }))
        .finally(() => setBrandingBusy(''));
    };
    reader.readAsDataURL(file);
  };

  // v26.51: Neues Logo-Video hochladen (SiteAssets, fester Name, Overwrite).
  const onVideoFileChosen = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file || !eventServiceRef) return;
    if (file.size > 100 * 1024 * 1024) {
      showAlert(
        isDe
          ? 'Die Datei ist größer als 100 MB — bitte das Video vorher komprimieren und erneut hochladen.'
          : 'The file is larger than 100 MB — please compress the video first and upload again.',
        { variant: 'error' });
      return;
    }
    setBrandingBusy('video');
    eventServiceRef.uploadBrandingVideo(file)
      .then(url => {
        if (url) {
          setBranding(prev => prev ? { ...prev, videoUrl: url, videoFileName: file.name } : { logoBase64: '', orbBase64: '', videoUrl: url, videoFileName: file.name });
          setVideoVer(v => v + 1);
          showAlert(isDe ? 'Neues Logo-Video hochgeladen.' : 'New logo video uploaded.', { variant: 'success' });
        } else {
          showAlert(isDe ? 'Video konnte nicht hochgeladen werden.' : 'The video could not be uploaded.', { variant: 'error' });
        }
      })
      .catch(() => showAlert(isDe ? 'Video konnte nicht hochgeladen werden.' : 'The video could not be uploaded.', { variant: 'error' }))
      .finally(() => setBrandingBusy(''));
  };

  const tools: Array<{ icon: React.ReactNode; title: string; desc: string; onClick: () => void }> = [
    { icon: <Users size={20} />, title: isDe ? 'Organizer Center' : 'Organizer center', desc: isDe ? 'Teilnehmer, Prozesse, Audit-Log, SharePoint-Liste — alle Event-Werkzeuge pro Event.' : 'Attendees, processes, audit log, SharePoint list — all per-event tools.', onClick: () => navigate('admin') },
    { icon: <Settings size={20} />, title: isDe ? 'Prozessübersicht' : 'Process overview', desc: isDe ? 'Wie die Abläufe in DEX funktionieren — verständlich erklärt.' : 'How the DEX processes work — explained simply.', onClick: () => navigate('flowcharts') },
    { icon: <FileText size={20} />, title: isDe ? 'Architektur' : 'Architecture', desc: isDe ? 'App, SharePoint-Listen, Power-Automate-Flows und M365-Dienste — mit PDF-Export.' : 'App, SharePoint lists, Power Automate flows and M365 services — with PDF export.', onClick: () => navigate('architecture') },
    { icon: <Users size={20} />, title: isDe ? 'Rollenverwaltung' : 'Role management', desc: isDe ? 'User, Organizer und Admins zuweisen oder entfernen.' : 'Assign or remove users, organizers and admins.', onClick: () => navigate('settings') },
    { icon: <Columns size={20} />, title: isDe ? 'Rollenmatrix' : 'Role matrix', desc: isDe ? 'Übersicht: wer welche Rechte hat (User, Organizer, Admin).' : 'Overview: who has which permissions (user, organizer, admin).', onClick: () => navigate('role-matrix') },
    { icon: <Mail size={20} />, title: isDe ? 'Mail-Vorlagen' : 'Mail templates', desc: isDe ? 'Globale Standard-Mails (Anmeldung, Warteliste, Abmeldung …) bearbeiten — mit Live-Vorschau.' : 'Edit the global default emails (registration, waitlist, cancellation …) — with live preview.', onClick: () => navigate('email-templates') },
    { icon: <BarChart3 size={20} />, title: isDe ? 'Statistik-Archiv' : 'Statistics archive', desc: isDe ? 'Kennzahlen gelöschter Teilnehmerlisten — welches Event, wann, von wem, wie viele (ohne Personendaten).' : 'KPIs of deleted participant lists — which event, when, by whom, how many (no personal data).', onClick: () => navigate('stats-archive') },
    { icon: <Book size={20} />, title: isDe ? 'Handbuch' : 'Manual', desc: isDe ? 'Ausführliche Anleitung zu allen Funktionen.' : 'Detailed guide for all features.', onClick: () => navigate('manual') },
    // v29.24: Onepager für die Einführungsveranstaltung — Zyklus, Einsatzbereich (Venn), Rollen, Kernfunktionen.
    { icon: <GraduationCap size={20} />, title: isDe ? 'Einführungs-Onepager' : 'Introduction one-pager', desc: isDe ? 'DEX auf einen Blick: Event-Zyklus, Einsatzbereich, Rollen, Kernfunktionen.' : 'DEX at a glance: event cycle, scope, roles, core functions.', onClick: () => navigate('intro-onepager') },
    // v30.25: Ersatz-Einstieg für Admins — die „Neu hier?"-Pille auf der
    // Startseite wird ihnen nicht mehr angeboten (s. Header). Das Tutorial
    // startet auf der Landing Page, deshalb erst dorthin navigieren und die
    // Tour anschließend anstoßen.
    { icon: <GraduationCap size={20} />, title: isDe ? 'DEX Tutorial' : 'DEX tutorial', desc: isDe ? 'Die geführte Tour durch die App starten — praktisch, um sie neuen Kolleg:innen zu zeigen.' : 'Start the guided tour through the app — handy for showing it to new colleagues.', onClick: () => { navigate('start'); window.setTimeout(() => { try { openTutorial(); } catch { /* Tour nicht verfügbar */ } }, 350); } },
  ];

  /**
   * v31.3: EINE Aktions-Kachel (`dex-ui-action`): Symbol, Titel, eine Zeile
   * Folge, optional eine Zähler-Pille. Vorher war jede Wartungs-Aktion eine
   * Karte mit vier Zeilen Erklärung und einem Knopf ganz unten — bei zehn
   * Karten liest man weder das eine noch das andere. Die ausführliche
   * Erklärung steht jetzt im Aufklapper unter der Gruppe.
   */
  const hubAction = (
    key: string,
    icon: React.ReactNode,
    title: string,
    desc: string,
    onClick: () => void,
    opts?: { disabled?: boolean; danger?: boolean; badge?: React.ReactNode },
  ): React.ReactElement => (
    <button
      key={key}
      type="button"
      className={cx('dex-ui-action', opts && opts.danger && 'dex-ui-action--danger')}
      disabled={!!(opts && opts.disabled)}
      onClick={onClick}
    >
      <span className="dex-ui-action-icon" aria-hidden="true">{icon}</span>
      <span className="dex-ui-action-body">
        <span className="dex-ui-action-title">{title}</span>
        <span className="dex-ui-action-desc">{desc}</span>
      </span>
      {opts && opts.badge ? <span className="dex-ui-action-badge">{opts.badge}</span> : null}
    </button>
  );

  // v31.3: Fortschrittsbalken einer laufenden Aktion — bis v31.2 stand derselbe
  // Balken sechsmal als Inline-Style im JSX.
  const runProgress = (pct: number, label: string): React.ReactElement => (
    <div>
      <div className="dex-ui-progress"><div className="dex-ui-progress-bar" style={{ width: `${pct}%` }} /></div>
      <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 4, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{label}</div>
    </div>
  );

  // v31.3: Eine Zeile im Aufklapper „Was diese Aktionen genau tun".
  const actionDetail = (title: string, text: string): React.ReactElement => (
    <p key={title} style={{ margin: '0 0 10px', fontSize: '0.8rem', lineHeight: 1.55, color: 'var(--dex-gray-600)' }}>
      <strong style={{ color: 'var(--dex-gray-800)' }}>{title}</strong> — {text}
    </p>
  );

  /**
   * v31.3: Die Fußzeilen der beiden Dialoge. `Modal` rendert sie seit v31.2
   * selbst (Trennlinie oben, rechtsbündig) — vorher baute jeder Dialog seine
   * eigene Knopfzeile mitten in den Inhalt. Primär-Knopf rechts außen.
   */
  const permCleanupFooter = permCleanupBusy ? undefined : (permCleanupReport ? (
    !permCleanupReport.apply && (permCleanupReport.strayWriteFound > 0 || permCleanupReport.ilsIssues > 0) ? (
      <>
        <button type="button" className="btn btn-outline dex-ui-btn-sm" onClick={() => { setPermCleanupOpen(false); setPermCleanupReport(null); }}>
          {isDe ? 'Schließen' : 'Close'}
        </button>
        <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => { void runPermCleanup(false); }}>
          {isDe ? 'Erneut prüfen' : 'Re-check'}
        </button>
        <button
          type="button"
          className="btn btn-primary dex-ui-btn-sm"
          onClick={() => {
            (async () => {
              const ok = await confirmDialog(isDe
                ? `${permCleanupReport.strayWriteFound} Über-Freigabe(n) entfernen und ${permCleanupReport.ilsIssues} Element-Sicherheit(en) korrigieren?\n\nLeserechte und Gruppen-Berechtigungen bleiben erhalten. Der Vorgang kann je nach Größe einige Minuten dauern.`
                : `Remove ${permCleanupReport.strayWriteFound} over-grant(s) and fix ${permCleanupReport.ilsIssues} item-security setting(s)?\n\nRead access and group permissions are preserved. Depending on size this can take a few minutes.`,
                { confirmLabel: isDe ? 'Jetzt korrigieren' : 'Fix now' });
              if (ok) await runPermCleanup(true);
            })().catch(() => { /* */ });
          }}
        >
          {isDe ? 'Jetzt korrigieren' : 'Fix now'}
        </button>
      </>
    ) : (
      <>
        <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => { void runPermCleanup(false); }}>
          {isDe ? 'Erneut prüfen' : 'Re-check'}
        </button>
        <button type="button" className="btn btn-primary dex-ui-btn-sm" onClick={() => { setPermCleanupOpen(false); setPermCleanupReport(null); }}>
          {isDe ? 'Schließen' : 'Close'}
        </button>
      </>
    )
  ) : (
    <>
      <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setPermCleanupOpen(false)}>
        {isDe ? 'Abbrechen' : 'Cancel'}
      </button>
      <button type="button" className="btn btn-primary dex-ui-btn-sm" onClick={() => { void runPermCleanup(false); }}>
        {isDe ? 'Prüfen (ohne Änderung)' : 'Check (no changes)'}
      </button>
    </>
  ));

  const orphanFooter = orphanBusy ? undefined : (orphanResult ? (
    <>
      <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => { void runOrphanScan(); }}>
        {isDe ? 'Erneut prüfen' : 'Re-check'}
      </button>
      <button type="button" className="btn btn-primary dex-ui-btn-sm" onClick={() => { setOrphanOpen(false); setOrphanResult(null); }}>
        {isDe ? 'Schließen' : 'Close'}
      </button>
    </>
  ) : (
    <>
      <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setOrphanOpen(false)}>
        {isDe ? 'Abbrechen' : 'Cancel'}
      </button>
      <button type="button" className="btn btn-primary dex-ui-btn-sm" onClick={() => { void runOrphanScan(); }}>
        {isDe ? 'Jetzt prüfen' : 'Check now'}
      </button>
    </>
  ));

  // v31.3: Aufklapper-Knopf je Aktions-Gruppe (nur einer offen).
  const detailToggle = (key: string): React.ReactElement => (
    <button
      type="button"
      className={cx('dex-ui-disclosure', openInfo === key && 'is-open')}
      onClick={() => setOpenInfo(openInfo === key ? '' : key)}
    >
      <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
      {isDe ? 'Was diese Aktionen genau tun' : 'What these actions do in detail'}
    </button>
  );

  return (
    <div className="page-container">
      {/* v31.3: Seitenkopf wie im Organizer Center — Titel und eine Zeile, was
          die Seite ist. */}
      <div className="dex-ui-page-head">
        <div style={{ minWidth: 0 }}>
          <h1 className="dex-ui-page-head-title" style={{ marginTop: 0 }}>{isDe ? 'Admin' : 'Admin'}</h1>
          <div className="dex-ui-page-head-meta">
            {isDe
              ? 'Zentrale Anlaufstelle für Admin-Themen: Werkzeuge öffnen, Daten prüfen, reparieren, archivieren.'
              : 'Central place for admin topics: open tools, check data, repair, archive.'}
          </div>
        </div>
      </div>

      {/* v31.3: Werkzeuge sind reine Navigation — die ganze Kachel ist der
          Knopf, deshalb ein <button> mit Hover aus der Klasse statt einer
          <div>-Karte mit onMouseEnter-Optik. */}
      <section className="dex-ui-section">
        <h2 className="dex-ui-section-title">{isDe ? 'Werkzeuge' : 'Tools'}</h2>
        <div className="dex-ui-grid-auto">
          {tools.map((t, i) => (
            <button key={i} type="button" className="dex-ui-tile" onClick={t.onClick}>
              <span className="dex-ui-tile-icon" aria-hidden="true">{t.icon}</span>
              <span className="dex-ui-tile-title">{t.title}</span>
              <span className="dex-ui-tile-desc">{t.desc}</span>
            </button>
          ))}
        </div>
      </section>

      {/* v31.3: Wartung als Aktions-Gruppen statt zwölf Karten — geordnet nach
          Absicht: erst was nur prüft, dann was repariert, dann Mails, zuletzt
          das Unwiderrufliche (Archiv & Löschen). Jede Kachel trägt eine Zeile
          Folge; die ausführliche Begründung steht im Aufklapper darunter. */}
      <section className="dex-ui-section">
        <h2 className="dex-ui-section-title">{isDe ? 'Wartung & Daten' : 'Maintenance & data'}</h2>
        <p className="dex-ui-section-desc">
          {isDe
            ? 'Jede Aktion sagt in einer Zeile, was sie tut. Vor jeder Änderung kommt eine Rückfrage; Fortschritt und Ergebnis stehen unter der Gruppe.'
            : 'Every action says in one line what it does. Every change asks first; progress and result appear below its group.'}
        </p>

        {/* Prüfen — findet und zeigt, ändert von sich aus nichts. */}
        <div className="dex-ui-action-group">
          <div className="dex-ui-action-group-title">{isDe ? 'Prüfen — ändert nichts' : 'Check — changes nothing'}</div>
          <div className="dex-ui-action-grid">
            {hubAction('permcheck', <Wrench size={16} />,
              isDe ? 'Berechtigungen prüfen' : 'Check permissions',
              isDe ? 'Findet Einzel-Freigaben mit Schreibrecht im ganzen SharePoint — erst der Bericht, korrigiert wird nur auf Knopfdruck.' : 'Finds individual write grants across the whole SharePoint — report first, fixes only on request.',
              () => { setPermCleanupReport(null); setPermCleanupOpen(true); },
              { disabled: busy !== '' })}
            {hubAction('orphanscan', <Search size={16} />,
              isDe ? 'Verwaiste Subsites suchen' : 'Find orphan subsites',
              isDe ? 'Zeigt Subsites, die zu keinem Event mehr gehören — löschen kannst du danach jede einzeln.' : 'Shows subsites that no longer belong to any event — you then delete each one individually.',
              () => { setOrphanResult(null); setOrphanOpen(true); },
              { disabled: busy !== '' })}
          </div>
          {detailToggle('check')}
          {openInfo === 'check' && (
            <div className="dex-ui-disclosure-body">
              {actionDetail(
                isDe ? 'Berechtigungen prüfen' : 'Check permissions',
                isDe
                  ? 'Prüft die gesamte SharePoint-Seite (Hauptseite, alle Listen/Bibliotheken und alle Event-Subsites) auf manuelle Einzel-Freigaben, die einzelnen Personen mehr Rechte geben als im Berechtigungskonzept vorgesehen (z.B. Schreib-/Vollzugriff auf ganze Listen). Erst kommt ein Bericht ohne Änderung, danach kannst du die Über-Freigaben mit einem Klick entfernen. Leserechte bleiben immer erhalten (auch für internationale Kolleg:innen); Schreiben ist danach nur über die Gruppen und für Admins/Organizer möglich.'
                  : 'Scans the whole SharePoint site (main site, all lists/libraries and every event subsite) for manual individual grants that give single people more rights than the permission concept allows (e.g. write/full control on entire lists). First a report without changes, then you can remove the over-grants with one click. Read access always stays (including for international colleagues); writing is afterwards only via the groups and for admins/organizers.')}
              {actionDetail(
                isDe ? 'Verwaiste Subsites suchen' : 'Find orphan subsites',
                isDe
                  ? 'Findet Event-Subsites, die noch existieren, aber zu KEINEM Event mehr gehören — z.B. Test-Subsites, deren Event bereits gelöscht wurde. Zeigt pro Rest, ob eine Teilnehmerliste (und wie viele Zeilen) vorhanden ist. Anschließend kannst du jeden Rest einzeln und bewusst löschen.'
                  : 'Finds event subsites that still exist but no longer belong to any event — e.g. test subsites whose event was already deleted. Shows per orphan whether a participant list exists (and how many rows). You can then delete each orphan individually and deliberately.')}
            </div>
          )}
        </div>

        {/* Reparieren — schreibt, fragt aber vorher. */}
        <div className="dex-ui-action-group">
          <div className="dex-ui-action-group-title">{isDe ? 'Reparieren — fragt vor jeder Änderung' : 'Repair — asks before every change'}</div>
          <div className="dex-ui-action-grid">
            {hubAction('fixcols', <Columns size={16} />,
              isDe ? 'Spalten fixen (alle Events)' : 'Fix columns (all events)',
              busy === 'fixcols'
                ? (isDe ? 'Wird geprüft…' : 'Checking…')
                : (isDe ? 'Legt fehlende Spalten an und trägt die Unternehmenszugehörigkeit nach — ohne sie scheitert jede Anmeldung mit diesem Feld.' : 'Adds missing columns and backfills the company — without them every registration using that field fails.'),
              () => { void doFixAllColumns(); },
              { disabled: busy !== '' })}
            {/* v30.39: Organizer-Berechtigungen über alle Events. Der Einzel-Fix im
                Organizer Center (v30.37) hilft nur dem, der von dem Problem schon
                weiß — und sichtbar wird es erst, wenn jemand vor einer leeren
                Teilnehmerliste steht. Diese Aktion geht über den Bestand. */}
            {hubAction('perms', <Users size={16} />,
              isDe ? 'Organizer-Rechte prüfen (alle Events)' : 'Check organizer permissions (all events)',
              busy === 'perms'
                ? (isDe ? 'Wird geprüft…' : 'Checking…')
                : (isDe ? 'Ergänzt fehlende Rechte auf jeder Teilnehmerliste — auch auf jedem Sub-Event. Es wird nichts entzogen.' : 'Adds missing rights on every participant list — including every sub-event. Nothing is revoked.'),
              () => { void doRepairPermissions(); },
              { disabled: busy !== '' })}
            {/* v28.26: Teilnehmer-Register bereinigen — Dubletten (mehrere Einträge
                zur selben E-Mail) zusammenführen. Sie entstehen, wenn der Lookup vor
                dem Schreiben scheitert (siehe v28.25): Ab da landen Anmeldungen mal
                im einen, mal im anderen Eintrag, und „Meine Events" zeigt je nach
                Treffer nur einen Teil der Events. Site-weit, daher hier statt im
                Organizer Center. */}
            {hubAction('regclean', <Users size={16} />,
              isDe ? 'Teilnehmer-Register bereinigen' : 'Clean up participant registry',
              regCleanBusy
                ? (isDe ? 'Wird geprüft…' : 'Checking…')
                : (isDe ? 'Führt Dubletten zusammen und entfernt Verweise ins Leere — Anmeldungen gehen dabei nicht verloren.' : 'Merges duplicates and removes dead references — no registrations are lost.'),
              doRegistryCleanup,
              { disabled: regCleanBusy || busy !== '' || !eventServiceRef })}
            {/* v30.70: Sammel-Heilung nach einem Ausfall des Flows
                DEX_IDReorder_TeilnehmerIDs (02.09.2026). Erst planen und im Dialog
                zeigen, wer nachrückt — dann ausführen. Logik in
                admin/logic/healAllEvents.ts. */}
            {hubAction('healall', <Wrench size={16} />,
              isDe ? 'Nachrücken & IDs nachholen' : 'Catch up promotions & IDs',
              healBusy
                ? (isDe ? 'Heilung läuft…' : 'Healing…')
                : (isDe ? 'Rückt überall dort nach, wo Plätze frei sind, und nummeriert lückenhafte IDs neu — Vorschau vor dem Versand.' : 'Promotes wherever seats are free and renumbers IDs with gaps — preview before anything is sent.'),
              doHealAll,
              { disabled: busy !== '' || healBusy || !eventServiceRef })}
            {/* v26.13: Feld-Eigenschaften aus der Versionshistorie — erst der
                Trockenlauf, dann das Auffüllen. */}
            {hubAction('restorepreview', <FileText size={16} />,
              isDe ? 'Feld-Beschreibungen: Vorschau' : 'Field descriptions: preview',
              isDe ? 'Zeigt aus der Versionshistorie, was zurückkäme — es wird nichts geändert.' : 'Shows from the version history what would come back — nothing is changed.',
              () => { void doPreviewDescriptions(); },
              { disabled: busy !== '' })}
            {hubAction('restoredesc', <FileText size={16} />,
              isDe ? 'Feld-Beschreibungen wiederherstellen' : 'Restore field descriptions',
              busy === 'restoredesc'
                ? (isDe ? 'Läuft…' : 'Running…')
                : (isDe ? 'Füllt nur FEHLENDE Eigenschaften der Abfragefelder auf — aktuelle Eingaben bleiben unangetastet.' : 'Fills in only MISSING properties of the form fields — current entries stay untouched.'),
              () => { void doRestoreDescriptions(); },
              { disabled: busy !== '' })}
            {/* v26.63: Startseiten-Zähler (Events/Teilnehmer) neu berechnen. */}
            {hubAction('kpi', <BarChart3 size={16} />,
              isDe ? 'Startseiten-Zähler neu berechnen' : 'Recompute landing-page counter',
              busy === 'kpi'
                ? (isDe ? 'Wird berechnet…' : 'Computing…')
                : (isDe ? 'Rechnet die Kennzahl „Events" frisch aus der Event-Liste — der Teilnehmer-Zähler bleibt unverändert.' : 'Recomputes the „Events" KPI straight from the event list — the attendee counter stays unchanged.'),
              () => { void doRecomputeKpi(); },
              { disabled: busy !== '' })}
          </div>

          {/* Fortschritt und Ergebnis der Gruppe — direkt unter den Kacheln,
              damit man nicht sucht, wo die Antwort steht. */}
          <div className="dex-ui-stack" style={{ marginTop: 10 }}>
            {busy === 'fixcols' && fixProgress && runProgress(
              fixProgress.total > 0 ? Math.round((fixProgress.done / fixProgress.total) * 100) : 0,
              `${isDe ? 'Spalten fixen' : 'Fix columns'} · ${fixProgress.done}/${fixProgress.total}${fixProgress.label ? ` · ${fixProgress.label}` : ''}`)}
            {busy === 'perms' && permProgress && runProgress(
              permProgress.total > 0 ? Math.round((permProgress.done / permProgress.total) * 100) : 0,
              `${isDe ? 'Organizer-Rechte' : 'Organizer permissions'} · ${permProgress.done}/${permProgress.total}${permProgress.label ? ` · ${permProgress.label}` : ''}`)}
            {/* Lese-Phase (total = 0): Gesamtzahl ist noch unbekannt, der Balken
                waechst mit den gelesenen Zeilen (2000 je Seite) und bleibt unter
                90 %, damit er nie faelschlich „fertig" wirkt. */}
            {regCleanBusy && regCleanProgress && runProgress(
              regCleanProgress.total > 0
                ? Math.min(100, Math.round((regCleanProgress.done / regCleanProgress.total) * 100))
                : Math.max(6, Math.min(90, Math.round(regCleanProgress.done / 100))),
              `${regCleanProgress.total > 0 ? `${regCleanProgress.done}/${regCleanProgress.total} · ` : ''}${regCleanProgress.label}`)}
            {busy === 'restoredesc' && restoreProgress && runProgress(
              restoreProgress.total > 0 ? Math.round((restoreProgress.done / restoreProgress.total) * 100) : 0,
              `${isDe ? 'Feld-Beschreibungen' : 'Field descriptions'} · ${restoreProgress.done}/${restoreProgress.total}${restoreProgress.label ? ` · ${restoreProgress.label}` : ''}`)}
            {healBusy && healProgress && (
              <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {healProgress}
              </div>
            )}
            {kpiResult !== null && (
              <div className="dex-ui-callout dex-ui-callout--success">
                <span className="dex-ui-callout-icon"><BarChart3 size={15} /></span>
                <span>{isDe ? `Ergebnis: ${kpiResult} Events` : `Result: ${kpiResult} events`}</span>
              </div>
            )}
            {regCleanResult && (
              <div className={cx('dex-ui-callout', regCleanIsError ? 'dex-ui-callout--danger' : 'dex-ui-callout--success')}>
                <span className="dex-ui-callout-icon"><Users size={15} /></span>
                <span>{regCleanResult}</span>
              </div>
            )}
            {healResult && !healBusy && (
              <div className={cx('dex-ui-callout', healResult.isError ? 'dex-ui-callout--danger' : 'dex-ui-callout--success')}>
                <span className="dex-ui-callout-icon"><Wrench size={15} /></span>
                <span>{healResult.text}</span>
              </div>
            )}
            {/* v30.58: Der Befund je Event. Das ist der Punkt des Laufs — eine
                Zahl allein beantwortet nicht, warum eine Anmeldung scheitert. */}
            {fixReport && (
              <div
                className={cx('dex-ui-callout', fixReport.some(d => d.stillMissing.length > 0 || d.listMissing || d.error) ? 'dex-ui-callout--danger' : 'dex-ui-callout--success')}
                style={{ display: 'block', maxHeight: 300, overflowY: 'auto' }}
              >
                {fixReport.length === 0 ? (
                  <strong>
                    {isDe
                      ? 'Kein Befund — auf allen Teilnehmerlisten sind alle Spalten der Abfragefelder vorhanden.'
                      : 'Nothing found — every participant list has all columns of the form fields.'}
                  </strong>
                ) : (
                  <>
                    <strong>
                      {isDe
                        ? `Befund (${fixReport.length} ${fixReport.length === 1 ? 'Event' : 'Events'})`
                        : `Findings (${fixReport.length} ${fixReport.length === 1 ? 'event' : 'events'})`}
                    </strong>
                    <p style={{ margin: '4px 0 8px' }}>
                      {isDe
                        ? <>Fehlt auf einer Liste die Spalte zu einem Abfragefeld, lehnt SharePoint die <strong>gesamte Anmeldung</strong> ab — aber nur bei den Personen, die dieses Feld ausfüllen. Deshalb sieht es aus wie ein Einzelfall.</>
                        : <>If a list is missing the column of a form field, SharePoint rejects the <strong>entire registration</strong> — but only for the people who fill in that field. That is why it looks like a one-off.</>}
                    </p>
                    {fixReport.map(d => (
                      <div key={d.eventId} style={{ marginBottom: 8, paddingBottom: 8, borderBottom: '1px solid var(--dex-gray-200)' }}>
                        <div style={{ fontWeight: 600 }}>
                          {d.eventTitle}{' '}
                          <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                            ({d.isParent ? (isDe ? 'Haupt-/Klammer-Event' : 'Main/umbrella event') : (isDe ? 'Sub-Event' : 'Sub-event')})
                          </span>
                        </div>
                        {d.listMissing && (
                          <div style={{ color: 'var(--dex-red, #da291c)' }}>
                            {isDe ? 'Teilnehmerliste existiert nicht (mehr).' : 'The participant list does not exist (any more).'}
                          </div>
                        )}
                        {d.fixedColumns.length > 0 && (
                          <div style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Ergänzt: ' : 'Added: '}{d.fixedColumns.join(', ')}</div>
                        )}
                        {d.stillMissing.length > 0 && (
                          <div style={{ color: 'var(--dex-red, #da291c)' }}>
                            <strong>{isDe ? 'Fehlt weiterhin:' : 'Still missing:'}</strong> {d.stillMissing.join(', ')}
                          </div>
                        )}
                        {d.error && (
                          <div style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>{isDe ? 'Hinweis: ' : 'Note: '}{d.error}</div>
                        )}
                      </div>
                    ))}
                  </>
                )}
              </div>
            )}
            {restorePreview && (
              <div className="dex-ui-card dex-ui-card--soft" style={{ maxHeight: 220, overflowY: 'auto', fontSize: '0.78rem' }}>
                <div style={{ fontWeight: 700, color: 'var(--dex-gray-800)', marginBottom: 6 }}>
                  {isDe ? 'Vorschau: Feld-Beschreibungen' : 'Preview: field descriptions'}
                </div>
                {restorePreview.length === 0 ? (
                  <div style={{ color: 'var(--dex-gray-500)' }}>{isDe ? 'Keine wiederherstellbaren Eigenschaften gefunden.' : 'No restorable properties found.'}</div>
                ) : restorePreview.map(ev => (
                  <div key={ev.eventId} style={{ marginBottom: 8 }}>
                    <div style={{ fontWeight: 700, color: 'var(--dex-gray-800)' }}>{ev.eventTitle}</div>
                    {ev.fields.map((f, i) => (
                      <div key={i} style={{ color: 'var(--dex-gray-600)', paddingLeft: 8 }}>• {f.label}: {f.props.join(', ')}</div>
                    ))}
                  </div>
                ))}
              </div>
            )}
          </div>

          {detailToggle('repair')}
          {openInfo === 'repair' && (
            <div className="dex-ui-disclosure-body">
              {actionDetail(
                isDe ? 'Spalten fixen (alle Events)' : 'Fix columns (all events)',
                isDe
                  ? 'Prüft die Teilnehmerlisten ALLER Events inkl. Sub-Events, legt fehlende Spalten an (z.B. „Unternehmen") und trägt die Unternehmenszugehörigkeit für bestehende Teilnehmer nach.'
                  : 'Checks the participant lists of ALL events incl. sub-events, adds missing columns (e.g. „Company") and backfills the company affiliation for existing attendees.')}
              {actionDetail(
                isDe ? 'Organizer-Rechte prüfen (alle Events)' : 'Check organizer permissions (all events)',
                isDe
                  ? 'Stellt sicher, dass jeder Organizer und Co-Organizer die Teilnehmerliste seiner Events lesen darf — auf dem Haupt-Event UND auf jedem Sub-Event. Bis v30.36 wurde die Berechtigung beim Speichern nur auf dem Haupt-Event gesetzt: Wer nachträglich als Organizer dazukam, sah bei einem Event mit mehreren Terminen überall 0 Teilnehmer, obwohl Anmeldungen vorlagen. Fehlende Rechte werden ergänzt, es wird nichts entzogen und nichts gelöscht.'
                  : 'Ensures every organizer and co-organizer can read the participant list of their events — on the main event AND on every sub-event. Until v30.36 permissions were set on the main event only: anyone added as organizer later saw 0 participants everywhere on multi-date events although registrations existed. Missing permissions are added; nothing is revoked or deleted.')}
              {actionDetail(
                isDe ? 'Teilnehmer-Register bereinigen' : 'Clean up participant registry',
                isDe
                  ? 'Sucht in der zentralen Teilnehmer-Übersicht (DEX_Participants) nach Dubletten — mehrere Einträge zur selben E-Mail — und führt sie zusammen: Der älteste Eintrag bleibt und bekommt ALLE Event-Nummern, die überzähligen Zeilen werden gelöscht. Es geht nichts verloren; wessen Anmeldungen auf zwei Einträge verteilt waren, sieht danach wieder alle Events unter „Meine Events". Prüft zuerst und fragt vor dem Zusammenführen nach. Danach gleicht die Aktion das Register gegen die TEILNEHMERLISTEN ab: Zeigt ein Verweis auf ein Event, in dessen Liste die Person gar nicht steht — typischerweise eine Abmeldung, bei der das Nachziehen scheiterte —, wird er auf Rückfrage entfernt. Genau solche Verweise lassen „Meine Events" eine Anmeldung anzeigen, die es nicht gibt.'
                  : 'Searches the central participant registry (DEX_Participants) for duplicates — several records for the same email — and merges them: the oldest record is kept and receives ALL event numbers, the surplus rows are deleted. Nothing is lost; anyone whose registrations were split across two records sees all their events in „My events" again. Checks first and asks before merging.')}
              {actionDetail(
                isDe ? 'Nachrücken & IDs nachholen' : 'Catch up promotions & IDs',
                isDe
                  ? 'Für den Fall, dass der Nachrück-Flow ausgefallen war: Prüft alle aktiven Events, lässt überall dort nachrücken, wo Plätze frei sind und Leute warten, nummeriert lückenhafte TeilnehmerIDs neu und gleicht alle Platzzähler ab. Zeigt VOR dem Ausführen, wer in welchem Event nachrücken würde — erst nach Bestätigung gehen Mails und Einladungen raus. Events ohne Vollzugriff und Events mit gemeinsamer Warteliste bei geteilten Gruppen werden namentlich ausgewiesen statt falsch gerechnet.'
                  : 'For when the promotion flow was down: checks all active events, promotes wherever seats are free and people are waiting, renumbers participant IDs with gaps and reconciles all seat counters. Shows BEFORE running who would move up in which event — emails and invites only go out after confirmation. Events without full access, and split-group events with a shared waitlist, are listed by name instead of being miscounted.')}
              {actionDetail(
                isDe ? 'Feld-Beschreibungen (Vorschau & Wiederherstellen)' : 'Field descriptions (preview & restore)',
                isDe
                  ? 'Stellt versehentlich verlorene Eigenschaften der Abfrage-/Auswahlfelder (Beschreibungen, Anzeige-Bedingungen, Mehrfachauswahl, Englisch-Varianten u.a.) aus der SharePoint-Versionshistorie wieder her. Füllt nur FEHLENDE Werte auf — aktuelle Eingaben bleiben erhalten. Die Vorschau zeigt dasselbe Ergebnis, ohne etwas zu schreiben.'
                  : 'Restores accidentally lost properties of the form/selection fields (descriptions, display conditions, multi-select, English variants, etc.) from the SharePoint version history. Only fills in MISSING values — current entries are preserved. The preview shows the same result without writing anything.')}
              {actionDetail(
                isDe ? 'Startseiten-Zähler neu berechnen' : 'Recompute landing-page counter',
                isDe
                  ? 'Berechnet die Kennzahl „Events" auf der Startseite frisch aus der Event-Liste (ohne Entwürfe, abgesagte und Sub-Events; abgelaufene zählen mit) — schnell, ohne die Teilnehmerlisten zu scannen. Der angezeigte Wert ist ein gespeicherter Zähler, der sonst nur einmal pro Admin-Sitzung automatisch aktualisiert wird. Der Teilnehmer-Zähler bleibt unverändert.'
                  : 'Recomputes the „Events" KPI on the landing page straight from the event list (excluding drafts, cancelled and sub-events; past ones count) — fast, without scanning the participant lists. The shown value is a stored counter that otherwise only refreshes once per admin session. The attendee counter is left unchanged.')}
            </div>
          )}
        </div>

        {/* v24.97: E-Mails & Berichte — globale Mail-Werkzeuge (Reseed + Wochenbericht) */}
        <div className="dex-ui-action-group">
          <div className="dex-ui-action-group-title">{isDe ? 'E-Mails & Berichte' : 'Emails & reports'}</div>
          <div className="dex-ui-action-grid">
            {hubAction('weekly', <Mail size={16} />,
              isDe ? 'Wochenbericht jetzt senden' : 'Send weekly report now',
              busy === 'weekly'
                ? (isDe ? 'Wird gesendet…' : 'Sending…')
                : (isDe ? 'Legt den Bericht sofort für alle Admins in die Mail-Warteschlange — nur zum Testen.' : 'Queues the report for all admins right away — for testing only.'),
              () => { void doWeekly(); },
              { disabled: busy !== '' })}
            {hubAction('reseed', <Mail size={16} />,
              isDe ? 'Default-Mail-Vorlagen zurücksetzen' : 'Reset default mail templates',
              busy === 'reseed'
                ? (isDe ? 'Wird zurückgesetzt…' : 'Resetting…')
                : (isDe ? 'Überschreibt alle Standard-Vorlagen mit den eingebauten Texten — eigene Anpassungen gehen verloren.' : 'Overwrites all default templates with the built-in texts — customizations are lost.'),
              () => { void doReseed(); },
              { disabled: busy !== '', danger: true })}
          </div>
          {detailToggle('mail')}
          {openInfo === 'mail' && (
            <div className="dex-ui-disclosure-body">
              {actionDetail(
                isDe ? 'Wochenbericht jetzt senden' : 'Send weekly report now',
                isDe
                  ? 'Löst den wöchentlichen Admin-Bericht sofort aus (überspringt die 7-Tage-Sperre) und legt ihn für alle Admins in die Mail-Warteschlange. Nur zum Testen.'
                  : 'Triggers the weekly admin report immediately (bypassing the 7-day lock) and queues it for all admins. For testing only.')}
              {actionDetail(
                isDe ? 'Default-Mail-Vorlagen zurücksetzen' : 'Reset default mail templates',
                isDe
                  ? 'Überschreibt alle Standard-Mail-Vorlagen (Anmeldung, Warteliste, Abmeldung, Nachrücken …) mit den eingebauten Texten aus dem aktuellen Stand der App. Achtung: eigene Anpassungen an den Standard-Vorlagen gehen verloren.'
                  : 'Overwrites all default mail templates with the built-in texts from the current app version. Note: customizations to the standard templates are lost.')}
            </div>
          )}
        </div>

        {/* Zuletzt das Unwiderrufliche — Löschen ist die einzige Aktion, die
            sich nicht zurücknehmen lässt. */}
        <div className="dex-ui-action-group">
          <div className="dex-ui-action-group-title">{isDe ? 'Archiv & Löschen' : 'Archive & deletion'}</div>
          <div className="dex-ui-action-grid">
            {hubAction('archive', <FileText size={16} />,
              isDe ? 'Jetzt archivieren' : 'Archive now',
              busy === 'arch'
                ? (isDe ? 'Wird archiviert…' : 'Archiving…')
                : (!countsRead.arch
                  ? (isDe ? 'Die Anzahl ist noch nicht gelesen — bis dahin ist unbekannt, ob etwas ansteht.' : 'The count has not been read yet — until then it is unknown whether anything is due.')
                  : (archTotal === 0
                    ? (isDe ? 'Zurzeit steht nichts zur Archivierung an.' : 'Nothing is waiting to be archived right now.')
                    : (isDe ? 'Verschiebt Zeilen abgelaufener und gelöschter Events aus den Arbeitslisten ins Archiv.' : 'Moves rows of expired and deleted events out of the working lists into the archive.'))),
              () => { void doArchive(); },
              {
                disabled: busy !== '' || archTotal === 0,
                badge: <span className="dex-ui-pill dex-ui-pill--gray">{countsRead.arch ? (isDe ? `${archTotal} Zeilen` : `${archTotal} rows`) : '–'}</span>,
              })}
            {hubAction('delarchive', <Trash2 size={16} />,
              isDe ? 'Alte Archiv-Einträge löschen' : 'Delete old archive entries',
              busy === 'del'
                ? (isDe ? 'Wird gelöscht…' : 'Deleting…')
                : (!countsRead.del
                  ? (isDe ? 'Die Anzahl ist noch nicht gelesen — bis dahin ist unbekannt, ob etwas zu löschen ist.' : 'The count has not been read yet — until then it is unknown whether anything is due.')
                  : (delTotal === 0
                    ? (isDe ? 'Zurzeit ist kein Eintrag älter als 1 Monat.' : 'No entry is older than 1 month right now.')
                    : (isDe ? 'Löscht Archiv-Einträge endgültig, deren Event länger als 1 Monat vorbei ist.' : 'Permanently deletes archive entries whose event is more than 1 month past.'))),
              () => { void doDelete(); },
              {
                disabled: busy !== '' || delTotal === 0,
                danger: true,
                badge: <span className="dex-ui-pill dex-ui-pill--gray">{countsRead.del ? (isDe ? `${delTotal} Einträge` : `${delTotal} entries`) : '–'}</span>,
              })}
          </div>
          {detailToggle('archive')}
          {openInfo === 'archive' && (
            <div className="dex-ui-disclosure-body">
              {actionDetail(
                isDe ? 'Jetzt archivieren' : 'Archive now',
                isDe
                  ? 'Zeilen aus abgelaufenen oder gelöschten Events wandern aus den Arbeitslisten ins Archiv, damit diese schlank bleiben. Die Pille auf der Kachel sagt, wie viele Zeilen gerade anstehen.'
                  : 'Rows from expired or deleted events move out of the working lists into the archive so they stay lean. The pill on the tile says how many rows are ready.')}
              {actionDetail(
                isDe ? 'Alte Archiv-Einträge löschen' : 'Delete old archive entries',
                isDe
                  ? 'Archiv-Einträge, deren Event länger als 1 Monat vorbei ist, werden endgültig gelöscht — das lässt sich nicht rückgängig machen. Die Pille auf der Kachel sagt, wie viele Einträge das gerade sind.'
                  : 'Archive entries whose event is more than 1 month past are deleted permanently — this cannot be undone. The pill on the tile says how many entries that currently is.')}
            </div>
          )}
        </div>
      </section>

      {/* v30.34: Der Einbettungs-Test aus v30.32 ist wieder raus — die
          Frage, die er beantworten sollte, ist beantwortet: Eine per
          iframe eingebettete Fremdseite bekommt in der SharePoint-App
          KEINE Kamera. Permissions Policy reicht Rechte nur abwaerts
          durch; was das WebView der Host-App nicht hat, kann kein
          Rahmen darin gewinnen. Damit ist auch eine selbst gehostete
          Scanner-Seite als Weg erledigt — nachgemessen 31.08.2026.
          Kein Werkzeug stehen lassen, das nur eine erledigte Frage
          stellt; der Befund steht in CLAUDE.md und den Release Notes. */}

      {/* v26.51: Logo & Branding — Default-Mail-Logo tauschen/herunterladen + Logo-Video.
          v31.3: hinter die Wartung gerückt — Branding stellt man einmal ein,
          Wartung braucht man laufend. */}
      {adminLike && (
        <section className="dex-ui-section">
          <h2 className="dex-ui-section-title">{isDe ? 'Logo & Branding' : 'Logo & branding'}</h2>
          <div className="dex-ui-grid-auto">
            {/* v26.58: DEX-Logo = der bunte Orb-Ring (vorher zeigte diese Karte
                fälschlich das Deloitte-Mail-Logo, dessen weißer Schriftzug auf
                weißem Grund unsichtbar war — „nur ein grüner Punkt"). */}
            <div className="dex-ui-card">
              <div className="dex-ui-card-head" style={{ marginBottom: 10 }}>
                <h3 className="dex-ui-card-head-title"><FileText size={16} /> {isDe ? 'DEX-Logo (Orb, PNG)' : 'DEX logo (orb, PNG)'}</h3>
                <InfoTooltip text={isDe
                  ? 'Der bunte DEX-Ring. Neue Mails nutzen nach einem Tausch automatisch das neue Bild. Zugleich die zentrale Download-Quelle, z. B. für Intranet-Artikel.'
                  : 'The colourful DEX ring. New emails automatically use the new image after a swap. Also the central download source, e.g. for intranet articles.'} />
              </div>
              {branding && branding.orbBase64 ? (
                <img src={branding.orbBase64} alt="DEX Orb" style={{ maxWidth: '100%', maxHeight: 90, display: 'block', margin: '0 auto 10px', background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 8 }} />
              ) : (
                <p className="dex-ui-muted" style={{ fontStyle: 'italic', margin: '0 0 10px' }}>{isDe ? 'Noch kein DEX-Logo hinterlegt.' : 'No DEX logo stored yet.'}</p>
              )}
              <p className="dex-ui-help" style={{ marginTop: 0 }}>
                {isDe
                  ? 'Standard-Bild in Mails von Events ohne eigenes Event-Bild.'
                  : 'Default image in emails of events without their own image.'}
              </p>
              <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                <button className="btn btn-outline dex-ui-btn-sm" disabled={brandingBusy !== '' || !eventServiceRef} onClick={() => { if (orbInputRef.current) orbInputRef.current.click(); }}>
                  {brandingBusy === 'orb' ? (isDe ? 'Wird gespeichert…' : 'Saving…') : (isDe ? 'Neues hochladen (PNG)' : 'Upload new (PNG)')}
                </button>
                <button className="btn btn-secondary dex-ui-btn-sm" disabled={!branding || !branding.orbBase64} onClick={doDownloadOrb}>
                  {isDe ? 'Herunterladen' : 'Download'}
                </button>
              </div>
              <input ref={orbInputRef} type="file" accept="image/png" style={{ display: 'none' }} onChange={onOrbFileChosen} />
            </div>
            <div className="dex-ui-card">
              <div className="dex-ui-card-head" style={{ marginBottom: 10 }}>
                <h3 className="dex-ui-card-head-title"><FileText size={16} /> {isDe ? 'Deloitte-Logo (E-Mail-Kopfzeile)' : 'Deloitte logo (email header)'}</h3>
                <InfoTooltip text={isDe
                  ? 'Nach einem Tausch tragen alle NEU versendeten Mails automatisch das neue Logo — bereits versendete bleiben unverändert. Die Vorschau steht auf Dunkel, weil der Schriftzug weiß ist.'
                  : 'After a swap, all NEWLY sent emails automatically carry the new logo — emails already sent remain unchanged. The preview is dark because the wordmark is white.'} />
              </div>
              {branding && branding.logoBase64 ? (
                // Dunkle Vorschau-Fläche: das Logo ist ein WEISSER Schriftzug für
                // den schwarzen Mail-Header — auf Weiß wäre nur der grüne Punkt sichtbar.
                <img src={branding.logoBase64} alt="Deloitte Logo" style={{ maxWidth: '100%', maxHeight: 90, display: 'block', margin: '0 auto 10px', background: '#0d0d0d', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 12 }} />
              ) : (
                <p className="dex-ui-muted" style={{ fontStyle: 'italic', margin: '0 0 10px' }}>{isDe ? 'Noch kein Logo hinterlegt.' : 'No logo stored yet.'}</p>
              )}
              <p className="dex-ui-help" style={{ marginTop: 0 }}>
                {isDe
                  ? 'Weißer Deloitte-Schriftzug in der schwarzen Kopfzeile aller App-Mails.'
                  : 'White Deloitte wordmark in the black header of all app emails.'}
              </p>
              <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                <button className="btn btn-outline dex-ui-btn-sm" disabled={brandingBusy !== '' || !eventServiceRef} onClick={() => { if (logoInputRef.current) logoInputRef.current.click(); }}>
                  {brandingBusy === 'logo' ? (isDe ? 'Wird gespeichert…' : 'Saving…') : (isDe ? 'Neues hochladen (PNG)' : 'Upload new (PNG)')}
                </button>
                <button className="btn btn-secondary dex-ui-btn-sm" disabled={!branding || !branding.logoBase64} onClick={doDownloadLogo}>
                  {isDe ? 'Herunterladen' : 'Download'}
                </button>
              </div>
              <input ref={logoInputRef} type="file" accept="image/png" style={{ display: 'none' }} onChange={onLogoFileChosen} />
            </div>
            <div className="dex-ui-card">
              <div className="dex-ui-card-head" style={{ marginBottom: 10 }}>
                <h3 className="dex-ui-card-head-title"><FileText size={16} /> {isDe ? 'DEX-Logo-Video' : 'DEX logo video'}</h3>
                <InfoTooltip text={isDe
                  ? 'Hier tauschen und herunterladen. Der animierte Ring in der App selbst ist KEIN Video, sondern wird von der App live gerendert — hier liegt die Video-Datei zum Weitergeben, sobald sie einmal hochgeladen wurde.'
                  : 'Swap and download it here. The animated ring in the app itself is NOT a video but rendered live by the app — this slot stores the shareable video file once uploaded.'} />
              </div>
              {branding && branding.videoUrl ? (
                <video key={videoVer} src={branding.videoUrl + (videoVer ? `?ver=${videoVer}` : '')} controls style={{ width: '100%', maxHeight: 160, borderRadius: 8, background: '#000', marginBottom: 10 }} />
              ) : (
                <p className="dex-ui-muted" style={{ fontStyle: 'italic', margin: '0 0 10px' }}>{isDe ? 'Noch kein Video hinterlegt.' : 'No video stored yet.'}</p>
              )}
              <p className="dex-ui-help" style={{ marginTop: 0 }}>
                {isDe
                  ? 'Zentral abgelegtes Logo-Video zum Weitergeben — z. B. für Intranet-Artikel und Präsentationen.'
                  : 'Centrally stored logo video for sharing — e.g. for intranet articles and presentations.'}
              </p>
              <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                <button className="btn btn-outline dex-ui-btn-sm" disabled={brandingBusy !== '' || !eventServiceRef} onClick={() => { if (videoInputRef.current) videoInputRef.current.click(); }}>
                  {brandingBusy === 'video' ? (isDe ? 'Wird hochgeladen…' : 'Uploading…') : (isDe ? 'Neues Video hochladen' : 'Upload new video')}
                </button>
                {branding && branding.videoUrl ? (
                  <a className="btn btn-secondary dex-ui-btn-sm" href={branding.videoUrl} download={branding.videoFileName || 'DEX_Logo_Video.mp4'} style={{ textDecoration: 'none' }}>
                    {isDe ? 'Video herunterladen' : 'Download video'}
                  </a>
                ) : null}
              </div>
              <input ref={videoInputRef} type="file" accept="video/mp4,video/webm,video/quicktime" style={{ display: 'none' }} onChange={onVideoFileChosen} />
            </div>
          </div>
        </section>
      )}

      {/* Listen-Erklärung.
          v31.3: Als Tabelle mit ruhigem Kopf — und die Sprung-Auswahl aus der
          alten Werkzeug-Kachel steht jetzt hier, wo sie hingehört. Zwei
          Bedienwege für dieselbe Sache an zwei Seitenenden waren die Sucherei,
          die dieser Umbau abstellen soll. Der Zeilen-Hover kommt aus der Klasse,
          nicht mehr aus onMouseEnter. */}
      <section className="dex-ui-section">
        <h2 className="dex-ui-section-title">{isDe ? 'SharePoint-Listen — was macht was' : 'SharePoint lists — what does what'}</h2>
        <p className="dex-ui-section-desc">
          {isDe ? 'Alle Hintergrund-Listen der DEX-Plattform und wofür sie da sind. Ein Klick auf die Zeile öffnet die Liste in SharePoint.' : 'All background lists of the DEX platform and what they are for. A click on the row opens the list in SharePoint.'}
        </p>
        {/* v23.44: direktes Springen in eine SharePoint-Liste. */}
        <div className="dex-ui-toolbar">
          <label className="dex-ui-muted" htmlFor="dex-hub-listjump">{isDe ? 'Direkt in eine Liste springen' : 'Jump straight into a list'}</label>
          <select
            id="dex-hub-listjump"
            className="dex-ui-select dex-ui-select--sm"
            style={{ width: 'auto', minWidth: 200 }}
            defaultValue=""
            onChange={e => { const v = e.target.value; if (v) { window.open(listUrl(v), '_blank', 'noopener'); e.target.value = ''; } }}
          >
            <option value="">{isDe ? 'Zu Liste springen…' : 'Jump to list…'}</option>
            {LIST_DOCS.map(l => <option key={l.name} value={l.name}>{l.name}</option>)}
          </select>
        </div>
        <div className="dex-ui-table-wrap">
          <table className="dex-ui-table dex-ui-table--compact">
            <thead>
              <tr>
                <th style={{ minWidth: 160 }}>{isDe ? 'Liste' : 'List'}</th>
                <th>{isDe ? 'Wofür sie da ist' : 'What it is for'}</th>
              </tr>
            </thead>
            <tbody>
              {LIST_DOCS.map(l => (
                <tr key={l.name} className="is-clickable" onClick={() => { window.open(listUrl(l.name), '_blank', 'noopener'); }}>
                  <td>
                    <a
                      href={listUrl(l.name)}
                      target="_blank"
                      rel="noopener noreferrer"
                      onClick={e => e.stopPropagation()}
                      style={{ fontFamily: 'Consolas, monospace', fontSize: '0.8rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, textDecoration: 'none' }}
                    >{l.name}</a>
                  </td>
                  <td style={{ lineHeight: 1.45 }}>{l.de}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* Release Notes / Neuerungen — vollständige, durchsuchbare Tabelle. */}
      <section className="dex-ui-section">
        <h2 className="dex-ui-section-title">{isDe ? 'Neuerungen (Release Notes)' : 'What’s new (release notes)'}</h2>
        <p className="dex-ui-section-desc">
          {isDe
            ? 'Alle Versionen — durchsuchbar und nach Bereich filterbar (neueste oben). Die lückenlose Historie ist ab v18.65 verfügbar; ältere Einträge sind die dokumentierten Meilensteine.'
            : 'All versions — searchable and filterable by area (newest first).'}
        </p>
        {/* v31.3: Filter als Werkzeugleiste — Suche links, Filter daneben, die
            Trefferzahl als Pille rechts. */}
        <div className="dex-ui-toolbar">
          <div className="dex-ui-searchbar">
            <span className="dex-ui-searchbar-icon" aria-hidden="true"><Search size={15} /></span>
            <input
              className="dex-ui-input dex-ui-input--sm"
              type="text"
              value={rnSearch}
              onChange={e => setRnSearch(e.target.value)}
              placeholder={isDe ? 'Suchen (Text, Bereich, Version) …' : 'Search …'}
              aria-label={isDe ? 'Neuerungen durchsuchen' : 'Search release notes'}
            />
          </div>
          <select className="dex-ui-select dex-ui-select--sm" style={{ width: 'auto' }} value={rnBereich} onChange={e => setRnBereich(e.target.value)} aria-label={isDe ? 'Bereich' : 'Area'}>
            <option value="">{isDe ? 'Alle Bereiche' : 'All areas'}</option>
            {RELEASE_BEREICHE.map(b => <option key={b} value={b}>{b}</option>)}
          </select>
          <select className="dex-ui-select dex-ui-select--sm" style={{ width: 'auto' }} value={rnType} onChange={e => setRnType(e.target.value)} aria-label={isDe ? 'Art' : 'Type'}>
            <option value="">{isDe ? 'Neu & Behoben' : 'All types'}</option>
            <option value="Feature">{isDe ? 'Nur Neu' : 'Features'}</option>
            <option value="Bugfix">{isDe ? 'Nur Behoben' : 'Fixes'}</option>
          </select>
          <span className="dex-ui-toolbar-spacer" />
          <span className="dex-ui-pill dex-ui-pill--gray">
            {isDe ? `${filteredNotes.length} von ${RELEASE_NOTES.length}` : `${filteredNotes.length} of ${RELEASE_NOTES.length}`}
          </span>
        </div>
      <div className="dex-ui-card" style={{ padding: 0, overflow: 'auto' }}>
        {/* Tabellenkopf */}
        <div style={{ display: isMobile ? 'none' : 'grid', gridTemplateColumns: '70px 92px 150px 78px 1fr', minWidth: 720, gap: 12, padding: '10px 16px', background: 'var(--dex-gray-50, #fafafa)', borderBottom: '1px solid var(--dex-gray-200)', fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.3 }}>
          <span>{isDe ? 'Version' : 'Version'}</span>
          <span>{isDe ? 'Datum' : 'Date'}</span>
          <span>{isDe ? 'Bereich' : 'Area'}</span>
          <span>{isDe ? 'Art' : 'Type'}</span>
          <span>{isDe ? 'Beschreibung' : 'Description'}</span>
        </div>
        {filteredNotes.length === 0 ? (
          <div className="dex-ui-empty" style={{ border: 'none', background: 'transparent' }}>
            <div className="dex-ui-empty-title">{isDe ? 'Keine Treffer für diese Filter.' : 'No matches for these filters.'}</div>
            {isDe ? 'Suchbegriff kürzen oder die Filter auf „Alle" stellen.' : 'Shorten the search term or set the filters back to „All".'}
          </div>
        ) : filteredNotes.map((n, i) => (
          <div key={`${n.version}-${i}`} style={{ display: 'grid', gridTemplateColumns: isMobile ? '64px 1fr' : '70px 92px 150px 78px 1fr', minWidth: isMobile ? 0 : 720, gap: isMobile ? '2px 10px' : 12, padding: '11px 16px', borderTop: i === 0 ? 'none' : '1px solid var(--dex-gray-100)', alignItems: 'baseline' }}>
            <code style={{ fontFamily: 'Consolas, monospace', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>v{n.version}</code>
            <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{fmtDate(n.date)}</span>
            <span style={{ fontSize: '0.74rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>{n.bereich}</span>
            <span className={cx('dex-ui-pill', n.type === 'Bugfix' ? 'dex-ui-pill--red' : 'dex-ui-pill--green')} style={{ justifySelf: 'start', alignSelf: 'start', fontSize: '0.68rem', padding: '2px 8px' }}>
              {n.type === 'Bugfix' ? (isDe ? 'Behoben' : 'Fix') : (isDe ? 'Neu' : 'New')}
            </span>
            {/* v30.60: Kernaussage fett, Einzelpunkte als Liste. Die Gliederung
                wird beim Anzeigen aus dem Text gelesen (siehe
                data/releaseNotes.splitReleaseNote) — dadurch gilt sie auch für
                alle Alt-Einträge, ohne dass 600 belegte Texte umgeschrieben
                werden müssten. */}
            <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
              {(() => {
                const p = splitReleaseNote(n.text);
                return (
                  <>
                    {p.lead && <strong style={{ display: 'block', color: 'var(--dex-gray-800)' }}>{p.lead}</strong>}
                    {p.rest && <span style={{ display: 'block', marginTop: p.lead ? 4 : 0 }}>{p.rest}</span>}
                    {p.points.length > 0 && (
                      <ul style={{ margin: p.lead ? '6px 0 0' : 0, paddingLeft: 18 }}>
                        {p.points.map((pt, pi) => <li key={pi} style={{ marginBottom: 4 }}>{pt}</li>)}
                      </ul>
                    )}
                  </>
                );
              })()}
            </span>
          </div>
        ))}
      </div>
      </section>

      {/* v26.81: Berechtigungen aufräumen — Prüf-/Korrektur-Modal. */}
      {permCleanupOpen && (
        <Modal
          open={true}
          onClose={() => { if (!permCleanupBusy) { setPermCleanupOpen(false); setPermCleanupReport(null); } }}
          dismissable={!permCleanupBusy}
          maxWidth={720}
          padding={24}
          ariaLabel={isDe ? 'Berechtigungen aufräumen' : 'Clean up permissions'}
          title={isDe ? 'Berechtigungen aufräumen' : 'Clean up permissions'}
          icon={<Wrench size={20} />}
          subtitle={isDe
            ? 'Der ganze SharePoint (Hauptseite, alle Listen/Bibliotheken und alle Event-Subsites) wird nach manuellen Einzel-Freigaben durchsucht, die einer Person mehr als Leserechte geben, obwohl sie laut Rollen-Konzept kein Admin/Organizer ist. Solche Über-Freigaben lassen sich hier entfernen — Leserechte und alle Gruppen-Berechtigungen bleiben unangetastet.'
            : 'The whole SharePoint (main site, all lists/libraries and every event subsite) is scanned for manual individual grants that give a person more than read access even though they are not an admin/organizer per the role concept. Such over-grants can be removed here — read access and all group permissions stay untouched.'}
          footer={permCleanupFooter}
        >
          {permCleanupBusy && permCleanupProgress && (() => {
            const { msg, done, total } = permCleanupProgress;
            const pct = Math.min(100, Math.round((done / Math.max(1, total)) * 100));
            return (
              <div>
                <p style={{ margin: '0 0 6px', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>{msg}</p>
                <div className="dex-ui-progress"><div className="dex-ui-progress-bar" style={{ width: `${pct}%` }} /></div>
                <p className="dex-ui-help" style={{ marginTop: 8 }}>
                  {isDe ? 'Bitte das Fenster geöffnet lassen, bis der Lauf abgeschlossen ist.' : 'Please keep this window open until the run completes.'}
                </p>
              </div>
            );
          })()}

          {!permCleanupBusy && permCleanupReport && (() => {
            const r = permCleanupReport;
            const hasIssues = r.strayWriteFound > 0 || r.ilsIssues > 0;
            const strayFindings = r.findings.filter(f => f.kind === 'stray-write');
            const ilsFindings = r.findings.filter(f => f.kind === 'ils');
            const errFindings = r.findings.filter(f => f.kind === 'error');
            const SHOW = 200;
            return (
              <div className="dex-ui-stack">
                <div className={cx('dex-ui-callout', hasIssues ? 'dex-ui-callout--warn' : 'dex-ui-callout--success')} style={{ display: 'block' }}>
                  <div style={{ fontWeight: 700, marginBottom: 4 }}>
                    {r.apply
                      ? (isDe ? 'Korrektur abgeschlossen' : 'Cleanup complete')
                      : (hasIssues ? (isDe ? 'Prüfung abgeschlossen — Abweichungen gefunden' : 'Check complete — deviations found') : (isDe ? 'Prüfung abgeschlossen — alles sauber' : 'Check complete — all clean'))}
                  </div>
                  {isDe ? (
                    <>Geprüft: {r.websScanned} Webs, {r.listsScanned} Listen mit eigenen Berechtigungen.<br />
                    Über-Freigaben (Schreib-/Vollzugriff einzelner Personen): <strong>{r.strayWriteFound}</strong>{r.apply ? ` — davon ${r.strayWriteRemoved} entfernt` : ''}.<br />
                    Element-Sicherheit falsch (sensible Listen): <strong>{r.ilsIssues}</strong>{r.apply ? ` — davon ${r.ilsFixed} korrigiert` : ''}.
                    {r.errors > 0 ? <><br />Nicht lesbar/Fehler: {r.errors}.</> : null}</>
                  ) : (
                    <>Scanned: {r.websScanned} webs, {r.listsScanned} lists with unique permissions.<br />
                    Over-grants (individual write/full control): <strong>{r.strayWriteFound}</strong>{r.apply ? ` — ${r.strayWriteRemoved} removed` : ''}.<br />
                    Item-security wrong (sensitive lists): <strong>{r.ilsIssues}</strong>{r.apply ? ` — ${r.ilsFixed} fixed` : ''}.
                    {r.errors > 0 ? <><br />Unreadable/errors: {r.errors}.</> : null}</>
                  )}
                </div>

                {(strayFindings.length > 0 || ilsFindings.length > 0 || errFindings.length > 0) && (
                  <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky" style={{ maxHeight: 300 }}>
                    <table className="dex-ui-table dex-ui-table--compact">
                      <thead>
                        <tr>
                          <th>{isDe ? 'Ort' : 'Location'}</th>
                          <th>{isDe ? 'Person' : 'Person'}</th>
                          <th>{isDe ? 'Befund' : 'Finding'}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...strayFindings, ...ilsFindings, ...errFindings].slice(0, SHOW).map((f, i) => (
                          <tr key={i}>
                            <td style={{ color: 'var(--dex-gray-600)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 200 }} title={f.scope}>{f.scope}</td>
                            <td style={{ color: 'var(--dex-gray-700)', wordBreak: 'break-all' }}>{f.principal || '—'}</td>
                            <td style={{ color: f.kind === 'error' ? 'var(--dex-red, #da291c)' : (f.fixed ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-orange-dark, #b35a00)') }}>{f.detail}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {r.findings.length > SHOW && (
                      <div className="dex-ui-table-foot">
                        {isDe ? `… und ${r.findings.length - SHOW} weitere (gekürzt).` : `… and ${r.findings.length - SHOW} more (truncated).`}
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })()}

          {!permCleanupBusy && !permCleanupReport && (
            <p className="dex-ui-muted" style={{ margin: 0 }}>
              {isDe
                ? 'Der Lauf liest nur — geändert wird erst, wenn du danach „Jetzt korrigieren" wählst. Je nach Größe der Seite dauert er einige Minuten.'
                : 'The run only reads — nothing changes until you pick „Fix now" afterwards. Depending on the site size it takes a few minutes.'}
            </p>
          )}
        </Modal>
      )}

      {/* v26.81: Verwaiste Subsites — Prüf-/Lösch-Modal. */}
      {orphanOpen && (
        <Modal
          open={true}
          onClose={() => { if (!orphanBusy) { setOrphanOpen(false); setOrphanResult(null); } }}
          dismissable={!orphanBusy}
          maxWidth={760}
          padding={24}
          ariaLabel={isDe ? 'Subsites prüfen' : 'Check subsites'}
          title={isDe ? 'Verwaiste Subsites' : 'Orphan subsites'}
          icon={<Trash2 size={20} />}
          subtitle={isDe
            ? 'Subsites, die noch existieren, aber zu keinem Event mehr gehören (z.B. Test-Subsites gelöschter Events). Prüfe pro Eintrag, ob wirklich ein Rest vorliegt, bevor du löschst — das Löschen ist endgültig.'
            : 'Subsites that still exist but no longer belong to any event (e.g. test subsites of deleted events). Check each entry before deleting — deletion is permanent.'}
          footer={orphanFooter}
        >
          {orphanBusy && orphanProgress && (() => {
            const { msg, done, total } = orphanProgress;
            const pct = Math.min(100, Math.round((done / Math.max(1, total)) * 100));
            return (
              <div>
                <p style={{ margin: '0 0 6px', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>{msg}</p>
                <div className="dex-ui-progress"><div className="dex-ui-progress-bar" style={{ width: `${pct}%` }} /></div>
              </div>
            );
          })()}

          {!orphanBusy && orphanResult && (() => {
            const r = orphanResult;
            const remaining = r.orphans.filter(o => !orphanDeleted[o.url]);
            return (
              <div className="dex-ui-stack">
                <div className={cx('dex-ui-callout', remaining.length > 0 ? 'dex-ui-callout--warn' : 'dex-ui-callout--success')} style={{ display: 'block' }}>
                  <div style={{ fontWeight: 700, marginBottom: 4 }}>
                    {remaining.length > 0
                      ? (isDe ? `${remaining.length} verwaiste Subsite(s) gefunden` : `${remaining.length} orphan subsite(s) found`)
                      : (isDe ? 'Keine verwaisten Subsites' : 'No orphan subsites')}
                  </div>
                  {isDe
                    ? <>{r.websScanned} Subsites geprüft, {r.eventSubsites} davon gehören zu Events.</>
                    : <>{r.websScanned} subsites scanned, {r.eventSubsites} belong to events.</>}
                </div>

                {remaining.length > 0 && (
                  <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky" style={{ maxHeight: 320 }}>
                    <table className="dex-ui-table dex-ui-table--compact">
                      <thead>
                        <tr>
                          <th>{isDe ? 'Subsite' : 'Subsite'}</th>
                          <th>{isDe ? 'Teilnehmerliste' : 'Participant list'}</th>
                          <th className="is-actions"><span className="dex-ui-sr-only">{isDe ? 'Aktion' : 'Action'}</span></th>
                        </tr>
                      </thead>
                      <tbody>
                        {remaining.map((o) => (
                          <tr key={o.url}>
                            <td style={{ color: 'var(--dex-gray-700)' }}>
                              <div style={{ fontWeight: 600 }}>{o.title || o.serverRel}</div>
                              <a href={o.url} target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark, #4a7c1f)', wordBreak: 'break-all', fontSize: '0.72rem' }}>{o.serverRel || o.url}</a>
                            </td>
                            <td>
                              <span className={cx('dex-ui-pill', o.hasParticipantList ? 'dex-ui-pill--orange' : 'dex-ui-pill--gray')}>
                                {o.hasParticipantList
                                  ? (isDe ? `ja · ${o.participantCount} Zeilen` : `yes · ${o.participantCount} rows`)
                                  : (isDe ? 'keine' : 'none')}
                              </span>
                            </td>
                            <td className="is-actions">
                              <button
                                type="button"
                                className="dex-ui-textbtn dex-ui-textbtn--danger"
                                disabled={!!orphanDeleting[o.url]}
                                onClick={() => { void deleteOrphan(o.url, o.title || o.serverRel, o.participantCount); }}
                              >
                                {orphanDeleting[o.url] ? (isDe ? 'Löscht…' : 'Deleting…') : (isDe ? 'Löschen' : 'Delete')}
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })()}

          {!orphanBusy && !orphanResult && (
            <p className="dex-ui-muted" style={{ margin: 0 }}>
              {isDe
                ? 'Die Prüfung liest nur — gelöscht wird nichts von allein, jeder Rest einzeln und mit Rückfrage.'
                : 'The check only reads — nothing is deleted on its own; each orphan goes individually and with a confirmation.'}
            </p>
          )}
        </Modal>
      )}
    </div>
  );
}
