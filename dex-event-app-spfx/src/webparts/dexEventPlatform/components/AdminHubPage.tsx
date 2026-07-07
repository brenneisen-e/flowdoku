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
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { useDialog } from '../context/DialogContext';
import { useIsMobile } from '../utils/useIsMobile';
import { Settings, Users, Mail, Book, FileText, Trash2, Columns, BarChart3, Wrench } from './Icons';
import { RELEASE_NOTES, RELEASE_BEREICHE } from '../data/releaseNotes';
import { EventService, PermCleanupReport, OrphanScanResult } from '../services/EventService';
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
  const listUrl = (name: string): string => `${siteUrl}/Lists/${name}`;
  const { events: allEvents, getArchivableCount, runArchiveExpired, getDeletableArchiveCount, runDeleteOldArchive, fixAllEventColumns, restoreCustomFieldDescriptions, reseedDefaultEmailTemplates, maybeSendWeeklyReport, recomputeEventKpiOnly } = useEvents();
  const { locale } = useLanguage();
  const { confirmDialog, showAlert } = useDialog();
  const isDe = locale === 'de';
  const isMobile = useIsMobile();
  const adminLike = isAdmin || originalIsAdmin;

  const [archTotal, setArchTotal] = React.useState(0);
  const [delTotal, setDelTotal] = React.useState(0);
  const [busy, setBusy] = React.useState<'' | 'arch' | 'del' | 'fixcols' | 'restoredesc' | 'reseed' | 'weekly' | 'kpi'>('');
  // v26.63: zuletzt neu berechnete Events-Zahl (für die Erfolgs-Anzeige).
  const [kpiResult, setKpiResult] = React.useState<number | null>(null);
  // v24.33: Fortschritt für das globale „Spalten fixen".
  const [fixProgress, setFixProgress] = React.useState<{ done: number; total: number; label: string } | null>(null);
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
    let cancelled = false;
    getArchivableCount().then(r => { if (!cancelled) setArchTotal(r.total); }).catch(() => { /* */ });
    getDeletableArchiveCount().then(n => { if (!cancelled) setDelTotal(n); }).catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [adminLike]);

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
      showAlert(msg, { variant: r.errors ? 'error' : 'success' });
    } catch { showAlert(isDe ? 'Spalten-Prüfung fehlgeschlagen.' : 'Column check failed.', { variant: 'error' }); }
    finally { setBusy(''); setFixProgress(null); }
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
    { icon: <Users size={28} />, title: isDe ? 'Organizer Center' : 'Organizer center', desc: isDe ? 'Teilnehmer, Prozesse, Audit-Log, SharePoint-Liste — alle Event-Werkzeuge pro Event.' : 'Attendees, processes, audit log, SharePoint list — all per-event tools.', onClick: () => navigate('admin') },
    { icon: <Settings size={28} />, title: isDe ? 'Prozessübersicht' : 'Process overview', desc: isDe ? 'Wie die Abläufe in DEX funktionieren — verständlich erklärt.' : 'How the DEX processes work — explained simply.', onClick: () => navigate('flowcharts') },
    { icon: <FileText size={28} />, title: isDe ? 'Architektur' : 'Architecture', desc: isDe ? 'Gesamtarchitektur (App, SharePoint-Listen, Power-Automate-Flows, M365-Dienste) — mit PDF-Export.' : 'Overall architecture (app, SharePoint lists, Power Automate flows, M365 services) — with PDF export.', onClick: () => navigate('architecture') },
    { icon: <Users size={28} />, title: isDe ? 'Rollenverwaltung' : 'Role management', desc: isDe ? 'User, Organizer und Admins zuweisen oder entfernen.' : 'Assign or remove users, organizers and admins.', onClick: () => navigate('settings') },
    { icon: <Columns size={28} />, title: isDe ? 'Rollenmatrix' : 'Role matrix', desc: isDe ? 'Übersicht: wer welche Rechte hat (User, Organizer, Admin).' : 'Overview: who has which permissions (user, organizer, admin).', onClick: () => navigate('role-matrix') },
    { icon: <Mail size={28} />, title: isDe ? 'Mail-Vorlagen' : 'Mail templates', desc: isDe ? 'Globale Standard-Mails (Anmeldung, Warteliste, Abmeldung …) ansehen und bearbeiten — mit Live-Vorschau.' : 'View and edit the global default emails — with live preview.', onClick: () => navigate('email-templates') },
    { icon: <BarChart3 size={28} />, title: isDe ? 'Statistik-Archiv' : 'Statistics archive', desc: isDe ? 'Kennzahlen gelöschter Teilnehmerlisten: welches Event, wann, von wem organisiert, mit welcher Teilnehmerzahl (ohne PII).' : 'KPIs of deleted participant lists: which event, when, organized by whom, with how many participants (no PII).', onClick: () => navigate('stats-archive') },
    { icon: <Book size={28} />, title: isDe ? 'Handbuch' : 'Manual', desc: isDe ? 'Ausführliche Anleitung zu allen Funktionen.' : 'Detailed guide for all features.', onClick: () => navigate('manual') },
  ];

  const cardStyle: React.CSSProperties = { background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 18, boxShadow: '0 2px 8px rgba(0,0,0,0.05)', transition: 'border-color 0.15s ease, box-shadow 0.15s ease, transform 0.15s ease' };
  // v24.70: Hover-Effekt für die klickbaren Kacheln — grüner Rand, leichter
  // Lift + Schatten (wie die Kacheln auf der Startseite).
  const onCardHover = (e: React.MouseEvent<HTMLDivElement>): void => {
    const el = e.currentTarget;
    el.style.borderColor = 'var(--dex-green, #86bc25)';
    el.style.boxShadow = '0 8px 22px rgba(134,188,37,0.20)';
    el.style.transform = 'translateY(-2px)';
  };
  const onCardLeave = (e: React.MouseEvent<HTMLDivElement>): void => {
    const el = e.currentTarget;
    el.style.borderColor = 'var(--dex-gray-200)';
    el.style.boxShadow = '0 2px 8px rgba(0,0,0,0.05)';
    el.style.transform = '';
  };

  return (
    <div className="page-container">
      <h1 style={{ marginTop: 0 }}>{isDe ? 'Admin' : 'Admin'}</h1>
      <p style={{ color: 'var(--dex-gray-600)', marginTop: 0 }}>
        {isDe ? 'Zentrale Anlaufstelle für Admin-Themen.' : 'Central place for admin topics.'}
      </p>

      {/* Werkzeuge */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 14, marginBottom: 28 }}>
        {tools.map((t, i) => (
          <div key={i} className="card-clickable" style={{ ...cardStyle, cursor: 'pointer', display: 'flex', gap: 14, alignItems: 'flex-start' }} onClick={t.onClick} onMouseEnter={onCardHover} onMouseLeave={onCardLeave}>
            <span style={{ color: 'var(--dex-green, #86bc25)', flexShrink: 0 }}>{t.icon}</span>
            <span>
              <span style={{ display: 'block', fontWeight: 700, color: 'var(--dex-gray-800)', marginBottom: 2 }}>{t.title}</span>
              <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.4 }}>{t.desc}</span>
            </span>
          </div>
        ))}
        {/* v23.44: eigene Kachel zum direkten Springen in eine SharePoint-Liste. */}
        <div style={{ ...cardStyle, display: 'flex', gap: 14, alignItems: 'flex-start' }}>
          <span style={{ color: 'var(--dex-green, #86bc25)', flexShrink: 0 }}><FileText size={28} /></span>
          <span style={{ flex: 1, minWidth: 0 }}>
            <span style={{ display: 'block', fontWeight: 700, color: 'var(--dex-gray-800)', marginBottom: 2 }}>{isDe ? 'SharePoint-Liste öffnen' : 'Open SharePoint list'}</span>
            <span style={{ display: 'block', fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.4, marginBottom: 8 }}>{isDe ? 'Direkt in eine der Hintergrund-Listen springen.' : 'Jump straight into one of the background lists.'}</span>
            <select
              defaultValue=""
              onChange={e => { const v = e.target.value; if (v) { window.open(listUrl(v), '_blank', 'noopener'); e.target.value = ''; } }}
              style={{ width: '100%', padding: '6px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300)', fontSize: '0.82rem', cursor: 'pointer' }}
            >
              <option value="">{isDe ? 'Zu Liste springen…' : 'Jump to list…'}</option>
              {LIST_DOCS.map(l => <option key={l.name} value={l.name}>{l.name}</option>)}
            </select>
          </span>
        </div>
      </div>

      {/* v26.51: Logo & Branding — Default-Mail-Logo tauschen/herunterladen + Logo-Video */}
      {adminLike && (
        <>
          <h2 style={{ fontSize: '1.15rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Logo & Branding' : 'Logo & branding'}</h2>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 14, marginBottom: 28 }}>
            {/* v26.58: DEX-Logo = der bunte Orb-Ring (vorher zeigte diese Karte
                fälschlich das Deloitte-Mail-Logo, dessen weißer Schriftzug auf
                weißem Grund unsichtbar war — „nur ein grüner Punkt"). */}
            <div style={cardStyle}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><FileText size={18} /></span>
                <span style={{ fontWeight: 700 }}>{isDe ? 'DEX-Logo (Orb, PNG)' : 'DEX logo (orb, PNG)'}</span>
              </div>
              {branding && branding.orbBase64 ? (
                <img src={branding.orbBase64} alt="DEX Orb" style={{ maxWidth: '100%', maxHeight: 90, display: 'block', margin: '0 auto 10px', background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 8 }} />
              ) : (
                <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: '0 0 10px' }}>{isDe ? 'Noch kein DEX-Logo hinterlegt.' : 'No DEX logo stored yet.'}</p>
              )}
              <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
                {isDe
                  ? 'Der bunte DEX-Ring — Standard-Bild in Mails von Events ohne eigenes Event-Bild und zentrale Download-Quelle (z. B. für Intranet-Artikel). Neue Mails nutzen nach einem Tausch automatisch das neue Bild.'
                  : 'The colourful DEX ring — default image in emails of events without their own image and central download source (e.g. for intranet articles). New emails automatically use the new image after a swap.'}
              </p>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1 }} disabled={!branding || !branding.orbBase64} onClick={doDownloadOrb}>
                  {isDe ? 'Herunterladen' : 'Download'}
                </button>
                <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1 }} disabled={brandingBusy !== '' || !eventServiceRef} onClick={() => { if (orbInputRef.current) orbInputRef.current.click(); }}>
                  {brandingBusy === 'orb' ? (isDe ? 'Wird gespeichert…' : 'Saving…') : (isDe ? 'Neues hochladen (PNG)' : 'Upload new (PNG)')}
                </button>
              </div>
              <input ref={orbInputRef} type="file" accept="image/png" style={{ display: 'none' }} onChange={onOrbFileChosen} />
            </div>
            <div style={cardStyle}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><FileText size={18} /></span>
                <span style={{ fontWeight: 700 }}>{isDe ? 'Deloitte-Logo (E-Mail-Kopfzeile)' : 'Deloitte logo (email header)'}</span>
              </div>
              {branding && branding.logoBase64 ? (
                // Dunkle Vorschau-Fläche: das Logo ist ein WEISSER Schriftzug für
                // den schwarzen Mail-Header — auf Weiß wäre nur der grüne Punkt sichtbar.
                <img src={branding.logoBase64} alt="Deloitte Logo" style={{ maxWidth: '100%', maxHeight: 90, display: 'block', margin: '0 auto 10px', background: '#0d0d0d', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 12 }} />
              ) : (
                <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: '0 0 10px' }}>{isDe ? 'Noch kein Logo hinterlegt.' : 'No logo stored yet.'}</p>
              )}
              <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
                {isDe
                  ? 'Weißer Deloitte-Schriftzug in der schwarzen Kopfzeile aller App-Mails (Vorschau deshalb auf Dunkel). Nach einem Tausch tragen alle NEU versendeten Mails automatisch das neue Logo — bereits versendete bleiben unverändert.'
                  : 'White Deloitte wordmark in the black header of all app emails (hence the dark preview). After a swap, all NEWLY sent emails automatically carry the new logo — emails already sent remain unchanged.'}
              </p>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1 }} disabled={!branding || !branding.logoBase64} onClick={doDownloadLogo}>
                  {isDe ? 'Herunterladen' : 'Download'}
                </button>
                <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1 }} disabled={brandingBusy !== '' || !eventServiceRef} onClick={() => { if (logoInputRef.current) logoInputRef.current.click(); }}>
                  {brandingBusy === 'logo' ? (isDe ? 'Wird gespeichert…' : 'Saving…') : (isDe ? 'Neues hochladen (PNG)' : 'Upload new (PNG)')}
                </button>
              </div>
              <input ref={logoInputRef} type="file" accept="image/png" style={{ display: 'none' }} onChange={onLogoFileChosen} />
            </div>
            <div style={cardStyle}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><FileText size={18} /></span>
                <span style={{ fontWeight: 700 }}>{isDe ? 'DEX-Logo-Video' : 'DEX logo video'}</span>
              </div>
              {branding && branding.videoUrl ? (
                <video key={videoVer} src={branding.videoUrl + (videoVer ? `?ver=${videoVer}` : '')} controls style={{ width: '100%', maxHeight: 160, borderRadius: 8, background: '#000', marginBottom: 10 }} />
              ) : (
                <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: '0 0 10px' }}>{isDe ? 'Noch kein Video hinterlegt.' : 'No video stored yet.'}</p>
              )}
              <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
                {isDe
                  ? 'Zentral abgelegtes Logo-Video (z. B. für Intranet-Artikel und Präsentationen) — hier tauschen und herunterladen. Hinweis: Der animierte Ring in der App selbst ist KEIN Video, sondern wird von der App live gerendert — hier liegt die Video-Datei zum Weitergeben, sobald sie einmal hochgeladen wurde.'
                  : 'Centrally stored logo video (e.g. for intranet articles and presentations) — swap and download it here. Note: the animated ring in the app itself is NOT a video but rendered live by the app — this slot stores the shareable video file once uploaded.'}
              </p>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                {branding && branding.videoUrl ? (
                  <a className="btn btn-secondary" href={branding.videoUrl} download={branding.videoFileName || 'DEX_Logo_Video.mp4'} style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1, textAlign: 'center', textDecoration: 'none' }}>
                    {isDe ? 'Video herunterladen' : 'Download video'}
                  </a>
                ) : null}
                <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1 }} disabled={brandingBusy !== '' || !eventServiceRef} onClick={() => { if (videoInputRef.current) videoInputRef.current.click(); }}>
                  {brandingBusy === 'video' ? (isDe ? 'Wird hochgeladen…' : 'Uploading…') : (isDe ? 'Neues Video hochladen' : 'Upload new video')}
                </button>
              </div>
              <input ref={videoInputRef} type="file" accept="video/mp4,video/webm,video/quicktime" style={{ display: 'none' }} onChange={onVideoFileChosen} />
            </div>
          </div>
        </>
      )}

      {/* Archiv & Löschung */}
      <h2 style={{ fontSize: '1.15rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Archiv & Löschung' : 'Archive & deletion'}</h2>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 14, marginBottom: 28 }}>
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><FileText size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Archivieren' : 'Archive'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? <><strong>{archTotal}</strong> Zeilen aus abgelaufenen oder gelöschten Events stehen zur Archivierung an. Sie wandern aus den Arbeitslisten ins Archiv.</>
              : <><strong>{archTotal}</strong> rows from expired or deleted events are ready to archive. They move from the working lists into the archive.</>}
          </p>
          <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }} disabled={busy !== '' || archTotal === 0} onClick={() => { void doArchive(); }}>
            {busy === 'arch' ? (isDe ? 'Wird archiviert…' : 'Archiving…') : (isDe ? 'Jetzt archivieren' : 'Archive now')}
          </button>
        </div>
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-red, #c00)', display: 'inline-flex' }}><Trash2 size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Altes Archiv löschen' : 'Delete old archive'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? <><strong>{delTotal}</strong> Archiv-Einträge sind älter als 1 Monat (Event vorbei) und können endgültig gelöscht werden.</>
              : <><strong>{delTotal}</strong> archive entries are older than 1 month and can be permanently deleted.</>}
          </p>
          <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%', color: 'var(--dex-red, #c00)' }} disabled={busy !== '' || delTotal === 0} onClick={() => { void doDelete(); }}>
            {busy === 'del' ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Alte Einträge löschen' : 'Delete old entries')}
          </button>
        </div>
      </div>

      {/* v24.33: Wartung — globales Spalten fixen (alle Events inkl. Sub-Events) */}
      <h2 style={{ fontSize: '1.15rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Wartung' : 'Maintenance'}</h2>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 14, marginBottom: 28 }}>
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Settings size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Spalten fixen (alle Events)' : 'Fix columns (all events)'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? 'Prüft die Teilnehmerlisten ALLER Events inkl. Sub-Events, legt fehlende Spalten an (z.B. „Unternehmen") und trägt die Unternehmenszugehörigkeit für bestehende Teilnehmer nach.'
              : 'Checks the participant lists of ALL events incl. sub-events, adds missing columns (e.g. „Company") and backfills the company affiliation for existing attendees.'}
          </p>
          {busy === 'fixcols' && fixProgress && (
            <div style={{ margin: '0 0 10px' }}>
              <div style={{ height: 8, background: 'var(--dex-gray-100)', borderRadius: 999, overflow: 'hidden' }}>
                <div style={{ height: '100%', width: `${fixProgress.total > 0 ? Math.round((fixProgress.done / fixProgress.total) * 100) : 0}%`, background: 'var(--dex-green, #86bc25)', transition: 'width 0.2s' }} />
              </div>
              <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 4, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {fixProgress.done}/{fixProgress.total}{fixProgress.label ? ` · ${fixProgress.label}` : ''}
              </div>
            </div>
          )}
          <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }} disabled={busy !== ''} onClick={() => { void doFixAllColumns(); }}>
            {busy === 'fixcols' ? (isDe ? 'Wird geprüft…' : 'Checking…') : (isDe ? 'Jetzt alle prüfen' : 'Check all now')}
          </button>
        </div>

        {/* v26.63: Startseiten-Zähler (Events/Teilnehmer) neu berechnen. */}
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><BarChart3 size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Startseiten-Zähler neu berechnen' : 'Recompute landing-page counter'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? 'Berechnet die Kennzahl „Events" auf der Startseite frisch aus der Event-Liste (ohne Entwürfe, abgesagte und Sub-Events; abgelaufene zählen mit) — schnell, ohne die Teilnehmerlisten zu scannen. Der angezeigte Wert ist ein gespeicherter Zähler, der sonst nur einmal pro Admin-Sitzung automatisch aktualisiert wird. Der Teilnehmer-Zähler bleibt unverändert.'
              : 'Recomputes the „Events" KPI on the landing page straight from the event list (excluding drafts, cancelled and sub-events; past ones count) — fast, without scanning the participant lists. The shown value is a stored counter that otherwise only refreshes once per admin session. The attendee counter is left unchanged.'}
          </p>
          {kpiResult !== null && (
            <div style={{ margin: '0 0 10px', padding: '8px 12px', background: '#f1f7e8', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
              {isDe ? `Ergebnis: ${kpiResult} Events` : `Result: ${kpiResult} events`}
            </div>
          )}
          <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }} disabled={busy !== ''} onClick={() => { void doRecomputeKpi(); }}>
            {busy === 'kpi' ? (isDe ? 'Wird berechnet…' : 'Computing…') : (isDe ? 'Events-Zähler neu berechnen' : 'Recompute events counter')}
          </button>
        </div>

        {/* v26.13: Feld-Eigenschaften aus der Versionshistorie wiederherstellen. */}
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Settings size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Feld-Beschreibungen wiederherstellen' : 'Restore field descriptions'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? 'Stellt versehentlich verlorene Eigenschaften der Abfrage-/Auswahlfelder (Beschreibungen, Anzeige-Bedingungen, Mehrfachauswahl, Englisch-Varianten u.a.) aus der SharePoint-Versionshistorie wieder her. Füllt nur FEHLENDE Werte auf — aktuelle Eingaben bleiben erhalten.'
              : 'Restores accidentally lost properties of the form/selection fields (descriptions, display conditions, multi-select, English variants, etc.) from the SharePoint version history. Only fills in MISSING values — current entries are preserved.'}
          </p>
          {busy === 'restoredesc' && restoreProgress && (
            <div style={{ margin: '0 0 10px' }}>
              <div style={{ height: 8, background: 'var(--dex-gray-100)', borderRadius: 999, overflow: 'hidden' }}>
                <div style={{ height: '100%', width: `${restoreProgress.total > 0 ? Math.round((restoreProgress.done / restoreProgress.total) * 100) : 0}%`, background: 'var(--dex-green, #86bc25)', transition: 'width 0.2s' }} />
              </div>
              <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 4, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {restoreProgress.done}/{restoreProgress.total}{restoreProgress.label ? ` · ${restoreProgress.label}` : ''}
              </div>
            </div>
          )}
          <div style={{ display: 'flex', gap: 8 }}>
            <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1 }} disabled={busy !== ''} onClick={() => { void doPreviewDescriptions(); }}>
              {isDe ? 'Vorschau (Trockenlauf)' : 'Preview (dry run)'}
            </button>
            <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 12px', flex: 1 }} disabled={busy !== ''} onClick={() => { void doRestoreDescriptions(); }}>
              {busy === 'restoredesc' ? (isDe ? 'Läuft…' : 'Running…') : (isDe ? 'Wiederherstellen' : 'Restore')}
            </button>
          </div>
          {restorePreview && (
            <div style={{ marginTop: 10, maxHeight: 220, overflowY: 'auto', fontSize: '0.78rem', borderTop: '1px solid var(--dex-gray-100)', paddingTop: 8 }}>
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

        {/* v26.81: Berechtigungen aufräumen (ganze Site-Collection). */}
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Wrench size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Berechtigungen aufräumen (ganzer SharePoint)' : 'Clean up permissions (whole SharePoint)'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? 'Prüft die gesamte SharePoint-Seite (Hauptseite, alle Listen/Bibliotheken und alle Event-Subsites) auf manuelle Einzel-Freigaben, die einzelnen Personen mehr Rechte geben als im Berechtigungskonzept vorgesehen (z.B. Schreib-/Vollzugriff auf ganze Listen). Erst kommt ein Bericht ohne Änderung, danach kannst du die Über-Freigaben mit einem Klick entfernen. Leserechte bleiben immer erhalten (auch für internationale Kolleg:innen); Schreiben ist danach nur über die Gruppen und für Admins/Organizer möglich.'
              : 'Scans the whole SharePoint site (main site, all lists/libraries and every event subsite) for manual individual grants that give single people more rights than the permission concept allows (e.g. write/full control on entire lists). First a report without changes, then you can remove the over-grants with one click. Read access always stays (including for international colleagues); writing is afterwards only via the groups and for admins/organizers.'}
          </p>
          <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }} disabled={busy !== ''} onClick={() => { setPermCleanupReport(null); setPermCleanupOpen(true); }}>
            {isDe ? 'Berechtigungen prüfen…' : 'Check permissions…'}
          </button>
        </div>

        {/* v26.81: Verwaiste Subsites prüfen (Reste gelöschter Events). */}
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Trash2 size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Subsites prüfen (verwaiste Reste)' : 'Check subsites (orphans)'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? 'Findet Event-Subsites, die noch existieren, aber zu KEINEM Event mehr gehören — z.B. Test-Subsites, deren Event bereits gelöscht wurde. Zeigt pro Rest, ob eine Teilnehmerliste (und wie viele Zeilen) vorhanden ist. Anschließend kannst du jeden Rest einzeln und bewusst löschen.'
              : 'Finds event subsites that still exist but no longer belong to any event — e.g. test subsites whose event was already deleted. Shows per orphan whether a participant list exists (and how many rows). You can then delete each orphan individually and deliberately.'}
          </p>
          <button className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }} disabled={busy !== ''} onClick={() => { setOrphanResult(null); setOrphanOpen(true); }}>
            {isDe ? 'Subsites prüfen…' : 'Check subsites…'}
          </button>
        </div>
      </div>

      {/* v24.97: E-Mails & Berichte — globale Mail-Werkzeuge (Reseed + Wochenbericht) */}
      <h2 style={{ fontSize: '1.15rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'E-Mails & Berichte' : 'Emails & reports'}</h2>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 14, marginBottom: 28 }}>
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Mail size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Default-Mail-Vorlagen zurücksetzen' : 'Reset default mail templates'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? 'Überschreibt alle Standard-Mail-Vorlagen (Anmeldung, Warteliste, Abmeldung, Nachrücken …) mit den eingebauten Texten aus dem aktuellen Stand der App. Achtung: eigene Anpassungen an den Standard-Vorlagen gehen verloren.'
              : 'Overwrites all default mail templates with the built-in texts from the current app version. Note: customizations to the standard templates are lost.'}
          </p>
          <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }} disabled={busy !== ''} onClick={() => { void doReseed(); }}>
            {busy === 'reseed' ? (isDe ? 'Wird zurückgesetzt…' : 'Resetting…') : (isDe ? 'Vorlagen zurücksetzen' : 'Reset templates')}
          </button>
        </div>
        <div style={cardStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><Mail size={18} /></span>
            <span style={{ fontWeight: 700 }}>{isDe ? 'Wochenbericht jetzt senden' : 'Send weekly report now'}</span>
          </div>
          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px', lineHeight: 1.45 }}>
            {isDe
              ? 'Löst den wöchentlichen Admin-Bericht sofort aus (überspringt die 7-Tage-Sperre) und legt ihn für alle Admins in die Mail-Warteschlange. Nur zum Testen.'
              : 'Triggers the weekly admin report immediately (bypassing the 7-day lock) and queues it for all admins. For testing only.'}
          </p>
          <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }} disabled={busy !== ''} onClick={() => { void doWeekly(); }}>
            {busy === 'weekly' ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? 'Wochenbericht senden' : 'Send weekly report')}
          </button>
        </div>
      </div>

      {/* Listen-Erklärung */}
      <h2 style={{ fontSize: '1.15rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'SharePoint-Listen — was macht was' : 'SharePoint lists — what does what'}</h2>
      <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', marginTop: 6 }}>
        {isDe ? 'Alle Hintergrund-Listen der DEX-Plattform und wofür sie da sind (Klick öffnet die Liste in SharePoint):' : 'All background lists of the DEX platform and what they are for (click opens the list in SharePoint):'}
      </p>
      <div style={{ ...cardStyle, padding: 0, overflow: 'hidden' }}>
        {LIST_DOCS.map((l, i) => (
          <div key={l.name} style={{ display: 'flex', gap: 14, flexWrap: 'wrap', padding: '12px 16px', borderTop: i === 0 ? 'none' : '1px solid var(--dex-gray-100)', alignItems: 'baseline' }}>
            <a href={listUrl(l.name)} target="_blank" rel="noopener noreferrer" style={{ flexShrink: 0, fontFamily: 'Consolas, monospace', fontSize: '0.82rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, minWidth: 150, textDecoration: 'none' }}
              onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
              onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
            >{l.name}</a>
            <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.45 }}>{l.de}</span>
          </div>
        ))}
      </div>

      {/* Release Notes / Neuerungen — vollständige, durchsuchbare Tabelle. */}
      <h2 style={{ fontSize: '1.15rem', color: 'var(--dex-green-dark, #4a7c1f)', marginTop: 28 }}>{isDe ? 'Neuerungen (Release Notes)' : 'What’s new (release notes)'}</h2>
      <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', marginTop: 6 }}>
        {isDe
          ? 'Alle Versionen — durchsuchbar und nach Bereich filterbar (neueste oben). Die lückenlose Historie ist ab v18.65 verfügbar; ältere Einträge sind die dokumentierten Meilensteine.'
          : 'All versions — searchable and filterable by area (newest first).'}
      </p>
      {/* Filter-Zeile: Suche + Bereich + Art */}
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, margin: '10px 0 12px', alignItems: 'center' }}>
        <input
          type="text"
          value={rnSearch}
          onChange={e => setRnSearch(e.target.value)}
          placeholder={isDe ? 'Suchen (Text, Bereich, Version) …' : 'Search …'}
          style={{ flex: '1 1 240px', minWidth: 180, padding: '8px 12px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.85rem' }}
        />
        <select value={rnBereich} onChange={e => setRnBereich(e.target.value)} style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.85rem' }}>
          <option value="">{isDe ? 'Alle Bereiche' : 'All areas'}</option>
          {RELEASE_BEREICHE.map(b => <option key={b} value={b}>{b}</option>)}
        </select>
        <select value={rnType} onChange={e => setRnType(e.target.value)} style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.85rem' }}>
          <option value="">{isDe ? 'Neu & Behoben' : 'All types'}</option>
          <option value="Feature">{isDe ? 'Nur Neu' : 'Features'}</option>
          <option value="Bugfix">{isDe ? 'Nur Behoben' : 'Fixes'}</option>
        </select>
        <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
          {filteredNotes.length} / {RELEASE_NOTES.length}
        </span>
      </div>
      <div style={{ ...cardStyle, padding: 0, overflow: 'auto' }}>
        {/* Tabellenkopf */}
        <div style={{ display: isMobile ? 'none' : 'grid', gridTemplateColumns: '70px 92px 150px 78px 1fr', minWidth: 720, gap: 12, padding: '10px 16px', background: 'var(--dex-gray-50, #f7f8f9)', borderBottom: '1px solid var(--dex-gray-200)', fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.3 }}>
          <span>{isDe ? 'Version' : 'Version'}</span>
          <span>{isDe ? 'Datum' : 'Date'}</span>
          <span>{isDe ? 'Bereich' : 'Area'}</span>
          <span>{isDe ? 'Art' : 'Type'}</span>
          <span>{isDe ? 'Beschreibung' : 'Description'}</span>
        </div>
        {filteredNotes.length === 0 ? (
          <div style={{ padding: '18px 16px', fontSize: '0.85rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
            {isDe ? 'Keine Treffer für diese Filter.' : 'No matches for these filters.'}
          </div>
        ) : filteredNotes.map((n, i) => (
          <div key={`${n.version}-${i}`} style={{ display: 'grid', gridTemplateColumns: isMobile ? '64px 1fr' : '70px 92px 150px 78px 1fr', minWidth: isMobile ? 0 : 720, gap: isMobile ? '2px 10px' : 12, padding: '11px 16px', borderTop: i === 0 ? 'none' : '1px solid var(--dex-gray-100)', alignItems: 'baseline' }}>
            <code style={{ fontFamily: 'Consolas, monospace', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>v{n.version}</code>
            <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{fmtDate(n.date)}</span>
            <span style={{ fontSize: '0.74rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>{n.bereich}</span>
            <span style={{ fontSize: '0.66rem', fontWeight: 700, padding: '2px 8px', borderRadius: 999, textAlign: 'center', alignSelf: 'start', background: n.type === 'Bugfix' ? 'rgba(218,41,28,0.12)' : 'rgba(134,188,37,0.15)', color: n.type === 'Bugfix' ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)' }}>
              {n.type === 'Bugfix' ? (isDe ? 'Behoben' : 'Fix') : (isDe ? 'Neu' : 'New')}
            </span>
            <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.45 }}>{n.text}</span>
          </div>
        ))}
      </div>

      {/* v26.81: Berechtigungen aufräumen — Prüf-/Korrektur-Modal. */}
      {permCleanupOpen && (
        <Modal
          open={true}
          onClose={() => { if (!permCleanupBusy) { setPermCleanupOpen(false); setPermCleanupReport(null); } }}
          dismissable={!permCleanupBusy}
          maxWidth={720}
          padding={24}
          ariaLabel={isDe ? 'Berechtigungen aufräumen' : 'Clean up permissions'}
        >
          <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem', display: 'flex', alignItems: 'center', gap: 8 }}>
            <Wrench size={18} /> {isDe ? 'Berechtigungen aufräumen' : 'Clean up permissions'}
          </h3>
          <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
            {isDe
              ? 'Der ganze SharePoint (Hauptseite, alle Listen/Bibliotheken und alle Event-Subsites) wird nach manuellen Einzel-Freigaben durchsucht, die einer Person mehr als Leserechte geben, obwohl sie laut Rollen-Konzept kein Admin/Organizer ist. Solche Über-Freigaben lassen sich hier entfernen — Leserechte und alle Gruppen-Berechtigungen bleiben unangetastet.'
              : 'The whole SharePoint (main site, all lists/libraries and every event subsite) is scanned for manual individual grants that give a person more than read access even though they are not an admin/organizer per the role concept. Such over-grants can be removed here — read access and all group permissions stay untouched.'}
          </p>

          {permCleanupBusy && permCleanupProgress && (() => {
            const { msg, done, total } = permCleanupProgress;
            const pct = Math.min(100, Math.round((done / Math.max(1, total)) * 100));
            return (
              <div style={{ marginBottom: 12 }}>
                <p style={{ margin: '0 0 6px', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>{msg}</p>
                <div style={{ background: 'var(--dex-gray-100, #f0f0f0)', borderRadius: 999, height: 10, overflow: 'hidden' }}>
                  <div style={{ width: `${pct}%`, height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease' }} />
                </div>
                <p style={{ margin: '8px 0 0', fontSize: '0.76rem', color: 'var(--dex-gray-400)' }}>
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
              <div>
                <div style={{ padding: 12, borderRadius: 'var(--dex-radius)', background: hasIssues ? 'rgba(237,139,0,0.08)' : 'rgba(134,188,37,0.10)', border: `1px solid ${hasIssues ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)'}`, marginBottom: 12, fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.6 }}>
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
                  <div style={{ maxHeight: 300, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)', marginBottom: 14 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.78rem' }}>
                      <thead>
                        <tr style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50, #fafafa)', textAlign: 'left' }}>
                          <th style={{ padding: '6px 8px', borderBottom: '1px solid var(--dex-gray-200)' }}>{isDe ? 'Ort' : 'Location'}</th>
                          <th style={{ padding: '6px 8px', borderBottom: '1px solid var(--dex-gray-200)' }}>{isDe ? 'Person' : 'Person'}</th>
                          <th style={{ padding: '6px 8px', borderBottom: '1px solid var(--dex-gray-200)' }}>{isDe ? 'Befund' : 'Finding'}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...strayFindings, ...ilsFindings, ...errFindings].slice(0, SHOW).map((f, i) => (
                          <tr key={i} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                            <td style={{ padding: '6px 8px', color: 'var(--dex-gray-600)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 200 }} title={f.scope}>{f.scope}</td>
                            <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)', wordBreak: 'break-all' }}>{f.principal || '—'}</td>
                            <td style={{ padding: '6px 8px', color: f.kind === 'error' ? 'var(--dex-red, #da291c)' : (f.fixed ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-orange-dark, #b35a00)') }}>{f.detail}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {r.findings.length > SHOW && (
                      <div style={{ padding: '6px 8px', fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>
                        {isDe ? `… und ${r.findings.length - SHOW} weitere (gekürzt).` : `… and ${r.findings.length - SHOW} more (truncated).`}
                      </div>
                    )}
                  </div>
                )}

                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                  <button className="btn btn-secondary" onClick={() => { void runPermCleanup(false); }} style={{ fontSize: '0.85rem', padding: '9px 16px' }}>
                    {isDe ? 'Erneut prüfen' : 'Re-check'}
                  </button>
                  {!r.apply && hasIssues && (
                    <button
                      className="btn btn-primary"
                      style={{ fontSize: '0.85rem', padding: '9px 18px' }}
                      onClick={() => {
                        (async () => {
                          const ok = await confirmDialog(isDe
                            ? `${r.strayWriteFound} Über-Freigabe(n) entfernen und ${r.ilsIssues} Element-Sicherheit(en) korrigieren?\n\nLeserechte und Gruppen-Berechtigungen bleiben erhalten. Der Vorgang kann je nach Größe einige Minuten dauern.`
                            : `Remove ${r.strayWriteFound} over-grant(s) and fix ${r.ilsIssues} item-security setting(s)?\n\nRead access and group permissions are preserved. Depending on size this can take a few minutes.`,
                            { confirmLabel: isDe ? 'Jetzt korrigieren' : 'Fix now' });
                          if (ok) await runPermCleanup(true);
                        })().catch(() => { /* */ });
                      }}
                    >
                      {isDe ? 'Jetzt korrigieren' : 'Fix now'}
                    </button>
                  )}
                  <button className="btn btn-outline" onClick={() => { setPermCleanupOpen(false); setPermCleanupReport(null); }} style={{ fontSize: '0.85rem', padding: '9px 16px' }}>
                    {isDe ? 'Schließen' : 'Close'}
                  </button>
                </div>
              </div>
            );
          })()}

          {!permCleanupBusy && !permCleanupReport && (
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
              <button className="btn btn-outline" onClick={() => setPermCleanupOpen(false)} style={{ fontSize: '0.85rem', padding: '9px 16px' }}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button className="btn btn-primary" onClick={() => { void runPermCleanup(false); }} style={{ fontSize: '0.85rem', padding: '9px 18px' }}>
                {isDe ? 'Prüfen (ohne Änderung)' : 'Check (no changes)'}
              </button>
            </div>
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
        >
          <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem', display: 'flex', alignItems: 'center', gap: 8 }}>
            <Trash2 size={18} /> {isDe ? 'Verwaiste Subsites' : 'Orphan subsites'}
          </h3>
          <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
            {isDe
              ? 'Subsites, die noch existieren, aber zu keinem Event mehr gehören (z.B. Test-Subsites gelöschter Events). Prüfe pro Eintrag, ob wirklich ein Rest vorliegt, bevor du löschst — das Löschen ist endgültig.'
              : 'Subsites that still exist but no longer belong to any event (e.g. test subsites of deleted events). Check each entry before deleting — deletion is permanent.'}
          </p>

          {orphanBusy && orphanProgress && (() => {
            const { msg, done, total } = orphanProgress;
            const pct = Math.min(100, Math.round((done / Math.max(1, total)) * 100));
            return (
              <div style={{ marginBottom: 12 }}>
                <p style={{ margin: '0 0 6px', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>{msg}</p>
                <div style={{ background: 'var(--dex-gray-100, #f0f0f0)', borderRadius: 999, height: 10, overflow: 'hidden' }}>
                  <div style={{ width: `${pct}%`, height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease' }} />
                </div>
              </div>
            );
          })()}

          {!orphanBusy && orphanResult && (() => {
            const r = orphanResult;
            const remaining = r.orphans.filter(o => !orphanDeleted[o.url]);
            return (
              <div>
                <div style={{ padding: 12, borderRadius: 'var(--dex-radius)', background: remaining.length > 0 ? 'rgba(237,139,0,0.08)' : 'rgba(134,188,37,0.10)', border: `1px solid ${remaining.length > 0 ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)'}`, marginBottom: 12, fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.6 }}>
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
                  <div style={{ maxHeight: 320, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)', marginBottom: 14 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.78rem' }}>
                      <thead>
                        <tr style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50, #fafafa)', textAlign: 'left' }}>
                          <th style={{ padding: '6px 8px', borderBottom: '1px solid var(--dex-gray-200)' }}>{isDe ? 'Subsite' : 'Subsite'}</th>
                          <th style={{ padding: '6px 8px', borderBottom: '1px solid var(--dex-gray-200)' }}>{isDe ? 'Teilnehmerliste' : 'Participant list'}</th>
                          <th style={{ padding: '6px 8px', borderBottom: '1px solid var(--dex-gray-200)' }} />
                        </tr>
                      </thead>
                      <tbody>
                        {remaining.map((o) => (
                          <tr key={o.url} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                            <td style={{ padding: '6px 8px', color: 'var(--dex-gray-700)' }}>
                              <div style={{ fontWeight: 600 }}>{o.title || o.serverRel}</div>
                              <a href={o.url} target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark, #4a7c1f)', wordBreak: 'break-all', fontSize: '0.72rem' }}>{o.serverRel || o.url}</a>
                            </td>
                            <td style={{ padding: '6px 8px', color: o.hasParticipantList ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-500)' }}>
                              {o.hasParticipantList
                                ? (isDe ? `ja · ${o.participantCount} Zeilen` : `yes · ${o.participantCount} rows`)
                                : (isDe ? 'keine' : 'none')}
                            </td>
                            <td style={{ padding: '6px 8px', textAlign: 'right' }}>
                              <button
                                className="btn btn-outline"
                                disabled={!!orphanDeleting[o.url]}
                                onClick={() => { void deleteOrphan(o.url, o.title || o.serverRel, o.participantCount); }}
                                style={{ fontSize: '0.74rem', padding: '5px 10px', color: 'var(--dex-red, #da291c)', borderColor: 'var(--dex-red, #da291c)' }}
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

                <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
                  <button className="btn btn-secondary" onClick={() => { void runOrphanScan(); }} style={{ fontSize: '0.85rem', padding: '9px 16px' }}>
                    {isDe ? 'Erneut prüfen' : 'Re-check'}
                  </button>
                  <button className="btn btn-outline" onClick={() => { setOrphanOpen(false); setOrphanResult(null); }} style={{ fontSize: '0.85rem', padding: '9px 16px' }}>
                    {isDe ? 'Schließen' : 'Close'}
                  </button>
                </div>
              </div>
            );
          })()}

          {!orphanBusy && !orphanResult && (
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
              <button className="btn btn-outline" onClick={() => setOrphanOpen(false)} style={{ fontSize: '0.85rem', padding: '9px 16px' }}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button className="btn btn-primary" onClick={() => { void runOrphanScan(); }} style={{ fontSize: '0.85rem', padding: '9px 18px' }}>
                {isDe ? 'Jetzt prüfen' : 'Check now'}
              </button>
            </div>
          )}
        </Modal>
      )}
    </div>
  );
}
