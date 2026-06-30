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
import { Settings, Users, Mail, Book, FileText, Trash2, Columns } from './Icons';
import { RELEASE_NOTES, RELEASE_BEREICHE } from '../data/releaseNotes';

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
  const { isAdmin, originalIsAdmin, siteUrl } = useRoles();
  const listUrl = (name: string): string => `${siteUrl}/Lists/${name}`;
  const { getArchivableCount, runArchiveExpired, getDeletableArchiveCount, runDeleteOldArchive, fixAllEventColumns, restoreCustomFieldDescriptions, reseedDefaultEmailTemplates, maybeSendWeeklyReport } = useEvents();
  const { locale } = useLanguage();
  const { confirmDialog, showAlert } = useDialog();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;

  const [archTotal, setArchTotal] = React.useState(0);
  const [delTotal, setDelTotal] = React.useState(0);
  const [busy, setBusy] = React.useState<'' | 'arch' | 'del' | 'fixcols' | 'restoredesc' | 'reseed' | 'weekly'>('');
  // v24.33: Fortschritt für das globale „Spalten fixen".
  const [fixProgress, setFixProgress] = React.useState<{ done: number; total: number; label: string } | null>(null);
  const [restoreProgress, setRestoreProgress] = React.useState<{ done: number; total: number; label: string } | null>(null);
  const [restorePreview, setRestorePreview] = React.useState<Array<{ eventId: string; eventTitle: string; fields: Array<{ label: string; props: string[] }> }> | null>(null);
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

  if (!adminLike) return <div className="page-container" />;

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

  const tools: Array<{ icon: React.ReactNode; title: string; desc: string; onClick: () => void }> = [
    { icon: <Users size={28} />, title: isDe ? 'Organizer Center' : 'Organizer center', desc: isDe ? 'Teilnehmer, Prozesse, Audit-Log, SharePoint-Liste — alle Event-Werkzeuge pro Event.' : 'Attendees, processes, audit log, SharePoint list — all per-event tools.', onClick: () => navigate('admin') },
    { icon: <Settings size={28} />, title: isDe ? 'Prozessübersicht' : 'Process overview', desc: isDe ? 'Wie die Abläufe in DEX funktionieren — verständlich erklärt.' : 'How the DEX processes work — explained simply.', onClick: () => navigate('flowcharts') },
    { icon: <FileText size={28} />, title: isDe ? 'Architektur' : 'Architecture', desc: isDe ? 'Gesamtarchitektur (App, SharePoint-Listen, Power-Automate-Flows, M365-Dienste) — mit PDF-Export.' : 'Overall architecture (app, SharePoint lists, Power Automate flows, M365 services) — with PDF export.', onClick: () => navigate('architecture') },
    { icon: <Users size={28} />, title: isDe ? 'Rollenverwaltung' : 'Role management', desc: isDe ? 'User, Organizer und Admins zuweisen oder entfernen.' : 'Assign or remove users, organizers and admins.', onClick: () => navigate('settings') },
    { icon: <Columns size={28} />, title: isDe ? 'Rollenmatrix' : 'Role matrix', desc: isDe ? 'Übersicht: wer welche Rechte hat (User, Organizer, Admin).' : 'Overview: who has which permissions (user, organizer, admin).', onClick: () => navigate('role-matrix') },
    { icon: <Mail size={28} />, title: isDe ? 'Mail-Vorlagen' : 'Mail templates', desc: isDe ? 'Globale Standard-Mails (Anmeldung, Warteliste, Abmeldung …) ansehen und bearbeiten — mit Live-Vorschau.' : 'View and edit the global default emails — with live preview.', onClick: () => navigate('email-templates') },
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
          <div key={l.name} style={{ display: 'flex', gap: 14, padding: '12px 16px', borderTop: i === 0 ? 'none' : '1px solid var(--dex-gray-100)', alignItems: 'baseline' }}>
            <a href={listUrl(l.name)} target="_blank" rel="noopener noreferrer" style={{ flexShrink: 0, fontFamily: 'Consolas, monospace', fontSize: '0.82rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, minWidth: 190, textDecoration: 'none' }}
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
      <div style={{ ...cardStyle, padding: 0, overflow: 'hidden' }}>
        {/* Tabellenkopf */}
        <div style={{ display: 'grid', gridTemplateColumns: '70px 92px 150px 78px 1fr', gap: 12, padding: '10px 16px', background: 'var(--dex-gray-50, #f7f8f9)', borderBottom: '1px solid var(--dex-gray-200)', fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.3 }}>
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
          <div key={`${n.version}-${i}`} style={{ display: 'grid', gridTemplateColumns: '70px 92px 150px 78px 1fr', gap: 12, padding: '11px 16px', borderTop: i === 0 ? 'none' : '1px solid var(--dex-gray-100)', alignItems: 'baseline' }}>
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
    </div>
  );
}
