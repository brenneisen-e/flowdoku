/**
 * v26.28: Architektur-Übersicht der DEX-Plattform (Admin-only).
 *
 * Zeigt die Gesamtarchitektur als geschichtete Skizze: Client (SPFx-Web-Part),
 * Datenhaltung (SharePoint-Listen + Event-Subsites), Automatisierung
 * (Power-Automate-Flows) und die genutzten Microsoft-365-Dienste — inkl. der
 * wichtigsten Datenflüsse. Erreichbar über eine Kachel im Admin-Hub.
 *
 * Bewusst rein deklarativ aus statischen Daten gebaut (keine Laufzeit-Abfrage):
 * die Architektur ist Code-Wissen, nicht Tenant-Zustand. Listen sind anklickbar
 * (öffnen die SharePoint-Liste). Zweisprachig DE/EN, Fluent-Look, kein Emoji.
 */
import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { Settings, Users, Mail, FileText, Columns, Book } from './Icons';
import { downloadArchitecturePdf, ArchPdfBlock } from '../utils/architecturePdf';
import { APP_VERSION } from '../version';

type Item = { name: string; de: string; en: string; link?: boolean };

const GREEN = 'var(--dex-green, #86bc25)';
const GREEN_DARK = 'var(--dex-green-dark, #4a7c1f)';

export default function ArchitecturePage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isAdmin, originalIsAdmin, siteUrl } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;
  const listUrl = (name: string): string => `${siteUrl}/Lists/${name}`;

  React.useEffect(() => {
    if (!adminLike) navigate('start');
  }, [adminLike, navigate]);
  if (!adminLike) return <div className="page-container" />;

  // ---- Datenhaltung: SharePoint-Listen (gruppiert) ----
  const coreLists: Item[] = [
    { name: 'DEX_Events', de: 'Ein Eintrag pro Event/Sub-Event (Titel, Datum, Ort, Sichtbarkeit, Kapazität, Kommunikation).', en: 'One entry per event/sub-event (title, date, location, visibility, capacity, communication).', link: true },
    { name: 'DEX_Roles', de: 'Rollen: User, Organizer, Admin.', en: 'Roles: user, organizer, admin.', link: true },
    { name: 'DEX_Participants', de: 'Übergreifendes Personen-Register (Statistik/KPI).', en: 'Cross-event person register (stats/KPI).', link: true },
    { name: 'DEX_EmailTemplates', de: 'Anpassbare Mail-Vorlagen + zentrale Konfiguration (Logo, Standardbild, KPI-Zähler).', en: 'Customizable mail templates + central config (logo, default image, KPI counter).', link: true },
    { name: 'DEX_Logos', de: 'Hinterlegte Logos/Bilder für Mails und Outlook.', en: 'Stored logos/images for mails and Outlook.', link: true },
  ];
  const subsiteItems: Item[] = [
    { name: '„Teilnehmer" (pro Event)', de: 'Eigene Subsite je Event mit der Anmeldeliste — Item-Level-Security: jeder sieht nur seine eigene Anmeldung.', en: 'Own subsite per event with the registration list — item-level security: each user sees only their own entry.' },
    { name: 'Event-Bild & -Dokumente', de: 'Als Anhänge am Event-/Teilnehmer-Element gespeichert.', en: 'Stored as attachments on the event/attendee item.' },
  ];
  const queueLists: Item[] = [
    { name: 'DEX_Emails', de: 'Warteschlange ausgehender Mails (Bestätigung, Erinnerung, Bericht …).', en: 'Outbound mail queue (confirmation, reminder, report …).', link: true },
    { name: 'DEX_Outlook', de: 'Kalender-Aufträge: ein-/ausladen, Termin aktualisieren/löschen.', en: 'Calendar jobs: add/remove attendees, update/delete event.', link: true },
    { name: 'DEX_OutlookLocks', de: 'Sperre, damit zwei Kalender-Läufe desselben Events sich nicht überschneiden.', en: 'Lock so two calendar runs for the same event do not collide.', link: true },
    { name: 'DEX_IDReorder', de: 'Aufträge zum Neu-Nummerieren der Teilnehmer-IDs + Nachrücken von der Warteliste.', en: 'Jobs to renumber attendee IDs + promote from the waitlist.', link: true },
    { name: 'DEX_AccessFix', de: 'Setzt bei stellvertretenden Anmeldungen den richtigen Zeilen-Autor.', en: 'Sets the correct row author for proxy registrations.', link: true },
    { name: 'DEX_AssistantAccess', de: 'Delegation/Zugriff für die „Meine Assistenz"-Funktion.', en: 'Delegation/access for the “my assistant” feature.', link: true },
    { name: 'DEX_TeamJoinRequests', de: 'Beitritts-Anfragen für Team-Anmeldungen (Kapitän bestätigt).', en: 'Join requests for team registrations (captain approves).', link: true },
    { name: 'DEX_OrganizerRequests', de: 'Anträge „Organizer werden" zur Admin-Freigabe.', en: '“Become organizer” requests for admin approval.', link: true },
    { name: 'DEX_Tickets', de: 'Fragen & Antworten (Support-Tickets) zu Events und zur App.', en: 'Questions & answers (support tickets) for events and the app.', link: true },
  ];
  const logLists: Item[] = [
    { name: 'DEX_TeilnehmerCounter', de: 'Atomarer Zähler für eindeutige Teilnehmer-IDs + Sitzplatz-Reservierung.', en: 'Atomic counter for unique attendee IDs + seat reservation.', link: true },
    { name: 'DEX_ChangeLog', de: 'Änderungsprotokoll (Audit): wer hat wann was getan.', en: 'Audit log: who did what and when.', link: true },
    { name: 'DEX_WeeklyReports', de: 'Protokoll des Wochenberichts (Doppelversand-Schutz).', en: 'Weekly-report log (duplicate-send protection).', link: true },
    { name: 'DEX_InactiveNotices', de: 'Merkt vorgenommene „Konto inaktiv"-Hinweise an Organizer.', en: 'Tracks “account inactive” notices sent to organizers.', link: true },
    { name: 'DEX_OrganizerArchived', de: 'Pro Organizer ausgeblendete (archivierte) Events.', en: 'Per-organizer hidden (archived) events.', link: true },
    { name: 'DEX_Archive', de: 'End-Ablage für alte Queue-/Log-Zeilen (hält die Arbeitslisten schlank).', en: 'Final store for old queue/log rows (keeps working lists lean).', link: true },
  ];

  // ---- Automatisierung: Power-Automate-Flows (Trigger → Aktion) ----
  const flows: Item[] = [
    { name: 'DEX_SEND_MAIL', de: 'Neues DEX_Emails-Element → versendet die Mail über das Sammelpostfach, setzt Status „Sent".', en: 'New DEX_Emails item → sends the mail via the shared mailbox, sets status “Sent”.' },
    { name: 'DEX_CreateOutlookEvent', de: 'Neues DEX_Events-Element → legt den Outlook-Kalendertermin an.', en: 'New DEX_Events item → creates the Outlook calendar event.' },
    { name: 'DEX_Outlook_Einladungen', de: 'Neues DEX_Outlook-Element → Teilnehmer ein-/ausladen, Termin aktualisieren/löschen (pro Event gesperrt).', en: 'New DEX_Outlook item → add/remove attendees, update/delete event (per-event lock).' },
    { name: 'DEX_IDReorder_TeilnehmerIDs', de: 'Neues DEX_IDReorder-Element → nummeriert IDs neu und rückt von der Warteliste nach.', en: 'New DEX_IDReorder item → renumbers IDs and promotes from the waitlist.' },
    { name: 'DEX_OutlookDeclineHandler', de: 'Termin-Absage im Postfach → Erinnerungs-Mail bzw. Auto-Abmeldung.', en: 'Calendar decline in the mailbox → reminder mail or auto-deregistration.' },
    { name: 'DEX_OutlookForwardHandler', de: 'Weitergeleitete Einladung → FYI an Organizer, wenn der Empfänger nicht angemeldet ist.', en: 'Forwarded invitation → FYI to organizer if the recipient is not registered.' },
    { name: 'DEX_AccessFix_Autor', de: 'Neues DEX_AccessFix-Element → setzt den Zeilen-Autor auf den Teilnehmer.', en: 'New DEX_AccessFix item → sets the row author to the attendee.' },
    { name: 'DEX_AssistantAccess_Grant', de: 'Neues DEX_AssistantAccess-Element → richtet den Assistenz-Zugriff ein.', en: 'New DEX_AssistantAccess item → grants the assistant access.' },
  ];

  // ---- Microsoft-365-Dienste ----
  const services: Item[] = [
    { name: 'Exchange / Outlook', de: 'Sammelpostfach no_reply.events@deloitte.de — versendet Mails und verwaltet Kalendereinladungen.', en: 'Shared mailbox no_reply.events@deloitte.de — sends mails and manages calendar invitations.' },
    { name: 'Microsoft Graph', de: 'Profil-, Gruppen-, Standort- und Kalender-Abfragen (Sichtbarkeit, People-Picker, Konto-Aktiv-Check).', en: 'Profile, group, location and calendar lookups (visibility, people picker, account-active check).' },
  ];

  // ---- Typische Datenflüsse (auch im PDF) ----
  const dataFlows: Item[] = [
    { name: isDe ? 'Anmeldung' : 'Registration', de: 'App reserviert atomar einen Platz (DEX_TeilnehmerCounter), schreibt das Teilnehmer-Item in die Event-Subsite, legt die Bestätigungs-Mail (DEX_Emails) und die Kalendereinladung (DEX_Outlook) an → DEX_SEND_MAIL versendet, DEX_Outlook_Einladungen lädt ein.', en: 'The app atomically reserves a seat (DEX_TeilnehmerCounter), writes the attendee item into the event subsite, queues the confirmation mail (DEX_Emails) and the calendar invite (DEX_Outlook) → DEX_SEND_MAIL sends, DEX_Outlook_Einladungen invites.' },
    { name: isDe ? 'Abmeldung & Nachrücken' : 'Cancellation & promotion', de: 'App setzt den Status, legt einen DEX_IDReorder-Auftrag an → der Flow nummeriert die IDs neu und rückt die erste Person von der Warteliste nach (inkl. Mail).', en: 'The app sets the status and queues a DEX_IDReorder job → the flow renumbers IDs and promotes the first person from the waitlist (incl. mail).' },
    { name: isDe ? 'Frage stellen (Ticket)' : 'Ask a question (ticket)', de: 'App legt ein DEX_Tickets-Element an und benachrichtigt Organizer bzw. Power-User per DEX_Emails — Antwort und Eingangs-Bestätigung laufen über dieselbe Mail-Warteschlange.', en: 'The app creates a DEX_Tickets item and notifies organizers or power users via DEX_Emails — answer and acknowledgement run through the same mail queue.' },
    { name: isDe ? 'Sichtbarkeit & People-Picker' : 'Visibility & people picker', de: 'App fragt Microsoft Graph nach Profil, Gruppenzugehörigkeit und Standort, um zu entscheiden, wer ein Event sieht, und um Personen auszuwählen.', en: 'The app queries Microsoft Graph for profile, group membership and location to decide who sees an event and to pick people.' },
    { name: isDe ? 'Co-Organizer-Freigabe' : 'Co-organizer approval', de: 'Wird jemand als Co-Organizer benannt, der noch kein Organizer ist, legt die App einen DEX_OrganizerRequests-Antrag an und mailt den Admins einen Freigabe-Link; nach der Freigabe erhält die Person Schreibrechte auf DEX_Events.', en: 'If someone is named co-organizer who is not yet an organizer, the app creates a DEX_OrganizerRequests entry and emails admins an approval link; after approval the person receives write access to DEX_Events.' },
  ];

  const [pdfBusy, setPdfBusy] = React.useState(false);
  const toPdf = (items: Item[]): Array<{ name: string; desc: string }> => items.map(i => ({ name: i.name, desc: isDe ? i.de : i.en }));
  const downloadPdf = async (): Promise<void> => {
    if (pdfBusy) return;
    setPdfBusy(true);
    try {
      const blocks: ArchPdfBlock[] = [
        { heading: isDe ? '1 · Nutzer & App (Client)' : '1 · Users & app (client)', items: toPdf([{ name: 'DEX Event Experience Platform', de: 'SPFx-Web-Part (React), eingebettet in SharePoint Online. Läuft im Browser des Nutzers; Anmeldung automatisch über Microsoft 365 (Single Sign-on). Liest und schreibt über die SharePoint-REST-API und Microsoft Graph.', en: 'SPFx web part (React) embedded in SharePoint Online. Runs in the user’s browser; sign-in automatically via Microsoft 365 (single sign-on). Reads and writes via the SharePoint REST API and Microsoft Graph.' }]) },
        { heading: isDe ? '2 · Daten — Kerndaten (SharePoint Online)' : '2 · Data — core data (SharePoint Online)', items: toPdf(coreLists) },
        { heading: isDe ? '2 · Daten — pro Event (Subsite)' : '2 · Data — per event (subsite)', items: toPdf(subsiteItems) },
        { heading: isDe ? '2 · Daten — Warteschlangen' : '2 · Data — queues', note: isDe ? 'Die App füllt diese Listen, die Power-Automate-Flows verarbeiten sie (Trigger: neues Element).' : 'The app fills these lists, the Power Automate flows process them (trigger: new item).', items: toPdf(queueLists) },
        { heading: isDe ? '2 · Daten — Zähler, Protokolle & Archiv' : '2 · Data — counter, logs & archive', items: toPdf(logLists) },
        { heading: isDe ? '3 · Automatisierung — Power-Automate-Flows' : '3 · Automation — Power Automate flows', note: isDe ? 'Jeder Flow wird durch ein neues Element in seiner Auslöser-Liste gestartet (Trigger → Aktion).' : 'Each flow is started by a new item in its trigger list (trigger → action).', items: toPdf(flows) },
        { heading: isDe ? '4 · Microsoft-365-Dienste (im Tenant)' : '4 · Microsoft 365 services (in tenant)', items: toPdf(services) },
        { heading: isDe ? 'Typische Datenflüsse' : 'Typical data flows', items: toPdf(dataFlows) },
      ];
      await downloadArchitecturePdf({
        isDe,
        version: APP_VERSION,
        title: isDe ? 'Architektur der DEX-Plattform' : 'DEX platform architecture',
        intro: isDe
          ? 'Die DEX Event Experience Platform läuft vollständig im Microsoft-365-Tenant von Deloitte — es werden keine externen Dienste genutzt. Dieses Dokument beschreibt die Gesamtarchitektur: die App, die Datenhaltung in SharePoint (alle Listen mit ihrer Funktion und die Event-Subsites), die Hintergrund-Automatisierung über Power Automate, die genutzten Microsoft-365-Dienste sowie die wichtigsten Datenflüsse und Zusammenhänge.'
          : 'The DEX Event Experience Platform runs entirely within Deloitte’s Microsoft 365 tenant — no external services are used. This document describes the overall architecture: the app, the data storage in SharePoint (all lists with their function and the event subsites), the background automation via Power Automate, the Microsoft 365 services used, and the key data flows and relationships.',
        blocks,
      });
    } catch (e) { console.warn('[DEX] Architektur-PDF fehlgeschlagen:', e); }
    finally { setPdfBusy(false); }
  };

  // ---- Styles ----
  const band: React.CSSProperties = { border: `1px solid var(--dex-gray-200)`, borderRadius: 12, background: '#fff', boxShadow: '0 2px 8px rgba(0,0,0,0.05)', overflow: 'hidden', marginBottom: 6 };
  const bandHead: React.CSSProperties = { display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', padding: '10px 16px', background: 'var(--dex-gray-50, #f7f8f9)', borderBottom: '1px solid var(--dex-gray-200)' };
  const bandTitle: React.CSSProperties = { fontWeight: 700, color: GREEN_DARK, fontSize: '1rem' };
  const bandSub: React.CSSProperties = { fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginLeft: 'auto', textAlign: 'right' };
  const grid: React.CSSProperties = { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(230px, 1fr))', gap: 10, padding: 14 };
  const box: React.CSSProperties = { border: `1px solid var(--dex-gray-200)`, borderRadius: 9, padding: '10px 12px', background: '#fff' };
  const boxName: React.CSSProperties = { fontFamily: 'Consolas, monospace', fontSize: '0.82rem', fontWeight: 700, color: GREEN_DARK, display: 'block', marginBottom: 3, wordBreak: 'break-word' };
  const boxDesc: React.CSSProperties = { fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.4 };
  const groupLabel: React.CSSProperties = { gridColumn: '1 / -1', fontSize: '0.72rem', fontWeight: 700, letterSpacing: 0.3, textTransform: 'uppercase', color: 'var(--dex-gray-500)', marginTop: 4 };

  const connector = (text: string): React.ReactElement => (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', margin: '2px 0' }}>
      <span style={{ color: GREEN, fontSize: 20, lineHeight: 1 }}>↓</span>
      <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', textAlign: 'center', maxWidth: 620 }}>{text}</span>
    </div>
  );

  const renderBox = (it: Item): React.ReactElement => {
    const inner = (
      <>
        <span style={boxName}>{it.name}</span>
        <span style={boxDesc}>{isDe ? it.de : it.en}</span>
      </>
    );
    if (it.link) {
      return (
        <a key={it.name} href={listUrl(it.name)} target="_blank" rel="noopener noreferrer"
          style={{ ...box, textDecoration: 'none', display: 'block', transition: 'border-color 0.15s, box-shadow 0.15s' }}
          onMouseEnter={e => { e.currentTarget.style.borderColor = GREEN; e.currentTarget.style.boxShadow = '0 4px 12px rgba(134,188,37,0.18)'; }}
          onMouseLeave={e => { e.currentTarget.style.borderColor = 'var(--dex-gray-200)'; e.currentTarget.style.boxShadow = 'none'; }}
        >{inner}</a>
      );
    }
    return <div key={it.name} style={box}>{inner}</div>;
  };

  return (
    <div className="page-container">
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center', marginBottom: 14 }}>
        <button className="btn btn-secondary" onClick={() => navigate('admin-hub')}>
          {isDe ? '← Zurück zum Admin' : '← Back to admin'}
        </button>
        <button className="btn btn-primary" style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 'auto' }} disabled={pdfBusy} onClick={() => { void downloadPdf(); }}>
          <FileText size={16} /> {pdfBusy ? (isDe ? 'PDF wird erstellt…' : 'Generating PDF…') : (isDe ? 'Als PDF herunterladen' : 'Download as PDF')}
        </button>
      </div>
      <h1 style={{ marginTop: 0 }}>{isDe ? 'Architektur der DEX-Plattform' : 'DEX platform architecture'}</h1>
      <p style={{ color: 'var(--dex-gray-600)', marginTop: 0, maxWidth: 820 }}>
        {isDe
          ? 'Die Plattform läuft vollständig im Microsoft-365-Tenant von Deloitte — keine externen Dienste. Oben die App, darunter die Daten in SharePoint, die Hintergrund-Automatisierung (Power Automate) und die genutzten Microsoft-365-Dienste.'
          : 'The platform runs entirely within Deloitte’s Microsoft 365 tenant — no external services. From top to bottom: the app, the data in SharePoint, the background automation (Power Automate) and the Microsoft 365 services used.'}
      </p>

      {/* Layer 1 — Client */}
      <div style={band}>
        <div style={bandHead}>
          <span style={{ color: GREEN, display: 'inline-flex' }}><Users size={20} /></span>
          <span style={bandTitle}>{isDe ? '1 · Nutzer & App (Client)' : '1 · Users & app (client)'}</span>
          <span style={bandSub}>{isDe ? 'Browser & SharePoint-Mobile · Anmeldung über Microsoft 365' : 'Browser & SharePoint mobile · sign-in via Microsoft 365'}</span>
        </div>
        <div style={grid}>
          {renderBox({ name: 'DEX Event Experience Platform', de: 'SPFx-Web-Part (React), eingebettet in SharePoint Online. Läuft im Browser des Nutzers; Anmeldung automatisch über Microsoft 365 (Single Sign-on).', en: 'SPFx web part (React) embedded in SharePoint Online. Runs in the user’s browser; sign-in automatically via Microsoft 365 (single sign-on).' })}
        </div>
      </div>
      {connector(isDe ? 'liest & schreibt über die SharePoint-REST-API und Microsoft Graph' : 'reads & writes via the SharePoint REST API and Microsoft Graph')}

      {/* Layer 2 — Daten */}
      <div style={band}>
        <div style={bandHead}>
          <span style={{ color: GREEN, display: 'inline-flex' }}><FileText size={20} /></span>
          <span style={bandTitle}>{isDe ? '2 · Datenhaltung — SharePoint Online' : '2 · Data — SharePoint Online'}</span>
          <span style={bandSub}>{isDe ? 'Site-Collection + Subsite pro Event' : 'Site collection + one subsite per event'}</span>
        </div>
        <div style={grid}>
          <span style={groupLabel}>{isDe ? 'Kerndaten' : 'Core data'}</span>
          {coreLists.map(renderBox)}
          <span style={groupLabel}>{isDe ? 'Pro Event (Subsite)' : 'Per event (subsite)'}</span>
          {subsiteItems.map(renderBox)}
          <span style={groupLabel}>{isDe ? 'Warteschlangen (App füllt, Flows verarbeiten)' : 'Queues (app fills, flows process)'}</span>
          {queueLists.map(renderBox)}
          <span style={groupLabel}>{isDe ? 'Zähler, Protokolle & Archiv' : 'Counter, logs & archive'}</span>
          {logLists.map(renderBox)}
        </div>
      </div>
      {connector(isDe ? 'ein neues Element in einer Warteschlange löst den passenden Flow aus (Trigger: „neues Element")' : 'a new item in a queue triggers the matching flow (trigger: “new item”)')}

      {/* Layer 3 — Automatisierung */}
      <div style={band}>
        <div style={bandHead}>
          <span style={{ color: GREEN, display: 'inline-flex' }}><Settings size={20} /></span>
          <span style={bandTitle}>{isDe ? '3 · Automatisierung — Power Automate' : '3 · Automation — Power Automate'}</span>
          <span style={bandSub}>{isDe ? 'Hintergrund-Flows im Tenant' : 'background flows in the tenant'}</span>
        </div>
        <div style={grid}>
          {flows.map(renderBox)}
        </div>
      </div>
      {connector(isDe ? 'die Flows rufen Microsoft-365-Dienste auf (und schreiben Status in die Listen zurück)' : 'the flows call Microsoft 365 services (and write status back into the lists)')}

      {/* Layer 4 — M365-Dienste */}
      <div style={band}>
        <div style={bandHead}>
          <span style={{ color: GREEN, display: 'inline-flex' }}><Mail size={20} /></span>
          <span style={bandTitle}>{isDe ? '4 · Microsoft-365-Dienste (im Tenant)' : '4 · Microsoft 365 services (in tenant)'}</span>
          <span style={bandSub}>{isDe ? 'keine externen Dienste' : 'no external services'}</span>
        </div>
        <div style={grid}>
          {services.map(renderBox)}
        </div>
      </div>

      {/* Datenfluss-Beispiele */}
      <h2 style={{ fontSize: '1.1rem', color: GREEN_DARK, marginTop: 26 }}>{isDe ? 'Typische Datenflüsse' : 'Typical data flows'}</h2>
      <div style={{ ...band, marginBottom: 20 }}>
        {dataFlows.map((f, i) => (
          <div key={i} style={{ display: 'flex', gap: 12, flexWrap: 'wrap', padding: '11px 16px', borderTop: i === 0 ? 'none' : '1px solid var(--dex-gray-100)', alignItems: 'baseline' }}>
            <span style={{ flexShrink: 0, minWidth: 150, fontWeight: 700, color: GREEN_DARK, fontSize: '0.85rem' }}>{f.name}</span>
            <span style={{ fontSize: '0.83rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>{isDe ? f.de : f.en}</span>
          </div>
        ))}
      </div>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
        <button className="btn btn-secondary" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => navigate('flowcharts')}>
          <Columns size={16} /> {isDe ? 'Detaillierte Prozessübersicht' : 'Detailed process overview'}
        </button>
        <button className="btn btn-secondary" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => navigate('manual')}>
          <Book size={16} /> {isDe ? 'Handbuch' : 'Manual'}
        </button>
      </div>
    </div>
  );
}
