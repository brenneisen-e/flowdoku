import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import AdminPage from '../../AdminPage';

export function manageParticipantsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'manage-participants',
    title: isDe ? 'Teilnehmer verwalten (Admin Center)' : 'Manage participants (Admin Center)',
    category: 'organizer',
    description: isDe
      ? 'Liste aller Teilnehmer einsehen, Status ändern, E-Mails kopieren, manuell einchecken.'
      : 'View the full participant list, change statuses, copy emails, check in manually.',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Admin Center öffnen' : 'Open the admin center',
            description: (
              <>
                {isDe
                  ? 'Im Admin-Bereich → Karte deines Events → "Admin Center". Du siehst oben vier Metrik-Kacheln (Angemeldet, Warteliste, Eingecheckt, Abgemeldet) und darunter Aktions-Buttons.'
                  : 'Admin area → your event card → "Admin Center". You\'ll see four metric tiles (Registered, Waitlist, Checked-in, Cancelled) and action buttons below.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Admin Center → KPI-Kacheln + Aktionen (echte Ansicht)' : 'Admin center → KPI tiles + actions (real view)'}
                role="Organizer"
                page="admin"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
              >
                <Header />
                <AdminPage />
              </AppPreview>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Teilnehmer durchsuchen & sortieren' : 'Search & sort participants',
            description: (
              <>
                {isDe
                  ? 'Über das Suchfeld filterst du nach Name, E-Mail oder TeilnehmerID. Mit Klick auf eine Spaltenüberschrift sortierst du auf- oder absteigend.'
                  : 'Use the search field to filter by name, email or attendee ID. Click a column header to sort ascending / descending.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Admin Center → Teilnehmertabelle mit Suche + Sortierung' : 'Admin center → participant table with search + sort'}
                role="Organizer"
                page="admin"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
              >
                <Header />
                <AdminPage />
              </AppPreview>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Status manuell ändern' : 'Change status manually',
            description: (
              <>
                {isDe
                  ? 'Je Teilnehmer-Zeile gibt es Aktionen: Ein-/Auschecken, Abmelden, Warteliste-Verschiebung. Der Organizer kann alle Aktionen für eigene Events durchführen; Admin für alle Events.'
                  : 'Each attendee row has actions: check in / out, cancel, move between registered / waitlist. Organizers can do all actions for their own events; admins for all events.'}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'E-Mail-Adressen kopieren' : 'Copy email addresses',
            description: (
              <>
                {isDe
                  ? 'Der Button "E-Mails kopieren" legt alle aktiven Teilnehmer-Mails semikolon-getrennt in die Zwischenablage. Direkt einfügbar in Outlook oder eine externe Mailinglist-App.'
                  : 'The "Copy emails" button puts all active attendee emails semicolon-separated into your clipboard. Paste directly into Outlook or an external mailing list app.'}
              </>
            ),
            mockup: (
              <Callout variant="tip">
                {isDe
                  ? 'Abgemeldete Teilnehmer werden NICHT exportiert — nur Angemeldete, Eingecheckte und Wartelistler.'
                  : 'Cancelled attendees are NOT exported — only Registered, Checked-in, and Waitlisted.'}
              </Callout>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Teilnehmer in der App bearbeiten' : 'Edit attendee in the app',
            description: (
              <>
                {isDe
                  ? 'Neben jedem Teilnehmer gibt es einen "Bearbeiten"-Button. Klick darauf öffnet ein Modal, in dem du Vor-/Nachname und E-Mail-Adresse korrigieren sowie alle event-spezifischen Felder anpassen kannst — ohne Umweg über die SharePoint-Liste. Praktisch z.B. wenn jemand einen Tippfehler in der Anmeldung hatte oder die falsche Mail-Adresse hinterlegt wurde. Beim Speichern einer geänderten E-Mail prüft die App, ob die Adresse zum Deloitte-Tenant gehört und die Person dort wirklich existiert — externe Adressen werden abgewiesen, bei Tippfehlern kommt eine klare Fehlermeldung statt einer kaputten Anmeldung. Phone, Department, Standort und Job Title kommen aus dem M365-Profil und werden bei einer Mail-Änderung automatisch nachgezogen. Jede Änderung wird sowohl im "ChangeLog" der Teilnehmer-Zeile als auch im zentralen Audit-Log mit Datum, deinem Namen und Vorher-/Nachher-Wert protokolliert.'
                  : 'Every attendee row has an "Edit" button. Clicking it opens a modal where you can fix first/last name and email and adjust all event-specific fields — no detour via the SharePoint list. Useful e.g. when someone had a typo during registration or the wrong email was used. When saving a changed email the app verifies that the address belongs to the Deloitte tenant and that the person actually exists there — external addresses are rejected, typos yield a clear error message instead of a broken registration. Phone, Department, Location and Job Title come from the M365 profile and are refreshed automatically when the email changes. Every change is logged both in the row\'s "ChangeLog" and in the central audit log with date, your name and before/after value.'}
              </>
            ),
            tip: isDe
              ? 'Vorteil gegenüber direktem Editieren in SharePoint: keine Datums-Format-Probleme, automatische Audit-Spur, Tenant-Validierung der Mail-Adresse, keine Schulung der Organizer auf SP-UI nötig.'
              : 'Advantage over direct editing in SharePoint: no date-format issues, automatic audit trail, tenant validation of the email address, no need to train organizers on the SP UI.',
          },
          {
            number: 6,
            title: isDe ? 'Aktualisieren-Button' : 'Refresh button',
            description: (
              <>
                {isDe
                  ? 'Oben rechts in jeder Übersicht (Eventliste, MyEvents, Admin Center) gibt es einen "Aktualisieren"-Button mit Refresh-Icon. Damit kannst du die Daten frisch aus SharePoint laden, ohne die App neu starten zu müssen — sinnvoll z.B. wenn ein anderer Organizer parallel etwas geändert hat oder wenn der Power-Automate-Flow gerade einen Nachrücker promotet hat und du das Ergebnis sehen willst. Im Admin Center werden bei selektiertem Event auch dessen Teilnehmer neu geladen.'
                  : 'In every overview (event list, MyEvents, admin center) there\'s a "Refresh" button with a reload icon at the top right. It pulls fresh data from SharePoint without restarting the app — useful e.g. when another organizer has changed something in parallel or when the Power Automate flow has just promoted a waitlister and you want to see the result. In the admin center with a selected event, that event\'s attendees are also reloaded.'}
              </>
            ),
          },
          {
            number: 7,
            title: isDe ? 'Deep-Link auf ein Event teilen' : 'Share a deep-link to an event',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center eines Events gibt es bei den Aktionen-Kacheln den Button "Deep-Link kopieren". Ein Klick legt den direkten Link auf genau diese Detail-Ansicht (?action=admin&event=<ID>) in die Zwischenablage. Den Link kannst du per Mail oder Teams an Co-Organizer, Helfer oder Check-In-Team schicken — sie landen nach Login direkt auf der gleichen Admin-Seite, ohne sich erst durch die Event-Auswahl klicken zu müssen. Funktioniert für Organizer (eigene Events) und Admin (alle Events). Hat der Empfänger keine Berechtigung für das Event, sieht er stattdessen die normale Event-Liste.'
                  : 'On every event\'s admin center the action tiles include a "Copy deep-link" button. One click puts the direct link to exactly this detail view (?action=admin&event=<ID>) into your clipboard. You can send the link by mail or Teams to co-organizers, helpers, or the check-in team — after login they land directly on the same admin page, without first navigating through the event picker. Available to organizers (own events) and admins (all events). If the recipient has no permission for the event, they see the regular event list instead.'}
              </>
            ),
            tip: isDe
              ? 'Der Page-ID-Indikator im Header-Avatar-Popup zeigt "admin-event" wenn du auf der Detail-Seite bist und "admin-center" auf der Event-Auswahl — praktisch, wenn du in einem Bug-Report exakt sagen willst, wo du gerade warst.'
              : 'The Page-ID indicator in the header avatar popup shows "admin-event" when you\'re on the detail view and "admin-center" on the event picker — handy when filing a bug report so you can name exactly where you were.',
          },
          {
            number: 8,
            title: isDe ? 'Events mit Sub-Events: konsolidierte Teilnehmer-Ansicht' : 'Events with sub-events: consolidated participant view',
            description: (
              <>
                {isDe
                  ? 'Bei einem Event mit Sub-Events (Sections) siehst du im Admin Center eine konsolidierte Tabelle: eine Zeile pro Person, mit Spalten je Sub-Event („angemeldet?") und den übergreifenden Hauptevent-Feldern. Seit v19.30 kannst du hier pro Person (1) die übergreifenden Hauptevent-Felder über den „Felder"-Button direkt bearbeiten und (2) über den „Abmelden"-Button ein Modal öffnen, in dem du auswählst, von welchen Sub-Events — oder von allen — die Person abgemeldet werden soll. Pro gewähltem Sub-Event laufen Abmelde-Mail, Outlook-Absage und Nachrücken automatisch (die Schalter des jeweiligen Sub-Events werden respektiert). In der „Abmeldungen"-Liste darunter kannst du einzelne abgemeldete Registrierungen nach einer Sicherheits-Abfrage endgültig löschen (z.B. Test-Anmeldungen).'
                  : 'For an event with sub-events (sections), the admin center shows a consolidated table: one row per person, with a column per sub-event („registered?") plus the cross-cutting main-event fields. Since v19.30 you can, per person, (1) edit the cross-cutting main-event fields directly via the „Fields" button and (2) open a modal via the „Cancel" button to choose which sub-events — or all — to deregister the person from. For each selected sub-event, the cancellation email, Outlook removal and waitlist promotion run automatically (each sub-event\'s switches are respected). In the „Cancellations" list below you can permanently delete individual cancelled registrations after a safety prompt (e.g. test registrations).'}
              </>
            ),
            tip: isDe
              ? 'Jede Abmeldung, Löschung und Feld-Änderung landet im Audit-Log. Öffne es über die Aktion „Audit-Log / Änderungsprotokoll" — es ist dann direkt auf dieses Event vorgefiltert und zeigt pro Eintrag wer, wann, was (bei Daten-Änderungen Vorher → Nachher je Feld).'
              : 'Every cancellation, deletion and field change is recorded in the audit log. Open it via the „Audit log / change log" action — it is then pre-filtered to this event and shows who, when and what per entry (for data changes: before → after per field).',
          },
        ],
      },
    ],
  };
}
