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
            title: isDe ? 'Teilnehmer in der App bearbeiten (v8.0)' : 'Edit attendee in the app (v8.0)',
            description: (
              <>
                {isDe
                  ? 'Neben jedem Teilnehmer findet sich seit v8.0 ein "Bearbeiten"-Button. Klick darauf öffnet ein Modal, in dem du Anrede, Vor-/Nachname, E-Mail, Telefon, Department, Standort, JobTitle, Status sowie alle event-spezifischen Felder direkt anpassen kannst — ohne Umweg über die SharePoint-Liste. Praktisch z.B. wenn ein Teilnehmer einen Tippfehler in der Anmeldung hatte oder seinen Standort geändert hat. Jede Änderung wird automatisch im "ChangeLog" der Teilnehmer-Zeile mit Datum + deinem Namen + Vorher-/Nachher-Wert protokolliert. Außerdem wird "LastModifiedDate" auf jetzt gesetzt.'
                  : 'Since v8.0 every attendee row has an "Edit" button. Clicking it opens a modal where you can change salutation, first/last name, email, phone, department, location, job title, status and all event-specific fields directly — no detour via the SharePoint list. Useful e.g. when an attendee had a typo during registration or has changed location. Every change is automatically logged in the row\'s "ChangeLog" with date + your name + before/after value, and "LastModifiedDate" is set to now.'}
              </>
            ),
            tip: isDe
              ? 'Vorteil gegenüber direktem Editieren in SharePoint: keine Datums-Format-Probleme, automatische Audit-Spur, keine Schulung der Organizer auf SP-UI nötig.'
              : 'Advantage over direct editing in SharePoint: no date-format issues, automatic audit trail, no need to train organizers on the SP UI.',
          },
          {
            number: 6,
            title: isDe ? 'Aktualisieren-Button (v7.37+)' : 'Refresh button (v7.37+)',
            description: (
              <>
                {isDe
                  ? 'Oben rechts in jeder Übersicht (Eventliste, MyEvents, Admin Center) gibt es einen "Aktualisieren"-Button mit Refresh-Icon. Damit kannst du die Daten frisch aus SharePoint laden, ohne die App neu starten zu müssen — sinnvoll z.B. wenn ein anderer Organizer parallel etwas geändert hat oder wenn der Power-Automate-Flow gerade einen Nachrücker promotet hat und du das Ergebnis sehen willst. Im Admin Center werden bei selektiertem Event auch dessen Teilnehmer neu geladen.'
                  : 'In every overview (event list, MyEvents, admin center) there\'s a "Refresh" button with a reload icon at the top right. It pulls fresh data from SharePoint without restarting the app — useful e.g. when another organizer has changed something in parallel or when the Power Automate flow has just promoted a waitlister and you want to see the result. In the admin center with a selected event, that event\'s attendees are also reloaded.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
