import * as React from 'react';
import { ManualSection } from '../types';
import { DemoAdminCenter, DemoParticipantTable, Callout } from '../ManualMockups';

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
            mockup: <DemoAdminCenter />,
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
            mockup: <DemoParticipantTable />,
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
        ],
      },
    ],
  };
}
