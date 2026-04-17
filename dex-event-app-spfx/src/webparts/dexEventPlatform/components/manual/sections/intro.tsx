import * as React from 'react';
import { ManualSection } from '../types';
import { Callout, DemoEventCard } from '../ManualMockups';

export function introSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'intro',
    title: isDe ? 'Willkommen in der Event Experience Platform' : 'Welcome to the Event Experience Platform',
    category: 'general',
    description: isDe
      ? 'Erste Schritte, Rollen-Übersicht und was die Plattform eigentlich macht.'
      : 'First steps, roles overview, and what the platform actually does.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        title: isDe ? 'Überblick für alle' : 'Overview for everyone',
        intro: (
          <>
            {isDe
              ? 'Die DEX Event Experience Platform ist die Event-Management-App von Deloitte. Sie vereint Anmeldung, Teilnehmerverwaltung, automatische E-Mails, Outlook-Termine und Check-In in einer App.'
              : 'The DEX Event Experience Platform is Deloitte\'s event management app. It combines registration, participant management, automated emails, Outlook invitations, and on-site check-in in one place.'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Drei Rollen — drei Sichten' : 'Three roles — three views',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Jede:r Mitarbeiter:in hat genau eine Rolle, die steuert, was in der App sichtbar ist:'
                    : 'Every employee has exactly one role that controls what they see in the app:'}
                </p>
                <ul style={{ paddingLeft: 18, margin: 0, lineHeight: 1.7 }}>
                  <li><strong>User</strong> — {isDe ? 'kann sich für Events anmelden, eigene Registrierungen verwalten und einchecken.' : 'registers for events, manages own sign-ups, checks in on-site.'}</li>
                  <li><strong>Organizer</strong> — {isDe ? 'kann zusätzlich eigene Events erstellen, Teilnehmer verwalten und kommunizieren.' : 'additionally creates events, manages participants, and communicates with them.'}</li>
                  <li><strong>Admin</strong> — {isDe ? 'verwaltet zusätzlich Rollen, globale E-Mail-Templates und technische Einstellungen.' : 'additionally manages roles, global email templates, and technical settings.'}</li>
                </ul>
              </>
            ),
            mockup: (
              <Callout variant="info" title={isDe ? 'Rollen werden zentral verwaltet' : 'Roles are managed centrally'}>
                {isDe
                  ? 'Ein Admin legt deine Rolle im Admin Center fest. Wenn du Events organisieren sollst, aber keine Möglichkeit dazu siehst, bitte deinen Admin um die Rolle "Organizer".'
                  : 'An admin assigns your role in the admin center. If you need to organize events but don\'t see the option, ask an admin to grant you the "Organizer" role.'}
              </Callout>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Sichtbarkeit von Events' : 'Event visibility',
            description: (
              <>
                {isDe
                  ? 'Events können standortbasiert gefiltert sein. Ein Event, das nur für Köln und Düsseldorf freigegeben ist, erscheint nicht in deiner Übersicht, wenn dein Office in Berlin ist. Events ohne Standort-Filter sind für alle sichtbar.'
                  : 'Events can be location-scoped. An event visible only for Cologne and Düsseldorf won\'t show up if your office is Berlin. Events without a location filter are visible to everyone.'}
              </>
            ),
            mockup: <DemoEventCard />,
          },
          {
            number: 3,
            title: isDe ? 'Sprache umstellen' : 'Change language',
            description: (
              <>
                {isDe
                  ? 'Auf der Startseite (Landing Page) kannst du zwischen Deutsch und Englisch wählen. Die Auswahl wird lokal gespeichert und bleibt über Sessions hinweg bestehen. Event-Mails gehen in der Sprache raus, die vom Organizer pro Event festgelegt wurde — das ist unabhängig von deiner UI-Sprache.'
                  : 'On the landing page you can switch between German and English. Your choice is stored locally and persists across sessions. Event emails go out in the language the organizer set per event — independent of your UI language.'}
              </>
            ),
            tip: isDe
              ? 'Die Sprach-Einstellung gilt nur für dich in der App. Andere Nutzer sehen ihre eigene Sprache.'
              : 'The language setting only applies to you. Other users see their own language preference.',
          },
        ],
      },
    ],
  };
}
