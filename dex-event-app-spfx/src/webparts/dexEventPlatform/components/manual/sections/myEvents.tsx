import * as React from 'react';
import { ManualSection } from '../types';
import { Callout, DemoEventCard } from '../ManualMockups';

export function myEventsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'my-events',
    title: isDe ? 'Meine Events verwalten' : 'Manage my events',
    category: 'general',
    description: isDe
      ? 'Eigene Anmeldungen einsehen, Angaben ändern und Abmeldung durchführen.'
      : 'Review your sign-ups, update answers and cancel.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        steps: [
          {
            number: 1,
            title: isDe ? 'Zur "Meine Events"-Seite navigieren' : 'Go to "My Events"',
            description: (
              <>
                {isDe
                  ? 'Vom Start-Screen oder über die Navigation gelangst du zu "Meine Events". Hier siehst du alle Events, für die du angemeldet bist — inklusive Warteliste-Einträge.'
                  : 'From the start screen or via the navigation you reach "My Events". Here you see every event you\'re registered for — including waitlist entries.'}
              </>
            ),
            mockup: <DemoEventCard status={isDe ? 'Du bist angemeldet' : 'You are registered'} />,
          },
          {
            number: 2,
            title: isDe ? 'Angaben ändern' : 'Update your answers',
            description: (
              <>
                {isDe
                  ? 'Solange das Event noch nicht begonnen hat, kannst du deine Event-spezifischen Antworten (T-Shirt-Größe, Zimmerpartner usw.) jederzeit anpassen. Deine Stamm-Daten (Name, E-Mail) sind nicht änderbar.'
                  : 'As long as the event hasn\'t started, you can update your event-specific answers (t-shirt size, roommate, etc.) at any time. Your master data (name, email) cannot be changed.'}
              </>
            ),
            mockup: (
              <Callout variant="tip" title={isDe ? 'Änderungen werden sofort gespeichert' : 'Changes are saved immediately'}>
                {isDe ? 'Du siehst eine Bestätigung sobald die Speicherung durch ist.' : 'You\'ll see a confirmation as soon as the save completes.'}
              </Callout>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Abmelden' : 'Cancel registration',
            description: (
              <>
                {isDe
                  ? 'Über "Abmelden" kannst du deine Teilnahme zurücknehmen. Die Abmeldung ist bis zur festgelegten Abmeldefrist möglich — danach blockiert die App die Abmeldung, weil Zimmer/Caterer/Busse bereits gebucht sind.'
                  : 'Via "Cancel" you can withdraw your participation. Cancellation is possible until the defined deadline — after that the app blocks self-cancellation because rooms/catering/buses are already booked.'}
              </>
            ),
            warning: isDe
              ? 'Nach Ablauf der Abmeldefrist kannst du dich nicht mehr selbst abmelden. Wende dich direkt an den Event-Organizer.'
              : 'After the cancellation deadline passes you can no longer cancel yourself. Please contact the event organizer directly.',
          },
          {
            number: 4,
            title: isDe ? 'Outlook-Termin nachvollziehen' : 'Outlook calendar sync',
            description: (
              <>
                {isDe
                  ? 'Dein Outlook-Termin wird automatisch mit der App synchronisiert. Meldest du dich ab, verschwindet der Termin — meldest du dich wieder an, kommt er zurück. Wenn du im Outlook selbst ablehnst, bekommst du eine Reminder-E-Mail der App, ob die Absage final sein soll.'
                  : 'Your Outlook appointment is automatically synced with the app. Cancel → the appointment disappears. Re-register → it comes back. If you decline in Outlook directly, you\'ll receive a reminder email asking whether the decline should be final.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
