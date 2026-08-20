import * as React from 'react';
import { ManualSection } from '../types';
import { Callout, DemoRoommatePicker, ClickPath } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import EventListPage from '../../EventListPage';
import RegistrationPage from '../../RegistrationPage';

export function findEventSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'find-event',
    title: isDe ? 'Event finden & anmelden' : 'Find & register for an event',
    category: 'general',
    description: isDe
      ? 'So meldest du dich für ein Event an — mit und ohne zusätzliche Fragen.'
      : 'How to register for an event — with and without extra questions.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    keywords: isDe
      ? 'Event finden anmelden registrieren Registrierung Anmeldung Event-Liste Übersicht Kachel Karten-Ansicht Listen-Ansicht Warteliste Zusatzfragen Zimmerpartner T-Shirt-Größe Anrede Bestätigungsmail Outlook-Termin ausgebucht Kapazität'
      : 'find event register registration sign up event list overview card list view waitlist wait list extra questions roommate t-shirt size salutation confirmation email outlook invite full capacity',
    perspectives: [
      {
        perspective: 'user',
        steps: [
          {
            number: 1,
            title: isDe ? 'Event-Kachel auf der Übersicht wählen' : 'Pick an event card from the overview',
            description: (
              <>
                {isDe
                  ? 'Über die Kachel "Aktuelle Events" (bzw. den Navigationspunkt) landest du auf der Event-Übersicht "Deine Events". Dort siehst du alle Events, die für dich sichtbar sind. Oben kannst du zwischen Karten- und Listen-Ansicht umschalten. Klick auf die Event-Karte, die dich interessiert — oder auf "Registrierung starten". Angemeldet bist du damit noch nicht, das passiert erst im nächsten Schritt.'
                  : 'Via the "Current events" tile (or navigation item) you reach the "Your events" overview. There you see every event visible to you. At the top you can switch between card and list view. Click the card you\'re interested in — or "Start registration". That does not register you yet; that happens in the next step.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Event-Übersicht (echte Ansicht)' : 'Event overview (real view)'}
                role="User"
                page="register"
                width={430}
                device="phone"
              >
                <Header />
                <EventListPage />
              </AppPreview>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Persönliche Daten bestätigen' : 'Confirm personal details',
            description: (
              <>
                {isDe
                  ? 'Vorname, Nachname und E-Mail sind aus deinem Deloitte-Profil vorbelegt und können nicht geändert werden. Die Anrede musst du einmalig auswählen.'
                  : 'First name, last name and email are pre-filled from your Deloitte profile and cannot be changed. You only need to pick a salutation.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Registrierungsformular (echte Ansicht)' : 'Registration form (real view)'}
                role="User"
                page="registration"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
              >
                <Header />
                <RegistrationPage />
              </AppPreview>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Event-spezifische Fragen ausfüllen' : 'Fill in event-specific questions',
            description: (
              <>
                {isDe
                  ? 'Pro Event können unterschiedliche Zusatzfragen vorhanden sein: T-Shirt-Größe, Ernährungsweise, Zimmerpartner, Einverständniserklärungen usw. Felder mit * sind Pflichtfelder.'
                  : 'Each event can have different custom questions: t-shirt size, diet, roommate, consents, etc. Fields marked with * are required.'}
              </>
            ),
            mockup: <DemoRoommatePicker />,
            tip: isDe
              ? 'Wenn du eine:n Zimmerpartner:in angibst, wird diese Person automatisch per E-Mail über deine Auswahl informiert.'
              : 'If you pick a preferred roommate, that person will automatically be notified by email.',
          },
          {
            number: 4,
            title: isDe ? 'Anmeldung absenden' : 'Submit your registration',
            description: (
              <>
                {isDe
                  ? 'Mit Klick auf "Anmelden" wird deine Registrierung gespeichert. Du bekommst eine Bestätigungs-E-Mail und — falls für das Event aktiviert — einen Outlook-Kalendereintrag.'
                  : 'Clicking "Register" saves your registration. You\'ll receive a confirmation email and — if enabled for the event — an Outlook calendar invite.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Anmelden' : 'Register'} hint={isDe ? 'E-Mail + Kalendereintrag kommen meist innerhalb weniger Minuten an.' : 'Email + calendar invite usually arrive within minutes.'} />,
          },
          {
            number: 5,
            title: isDe ? 'Wenn das Event ausgebucht ist' : 'When the event is full',
            description: (
              <>
                {isDe
                  ? 'Ist die maximale Teilnehmerzahl erreicht, landest du automatisch auf der Warteliste (sofern diese aktiviert ist). Sobald sich jemand abmeldet, rutschst du nach und bekommst automatisch eine Bestätigungsmail.'
                  : 'If the max capacity is reached you automatically go to the waitlist (if enabled). When someone cancels, you move up and receive a confirmation email automatically.'}
              </>
            ),
            mockup: (
              <Callout variant="info" title={isDe ? 'Warteliste' : 'Waitlist'}>
                {isDe
                  ? 'Du musst nichts weiter tun. Das Nachrücken läuft vollautomatisch über einen Power-Automate-Flow.'
                  : 'You don\'t have to do anything. Promotion from waitlist to attendee runs automatically via a Power Automate flow.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
