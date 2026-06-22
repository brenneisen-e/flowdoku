import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import RegistrationPage from '../../RegistrationPage';

export function registerForOtherSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'register-for-other',
    title: isDe ? 'Für andere Person registrieren' : 'Register for someone else',
    category: 'organizer',
    description: isDe
      ? 'Wie du als Organizer oder Assistant andere Personen anmeldest.'
      : 'How organizers and assistants register other people.',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Modus wechseln' : 'Switch mode',
            description: (
              <>
                {isDe
                  ? 'Auf der Registrierungsseite eines Events klickst du auf "Für jemand anderen registrieren". Deine eigenen Daten verschwinden, und ein Suchfeld erscheint.'
                  : 'On an event\'s registration page, click "Register for someone else". Your own data clears and a search field appears.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Registrierung mit "Für andere Person anmelden" (echte Ansicht)' : 'Registration with "Register for another person" (real view)'}
                role="Organizer"
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
            number: 2,
            title: isDe ? 'Person wählen — frühe Checks' : 'Pick a person — early checks',
            description: (
              <>
                {isDe
                  ? 'Ab zwei Zeichen zeigt die App passende Kolleg:innen an. Sobald du jemand auswählst, prüft die App SOFORT zwei Dinge und zeigt Hinweise, falls: a) die Person bereits angemeldet ist, b) die Person nicht zum Verteilerkreis des Events gehört.'
                  : 'From two characters the app shows matching colleagues. As soon as you pick someone, the app immediately checks and warns if: a) the person is already registered, or b) the person is not in the event audience.'}
              </>
            ),
            mockup: (
              <Callout variant="warning" title={isDe ? 'Beispiel-Warnung' : 'Example warning'}>
                {isDe
                  ? '"Diese Person ist bereits für das Event angemeldet." oder "Diese Person gehört nicht zum festgelegten Verteilerkreis dieses Events."'
                  : '"This person is already registered for this event." or "This person is not part of the defined audience for this event."'}
              </Callout>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Sonderregel Assistants' : 'Assistants-only rule',
            description: (
              <>
                {isDe
                  ? 'Nutzer mit JobTitle "Assistant" oder "Senior Assistant" können NUR Partner oder Directors anmelden. Andere Treffer in der Suche werden ausgegraut. So ist gewährleistet, dass z.B. Business Support nur Führungskräfte anmeldet, wofür sie zuständig sind.'
                  : 'Users with job title "Assistant" or "Senior Assistant" can register ONLY Partners or Directors. Other matches in the search are greyed out. This ensures that e.g. business support only signs up the leaders they\'re responsible for.'}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Rest wie bei Selbst-Anmeldung' : 'Rest works like self-registration',
            description: (
              <>
                {isDe
                  ? 'Sobald die Person gewählt ist, verhält sich die Anmeldung genau wie die eigene: Anrede, Custom Fields, Absenden. Die E-Mail-Bestätigung geht an die angemeldete Person — nicht an dich.'
                  : 'Once the person is picked, the registration behaves exactly like self-registration: salutation, custom fields, submit. The confirmation email goes to the registered person — not to you.'}
              </>
            ),
            tip: isDe
              ? 'Im Audit-Trail des Teilnehmer-Eintrags steht, wer die Anmeldung ausgelöst hat — für Rückfragen.'
              : 'The participant record\'s audit trail logs who triggered the registration — useful for follow-up questions.',
          },
          {
            number: 5,
            title: isDe ? 'Zustimmungs-Nachweis (v18.74)' : 'Consent record (v18.74)',
            description: (
              <>
                {isDe
                  ? 'Bei jeder stellvertretenden Anmeldung bestätigst du per Pflicht-Checkbox, dass die Person zugestimmt hat. Diese Bestätigung wird zusätzlich als Nachweis in der Teilnehmerliste gespeichert (Spalte „ProxyConsent") — mit deinem Namen und dem Datum. So lässt sich im Nachhinein belegen, dass die Zustimmung vorlag.'
                  : 'For every registration on someone\'s behalf you confirm via a mandatory checkbox that the person consented. This confirmation is additionally stored as a record in the participant list (column „ProxyConsent") — with your name and the date. This proves later that consent was in place.'}
              </>
            ),
          },
          {
            number: 6,
            title: isDe ? 'Person außerhalb Deloitte (v18.74)' : 'Person outside Deloitte (v18.74)',
            description: (
              <>
                {isDe
                  ? 'Mit der Checkbox „Person außerhalb Deloitte" meldest du eine externe Person an. Der People-Picker verschwindet — du trägst Vorname, Nachname und E-Mail selbst ein. Die Zustimmung musst du in diesem Fall SCHRIFTLICH einholen. Die App prüft die E-Mail auf Tippfehler und zeigt sie vor dem Absenden noch einmal groß zur Kontrolle an.'
                  : 'With the „Person outside Deloitte" checkbox you register an external person. The people picker disappears — you enter first name, last name and email yourself. In this case consent must be obtained IN WRITING. The app checks the email for typos and shows it once more, large, for review before submitting.'}
              </>
            ),
            tip: isDe
              ? 'Die Bestätigungs-Mail geht direkt an die externe Person, mit den Organizern auf CC. Ein Outlook-Termin wird an externe Adressen NICHT versendet — der Termin muss manuell in den Kalender eingetragen werden.'
              : 'The confirmation email is sent directly to the external person, with the organizers on CC. An Outlook invite is NOT sent to external addresses — the appointment must be added to the calendar manually.',
          },
          {
            number: 7,
            title: isDe ? 'Später verwalten: Kachel „Assistenz" (v24.36)' : 'Manage later: „Assistant" tile (v24.36)',
            description: (
              <>
                {isDe
                  ? 'Anmeldungen, die du stellvertretend für andere durchgeführt hast, verwaltest du danach gebündelt über die Startseiten-Kachel „Assistenz": Angaben einsehen und anpassen, die Person bei Sub-Events nachträglich an- oder abmelden und ganz abmelden — ähnlich wie „Meine Events", nur eben für die Personen, die du angemeldet hast. Die Kachel erscheint nur, wenn du wirklich solche Anmeldungen hast (Admins sehen sie immer).'
                  : 'Registrations you made on behalf of others can afterwards be managed in one place via the start-page tile „Assistant": view and edit details, register or cancel the person for sub-events, or cancel entirely — similar to „My events", but for the people you registered. The tile only appears if you actually have such registrations (admins always see it).'}
              </>
            ),
            tip: isDe
              ? 'Wichtig: Hier siehst du NUR Anmeldungen, die du selbst stellvertretend gemacht hast. Hat sich die Person (z.B. dein Partner oder Director) selbst angemeldet, ist diese Anmeldung hier nicht sichtbar — dann verwaltet die Person sie selbst über „Meine Events".'
              : 'Important: you only see registrations you made on someone\'s behalf yourself. If the person (e.g. your partner or director) registered themselves, that registration is not visible here — they manage it themselves via „My events".',
          },
        ],
      },
    ],
  };
}
