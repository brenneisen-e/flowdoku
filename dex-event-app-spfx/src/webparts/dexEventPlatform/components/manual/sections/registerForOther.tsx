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
        ],
      },
    ],
  };
}
