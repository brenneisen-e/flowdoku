import * as React from 'react';
import { ManualSection } from '../types';

export function peoplePickerScopeSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'people-picker-scope',
    title: isDe ? 'Personensuche: DE vs. international' : 'People search: DE vs. international',
    category: 'general',
    description: isDe
      ? 'Wie die Personensuche in der App funktioniert — Standard nur Deloitte Deutschland, optional mit internationalen Member-Firms.'
      : 'How the people search in the app works — German member firm by default, international member firms on demand.',
    visibleFor: ['Admin', 'Organizer'],
    keywords: isDe
      ? 'Personensuche People-Picker Personenfinder Suche Member-Firm international @deloitte.de @deloitte.com DEALL Toggle Checkbox DECH DEUS DECEMEA Co-Organizer Speaker'
      : 'people search people picker person finder search member firm international @deloitte.de @deloitte.com DEALL toggle checkbox DECH DEUS DECEMEA co-organizer speaker',
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Standardverhalten: nur deutsche Member-Firm' : 'Default: German member firm only',
            description: (
              <>
                {isDe
                  ? 'Jeder Personenfinder in der App (z.B. Organizer-Auswahl im Wizard, „Für andere anmelden", Team-Mitglied hinzufügen, Demo-Modus) sucht ausschließlich in der deutschen Member-Firm (alle @deloitte.de-Adressen, entspricht dem DEALL-Verteiler).'
                  : 'Every people picker in the app (organizer selection in the wizard, „register for other", add team member, demo mode) searches the German member firm only by default — all @deloitte.de addresses, equivalent to the DEALL distribution list.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Auswirkung für dich' : 'What this means for you',
            description: (
              <>
                {isDe
                  ? 'Treffer aus anderen Member-Firms (USA, Schweiz, EMEA-Hubs etc., die alle @deloitte.com nutzen) tauchen nicht mehr in den Vorschlägen auf. Das löst das Problem, dass eine Suche nach „Nils" früher auch „Nilesh" oder „Niladri" aus anderen Ländern angezeigt hat.'
                  : 'Hits from other member firms (USA, Switzerland, EMEA hubs etc., all using @deloitte.com) no longer show up in the suggestions. This fixes the issue where typing „Nils" used to also surface „Nilesh" or „Niladri" from other countries.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'International suchen: Toggle umlegen' : 'Searching internationally: flip the toggle',
            description: (
              <>
                {isDe
                  ? 'Direkt unter jedem Suchfeld findest du eine kleine Checkbox „Auch international suchen (@deloitte.com)". Hak sie an, wenn du explizit jemanden aus einer anderen Member-Firm finden willst — z.B. einen Co-Organizer aus DECH oder einen Speaker aus DEUS. Die laufende Suche wird sofort wiederholt und liefert dann auch internationale Treffer.'
                  : 'Right below every search field there is a small checkbox „Search internationally too (@deloitte.com)". Tick it when you explicitly need to find someone from another member firm — e.g. a co-organizer from DECH or a speaker from DEUS. The current query is immediately re-run and now also returns international hits.'}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Toggle gilt nur lokal' : 'Toggle is local-only',
            description: (
              <>
                {isDe
                  ? 'Jeder Finder hält seinen eigenen Toggle-Zustand. Wenn du z.B. im Wizard internationale Organizer suchst, bleibt die Personensuche in „Für andere anmelden" trotzdem auf Deutschland. Das verhindert versehentliche Verbreiterung der Suche an Stellen, wo du eigentlich nur DE-Kollegen wolltest.'
                  : 'Every picker keeps its own toggle state. If you search for international organizers in the wizard, the people search in „register for other" still stays on Germany. This prevents accidental widening of the search where you actually only wanted DE colleagues.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
