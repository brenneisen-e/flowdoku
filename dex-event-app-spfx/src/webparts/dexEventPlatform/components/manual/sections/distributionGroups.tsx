import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v26.33: Verteilergruppen / Mail-Verteiler als Zielgruppe (Sichtbarkeit).
 * Beantwortet die häufige Frage „Kann man einen Outlook-Verteiler einsetzen,
 * statt alle Personen einzeln einzutragen?". Reich mit Such-Stichwörtern
 * (keywords) versehen, damit Suche UND Frage-Vorschläge greifen.
 */
export function distributionGroupsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'distribution-groups',
    title: isDe ? 'Verteilergruppen / Mail-Verteiler als Zielgruppe' : 'Distribution groups as audience',
    category: 'organizer',
    description: isDe
      ? 'Statt alle Personen einzeln einzutragen kannst du einen Outlook-/Mail-Verteiler als Zielgruppe (Sichtbarkeit) hinterlegen — das System löst die Mitglieder automatisch auf.'
      : 'Instead of adding every person individually you can set an Outlook / mail distribution group as the audience (visibility) — the system resolves the members automatically.',
    visibleFor: ['Admin', 'Organizer'],
    // Such-Stichwörter/Synonyme, damit Suche + Frage-Vorschläge den Artikel finden.
    keywords: 'Verteilergruppen Verteilergruppe Mailverteiler Outlookverteiler Outlook-Verteiler Verteiler Mailgruppen Mailgruppe Verteilerliste E-Mail-Liste Email-Liste Emailverteiler distribution list distribution group mailing list Sichtbarkeit Audience Zielgruppe Gruppe',
    perspectives: [
      {
        perspective: 'organizer',
        intro: isDe
          ? 'Ja — du musst nicht jede Person einzeln eintragen. Trage einfach den Namen eines Outlook-/Mail-Verteilers ein, dann greift das System auf genau die im Verteiler hinterlegten Mitglieder zu.'
          : 'Yes — you do not need to add every person individually. Just enter the name of an Outlook / mail distribution group and the system uses exactly the members stored in that group.',
        steps: [
          {
            number: 1,
            title: isDe ? 'Wo? Schritt „Kapazität & Sichtbarkeit"' : 'Where? Step "Capacity & visibility"',
            description: (
              <>
                {isDe
                  ? 'Im Event-Wizard auf dem Schritt „Kapazität & Sichtbarkeit" (Schritt 5) findest du unter dem Standort-Filter das Feld „Audience" (Zielgruppe/Sichtbarkeit). Damit schränkst du ein, wer das Event sehen und sich anmelden kann. Lässt du es leer, ist das Event für alle Mitarbeiter sichtbar.'
                  : 'In the event wizard on the "Capacity & visibility" step (step 5) you find the "Audience" field below the location filter. It restricts who can see and register for the event. Leave it empty and the event is visible to everyone.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Verteiler statt Einzelpersonen eintragen' : 'Enter a distribution group instead of individuals',
            description: (
              <>
                {isDe
                  ? 'Gib einfach den Namen des Outlook-/Mail-Verteilers ein — z.B. „DE TT DUESSELDORF". Das System löst den Verteiler auf und übernimmt automatisch alle darin hinterlegten Mitglieder. Du musst also nicht jede Person einzeln suchen und hinzufügen. Mehrere Verteiler kannst du kommasepariert angeben.'
                  : 'Just type the name of the Outlook / mail distribution group — e.g. "DE TT DUESSELDORF". The system resolves the group and automatically pulls in all of its members. So you do not have to search and add each person individually. You can enter several groups separated by commas.'}
              </>
            ),
            tip: isDe
              ? 'Standort-Filter (Pills) und Verteiler lassen sich kombinieren. Es zählt die Schnittmenge — z.B. „nur Düsseldorf" plus Verteiler „DE TT DUESSELDORF".'
              : 'Location pills and distribution groups can be combined — e.g. "Düsseldorf only" plus the "DE TT DUESSELDORF" group.',
          },
          {
            number: 3,
            title: isDe ? 'Wirkung: nur Verteiler-Mitglieder sehen das Event' : 'Effect: only group members see the event',
            description: (
              <>
                {isDe
                  ? 'Nach dem Speichern sehen nur noch die Mitglieder des hinterlegten Verteilers das Event in ihrer Liste — und nur sie können sich anmelden. Ändert sich der Verteiler in Outlook, greift die App beim nächsten Öffnen auf den aktuellen Stand zu.'
                  : 'After saving, only the members of the stored distribution group see the event in their list — and only they can register. If the group changes in Outlook, the app uses the current membership the next time it is opened.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
