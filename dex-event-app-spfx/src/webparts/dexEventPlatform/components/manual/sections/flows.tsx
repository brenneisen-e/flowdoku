import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';

export function flowsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'flows',
    title: isDe ? 'Power Automate Flows' : 'Power Automate Flows',
    category: 'architecture',
    description: isDe
      ? 'Die sechs Hintergrund-Flows, die E-Mails und Outlook-Termine steuern.'
      : 'The six background flows that drive emails and Outlook appointments.',
    visibleFor: ['Admin'],
    perspectives: [
      {
        perspective: 'admin',
        intro: (
          <>
            {isDe
              ? 'Die App speichert E-Mail- und Outlook-Intents in SharePoint-Queues (DEX_Emails, DEX_Outlook). Sechs Power-Automate-Flows arbeiten diese ab. Details findest du in docs/flow-jsons.md.'
              : 'The app writes email and Outlook intents into SharePoint queues (DEX_Emails, DEX_Outlook). Six Power Automate flows drain them. Details are in docs/flow-jsons.md.'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: 'DEX_SEND_MAIL',
            description: (
              <>
                {isDe
                  ? 'Holt neue Einträge aus DEX_Emails (Status: Pending) und verschickt die HTML-Mail über die Shared Mailbox no_reply.events@deloitte.de. Setzt danach Status auf "Sent" oder "Failed".'
                  : 'Picks pending entries from DEX_Emails and sends the HTML mail via the shared mailbox no_reply.events@deloitte.de. Sets status to "Sent" or "Failed".'}
              </>
            ),
          },
          {
            number: 2,
            title: 'DEX_CreateOutlookEvent',
            description: (
              <>
                {isDe
                  ? 'Legt pro Event initial einen Outlook-Kalendereintrag im Postfach des Hauptorganisators an, sobald das Event erstellt wird.'
                  : 'On event creation, creates an initial Outlook calendar entry in the main organizer\'s mailbox.'}
              </>
            ),
          },
          {
            number: 3,
            title: 'DEX_Outlook_Einladungen',
            description: (
              <>
                {isDe
                  ? 'Reagiert auf Einträge in DEX_Outlook (Einladen / Entfernen / Update / Delete) und fügt Teilnehmer zum Outlook-Termin hinzu oder entfernt sie. Aktualisiert auch Zeit, Ort, Beschreibung.'
                  : 'Reacts to DEX_Outlook entries (invite / remove / update / delete) and adds or removes attendees from the Outlook appointment. Also updates time, location, description.'}
              </>
            ),
          },
          {
            number: 4,
            title: 'DEX_OutlookDeclineHandler',
            description: (
              <>
                {isDe
                  ? 'Erkennt, wenn ein Teilnehmer die Einladung in Outlook ablehnt. Statt automatisch abzumelden, wird eine Reminder-Mail an den Nutzer gesendet ("war das Absicht?") — Nutzer bleibt in der App angemeldet.'
                  : 'Detects when an attendee declines the invitation in Outlook. Instead of auto-cancelling, sends a reminder email ("did you mean it?") — user stays registered in the app.'}
              </>
            ),
          },
          {
            number: 5,
            title: 'DEX_OutlookForwardHandler',
            description: (
              <>
                {isDe
                  ? 'Erkennt Meeting-Forwards. Wenn der weitergeleitete Empfänger NICHT im Event angemeldet ist, schickt eine FYI-Mail an den Organizer, damit der reagieren kann.'
                  : 'Detects meeting forwards. If the forwarded-to person is NOT registered in the event, sends an FYI to the organizer so they can react.'}
              </>
            ),
          },
          {
            number: 6,
            title: 'DEX_IDReorder_TeilnehmerIDs',
            description: (
              <>
                {isDe
                  ? 'Wird durch den Admin-Button "IDs neu vergeben" getriggert. Nummeriert alle aktiven Teilnehmer eines Events sequentiell und rückt Wartelistler entsprechend nach.'
                  : 'Triggered by the admin "Renumber IDs" button. Renumbers active attendees of an event sequentially and promotes waitlist members accordingly.'}
              </>
            ),
          },
          {
            number: 7,
            title: isDe ? 'Monitoring' : 'Monitoring',
            description: (
              <>
                {isDe
                  ? 'Alle Flows laufen in der Shared Mailbox von no_reply.events@deloitte.de. Fehler siehst du im Power-Automate-Admin-Center. Die Queue-Listen (DEX_Emails, DEX_Outlook) zeigen in Spalte Status auch "Failed"-Einträge — regelmäßig sichten.'
                  : 'All flows run in the shared mailbox of no_reply.events@deloitte.de. Errors are visible in the Power Automate admin center. The queue lists (DEX_Emails, DEX_Outlook) also show "Failed" in the Status column — review regularly.'}
              </>
            ),
            mockup: (
              <Callout variant="warning">
                {isDe
                  ? 'Bei Änderungen an einem Flow: IMMER das aktuelle JSON in docs/flow-jsons.md pflegen. Das ist die einzige Quelle der Wahrheit für den aktuellen Stand.'
                  : 'When modifying a flow: ALWAYS update the current JSON in docs/flow-jsons.md. That file is the single source of truth for the current state.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
