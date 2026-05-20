import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';

export function inviteMailSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'invite-mail',
    title: isDe ? 'Einladungsmail mit Anmelde-Link' : 'Invitation email with registration link',
    category: 'organizer',
    description: isDe
      ? 'Eine Einladungs-Mail an dich selbst oder direkt an den Mailverteiler des Events verschicken — mit Anmelde-Link, im Deloitte-Template, frei editierbar.'
      : 'Send an invitation email to yourself or directly to the event mail distribution list — with registration link, in the Deloitte template, freely editable.',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Einladungsmail-Dialog öffnen' : 'Open the invitation dialog',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center eines Events findest du neben der Aktion „Massenmail" eine zweite Kachel „Einladungsmail". Ein Klick öffnet ein modales Fenster mit Betreff, Überschrift und Inhalt — alles mit einem sinnvollen Standard-Text vorbefüllt, der schon den Anmelde-Link enthält und einen Hinweis, dass eine Abmeldung jederzeit über die App möglich ist.'
                  : 'In the Admin Center of an event you will find a second tile „Invitation email" next to „Mass email". Clicking opens a modal with subject, heading and content — pre-populated with a sensible default text that already contains the registration link and a hint that cancellation is possible anytime via the app.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Empfänger wählen' : 'Choose recipient',
            description: (
              <>
                {isDe ? (
                  <>
                    Oben im Dialog stehen zwei Optionen zur Auswahl:
                    <ul>
                      <li><strong>An mich — zum Weiterleiten:</strong> Die Mail geht nur an deine eigene E-Mail-Adresse. Du bekommst sie sauber im Deloitte-Template gerendert und kannst sie aus Outlook per Forward an beliebige Verteiler / Kollegen / externe Adressen schicken.</li>
                      <li><strong>An Mailverteiler des Events:</strong> Geht direkt an alle E-Mail-Adressen, die du im Event-Edit unter <em>Schritt 3 (Kapazität &amp; Sichtbarkeit)</em> als Mailverteiler hinterlegt hast. Die App zeigt dir die Liste der Empfänger direkt im Dialog, damit du siehst wer es bekommt.</li>
                    </ul>
                    Ist auf dem Event kein Mailverteiler hinterlegt, ist die zweite Option ausgegraut.
                  </>
                ) : (
                  <>
                    The top of the dialog offers two options:
                    <ul>
                      <li><strong>To me — for forwarding:</strong> The mail goes only to your own email address. You receive it rendered in the Deloitte template and can forward it from Outlook to any distribution list / colleagues / external addresses.</li>
                      <li><strong>To event mail distribution:</strong> Goes directly to all email addresses configured as mail distribution in <em>step 3 (Capacity &amp; Visibility)</em> of the event edit. The app shows the list of recipients in the dialog so you see who receives it.</li>
                    </ul>
                    If no mail distribution is configured on the event, the second option is greyed out.
                  </>
                )}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Text anpassen' : 'Adjust the text',
            description: (
              <>
                {isDe
                  ? 'Der Standardtext ist editierbar — du kannst Begrüßung, Beschreibung, Hinweise zur Kleiderordnung / Anfahrt / Agenda etc. ergänzen oder den Anmelde-Link mit der Platzhalter-Variable {{Link}} an anderer Stelle einfügen. Eine Live-Vorschau rechts zeigt dir, wie die Mail im Outlook-Postfach des Empfängers aussieht.'
                  : 'The default text is editable — you can add greeting, description, notes about dress code / directions / agenda etc., or place the registration link via the placeholder variable {{Link}} at a different position. A live preview on the right shows how the mail looks in the recipient\'s Outlook inbox.'}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Versenden' : 'Send',
            description: (
              <>
                {isDe
                  ? 'Der Senden-Button unten zeigt dir zusätzlich nochmal das gewählte Ziel an („Senden — An mich" oder „Senden — An Mailverteiler"). Die Mail landet in der DEX_Emails-Queue und wird vom Power-Automate-Flow „DEX_SEND_MAIL" innerhalb weniger Minuten versendet.'
                  : 'The send button at the bottom also shows the selected target („Send — To me" or „Send — To mail distribution"). The mail goes into the DEX_Emails queue and is sent by the Power Automate flow „DEX_SEND_MAIL" within minutes.'}
              </>
            ),
            mockup: (
              <Callout variant="info">
                {isDe
                  ? 'Die Mail wird wie alle App-Mails vom Shared-Mailbox-Absender no_reply.events@deloitte.de verschickt — nicht von deiner persönlichen Adresse. Wenn du den Verteiler-Modus nutzt und eine Antwort möchtest, ergänze im Mail-Text einen Hinweis auf deine eigene Adresse.'
                  : 'Like all app mails, this is sent from the shared mailbox no_reply.events@deloitte.de — not from your personal address. If you use distribution mode and want a reply, add a note with your own address inside the mail body.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
