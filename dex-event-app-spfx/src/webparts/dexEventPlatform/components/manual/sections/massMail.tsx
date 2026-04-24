import * as React from 'react';
import { ManualSection } from '../types';
import { DemoEmailPreview, Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import AdminPage from '../../AdminPage';

export function massMailSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'mass-mail',
    title: isDe ? 'Massenmail an Teilnehmer' : 'Mass email to attendees',
    category: 'organizer',
    description: isDe
      ? 'Alle Angemeldeten auf einmal anschreiben — im Deloitte-Template, mit Rich-Text.'
      : 'Reach all registered attendees at once — in the Deloitte template, with rich text.',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Massenmail-Dialog öffnen' : 'Open the mass email dialog',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center eines Events findest du die Aktion "Massenmail". Ein modaler Dialog öffnet sich mit Betreff-, Überschrift- und Inhalt-Feld. Für den Inhalt steht dir ein Rich-Text-Editor zur Verfügung (fett, kursiv, Links, Listen).'
                  : 'The event admin center has an action "Mass email". A modal opens with subject, heading and content fields. The content uses a rich-text editor (bold, italic, links, lists).'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Admin Center → Massenmail-Button (echte Ansicht)' : 'Admin center → mass email button (real view)'}
                role="Organizer"
                page="admin"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
              >
                <Header />
                <AdminPage />
              </AppPreview>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Empfängerkreis wählen' : 'Choose recipients',
            description: (
              <>
                {isDe
                  ? 'Standardmäßig gehen Mails an alle aktiven Teilnehmer (Angemeldet + Eingecheckt). Mit Checkboxen kannst du zusätzlich Warteliste oder nur eine Untergruppe auswählen.'
                  : 'By default mails go to all active attendees (Registered + Checked-in). Checkboxes let you additionally include the waitlist or only a subset.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Vorschau' : 'Preview',
            description: (
              <>
                {isDe
                  ? 'Vor dem Versand siehst du eine Vorschau im Deloitte-Mail-Template mit Event-Logo (falls hinterlegt) und Signatur-Block. So vermeidest du Überraschungen.'
                  : 'Before sending you see a preview in the Deloitte mail template with event logo (if set) and signature block. Avoids surprises.'}
              </>
            ),
            mockup: <DemoEmailPreview heading={isDe ? 'Info zum Event' : 'Info about the event'} />,
          },
          {
            number: 4,
            title: isDe ? 'Versenden' : 'Send',
            description: (
              <>
                {isDe
                  ? 'Die Mails landen in der DEX_Emails-Queue und werden vom Power-Automate-Flow "DEX_SEND_MAIL" abgearbeitet (meist innerhalb weniger Minuten).'
                  : 'Emails go into the DEX_Emails queue and are processed by the Power Automate flow "DEX_SEND_MAIL" (usually within minutes).'}
              </>
            ),
            mockup: (
              <Callout variant="info">
                {isDe
                  ? 'Die Mail geht vom Shared-Mailbox-Absender "no_reply.events@deloitte.de" raus — nicht von dir persönlich.'
                  : 'Emails are sent from the shared mailbox "no_reply.events@deloitte.de" — not from your personal address.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
