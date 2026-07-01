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
    title: isDe ? 'E-Mail an alle Teilnehmer senden' : 'Send an email to all attendees',
    category: 'organizer',
    description: isDe
      ? 'Alle Angemeldeten eines Events auf einmal anschreiben — im Deloitte-Layout, mit Rich-Text-Editor und Test-Versand an dich selbst.'
      : 'Reach all registered attendees of an event at once — in the Deloitte layout, with a rich-text editor and a test send to yourself.',
    visibleFor: ['Organizer', 'Admin'],
    keywords: 'email senden mail senden e-mail versenden massenmail rundmail rund-mail email an alle mail an alle teilnehmer anschreiben nachricht info-mail infomail serienmail newsletter benachrichtigen',
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Aktion „E-Mail versenden" öffnen' : 'Open the „Send email" action',
            description: (
              <>
                {isDe
                  ? 'Öffne dein Event im Organizer Center (Kachel „Organizer" auf der Startseite → Event anklicken). Klicke dort auf das Dropdown „Aktionen" und wähle in der Kategorie „E-Mails" den Eintrag „E-Mail versenden". Hinweis: Du musst nichts in der Teilnehmerliste markieren — wen die Mail erreicht, legst du im nächsten Schritt fest.'
                  : 'Open your event in the Organizer Center (the „Organizer" tile on the start page → click the event). Open the „Actions" dropdown and pick „Send email" under the „Emails" category. Note: you do not need to select anyone in the attendee list — you choose the recipients in the next step.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Organizer Center → Aktionen → E-Mails (echte Ansicht)' : 'Organizer Center → Actions → Emails (real view)'}
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
            title: isDe ? 'Empfängerkreis wählen' : 'Choose the recipients',
            description: (
              <>
                {isDe
                  ? 'Zuerst fragt dich ein kleiner Dialog „An wen soll die Mail gehen?". Du hast fünf Möglichkeiten:'
                  : 'First a small dialog asks „Who should receive the email?". You have five options:'}
                <ul style={{ margin: '10px 0 0', paddingLeft: 20, lineHeight: 1.6 }}>
                  <li>{isDe
                    ? <><strong>Teilnehmer (alle aktiven)</strong> — der Normalfall: alle, die angemeldet, mit QR-Code versorgt oder schon eingecheckt sind (also NICHT Warteliste und NICHT Abgemeldete).</>
                    : <><strong>Attendees (all active)</strong> — the usual case: everyone registered, QR-sent or already checked in (NOT the waitlist and NOT cancellations).</>}</li>
                  <li>{isDe
                    ? <><strong>Teilnehmer + Warteliste</strong> — zusätzlich die Wartelistler, z.B. wenn du auch sie vorwarnen willst.</>
                    : <><strong>Attendees + waitlist</strong> — additionally the waitlisted people, e.g. to give them a heads-up too.</>}</li>
                  <li>{isDe
                    ? <><strong>Nur Warteliste</strong> — ausschließlich die Wartelistler.</>
                    : <><strong>Waitlist only</strong> — only the waitlisted people.</>}</li>
                  <li>{isDe
                    ? <><strong>Nachrücker (manueller Abgleich)</strong> — du fügst im nächsten Schritt eine Liste von E-Mail-Adressen ein; die Mail geht dann an alle aktiven Teilnehmer, die NICHT in deiner Liste stehen.</>
                    : <><strong>No-shows / top-ups (manual match)</strong> — you paste a list of email addresses in the next step; the mail then goes to all active attendees who are NOT on your list.</>}</li>
                  <li>{isDe
                    ? <><strong>Eigene Auswahl (nach Status)</strong> — du hakst genau die Status an, die die Mail bekommen sollen (z.B. nur „QR versendet“).</>
                    : <><strong>Custom selection (by status)</strong> — tick exactly the statuses that should receive the mail (e.g. only „QR sent“).</>}</li>
                </ul>
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Betreff und Text schreiben' : 'Write subject and text',
            description: (
              <>
                {isDe
                  ? 'Danach öffnet sich der Editor mit Betreff, Überschrift, Unter-Überschrift und dem eigentlichen Inhalt. Der Inhalt läuft über einen Rich-Text-Editor (fett, kursiv, Links, Aufzählungen). Weil die Massenmail als EIN Anschreiben an die ganze Empfängergruppe geht (nicht personalisiert pro Person), werden nur event-weite Platzhalter ersetzt: '
                  : 'Then the editor opens with subject, heading, sub-heading and the actual content. The content uses a rich-text editor (bold, italic, links, bullet lists). Because the mass mail goes out as ONE message to the whole recipient group (not personalised per person), only event-wide placeholders are replaced: '}
                <code>{'{{EventTitle}}'}</code> {isDe ? 'und' : 'and'} <code>{'{{Organizer}}'}</code>.{' '}
                {isDe
                  ? 'Persönliche Anreden wie {{Vorname}} funktionieren hier NICHT (sie blieben als Text stehen) — dafür sind die System-Mails (Anmeldung, Warteliste, …) gedacht.'
                  : 'Personal salutations like {{Vorname}} do NOT work here (they would remain as literal text) — the system mails (registration, waitlist, …) are meant for that.'}{' '}
                {isDe
                  ? 'Dein Entwurf wird pro Event automatisch gespeichert — du kannst den Dialog also schließen und später weiterschreiben, ohne den Text zu verlieren.'
                  : 'Your draft is saved automatically per event — so you can close the dialog and continue later without losing the text.'}
              </>
            ),
            mockup: <DemoEmailPreview heading={isDe ? 'Info zum Event' : 'Info about the event'} />,
          },
          {
            number: 4,
            title: isDe ? 'Test an dich selbst' : 'Test to yourself',
            description: (
              <>
                {isDe
                  ? 'Bevor du an alle sendest, klicke auf „Test an mich". Die App schickt die Mail genau so, wie sie rausgehen würde, an die hinterlegten Organizer — also an dich. So siehst du in deinem Postfach das fertige Deloitte-Layout (Logo, grüner Balken, Signatur) und kannst Tippfehler oder falsche Links vor dem echten Versand korrigieren.'
                  : 'Before sending to everyone, click „Test to me". The app sends the email exactly as it would go out to the configured organizers — i.e. to you. That way you see the finished Deloitte layout (logo, green bar, signature) in your mailbox and can fix typos or wrong links before the real send.'}
              </>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Absenden' : 'Send',
            description: (
              <>
                {isDe
                  ? 'Bist du zufrieden, klickst du auf Senden. Die App fragt kurz nach („E-Mail an N Teilnehmer senden?") und legt danach EINE Mail in die Versand-Warteschlange (DEX_Emails), die an alle gewählten Empfänger geht; die Organizer werden automatisch auf CC gesetzt (sofern sie nicht ohnehin schon Empfänger sind). Zustellung meist innerhalb weniger Minuten. Die Mail kommt vom gemeinsamen Absender „no_reply.events@deloitte.de", nicht von deiner persönlichen Adresse; Rückfragen der Teilnehmer landen daher nicht automatisch bei dir, sondern du wirst im Text als Ansprechpartner genannt.'
                  : 'When you are happy, click Send. The app confirms briefly („Send email to N attendees?") and then places ONE mail into the send queue (DEX_Emails) addressed to all selected recipients; the organizers are automatically added on CC (unless they are already recipients). Delivery usually within minutes. The mail comes from the shared sender „no_reply.events@deloitte.de", not from your personal address; replies therefore do not land with you automatically — instead you are named as the contact in the text.'}
              </>
            ),
            mockup: (
              <Callout variant="info">
                {isDe
                  ? 'Tipp: Für eine Einladung an Personen, die noch NICHT angemeldet sind, nutze stattdessen die Aktion „Einladungsmail" (Anmelde-Link an dich zum Weiterleiten oder direkt an den Mailverteiler des Events).'
                  : 'Tip: To invite people who are NOT registered yet, use the „Invitation email" action instead (registration link to yourself for forwarding or directly to the event mail distribution list).'}
              </Callout>
            ),
          },
          {
            number: 6,
            title: isDe ? 'Wer darf das?' : 'Who is allowed to do this?',
            description: (
              <>
                {isDe
                  ? 'Den Massen-E-Mail-Versand können Organizer für ihre eigenen Events und Admins für alle Events nutzen. Normale Teilnehmer sehen das Organizer Center nicht und haben die Aktion daher nicht.'
                  : 'Mass email sending is available to organizers for their own events and to admins for all events. Regular attendees do not see the Organizer Center and therefore do not have this action.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
