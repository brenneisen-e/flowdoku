import * as React from 'react';
import { ManualSection } from '../types';
import { Callout, ClickPath } from '../ManualMockups';

export function subEventsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'sub-events',
    title: isDe ? 'Sub-Events (Trainingssessions)' : 'Sub-events (training sessions)',
    category: 'organizer',
    description: isDe
      ? 'Innerhalb eines Hauptevents mehrere Sessions anlegen — jede mit eigenem Termin, Ort, Kapazität, eigener An-/Abmelde-Mail und eigenem Outlook-Kalendereintrag. Seit v6.4 sind Sub-Events eigenständige DEX_Events-Items mit parentEventId — dadurch funktionieren Teilnehmerlisten, Warteliste, Outlook-Termine, Custom-Fields und Mail-Templates identisch zu Top-Level-Events (ohne Sonderlogik).'
      : 'Create multiple sessions inside a main event — each with its own date, location, capacity, own registration/cancellation emails and dedicated Outlook calendar entry. Since v6.4 sub-events are standalone DEX_Events items with parentEventId — so participant lists, waitlist, Outlook invites, custom fields and email templates work identically to top-level events (no special logic).',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Wofuer sind Sub-Events?' : 'What are sub-events for?',
            description: (
              <>
                {isDe
                  ? 'Beispiel B2Run: Das Hauptevent ist der eigentliche Lauf. Dazu gibt es mehrere optionale Trainingssessions, fuer die sich Teilnehmer zusaetzlich anmelden koennen. Jede Session bekommt einen eigenen Outlook-Termin und eine eigene Bestaetigungsmail. Teilnehmer muessen sich zuerst zum Hauptevent anmelden und buchen danach beliebige Sessions einzeln dazu.'
                  : 'Example B2Run: The main event is the actual run. In addition there are several optional training sessions attendees can sign up for. Each session gets its own Outlook invite and a dedicated confirmation email. Attendees register for the main event first and then opt in to any sessions individually.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Sub-Events anlegen (beim Event-Erstellen oder -Bearbeiten)' : 'Create sub-events (during event create/edit)',
            description: (
              <>
                {isDe
                  ? 'Im Event-Wizard zu Reiter 5 "Kommunikation" wechseln — dort gibt es unterhalb der E-Mail-Templates die Sektion "Sub-Events". Pro Session legst du Titel, Beschreibung, Ort, Start-/Endzeit und optional eine Max-Teilnehmerzahl sowie einen Anmeldeschluss fest. Zwei Schalter pro Session steuern, ob eigene Mails bzw. ein eigener Outlook-Termin erzeugt werden.'
                  : 'In the event wizard, go to step 5 "Communication" — below the email template list you will find the "Sub-events" section. For each session define title, description, location, start/end time and optionally a capacity and registration deadline. Two toggles per session control whether dedicated emails and/or a dedicated Outlook calendar entry are created.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Reiter 5 Kommunikation' : 'Step 5 Communication'} label2={isDe ? '+ Sub-Event hinzufuegen' : '+ Add sub-event'} />,
          },
          {
            number: 3,
            title: isDe ? 'Kommunikation & Deloitte-Layout' : 'Communication & Deloitte layout',
            description: (
              <>
                {isDe
                  ? 'Alle Mails zu Sub-Events (An-/Abmeldung) und Outlook-Termine nutzen automatisch das Deloitte-Layout (Logo, gruener Balken, Footer). Der Body wird pro Session aus Titel, Beschreibung, Datum und Ort zusammengesetzt und in die Deloitte-Vorlage gewrappt. Teilnehmer bekommen also fuer jede gebuchte Session eine separate, sauber gestaltete Mail und einen eigenen Kalendereintrag.'
                  : 'All sub-event emails (registration/cancellation) and Outlook invites automatically use the Deloitte layout (logo, green bar, footer). The body per session is composed from title, description, date and location and wrapped into the Deloitte template. Attendees receive a clean, branded email and a dedicated calendar entry for each booked session.'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Kein Power-Automate-Flow-Update noetig — Sub-Event-Details nutzen die bestehende DEX_Outlook-Queue mit Override-Feldern.' : 'No Power Automate flow update required — sub-event details use the existing DEX_Outlook queue with override fields.'}</Callout>,
          },
          {
            number: 4,
            title: isDe ? 'Teilnehmer verwalten pro Session' : 'Manage attendees per session',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center erscheint in der Teilnehmerliste ein zusaetzliches Dropdown, sobald das Event Sub-Events hat. Du kannst filtern nach: "Alle" (komplette Teilnehmerliste), "Nur Hauptevent" (ohne Session-Anmeldung) oder einem konkreten Sub-Event. Im Dropdown siehst du pro Session direkt die Auslastung (z.B. "12/25").'
                  : 'The admin center shows an additional dropdown above the attendee list as soon as the event has sub-events. Filter by: "All" (complete list), "Main event only" (no session) or a specific sub-event. The dropdown displays the current capacity per session inline (e.g. "12/25").'}
              </>
            ),
            mockup: <ClickPath label="Admin" label2={isDe ? 'Sub-Event-Filter' : 'Sub-event filter'} />,
          },
        ],
      },
      {
        perspective: 'user',
        steps: [
          {
            number: 1,
            title: isDe ? 'Sessions sind unabhängig vom Hauptevent' : 'Sessions are independent from the main event',
            description: (
              <>
                {isDe
                  ? 'Seit v6.13 kannst du dich für Sessions an- und abmelden, ohne zwangsläufig am Hauptevent teilzunehmen. Auf der Registrierungsseite eines Events siehst du im linken Bereich die Session-Liste mit einem "Anmelden"-Button pro Session — du kannst dich direkt für einzelne Sessions anmelden, auch wenn du nicht auf "Jetzt anmelden" für das Hauptevent klickst. Umgekehrt kannst du das Hauptevent buchen ohne eine Session dazu.'
                  : 'Since v6.13 you can register and cancel sessions independently of the main event. On the event registration page the session list appears in the left column with a "Register" button per session — you can sign up for individual sessions without submitting the main-event form. Conversely, you can register for the main event without picking any session.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Sessions direkt auf der Registrierungsseite oder danach buchen' : 'Pick sessions on the registration page or right after',
            description: (
              <>
                {isDe
                  ? 'Die Session-Liste erscheint sowohl vor dem Absenden der Hauptevent-Anmeldung als auch auf dem Bestätigungsbildschirm — beide Stellen führen zum selben Ergebnis. Pro Klick bekommst du eine eigene Bestätigungsmail und einen separaten Outlook-Kalendereintrag — beides im Deloitte-Layout.'
                  : 'The session list appears both before submitting the main registration and on the confirmation screen — both locations lead to the same result. Each click triggers a dedicated confirmation email and a separate Outlook calendar entry — both in Deloitte layout.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Sessions unter "My Events" verwalten' : 'Manage sessions in "My Events"',
            description: (
              <>
                {isDe
                  ? 'Unter "My Events" siehst du pro Event deine gebuchten Sessions. Du kannst jederzeit zusaetzliche Sessions buchen oder einzelne wieder stornieren — das Hauptevent bleibt davon unberuehrt.'
                  : 'Under "My Events" each event card lists your booked sessions. You can add or cancel individual sessions at any time — the main-event registration is not affected.'}
              </>
            ),
            mockup: <ClickPath label="My Events" label2={isDe ? 'Session abmelden' : 'Cancel session'} />,
          },
        ],
      },
    ],
  };
}
