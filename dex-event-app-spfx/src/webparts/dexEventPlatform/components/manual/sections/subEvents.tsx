import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import EventCreationPage from '../../EventCreationPage';
import AdminPage from '../../AdminPage';
import MyEventsPage from '../../MyEventsPage';

export function subEventsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'sub-events',
    title: isDe ? 'Sub-Events (Sessions, Workshops, Side-Events)' : 'Sub-events (sessions, workshops, side-events)',
    category: 'organizer',
    description: isDe
      ? 'Innerhalb eines Hauptevents mehrere Sessions anlegen — jede mit eigenem Termin, Ort, Kapazität, eigener An-/Abmelde-Mail und eigenem Outlook-Kalendereintrag. Seit v6.4 sind Sub-Events eigenständige DEX_Events-Items mit parentEventId — dadurch funktionieren Teilnehmerlisten, Warteliste, Outlook-Termine, Custom-Fields und Mail-Templates identisch zu Top-Level-Events (ohne Sonderlogik).'
      : 'Create multiple sessions inside a main event — each with its own date, location, capacity, own registration/cancellation emails and dedicated Outlook calendar entry. Since v6.4 sub-events are standalone DEX_Events items with parentEventId — so participant lists, waitlist, Outlook invites, custom fields and email templates work identically to top-level events (no special logic).',
    visibleFor: ['User', 'Organizer', 'Admin'],
    keywords: isDe
      ? 'Sub-Event Sub-Events Subevent Sessions Workshop Workshops Side-Event Nebenveranstaltung Programmpunkt Programmbaustein Track Breakout Break-out Session mehrtägig Offsite Networking-Dinner Kick-off Session buchen dazubuchen Session-Auswahl Nur Sessions Haupt-Event Bezeichnung umbenennen parentEventId Schritt 3 Sub-Events nutzen Sub-Event-Filter Auslastung pro Session'
      : 'sub-event sub-events subevent session sessions workshop workshops side-event side event breakout break-out track programme item building block multi-day offsite networking dinner kick-off session book add-on session selection sessions only main event label rename parentEventId step 3 use sub-events sub-event filter per-session capacity',
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Wofür sind Sub-Events?' : 'What are sub-events for?',
            description: (
              <>
                {isDe
                  ? 'Sub-Events sind eigenständige Sessions innerhalb eines Hauptevents — z.B. ein mehrtägiges Offsite mit Tagesblocks, ein Workshop-Track mit mehreren Themen-Sessions, ein Networking-Dinner als optionaler Add-on-Termin oder eine Kick-off-Session vor dem eigentlichen Event. Jedes Sub-Event hat einen eigenen Termin, Ort, Kapazität, eigene An-/Abmelde-Mail und einen eigenen Outlook-Kalendereintrag. Teilnehmer können sich zum Hauptevent anmelden und beliebige Sessions einzeln dazubuchen — entweder gleich bei der initialen Anmeldung oder später ueber „Meine Events".'
                  : 'Sub-events are standalone sessions inside a main event — e.g. a multi-day offsite with daily blocks, a workshop track with several theme sessions, a networking dinner as an optional add-on, or a kick-off session ahead of the main event. Each sub-event has its own date, location, capacity, own registration/cancellation email and dedicated Outlook calendar entry. Attendees register for the main event and then opt in to any sessions individually — either right at initial registration or later via "My Events".'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Sub-Events anlegen (beim Event-Erstellen oder -Bearbeiten)' : 'Create sub-events (during event create/edit)',
            description: (
              <>
                {isDe
                  ? 'Im Event-Wizard zu Schritt 3 "Sub-Events" wechseln und den Schalter "Sub-Events nutzen?" aktivieren — die gesamte Sub-Event-Konfiguration klappt darunter auf. Pro Session legst du Titel, Beschreibung, Ort, Start-/Endzeit und optional eine Max-Teilnehmerzahl sowie einen Anmeldeschluss fest. Zwei Schalter pro Session steuern, ob eigene Mails bzw. ein eigener Outlook-Termin erzeugt werden.'
                  : 'In the event wizard, go to step 3 "Sub-events" and flip the "Use sub-events?" toggle — the entire sub-event configuration unfolds below it. For each session define title, description, location, start/end time and optionally a capacity and registration deadline. Two toggles per session control whether dedicated emails and/or a dedicated Outlook calendar entry are created.'}
                <br /><br />
                <strong>{isDe ? 'Outlook nachträglich aktivieren — ohne Teilnehmer-Verlust (v11.69):' : 'Enabling Outlook after the fact — without losing attendees (v11.69):'}</strong>{' '}
                {isDe
                  ? 'Wenn du auf einem bereits gespeicherten Sub-Event den Outlook-Schalter nachträglich einschaltest oder den Termin-Text änderst, erkennt die App das beim Speichern und öffnet das Outlook-Update-Modal. Pro betroffenes Sub-Event setzt du einen Haken: für Sub-Events, die noch keinen Outlook-Termin haben, wird die DEX_Events-Zeile in der Eventverwaltung neu angelegt (damit der Termin entsteht), die bestehende Teilnehmer-Subsite mit allen Anmeldungen, Teilnehmer-IDs und Custom-Field-Antworten bleibt aber unangetastet erhalten. Nur die Event-ID in der Verwaltung ändert sich — für Teilnehmer ist das transparent. Vor v11.69 war das ein destruktives delete-and-recreate; jetzt ist es safe.'
                  : 'If you turn the Outlook toggle on for a sub-event that has already been saved (or change the calendar body), the app detects this when you save and opens the Outlook update modal. For each affected sub-event you tick a box: for sub-events that do not yet have an Outlook event, the DEX_Events row in event admin is re-created (so the calendar entry can be generated), but the existing participant subsite — with all registrations, participant IDs and custom-field answers — stays intact. Only the event ID in admin changes; for attendees this is transparent. Before v11.69 this was a destructive delete-and-recreate; it is safe now.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Wizard Schritt 3: Sub-Events verwalten (echte Ansicht)' : 'Wizard step 3: Manage sub-events (real view)'}
                role="Organizer"
                page="edit-event"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
                initialStep={2}
              >
                <Header />
                <EventCreationPage />
              </AppPreview>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Kommunikation & Deloitte-Layout' : 'Communication & Deloitte layout',
            description: (
              <>
                {isDe
                  ? 'Alle Mails zu Sub-Events (An-/Abmeldung) und Outlook-Termine nutzen automatisch das Deloitte-Layout (Logo, grüner Balken, Footer). Der Body wird pro Session aus Titel, Beschreibung, Datum und Ort zusammengesetzt und in die Deloitte-Vorlage gewrappt. Teilnehmer bekommen also für jede gebuchte Session eine separate, sauber gestaltete Mail und einen eigenen Kalendereintrag.'
                  : 'All sub-event emails (registration/cancellation) and Outlook invites automatically use the Deloitte layout (logo, green bar, footer). The body per session is composed from title, description, date and location and wrapped into the Deloitte template. Attendees receive a clean, branded email and a dedicated calendar entry for each booked session.'}
                <br /><br />
                <strong>{isDe ? 'Bezeichnung des Haupt-Events in der Auswahl (v24.58):' : 'Main-event label in the selection (v24.58):'}</strong>{' '}
                {isDe
                  ? 'Hat dein Event Sub-Events, sieht der Teilnehmer auf der Anmeldeseite mehrere wählbare Bereiche. Über dem Hauptevent steht standardmäßig „Haupt-Event". Im Sub-Events-Schritt kannst du das umbenennen (z.B. „Konferenz", „Hauptprogramm") oder ganz weglassen, sodass nur der Event-Titel erscheint. Eine Live-Vorschau zeigt dir sofort, wie es der Teilnehmer sieht.'
                  : 'If your event has sub-events, attendees see several selectable areas on the registration page. By default the main event is labelled „Main event". In the sub-events step you can rename it (e.g. „Conference", „Main programme") or drop it entirely so only the event title is shown. A live preview shows exactly what the attendee will see.'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Kein Power-Automate-Flow-Update nötig — Sub-Event-Details nutzen die bestehende DEX_Outlook-Queue mit Override-Feldern.' : 'No Power Automate flow update required — sub-event details use the existing DEX_Outlook queue with override fields.'}</Callout>,
          },
          {
            number: 4,
            title: isDe ? 'Teilnehmer verwalten pro Session' : 'Manage attendees per session',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center erscheint in der Teilnehmerliste ein zusätzliches Dropdown, sobald das Event Sub-Events hat. Du kannst filtern nach: "Alle" (komplette Teilnehmerliste), "Nur Hauptevent" (ohne Session-Anmeldung) oder einem konkreten Sub-Event. Im Dropdown siehst du pro Session direkt die Auslastung (z.B. "12/25").'
                  : 'The admin center shows an additional dropdown above the attendee list as soon as the event has sub-events. Filter by: "All" (complete list), "Main event only" (no session) or a specific sub-event. The dropdown displays the current capacity per session inline (e.g. "12/25").'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Admin Center → Sub-Event-Filter (echte Ansicht)' : 'Admin center → sub-event filter (real view)'}
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
        ],
      },
      {
        perspective: 'user',
        steps: [
          {
            number: 1,
            title: isDe ? 'Integrierte Event-Auswahl auf der Registrierungsseite' : 'Integrated event selection on the registration page',
            description: (
              <>
                {isDe
                  ? 'Seit v6.14 wählst du auf der Registrierungsseite direkt per Checkbox aus, wofür du dich anmelden möchtest — das Haupt-Event und/oder einzelne Sessions. Jede Checkbox kann unabhängig an- oder abgewählt werden. Beim Absenden wird für jede aktivierte Auswahl eine eigene Registrierung durchgeführt — pro Session eine eigene Bestätigungsmail und ein eigener Outlook-Kalendereintrag im Deloitte-Layout.'
                  : 'Since v6.14 you pick directly on the registration page — via checkboxes — what to register for: the main event and/or individual sessions. Each checkbox can be toggled independently. On submit, a separate registration is created for every selected item — each session gets its own confirmation email and its own Outlook calendar entry, both in the Deloitte layout.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Nur Sessions ohne Haupt-Event' : 'Sessions only (no main event)',
            description: (
              <>
                {isDe
                  ? 'Du kannst die Haupt-Event-Checkbox abwählen und dich ausschließlich für eine oder mehrere Sessions anmelden. In dem Fall erscheint auf der Bestätigungsseite der Hinweis "Du hast dich ausschließlich für die ausgewählten Sessions angemeldet". In "My Events" taucht das Haupt-Event mit orangefarbenem Badge "Nur Sessions" auf, damit du die Session-Anmeldungen weiterhin einsehen und verwalten kannst — ohne dass du dadurch automatisch am Haupt-Event teilnimmst.'
                  : 'You can uncheck the main-event checkbox and register only for one or more sessions. In that case the confirmation screen explains "You have registered for the selected sessions only". In "My Events" the main event card shows up with an orange "Sessions only" badge so you can keep viewing and managing your session registrations — without automatically participating in the main event.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'B2Run: Durchstarter / Funstarter pro Session' : 'B2Run: Durchstarter / Funstarter per session',
            description: (
              <>
                {isDe
                  ? 'Bei B2Run-Events mit getrennten Durchstarter-/Funstarter-Kapazitäten fragt die Registrierungsseite auch pro Session nach dem Starter-Typ — der Wert wird in die Session-Teilnehmerliste geschrieben. Wenn du dich gleichzeitig für das Haupt-Event anmeldest, erbt die Session automatisch den dort gewählten Starter-Typ — dann entfällt die Einzel-Auswahl pro Session (keine doppelte Abfrage).'
                  : 'For B2Run events with split Durchstarter/Funstarter capacities, the registration page asks for the starter type per session — the value is written to the session participant list. If you register for the main event at the same time, the session inherits that starter type automatically — no duplicate prompt per session.'}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Sessions unter "My Events" verwalten' : 'Manage sessions in "My Events"',
            description: (
              <>
                {isDe
                  ? 'Unter "My Events" siehst du pro Event deine gebuchten Sessions. Du kannst jederzeit zusätzliche Sessions buchen oder einzelne wieder stornieren — das Hauptevent bleibt davon unberührt.'
                  : 'Under "My Events" each event card lists your booked sessions. You can add or cancel individual sessions at any time — the main-event registration is not affected.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? '"Meine Events" mit Sessions (echte Ansicht)' : '"My Events" with sessions (real view)'}
                role="User"
                page="my-events"
                width={430}
                device="phone"
              >
                <Header />
                <MyEventsPage />
              </AppPreview>
            ),
          },
        ],
      },
    ],
  };
}
