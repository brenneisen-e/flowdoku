import * as React from 'react';
import { ManualSection } from '../types';
import { ClickPath } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import Header from '../../Header';
import EventCreationPage from '../../EventCreationPage';

export function createEventSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';

  // Kleine Helfer-Factory, damit die 7 Wizard-Previews einheitlich aussehen.
  // Jeder Preview öffnet die echte EventCreationPage im Create-Modus, setzt
  // via initialStep den gewünschten Wizard-Schritt (0..6). Laptop-Frame, weil
  // das Event-Anlegen auf dem Desktop stattfindet.
  const wizardPreview = (stepIndex: number, stepLabel: string): React.ReactElement => (
    <AppPreview
      label={stepLabel}
      role="Organizer"
      page="create-event"
      width={1024}
      device="laptop"
      initialStep={stepIndex}
    >
      <Header />
      <EventCreationPage />
    </AppPreview>
  );

  return {
    id: 'create-event',
    title: isDe ? 'Event erstellen (7-Schritte-Wizard)' : 'Create an event (7-step wizard)',
    category: 'organizer',
    description: isDe
      ? 'Schritt-für-Schritt-Anleitung durch alle sieben Stufen der Event-Erstellung.'
      : 'Step-by-step walkthrough of the seven event-creation stages.',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        intro: (
          <>
            {isDe
              ? 'Die Event-Erstellung führt dich durch sieben Schritte. Nur Pflichtfelder (mit *) sind zum Weiterklicken erforderlich — den Rest kannst du jederzeit später ergänzen. Jeder Schritt unten hat unter "Vorschau der echten App" einen Button, der dir den jeweiligen Wizard-Step direkt öffnet.'
              : 'Event creation walks you through seven stages. Only required fields (marked *) need to be filled before advancing — everything else can be added later. Each step below has a "Real app preview" button that opens the matching wizard step.'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Schritt 1: Grundlagen' : 'Step 1: Basics',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Wähle zuerst eine Vorlage: "Leer" für eigene Felder oder "B2Run" für alle Pflichtfelder einer B2Run-Anmeldung. Danach: Titel, Event-Typ und Organisatoren.'
                    : 'Pick a template first: "Blank" for custom fields, or "B2Run" to pre-fill all required fields for a B2Run registration. Then: title, event type, organizers.'}
                </p>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Organisatoren fügst du als Chips hinzu: ins Suchfeld tippen, aus der Liste wählen — der Chip erscheint oben. Reihenfolge per ◀ ▶, Entfernen per ×. Standardmäßig bist du selbst als erster Organizer eingetragen. In derselben Sektion findest du auch den QR-Code-Scanner-Picker (orange Pills) für Leute, die am Event-Tag nur einchecken sollen.'
                    : 'Add organizers as chips: type in the search field, pick from the list — the chip appears above. Reorder via ◀ ▶, remove via ×. You\'re added as the first organizer by default. The same section holds the QR code scanner picker (orange pills) for people who should only handle check-in on event day.'}
                </p>
              </>
            ),
            mockup: wizardPreview(0, isDe ? 'Wizard Schritt 1: Grundlagen (echte Ansicht)' : 'Wizard step 1: Basics (real view)'),
            tip: isDe
              ? 'Nur Nutzer mit Rolle "Organizer" oder "Admin" erscheinen in der Organizer- und Scanner-Suche. Wenn jemand fehlt → Admin bitten, die Rolle zu setzen.'
              : 'Only users with role "Organizer" or "Admin" appear in the organizer and scanner search. If someone\'s missing → ask an admin to grant the role.',
          },
          {
            number: 2,
            title: isDe ? 'Grundlagen: Standort-Filter & Audience' : 'Basics: location filter & audience',
            description: (
              <>
                {isDe
                  ? 'Weiter unten auf derselben Seite: Möchtest du das Event nur für bestimmte Standorte sichtbar machen, wähl die entsprechenden Pills aus. Ohne Auswahl sehen alle Mitarbeiter das Event. Mit dem "Audience"-Feld kannst du zusätzlich E-Mail-Listen oder Unternehmensbereiche beschränken.'
                  : 'Further down the same page: to scope the event to specific locations, select the corresponding pills. Without selection, every employee sees the event. The "Audience" field lets you additionally restrict by email lists or business units.'}
              </>
            ),
            mockup: wizardPreview(0, isDe ? 'Wizard Schritt 1: Standort-Filter (runterscrollen)' : 'Wizard step 1: Location filter (scroll down)'),
          },
          {
            number: 3,
            title: isDe ? 'Schritt 2: Zeit & Ort' : 'Step 2: Date & Location',
            description: (
              <>
                {isDe
                  ? 'Start- und Enddatum legen den Zeitraum fest (inklusive Uhrzeit für den Outlook-Kalender). Die App verhindert, dass Start nach Ende liegt. Der Ort fließt in die Event-Karte und in die Anmeldebestätigungs-Mail.'
                  : 'Start and end date define the timespan (including time for the Outlook calendar). The app prevents start being after end. The location appears on the event card and in the confirmation email.'}
              </>
            ),
            mockup: wizardPreview(1, isDe ? 'Wizard Schritt 2: Zeit & Ort (echte Ansicht)' : 'Wizard step 2: Date & Location (real view)'),
            warning: isDe
              ? 'Datum und Uhrzeit gehen direkt in den Outlook-Termin — also sorgfältig prüfen, bevor Einladungen rausgehen.'
              : 'Date and time flow directly into the Outlook invite — double-check before invitations go out.',
          },
          {
            number: 4,
            title: isDe ? 'Schritt 3: Kapazität & Fristen' : 'Step 3: Capacity & Deadlines',
            description: (
              <>
                {isDe
                  ? 'Lege die maximale Teilnehmerzahl fest oder wähle "Unbegrenzt". Aktiviere optional die Warteliste — dann werden Anmeldungen nach Erreichen der Kapazität auf eine Warteschlange gelegt und automatisch nachgerückt, wenn jemand absagt. An- und Abmeldefristen erfassen Datum + Uhrzeit. Für B2Run-Events findest du hier zusätzlich getrennte Durchstarter-/Funstarter-Kapazitäten + Starter-Typ→Startblock-Zuordnung + optional Leistungsnachweis-Pflicht.'
                  : 'Define the max participants or pick "Unlimited". Optionally enable the waitlist — extra registrations queue up and auto-promote when someone cancels. Registration and cancellation deadlines include both date and time. For B2Run events you additionally find split Durchstarter/Funstarter capacities + starter-type → start-block mapping + optional proof-of-performance requirement.'}
              </>
            ),
            mockup: wizardPreview(2, isDe ? 'Wizard Schritt 3: Kapazität (echte Ansicht)' : 'Wizard step 3: Capacity (real view)'),
          },
          {
            number: 5,
            title: isDe ? 'Schritt 4: Custom Fields' : 'Step 4: Custom Fields',
            description: (
              <>
                {isDe
                  ? 'Hier definierst du die Zusatzfragen, die Teilnehmer bei der Anmeldung sehen. Pro Feld stellst du folgendes ein: '
                  : 'Here you define the additional questions attendees see during registration. Per field you can configure: '}
                <ul style={{ marginTop: 6, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe
                    ? 'Feldtyp: Text, Zahl, Dropdown, Checkbox, Personen-Suche oder Roommate (Doppelzimmer). Roommate ist seit v7.17 ein eigener Typ — nur er löst beim Anmelden eine automatische "Zimmerpartner-Anfrage"-Mail an die ausgewählte Person aus.'
                    : 'Field type: text, number, dropdown, checkbox, people search or roommate (double room). Roommate is its own type since v7.17 — only it triggers an automatic "roommate request" email to the picked person on registration.'}</li>
                  <li>{isDe
                    ? 'Mehrfachauswahl bei Dropdowns (z.B. mehrere Allergien anhaken). Werte werden " | "-getrennt gespeichert.'
                    : 'Multi-select for dropdowns (e.g. tick multiple allergies). Values are stored " | "-separated.'}</li>
                  <li>{isDe
                    ? 'Pflichtfeld (rotes Sternchen, Anmeldung blockiert wenn leer)'
                    : 'Required (red asterisk, blocks submit when empty)'}</li>
                  <li>{isDe
                    ? 'Beschreibung pro Feld — landet als „i"-Tooltip neben dem Label'
                    : 'Description per field — appears as „i" tooltip next to the label'}</li>
                  <li>{isDe
                    ? 'Sichtbarkeitsbedingung: Feld nur dann anzeigen wenn eine andere Frage einen bestimmten Wert hat. Beispiel: „Roommate" nur fragen wenn „Zimmerart = Doppelzimmer". Versteckte Felder blockieren auch die Pflichtfeld-Validierung nicht.'
                    : 'Visibility condition: only show the field when another question has a specific value. Example: "Roommate" only asked when "Room type = Double". Hidden fields also do not block required-field validation.'}</li>
                  <li>{isDe
                    ? 'Reihenfolge per Drag-Handle oder ▲▼-Pfeile — Nummerierung passt sich automatisch an'
                    : 'Reordering via drag handle or ▲▼ arrows — numbering updates automatically'}</li>
                  <li>{isDe
                    ? 'Vorlagen: „Vorgeschlagene Felder" (Katalog mit T-Shirt, Allergien, Hotel, Roommate, …) und das vollständige B2Run-Template mit Startblock, Gruppe, Altersklasse, Datenschutz-Checkbox + Links.'
                    : 'Templates: "Suggested fields" (catalogue with t-shirt, allergies, hotel, roommate, …) and the full B2Run template with start block, group, age category, privacy checkbox + links.'}</li>
                </ul>
                {isDe
                  ? 'Datenschutz-Hinweis oben: Erhebe nur Daten die zwingend für das Event nötig sind — bei Unklarheiten ist immer Rücksprache mit dem Datenschutz durchzuführen.'
                  : 'Privacy notice on top: only collect data strictly necessary for the event — in case of doubt, always check with the data-protection officer.'}
              </>
            ),
            mockup: wizardPreview(3, isDe ? 'Wizard Schritt 4: Felder (echte Ansicht)' : 'Wizard step 4: Fields (real view)'),
            tip: isDe
              ? 'Mit dem Vorschau-Button rechts oben siehst du jederzeit live wie das Anmeldeformular für den Teilnehmer aussehen wird — inklusive Sichtbarkeitsbedingungen und i-Tooltips.'
              : 'Use the preview button top right to see live what the attendee\'s registration form will look like — including visibility conditions and i-tooltips.',
          },
          {
            number: 6,
            title: isDe ? 'Schritt 5: Kommunikation' : 'Step 5: Communication',
            description: (
              <>
                {isDe
                  ? 'Stelle die E-Mail-Sprache ein (DE oder EN), hinterlege einen optionalen Event-Logo für den Mail-Header, passe die Outlook-Terminbeschreibung an und editiere falls nötig die vier System-Templates: Anmeldung, Warteliste, Abmeldung, Nachrücken. Darunter findest du die Sub-Event-Verwaltung (Trainingssessions etc.).'
                  : 'Set the email language (DE or EN), upload an optional event logo for the email header, customize the Outlook invite description, and edit the four system templates if needed: Registration, Waitlist, Cancellation, Promotion. Below you find the sub-event (training session) management.'}
              </>
            ),
            mockup: wizardPreview(4, isDe ? 'Wizard Schritt 5: Kommunikation (echte Ansicht)' : 'Wizard step 5: Communication (real view)'),
          },
          {
            number: 7,
            title: isDe ? 'Schritt 6: Dokumente' : 'Step 6: Documents',
            description: (
              <>
                {isDe
                  ? 'Lade PDFs oder Dokumente hoch, die Teilnehmer vor dem Event sehen sollen: Agenda, Anreise-Infos, Sicherheitshinweise. Die Dokumente werden als Attachments am Event-Item gespeichert.'
                  : 'Upload PDFs or documents that attendees see before the event: agenda, travel info, safety notes. Documents are stored as attachments on the event item.'}
              </>
            ),
            mockup: wizardPreview(5, isDe ? 'Wizard Schritt 6: Dokumente (echte Ansicht)' : 'Wizard step 6: Documents (real view)'),
          },
          {
            number: 8,
            title: isDe ? 'Schritt 7: Fun-Zone (Quiz)' : 'Step 7: Fun-Zone (Quiz)',
            description: (
              <>
                {isDe
                  ? 'Optional: Erstelle ein Quiz, das Teilnehmer vor oder während des Events spielen. Jede Frage hat 4 Antwortoptionen und eine markierte richtige Antwort. Siehe Kapitel "Quiz / Fun-Zone" für Details.'
                  : 'Optional: create a quiz for attendees to play before or during the event. Each question has 4 answer options and one marked correct. See "Quiz / Fun-Zone" for details.'}
              </>
            ),
            mockup: wizardPreview(6, isDe ? 'Wizard Schritt 7: Fun-Zone (echte Ansicht)' : 'Wizard step 7: Fun-Zone (real view)'),
          },
          {
            number: 9,
            title: isDe ? 'Vorschau & speichern' : 'Preview & save',
            description: (
              <>
                {isDe
                  ? 'Klick "Vorschau anzeigen" — du siehst, wie das Event auf der Registrierungsseite wirken wird. Mit "Event erstellen" legst du es an: im Hintergrund entsteht automatisch eine SharePoint-Subsite mit der Teilnehmerliste, und dein Event taucht in der Übersicht auf.'
                  : 'Click "Preview" — you\'ll see how the event looks on the registration page. Click "Create event" to commit: a SharePoint subsite with the participant list is auto-provisioned, and your event appears in the overview.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Vorschau' : 'Preview'} label2={isDe ? 'Event erstellen' : 'Create event'} />,
            warning: isDe
              ? 'Die Subsite wird bei erstmaligem Speichern angelegt — das dauert einige Sekunden. Bitte nicht die Seite neu laden während "Event wird erstellt…" angezeigt wird.'
              : 'The subsite is provisioned on first save — this takes a few seconds. Do not reload while "Creating event..." is shown.',
          },
        ],
      },
    ],
  };
}

