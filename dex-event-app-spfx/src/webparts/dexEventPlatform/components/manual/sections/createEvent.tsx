import * as React from 'react';
import { ManualSection } from '../types';
import {
  DemoWizardProgress, DemoFormField, DemoOrganizerChips, DemoPillToggles, DemoToggle,
  DemoDateTimePicker, Callout, ClickPath, DemoStepper,
} from '../ManualMockups';

export function createEventSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
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
              ? 'Die Event-Erstellung führt dich durch sieben Schritte. Nur Pflichtfelder (mit *) sind zum Weiterklicken erforderlich — den Rest kannst du jederzeit später ergänzen.'
              : 'Event creation walks you through seven stages. Only required fields (marked *) need to be filled before advancing — everything else can be added later.'}
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
                    ? 'Organisatoren fügst du als Chips hinzu: ins Suchfeld tippen, aus der Liste wählen — der Chip erscheint oben. Reihenfolge per ◀ ▶, Entfernen per ×. Standardmäßig bist du selbst als erster Organizer eingetragen.'
                    : 'Add organizers as chips: type in the search field, pick from the list — the chip appears above. Reorder via ◀ ▶, remove via ×. You\'re added as the first organizer by default.'}
                </p>
              </>
            ),
            mockup: (
              <div>
                <DemoWizardProgress activeStep={0} />
                <div style={{ marginTop: 12, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 14 }}>
                  <DemoFormField label={isDe ? 'Titel' : 'Title'} value="B2Run Frankfurt 2026" required info={isDe ? 'Name des Events' : 'Name of the event'} />
                  <DemoFormField label={isDe ? 'Event-Typ' : 'Event type'} value="B2Run" required type="select" />
                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: '0.82rem', fontWeight: 600, marginBottom: 6 }}>{isDe ? 'Organisatoren' : 'Organizers'}</div>
                    <DemoOrganizerChips />
                  </div>
                </div>
              </div>
            ),
            tip: isDe
              ? 'Nur Nutzer mit Rolle "Organizer" oder "Admin" erscheinen in der Organizer-Suche. Wenn jemand fehlt → Admin bitten, die Rolle zu setzen.'
              : 'Only users with role "Organizer" or "Admin" appear in the search. If someone\'s missing → ask an admin to grant the role.',
          },
          {
            number: 2,
            title: isDe ? 'Grundlagen: Standort-Filter & Audience' : 'Basics: location filter & audience',
            description: (
              <>
                {isDe
                  ? 'Möchtest du das Event nur für bestimmte Standorte sichtbar machen, wähl die entsprechenden Pills aus. Ohne Auswahl sehen alle Mitarbeiter das Event. Mit dem "Audience"-Feld kannst du zusätzlich E-Mail-Listen oder Unternehmensbereiche beschränken.'
                  : 'To scope the event to specific locations, select the corresponding pills. Without selection, every employee sees the event. The "Audience" field lets you additionally restrict by email lists or business units.'}
              </>
            ),
            mockup: (
              <div style={{ background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 14 }}>
                <div style={{ fontSize: '0.82rem', fontWeight: 600, marginBottom: 6 }}>{isDe ? 'Standort-Filter' : 'Location filter'}</div>
                <DemoPillToggles
                  options={['Berlin', 'Frankfurt', 'Köln', 'München', 'Hamburg', 'Düsseldorf', 'Stuttgart']}
                  selected={['Frankfurt', 'Köln']}
                />
              </div>
            ),
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
            mockup: (
              <div>
                <DemoWizardProgress activeStep={1} />
                <div style={{ marginTop: 12 }}>
                  <DemoDateTimePicker
                    label={isDe ? 'Start' : 'Start'}
                    value="12.06.2026, 18:30"
                    afterLabel={isDe ? 'Ende' : 'End'}
                    afterValue="12.06.2026, 22:00"
                  />
                </div>
              </div>
            ),
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
                  ? 'Lege die maximale Teilnehmerzahl fest oder wähle "Unbegrenzt". Aktiviere optional die Warteliste — dann werden Anmeldungen nach Erreichen der Kapazität auf eine Warteschlange gelegt und automatisch nachgerückt, wenn jemand absagt. An- und Abmeldefristen erfassen Datum + Uhrzeit.'
                  : 'Define the max participants or pick "Unlimited". Optionally enable the waitlist — extra registrations queue up and auto-promote when someone cancels. Registration and cancellation deadlines include both date and time.'}
              </>
            ),
            mockup: (
              <div>
                <DemoWizardProgress activeStep={2} />
                <div style={{ marginTop: 12, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 14, display: 'flex', flexDirection: 'column', gap: 10 }}>
                  <DemoToggle label={isDe ? 'Unbegrenzte Teilnehmer' : 'Unlimited participants'} on={false} hint="max 100" />
                  <DemoToggle label={isDe ? 'Warteliste aktiviert' : 'Waitlist enabled'} on />
                  <DemoDateTimePicker label={isDe ? 'Anmeldefrist' : 'Registration deadline'} value="10.06.2026, 23:59" />
                  <DemoDateTimePicker label={isDe ? 'Letzte Abmeldung' : 'Last cancellation' } value="08.06.2026, 12:00" />
                </div>
              </div>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Schritt 4: Custom Fields' : 'Step 4: Custom Fields',
            description: (
              <>
                {isDe
                  ? 'Hier definierst du die Zusatzfragen, die Teilnehmer bei der Anmeldung sehen: Textfelder, Dropdowns, Checkboxen, Datumsfelder, User-Picker (z.B. für Zimmerpartner). Pro Feld wählst du Typ, Label, Pflicht ja/nein und Sichtbarkeit.'
                  : 'Here you define the additional questions attendees see during registration: text fields, dropdowns, checkboxes, date fields, user pickers (e.g. for roommates). Per field you choose: type, label, required yes/no, and visibility.'}
              </>
            ),
            mockup: (
              <Callout variant="tip" title={isDe ? 'Typen' : 'Field types'}>
                {isDe
                  ? 'Text / Textarea / Select / Date / Checkbox / User (Personen-Suche). Leere Labels werden automatisch nicht gespeichert.'
                  : 'Text / Textarea / Select / Date / Checkbox / User (person search). Fields with empty labels are automatically dropped.'}
              </Callout>
            ),
          },
          {
            number: 6,
            title: isDe ? 'Schritt 5: Kommunikation' : 'Step 5: Communication',
            description: (
              <>
                {isDe
                  ? 'Stelle die E-Mail-Sprache ein (DE oder EN), hinterlege einen optionalen Event-Logo für den Mail-Header, passe die Outlook-Terminbeschreibung an und editiere falls nötig die vier System-Templates: Anmeldung, Warteliste, Abmeldung, Nachrücken.'
                  : 'Set the email language (DE or EN), upload an optional event logo for the email header, customize the Outlook invite description, and edit the four system templates if needed: Registration, Waitlist, Cancellation, Promotion.'}
              </>
            ),
            mockup: (
              <Callout variant="info" title={isDe ? 'Sprach-Default' : 'Language default'}>
                {isDe
                  ? 'Seit v4.11 wird die E-Mail-Sprache automatisch auf deine UI-Sprache vorbelegt — du kannst sie pro Event überschreiben.'
                  : 'Since v4.11 email language defaults to your UI locale — you can override per event.'}
              </Callout>
            ),
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
            mockup: <DemoStepper activeStep={6} />,
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
