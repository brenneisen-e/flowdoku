import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import EventCreationPage from '../../EventCreationPage';
import AdminPage from '../../AdminPage';

export function editEventSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'edit-event',
    title: isDe ? 'Event bearbeiten & löschen' : 'Edit & delete events',
    category: 'organizer',
    description: isDe
      ? 'Bestehende Events anpassen, dupliziert, archivieren oder endgültig löschen.'
      : 'Modify, duplicate, archive or permanently delete existing events.',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Event zum Bearbeiten öffnen' : 'Open event for editing',
            description: (
              <>
                {isDe
                  ? 'Im Admin-Bereich findest du alle Events die du organisiert hast (Admin: alle Events). Per Klick auf "Bearbeiten" landest du im gleichen Wizard wie bei der Erstellung — alle Felder vorbelegt.'
                  : 'The admin area lists every event you organize (Admin: all events). Clicking "Edit" opens the same wizard as creation — all fields pre-populated.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Admin Center → Event-Detailansicht (echte Ansicht)' : 'Admin center → event detail view (real view)'}
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
            title: isDe ? 'Änderungen speichern' : 'Save your changes',
            description: (
              <>
                {isDe
                  ? 'Jede Änderung wirkt sich sofort auf die Anmeldung aus — eine Titel-Änderung erscheint beim nächsten Laden der Registrierungsseite. Änderungen an Start-/Endzeit aktualisieren die Outlook-Termine aller bereits angemeldeten Teilnehmer.'
                  : 'Every change takes effect immediately for registrations — a title change appears on the next reload of the registration page. Start/end time changes update Outlook invites for all already-registered attendees.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Event-Wizard im Edit-Modus (echte Ansicht)' : 'Event wizard in edit mode (real view)'}
                role="Organizer"
                page="edit-event"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
              >
                <Header />
                <EventCreationPage />
              </AppPreview>
            ),
            warning: isDe
              ? 'Löschen eines Custom Fields löscht auch die Teilnehmer-Antworten dazu. Überlege gut, ob du ein Feld wirklich entfernen willst.'
              : 'Deleting a custom field also removes attendee answers to it. Think twice before removing a field.',
          },
          {
            number: 3,
            title: isDe ? 'Outlook-Update beim Speichern bewusst entscheiden' : 'Decide explicitly on Outlook updates when saving',
            description: (
              <>
                {isDe
                  ? 'Seit v11.57: Wenn du beim Bearbeiten Titel, Start-/Endzeit oder den Text im Outlook-Termin änderst und das Event bereits einen Outlook-Termin hat, fragt dich die App vor dem Speichern: „Outlook-Termin der Teilnehmer aktualisieren?". Du wirst nicht mehr global gefragt, sondern siehst eine Liste aller geänderten Termine — Hauptevent und betroffene Sub-Events. Pro Termin setzt du einen Haken: angehakte Termine bekommen sofort eine „Aktualisierter Termin"-Mail von Outlook an alle Teilnehmer, nicht angehakte Termine werden zurückgestellt und für später als „ausstehend" markiert. So kannst du z.B. eine Titel-Änderung am Hauptevent rausschicken, aber eine reine Beschreibungs-Korrektur am Sub-Event nicht in die Postfächer der Teilnehmer schicken.'
                  : 'Since v11.57: when editing the title, start/end time or the appointment text and the event already has an Outlook invite, the app asks before saving: “Update the attendees’ Outlook invite?”. You are no longer asked globally — instead you see a list of every changed invite (main event and affected sub-events). You tick a box per invite: ticked invites immediately trigger an “updated meeting” mail from Outlook to all attendees, unticked invites are deferred and flagged as “pending” for later. This lets you e.g. push out a title change on the main event but skip a pure description tweak on a sub-event without spamming attendee mailboxes.'}
                <br /><br />
                {isDe
                  ? 'Solltest du beim Speichern einen Haken nicht setzen, merkt sich die App den ausstehenden Outlook-Sync pro Termin und zeigt dir beim nächsten Bearbeiten oben in Schritt 1 (Grundlagen) einen gelben Hinweis — getrennt für das Hauptevent und/oder die jeweiligen Sub-Events. Beim nächsten Save kannst du den Sync nachholen oder weiter offen lassen.'
                  : 'If you leave a box unticked while saving, the app remembers the pending Outlook sync per invite and shows a yellow notice at the top of Step 1 (Basics) on the next edit — separately for the main event and/or the affected sub-events. On the next save you can resend or keep deferring.'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Hat das Event Sub-Events mit eigenen Outlook-Terminen, listet das Modal jeden Termin einzeln mit einer eigenen Checkbox — kein globaler „alles oder nichts"-Haken mehr.' : 'If the event has sub-events with their own Outlook invites, the modal lists each invite individually with its own checkbox — no more global “all or nothing” toggle.'}</Callout>,
          },
          {
            number: 4,
            title: isDe ? 'Kommunikations-Einstellungen pro Sub-Event' : 'Per-sub-event communication settings',
            description: (
              <>
                {isDe
                  ? 'Seit v11.57: In Schritt 6 (Kommunikation) findest du oben eine Tab-Leiste, sobald dein Event Sub-Events hat. Der erste Tab steht für das Haupt-Event, jeder weitere Tab für einen Sub-Event. Pro Tab kannst du eigene Werte für Mail-Sprache, Outlook-Termin-Text, Mail- und Outlook-Logo sowie die Schalter „E-Mails versenden" / „Outlook-Termin erstellen" pflegen. So lassen sich z.B. ein deutsches und ein englisches Sub-Event sauber nebeneinander pflegen.'
                  : 'Since v11.57: Step 6 (Communication) shows a tab bar at the top whenever the event has sub-events. The first tab is for the main event, every other tab is for a sub-event. Each tab keeps its own values for email language, Outlook appointment text, email and Outlook logo, and the “send emails” / “create Outlook invite” switches. This lets you cleanly run e.g. a German and an English sub-event side by side.'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Beim Tab-Wechsel werden die Werte automatisch zwischengespeichert — Speichern brauchst du erst am Ende des Wizards.' : 'Switching tabs auto-stashes the values — you only need to save at the very end of the wizard.'}</Callout>,
          },
          {
            number: 5,
            title: isDe ? 'Event löschen' : 'Delete event',
            description: (
              <>
                {isDe
                  ? 'Das Löschen eines Events entfernt das DEX_Events-Item und (optional) die zugehörige Subsite. Teilnehmerdaten gehen verloren. Für vergangene Events besser den Status auf "Archiviert" setzen.'
                  : 'Deleting an event removes the DEX_Events item and (optionally) the corresponding subsite. Attendee data is lost. For past events, better set the status to "Archived".'}
              </>
            ),
            mockup: <Callout variant="warning">{isDe ? 'Löschen ist nicht rückgängig machbar.' : 'Deletion is permanent — cannot be undone.'}</Callout>,
          },
        ],
      },
    ],
  };
}
