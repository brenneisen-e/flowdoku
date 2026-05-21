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
                  ? 'Seit v11.57: Wenn du beim Bearbeiten Titel, Start-/Endzeit oder den Text im Outlook-Termin änderst und das Event bereits einen Outlook-Termin hat, fragt dich die App vor dem Speichern: „Outlook-Termin der Teilnehmer aktualisieren?". Setze den Haken nur, wenn du wirklich willst, dass alle angemeldeten Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung bekommen. Bleibt der Haken aus, wird die Outlook-Einladung in Ruhe gelassen — praktisch wenn du nur einen Tippfehler in der Beschreibung korrigierst.'
                  : 'Since v11.57: when editing the title, start/end time or the appointment text and the event already has an Outlook invite, the app asks before saving: “Update the attendees’ Outlook invite?”. Only tick the box if you really want every registered attendee to receive an “updated meeting” notification. Leaving it unchecked leaves the Outlook invite alone — handy when you only fix a typo in the description.'}
                <br /><br />
                {isDe
                  ? 'Solltest du beim Speichern den Haken nicht setzen, merkt sich die App den ausstehenden Outlook-Sync und zeigt dir beim nächsten Bearbeiten dieses Events oben in Schritt 1 (Grundlagen) einen gelben Hinweis. Sobald du das Event wieder aufrufst, kannst du den Sync nachholen — oder absichtlich offen lassen, falls du nicht möchtest, dass die Teilnehmer eine Update-Mail bekommen.'
                  : 'If you leave the box unchecked while saving, the app remembers the pending Outlook sync and shows a yellow notice at the top of Step 1 (Basics) the next time you edit this event. The next time you open the wizard you can either trigger the sync or deliberately leave it pending if you don’t want to send an update mail to the attendees.'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Hat das Event Sub-Events mit eigenen Outlook-Terminen, aktualisiert das Modal sie alle gemeinsam — du musst nicht pro Sub-Event extra speichern.' : 'If the event has sub-events with their own Outlook invites, the modal updates them all together — no need to save each sub-event separately.'}</Callout>,
          },
          {
            number: 4,
            title: isDe ? 'Kommunikations-Einstellungen pro Sub-Event' : 'Per-sub-event communication settings',
            description: (
              <>
                {isDe
                  ? 'Seit v11.57: In Schritt 5 (Kommunikation) findest du oben eine Tab-Leiste, sobald dein Event Sub-Events hat. Der erste Tab steht für das Haupt-Event, jeder weitere Tab für einen Sub-Event. Pro Tab kannst du eigene Werte für Mail-Sprache, Outlook-Termin-Text, Mail- und Outlook-Logo sowie die Schalter „E-Mails versenden" / „Outlook-Termin erstellen" pflegen. So lassen sich z.B. ein deutsches und ein englisches Sub-Event sauber nebeneinander pflegen.'
                  : 'Since v11.57: Step 5 (Communication) shows a tab bar at the top whenever the event has sub-events. The first tab is for the main event, every other tab is for a sub-event. Each tab keeps its own values for email language, Outlook appointment text, email and Outlook logo, and the “send emails” / “create Outlook invite” switches. This lets you cleanly run e.g. a German and an English sub-event side by side.'}
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
