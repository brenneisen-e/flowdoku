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
