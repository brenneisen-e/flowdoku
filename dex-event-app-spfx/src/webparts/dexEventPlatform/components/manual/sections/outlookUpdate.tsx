import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import EventCreationPage from '../../EventCreationPage';

export function outlookUpdateSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'outlook-update',
    title: isDe ? 'Outlook-Termin: Bild ändern & aktualisieren' : 'Outlook appointment: change image & update',
    category: 'organizer',
    description: isDe
      ? 'Das Bild im Outlook-Kalendereintrag der Teilnehmer austauschen und einen bereits verschickten Termin nachträglich aktualisieren, wenn er noch veraltet ist.'
      : 'Swap the image in the attendees’ Outlook calendar entry and refresh an already-sent appointment when it is still outdated.',
    visibleFor: ['Organizer', 'Admin'],
    keywords: 'outlook termin bild foto invite einladung kalender aktualisieren update kalendereintrag termin aktualisieren outlook bild ändern foto im invite kalender-update calendar appointment refresh resend erneut senden neu verschicken kommunikation schritt 6 outlook-termin jetzt aktualisieren sub-event tab veraltet outdated header-bild',
    perspectives: [
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Wo finde ich das?' : 'Where do I find this?',
            description: (
              <>
                {isDe
                  ? 'Öffne dein Event zum Bearbeiten und wechsle zu Schritt 6 „Kommunikation". Dort gibt es zwei aufklappbare Bereiche rund um den Outlook-Termin: „Bild für den Outlook-Termin" (das Bild oben im Kalendereintrag) und „Text im Outlook-Termin" (der Beschreibungstext). Hat dein Event Sub-Events, kannst du oben zwischen dem Hauptevent und jedem Sub-Event umschalten — die Einstellungen gelten dann jeweils für den gewählten Tab.'
                  : 'Open your event for editing and go to step 6 „Communication". There are two collapsible areas around the Outlook appointment: „Image for the Outlook appointment" (the image at the top of the calendar entry) and „Text in the Outlook appointment" (the description text). If your event has sub-events, you can switch between the main event and each sub-event at the top — the settings then apply to the selected tab.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Event bearbeiten → Schritt 6 „Kommunikation" (echte Ansicht)' : 'Edit event → step 6 „Communication" (real view)'}
                role="Organizer"
                page="edit-event"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
                initialStep={5}
              >
                <Header />
                <EventCreationPage />
              </AppPreview>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Bild ändern' : 'Change the image',
            description: (
              <>
                {isDe
                  ? 'Klappe „Bild für den Outlook-Termin" auf und klicke auf „Bild auswählen". Du kannst z.B. ein Foto vom Veranstaltungsort oder ein Event-Logo hochladen. Mit „Entfernen" setzt du das Bild wieder auf das Standard-Logo zurück. Das Deloitte-Branding (Logo, grüner Balken, Footer) bleibt in jedem Fall erhalten — du tauschst nur das obere Bild.'
                  : 'Expand „Image for the Outlook appointment" and click „Select image". You can upload e.g. a photo of the venue or an event logo. „Remove" resets the image back to the default logo. The Deloitte branding (logo, green bar, footer) always stays — you only swap the top image.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Speichern — und Teilnehmer informieren' : 'Save — and inform attendees',
            description: (
              <>
                {isDe
                  ? 'Wenn du das Event speicherst und es bereits einen Outlook-Termin gibt, fragt dich die App, ob die Teilnehmer eine aktualisierte Termin-Einladung bekommen sollen. Setze pro betroffenem Event bzw. Sub-Event einen Haken — nur dann wird der Kalendereintrag bei allen Teilnehmern neu verschickt. Ohne Haken wird gespeichert, aber der Kalender der Teilnehmer bleibt erst einmal unverändert (die App merkt sich, dass noch ein Update aussteht).'
                  : 'When you save the event and an Outlook appointment already exists, the app asks whether attendees should receive an updated invitation. Tick the box per affected event or sub-event — only then is the calendar entry re-sent to all attendees. Without a tick the event is saved, but the attendees’ calendars stay unchanged for now (the app remembers that an update is still pending).'}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Nachträglich aktualisieren (wenn der Termin noch veraltet ist)' : 'Update afterwards (if the appointment is still outdated)',
            description: (
              <>
                {isDe
                  ? 'In beiden Bereichen („Bild für den Outlook-Termin" und „Text im Outlook-Termin") gibt es zusätzlich den Button „Outlook-Termin jetzt aktualisieren". Falls der Kalendereintrag der Teilnehmer trotz Speichern noch den alten Stand zeigt, klickst du einfach hier — die Einladung wird mit dem zuletzt gespeicherten Stand erneut verschickt. Wichtig: Hast du gerade etwas geändert, speichere bitte zuerst und klicke dann auf den Button (er verschickt den GESPEICHERTEN Stand, nicht ungespeicherte Änderungen).'
                  : 'In both areas („Image for the Outlook appointment" and „Text in the Outlook appointment") there is also a „Update Outlook appointment now" button. If the attendees’ calendar entry still shows the old state despite saving, simply click here — the invitation is re-sent with the last saved state. Important: if you just changed something, please save first and then click the button (it sends the SAVED state, not unsaved changes).'}
              </>
            ),
            mockup: (
              <Callout variant="info">
                {isDe
                  ? 'Den Button gibt es nur beim Bearbeiten eines Events, das schon einen Outlook-Termin hat. Bei einem brandneuen Event entsteht der Termin erst beim Speichern — danach steht der Aktualisieren-Button zur Verfügung. Der Button wirkt jeweils auf den oben gewählten Tab (Hauptevent oder ein bestimmtes Sub-Event).'
                  : 'The button only appears when editing an event that already has an Outlook appointment. For a brand-new event the appointment is created on save — after that the update button is available. The button always acts on the tab selected above (main event or a specific sub-event).'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
