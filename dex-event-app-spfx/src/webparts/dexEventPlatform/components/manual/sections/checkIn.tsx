import * as React from 'react';
import { ManualSection } from '../types';
import { DemoQRCode, DemoCheckInScanner, Callout, ClickPath } from '../ManualMockups';

export function checkInSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'check-in',
    title: isDe ? 'Check-In & QR-Codes' : 'Check-in & QR codes',
    category: 'general',
    description: isDe
      ? 'Kompletter Ablauf für das QR-basierte Check-In am Veranstaltungstag — inklusive der wichtigen Hinweise, warum der Scanner nur im Mobil-Browser (nicht in der SharePoint-App!) funktioniert und wie du die Kamera-Berechtigung korrekt erteilst.'
      : 'Full guide for QR-based check-in on event day — including the important notes on why the scanner only works in the mobile browser (NOT in the SharePoint app!) and how to grant camera permissions correctly.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        title: isDe ? 'Als Teilnehmer:in' : 'As attendee',
        steps: [
          {
            number: 1,
            title: isDe ? 'QR-Code aus der E-Mail' : 'QR code from the email',
            description: (
              <>
                {isDe
                  ? 'Wenn der Organizer QR-Codes versendet hat, findest du in der Mail einen persönlichen QR-Code. Diesen kannst du als Bild speichern oder einfach in der Mail offen lassen.'
                  : 'Once the organizer sends QR codes, you\'ll find a personal QR code in the email. Save it as an image or simply keep the email open.'}
              </>
            ),
            mockup: <DemoQRCode label={isDe ? 'Dein Check-In-Code' : 'Your check-in code'} />,
          },
          {
            number: 2,
            title: isDe ? 'Beim Einlass zeigen' : 'Show at the entrance',
            description: (
              <>
                {isDe
                  ? 'Vor Ort zeigst du deinen QR-Code dem Orga-Team. Die scannen ihn mit ihrem Handy, und du bist eingecheckt.'
                  : 'At the venue, show your QR code to the organizer team. They\'ll scan it with their phone and you\'re checked in.'}
              </>
            ),
          },
        ],
      },
      {
        perspective: 'organizer',
        title: isDe ? 'Als Organizer:in / QR-Scanner:in' : 'As organizer / QR scanner',
        steps: [
          {
            number: 1,
            title: isDe ? 'WICHTIG: App am Event-Tag richtig öffnen' : 'IMPORTANT: open the app correctly on event day',
            description: (
              <>
                {isDe
                  ? 'Der Check-In läuft ausschließlich über dein Smartphone, weil nur das Handy die Kamera zum QR-Scannen liefert. Aber: Öffne den Link UNBEDINGT im Mobil-Browser (Edge oder Safari) — NICHT in der SharePoint-App. Die SharePoint-Mobile-App stellt aus Sicherheitsgründen keinen Kamera-Zugriff bereit; der Scanner bleibt dort dunkel und du bekommst die Meldung "Kamera-Zugriff nicht verfügbar".'
                  : 'Check-in only runs on your smartphone (the phone provides the camera for scanning). But make sure to open the link in the mobile browser (Edge or Safari) — NOT inside the SharePoint app. The SharePoint mobile app blocks camera access for security reasons; the scanner stays dark and you get "camera access not available".'}
              </>
            ),
            mockup: (
              <Callout variant="warning" title={isDe ? 'SharePoint-App vorher deaktivieren oder löschen' : 'Disable or remove the SharePoint app beforehand'}>
                {isDe
                  ? 'Wenn du die SharePoint-App im Arbeitsprofil installiert hast, klickt Microsoft den Link automatisch DORT rein — am Browser vorbei. Lösung vor dem Event-Tag: SharePoint-App im Arbeitsprofil deaktivieren (Android: App-Einstellungen → Deaktivieren) oder komplett löschen. Danach öffnet der Link sauber in Edge/Safari.'
                  : 'If the SharePoint app is installed on your work profile, Microsoft opens the link THERE by default — bypassing the browser. Fix before event day: disable the SharePoint app in the work profile (Android: app settings → Disable) or uninstall it entirely. The link then opens cleanly in Edge/Safari.'}
              </Callout>
            ),
            tip: isDe
              ? 'Tipp: Sende dir den Check-In-Link per Mail aufs Handy und klicke ihn ERST nach dem Deaktivieren der SharePoint-App an.'
              : 'Tip: Email the check-in link to yourself on your phone, and open it ONLY after disabling the SharePoint app.',
          },
          {
            number: 2,
            title: isDe ? 'QR-Codes an alle Teilnehmer versenden' : 'Send QR codes to attendees',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center deines Events gibt es den Button "QR-Codes versenden". Damit werden alle angemeldeten Teilnehmer angeschrieben und bekommen ihren persönlichen QR-Code. Die QR-Codes sind reine Text-Codes (kein sensibler Datenspeicher), sondern verweisen auf den SharePoint-Eintrag.'
                  : 'The event\'s admin center has a "Send QR codes" button. All registered attendees receive a personal QR code via email. QR codes are plain text identifiers (no sensitive data), pointing to the SharePoint record.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Admin Center' : 'Admin Center'} label2={isDe ? 'QR versenden' : 'Send QR'} hint={isDe ? 'Dauert pro Teilnehmer ca. 2-3 Sekunden.' : 'Takes about 2-3 seconds per attendee.'} />,
          },
          {
            number: 3,
            title: isDe ? 'Scanner am Event-Tag öffnen' : 'Open the scanner on event day',
            description: (
              <>
                {isDe
                  ? 'Auf dem Smartphone öffnest du im Admin Center die Check-In-Seite. Die App aktiviert deine Kamera und sucht automatisch nach QR-Codes.'
                  : 'On your phone, open the check-in page from the admin center. The app activates your camera and automatically scans for QR codes.'}
              </>
            ),
            mockup: <DemoCheckInScanner />,
          },
          {
            number: 4,
            title: isDe ? 'Eingecheckt — fertig' : 'Checked in — done',
            description: (
              <>
                {isDe
                  ? 'Sobald ein Code erkannt wird, wird der Teilnehmer auf Status "Eingecheckt" gesetzt. Die Auslastungs-Zahl im Admin Center erhöht sich in Echtzeit.'
                  : 'As soon as a code is recognized, the attendee is switched to status "Checked-in". The admin center attendance counter updates in real time.'}
              </>
            ),
            mockup: <DemoCheckInScanner scanned />,
            tip: isDe
              ? 'Du kannst mehrere Handys gleichzeitig für Check-In nutzen — alle sehen den aktuellen Stand.'
              : 'You can run multiple phones as scanners in parallel — they all see the current state.',
          },
          {
            number: 5,
            title: isDe ? 'Manueller Check-In' : 'Manual check-in',
            description: (
              <>
                {isDe
                  ? 'Hat jemand seinen QR-Code vergessen, kannst du im Admin Center per Suche den Teilnehmer finden und manuell auf "Eingecheckt" setzen.'
                  : 'If someone forgot their QR code, you can find them in the admin center by search and manually set them to "Checked-in".'}
              </>
            ),
            mockup: (
              <Callout variant="tip">
                {isDe
                  ? 'Name, E-Mail oder TeilnehmerID funktionieren alle als Suchbegriff.'
                  : 'Name, email or attendee ID all work as search terms.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
