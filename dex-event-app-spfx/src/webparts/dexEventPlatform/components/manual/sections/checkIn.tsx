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
            title: isDe ? 'QR-Code aus der Bestätigungsmail' : 'QR code from the confirmation email',
            description: (
              <>
                {isDe
                  ? 'Sobald der Organizer die QR-Codes versendet (typischerweise 1-2 Tage vor dem Event), bekommst du eine Mail mit deinem persönlichen QR-Code. Der Code identifiziert dich als registrierten Teilnehmer und enthält KEINE sensiblen Daten — nur deine Event- und Teilnehmer-ID.'
                  : 'As soon as the organizer sends out the QR codes (typically 1-2 days before the event), you receive an email containing your personal QR code. The code identifies you as a registered attendee and contains NO sensitive data — only your event and attendee ID.'}
              </>
            ),
            mockup: <DemoQRCode label={isDe ? 'Dein Check-In-Code' : 'Your check-in code'} />,
            tip: isDe
              ? 'Speichere den QR-Code am besten gleich als Screenshot auf dem Handy ab, damit du ihn am Event-Tag offline parat hast — nützlich, falls am Veranstaltungsort schlechter Empfang herrscht.'
              : 'Save the QR code as a screenshot on your phone so you have it offline on event day — useful if reception at the venue is poor.',
          },
          {
            number: 2,
            title: isDe ? 'Beim Einlass zeigen' : 'Show at the entrance',
            description: (
              <>
                {isDe
                  ? 'Vor Ort zeigst du deinen QR-Code dem Check-In-Team. Kamera auf den Code halten, "Pling" — eingecheckt. Kein Wartezimmer, kein Namensabgleich, keine Papierliste.'
                  : 'At the venue, show your QR code to the check-in team. Camera pointed at the code, "beep" — checked in. No waiting, no name matching, no paper list.'}
              </>
            ),
            mockup: (
              <Callout variant="info">
                {isDe
                  ? 'Kein QR-Code erhalten oder versehentlich gelöscht? Das Check-In-Team kann dich auch manuell per Name oder E-Mail einchecken — siehe Organizer-Schritt "Manueller Check-In".'
                  : 'No QR code received or accidentally deleted? The check-in team can also check you in manually by name or email — see the organizer step "Manual check-in".'}
              </Callout>
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
            title: isDe ? 'Optional: QR-Scanner-Rolle pro Event vergeben' : 'Optional: assign QR scanner role per event',
            description: (
              <>
                {isDe
                  ? 'Seit v6.19 kannst du pro Event zusätzlich zu den Organizern explizite "QR-Code-Scanner" definieren. Das sind User, die am Event-Tag nur beim Einchecken helfen sollen — sie haben KEINE weiteren Rechte (kein Teilnehmer-Bearbeiten, kein Mail-Versand, kein Event-Edit). Scanner bekommen auch keine Organizer-Mails und tauchen nicht in der Organizer-Liste auf MyEvents/Registrierungsseite auf. Der Scanner-Pool besteht ausschließlich aus Usern, die bereits Organizer- oder Admin-Rolle in DEX_Roles haben — so wird verhindert, dass versehentlich ein einfacher User zu viele Rechte bekommt.'
                  : 'Since v6.19 you can define explicit "QR code scanners" per event — in addition to the organizers. These are users who only help with check-in on event day; they have NO further rights (no participant editing, no email sending, no event edit). Scanners also receive no organizer emails and do not appear in the organizer list on MyEvents / registration page. The scanner pool only contains users who already have Organizer or Admin role in DEX_Roles — preventing accidental escalation of a plain user.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Event bearbeiten' : 'Edit event'} label2={isDe ? 'Reiter 3 → QR-Code-Scanner' : 'Step 3 → QR code scanners'} hint={isDe ? 'Orange Pills unter den grünen Organizer-Pills.' : 'Orange pills below the green organizer pills.'} />,
            tip: isDe
              ? 'Ideal, wenn du das Check-In an eine Assistentin oder einen Event-Helfer delegieren möchtest, ohne ihnen parallel Event-Bearbeitungsrechte geben zu müssen.'
              : 'Ideal if you want to delegate check-in to an assistant or event helper without granting event-editing rights at the same time.',
          },
          {
            number: 3,
            title: isDe ? 'QR-Codes an alle Teilnehmer versenden' : 'Send QR codes to attendees',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center deines Events klickst du auf "QR-Codes versenden". Nur Teilnehmer mit Status "Angemeldet" werden angeschrieben — Wartelisten-Einträge und bereits Abgemeldete werden übersprungen. Vor dem Versand kommt eine Bestätigungs-Abfrage mit der Anzahl (z.B. "QR-Codes an 42 Teilnehmer versenden?"). Danach wird pro Teilnehmer eine Mail mit personalisiertem QR-Code aus DEX_Emails rausgefeuert, und der Status wird von "Angemeldet" auf "QR versendet" umgesetzt (sichtbar als lila KPI-Kachel). Das dauert ca. 2-3 Sekunden pro Teilnehmer — der Fortschritt wird live neben dem Button angezeigt. Die QR-Codes enthalten KEINE sensiblen Daten, sondern sind reine Event+Teilnehmer-Identifier der Form "DEX|<EventNr>|<E-Mail>".'
                  : 'In the event\'s admin center click "Send QR codes". Only participants with status "Registered" are emailed — waitlist entries and cancelled ones are skipped. A confirmation dialog shows the count (e.g. "Send QR codes to 42 participants?") before the send. Then a personalized QR-code email is queued per participant via DEX_Emails, and the status moves from "Registered" to "QR sent" (visible in the purple KPI tile). Takes ~2-3 seconds per participant — progress is shown live next to the button. The QR codes contain NO sensitive data, they are plain event+attendee identifiers of the form "DEX|<EventNr>|<email>".'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Admin Center' : 'Admin Center'} label2={isDe ? 'QR-Codes versenden' : 'Send QR codes'} hint={isDe ? 'Status wechselt danach von Angemeldet → QR versendet.' : 'Status moves from Registered → QR sent afterwards.'} />,
            tip: isDe
              ? 'Frühestens 1-2 Tage vor dem Event versenden, damit die Mail nicht in den älteren Posteingängen untergeht. Danach angemeldete Teilnehmer bekommen beim nächsten Klick auf den Button ebenfalls einen QR-Code — bereits verschickte werden übersprungen.'
              : 'Send no earlier than 1-2 days before the event, otherwise the email drowns in older inboxes. Participants who register afterwards get their QR code on the next click — already-sent ones are skipped.',
          },
          {
            number: 4,
            title: isDe ? 'Scanner am Event-Tag öffnen (Mobile-Shortcut)' : 'Open the scanner on event day (mobile shortcut)',
            description: (
              <>
                {isDe
                  ? 'Du öffnest die DEX-App im Mobil-Browser (nochmal: NICHT in der SharePoint-App, siehe Schritt 1). Seit v6.22 erscheint für User mit Check-In-Berechtigung (Admin / Organizer / QR-Scanner) direkt oben rechts auf der Landing-Page eine grüne Sprechblase "Geht\'s zum Check-in" mit einem QR-Icon daneben. Ein einziger Tap führt dich — ohne Umweg über Start → Admin-Center — direkt zum Scanner. Wenn du genau EIN Event einchecken darfst, springt der Shortcut sofort in die Scanner-Ansicht dieses Events; bei Admins oder Usern mit mehreren Events kommt erst der Event-Picker (siehe Schritt 5).'
                  : 'You open the DEX app in the mobile browser (again: NOT in the SharePoint app, see step 1). Since v6.22, users with check-in permission (Admin / Organizer / QR scanner) see a green speech bubble "Go to check-in" with a QR icon next to it in the top-right corner of the landing page. A single tap — without the detour via Start → Admin center — takes you directly to the scanner. If you\'re only allowed to check in for ONE event, the shortcut jumps straight to the scanner view; for admins or users with multiple events, the event picker appears first (see step 5).'}
              </>
            ),
            mockup: <DemoCheckInScanner />,
            tip: isDe
              ? 'Der Shortcut erscheint ausschließlich auf Geräten mit Viewport ≤ 768px. Auf dem Desktop bleibt er versteckt — dort benutzt du weiter das Admin-Center (Button "Check-In").'
              : 'The shortcut only appears on devices with viewport ≤ 768px. On desktop it stays hidden — there you use the admin center as before (button "Check-in").',
          },
          {
            number: 5,
            title: isDe ? 'Event auswählen (nur bei mehreren)' : 'Pick event (only if multiple)',
            description: (
              <>
                {isDe
                  ? 'Darfst du mehrere Events einchecken (z.B. als Admin oder Organizer mehrerer Events), zeigt die Check-In-Seite erst eine Liste aller für dich relevanten Events — jeweils mit Titel, Datum und Ort als Karte. Per Tap wählst du das Event, für das du jetzt scannen möchtest. Genau dieses Event wird dann in der Scanner-Maske oben im Titel geführt ("Check-In — <Event-Titel>"), und die KPI-Kacheln + der Live-Counter beziehen sich exakt auf die Teilnehmerliste dieses Events. Bei nur einem zugänglichen Event wird dieser Picker übersprungen.'
                  : 'If you can check in for multiple events (e.g. as admin or organizer of several events), the check-in page first shows a list of all events relevant to you — each as a card with title, date and location. Tap to pick the event you want to scan for now. That event is then shown in the scanner header ("Check-in — <event title>"), and the KPI tiles + live counter refer exactly to the participant list of that event. With only one accessible event this picker is skipped.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Check-In' : 'Check-in'} label2={isDe ? 'Event-Karte antippen' : 'Tap event card'} hint={isDe ? 'Bei einem einzigen Event wird automatisch weitergeleitet.' : 'With a single event you are redirected automatically.'} />,
          },
          {
            number: 6,
            title: isDe ? 'Kamera-Berechtigung erteilen' : 'Grant camera permission',
            description: (
              <>
                {isDe
                  ? 'Beim ersten Start fragt der Browser "Möchte diese Seite deine Kamera nutzen?" → bitte erlauben. Der Scanner nutzt bevorzugt die Rückkamera (environment facing). Wenn der Scanner dunkel bleibt, hast du entweder beim Permission-Prompt "Blockieren" angetippt oder eine andere App belegt die Kamera (z.B. laufender Teams-Anruf).'
                  : 'On first launch the browser asks "Does this site want to use your camera?" → please allow. The scanner prefers the rear camera (environment facing). If the scanner stays dark, you either tapped "Block" on the permission prompt or another app is holding the camera (e.g. an ongoing Teams call).'}
              </>
            ),
            mockup: (
              <Callout variant="tip" title={isDe ? 'Kamera nachträglich freigeben' : 'Re-enable camera afterwards'}>
                {isDe
                  ? 'iOS Safari: aA-Icon links in der Adresszeile → "Website-Einstellungen" → Kamera: "Erlauben". Android Chrome: Schloss-Icon in der Adresszeile → "Berechtigungen" → Kamera: "Zulassen". Danach Seite einmal neu laden. Die App zeigt beim Fehlschlag eine Fehlermeldung mit exakt diesen Anleitungen an.'
                  : 'iOS Safari: "aA" icon on the left of the address bar → "Website settings" → Camera: "Allow". Android Chrome: lock icon in the address bar → "Permissions" → Camera: "Allow". Then reload the page. The app shows the exact instructions as part of the error message if camera start fails.'}
              </Callout>
            ),
            tip: isDe
              ? 'Teams-Anruf, WhatsApp-Videoanruf oder eine andere Kamera-App geöffnet? Erst schließen, dann Scanner-Seite neu laden. Die App erkennt "NotReadableError" und zeigt dir den Grund dann auch explizit an.'
              : 'Teams call, WhatsApp video call, or another camera app open? Close it first, then reload the scanner page. The app detects "NotReadableError" and shows the reason explicitly.',
          },
          {
            number: 7,
            title: isDe ? 'Scannen & Live-Ergebnis (grün / orange / rot)' : 'Scan & live result (green / orange / red)',
            description: (
              <>
                {isDe
                  ? 'Du hältst die Kamera auf den QR-Code — die Erkennung passiert sofort, ohne auf einen Auslöser zu warten. Je nach Zustand des Teilnehmers gibt es drei mögliche Ergebnisse, die jeweils in einer farbigen Ergebnis-Karte oben in der Maske erscheinen: GRÜN = frisch eingecheckt (der Teilnehmer wird auf Status "Eingecheckt" gesetzt, CheckedIn-Counter erhöht sich um 1, Name + Foto werden zur Bestätigung angezeigt). ORANGE = war bereits eingecheckt (kein doppelter Counter-Sprung, nur Hinweis "Bereits eingecheckt am …"). ROT = Code nicht zugeordnet — z.B. QR eines anderen Events oder eines Abgemeldeten. Die Karte nennt dann den Grund. Der Live-Counter oben rechts ("7 eingecheckt") zählt pro Scanner-Session lokal mit, damit du auf einen Blick siehst wie viele du schon durchgecheckt hast.'
                  : 'Point the camera at the QR code — recognition is instant, no shutter button. Depending on the attendee\'s state there are three possible outcomes, each shown as a colored result card at the top of the screen: GREEN = freshly checked in (status "Checked-in", CheckedIn counter +1, name + photo shown as confirmation). ORANGE = was already checked in (no double counter bump, just "Already checked in at …"). RED = code not recognized — e.g. QR from a different event or a cancelled registration. The card names the reason. The live counter top-right ("7 checked in") tracks per scanner session locally so you see at a glance how many you\'ve processed.'}
              </>
            ),
            mockup: <DemoCheckInScanner scanned />,
            tip: isDe
              ? 'Neben dem lokalen Scanner-Counter aktualisieren sich im Admin Center in Echtzeit auch die KPI-Kacheln ("Eingecheckt" wird grün). Auf einem Tablet am Stand kannst du parallel die KPIs zeigen, während auf dem Handy gescannt wird.'
              : 'Besides the local scanner counter, the admin-center KPI tiles update in real time as well ("Checked-in" turns green). On a tablet at the booth you can show the KPIs while scanning on the phone.',
          },
          {
            number: 8,
            title: isDe ? 'Fallback 1: Foto-Upload' : 'Fallback 1: photo upload',
            description: (
              <>
                {isDe
                  ? 'Wenn die Kamera partout nicht startet (z.B. eingebetteter Browser ohne Live-Video-Zugriff, eingeschränktes Firmen-MDM), kannst du stattdessen per "Foto hochladen" arbeiten: Du machst mit der Standard-Kamera-App ein Foto vom QR-Code, wählst es in der Upload-Box aus — die App liest den QR-Code aus dem Bild aus und verarbeitet das Scan-Ergebnis genau wie beim Live-Scan. Ergebnis-Karte grün/orange/rot identisch zu Schritt 7.'
                  : 'If the camera absolutely refuses to start (e.g. embedded browser without live-video access, restricted corporate MDM), you can use "Upload photo" instead: take a photo of the QR code with the default camera app, pick it in the upload box — the app decodes the QR code from the image and processes the result just like a live scan. Result card green/orange/red identical to step 7.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Scanner-Maske' : 'Scanner page'} label2={isDe ? 'Foto hochladen' : 'Upload photo'} hint={isDe ? 'Langsamer als Live-Scan, aber funktioniert überall wo Dateizugriff geht.' : 'Slower than live scan, but works everywhere file access is allowed.'} />,
          },
          {
            number: 9,
            title: isDe ? 'Fallback 2: Manuelle Code-Eingabe' : 'Fallback 2: manual code entry',
            description: (
              <>
                {isDe
                  ? 'Auch der Foto-Upload klappt nicht? Dann tippst du den QR-Code-Inhalt per Hand ein. Der Code hat das feste Format "DEX|<EventNr>|<E-Mail>" — z.B. "DEX|7|max.mustermann@deloitte.de". Der Teilnehmer kann das direkt aus seiner Bestätigungsmail ablesen (steht unter dem QR-Bild als Monospace-Text). Nach Enter läuft die gleiche Check-In-Logik durch wie beim Scan.'
                  : 'Photo upload also fails? Then type the QR code content manually. The code has a fixed format "DEX|<EventNr>|<email>" — e.g. "DEX|7|max.mustermann@deloitte.de". The attendee can read it straight from their confirmation email (shown as monospace text below the QR image). After Enter, the same check-in logic runs as for the scan.'}
              </>
            ),
            mockup: (
              <Callout variant="info">
                {isDe
                  ? 'Die App kürzt nichts ab — wirklich den kompletten Code inkl. der beiden Pipe-Zeichen ("|") eingeben. Groß-/Kleinschreibung bei der E-Mail ist egal.'
                  : 'The app does not auto-correct — enter the full code including both pipe characters ("|"). Email case does not matter.'}
              </Callout>
            ),
          },
          {
            number: 10,
            title: isDe ? 'Manueller Check-In ohne QR-Code' : 'Manual check-in without QR code',
            description: (
              <>
                {isDe
                  ? 'Ist ein Teilnehmer ohne QR-Code aufgeschlagen (Mail gelöscht, kein Handy-Empfang, vergessen), findest du ihn im Admin Center über das Suchfeld oberhalb der Teilnehmertabelle. Die Suche greift parallel auf Name, Vorname, E-Mail und TeilnehmerID — ein Suchbegriff reicht. In der Tabelle ist pro Zeile rechts in der Aktionen-Spalte ein "Einchecken"-Button → ein Klick und der Teilnehmer ist auf Status "Eingecheckt". Der Live-Counter und die KPI-Kachel aktualisieren sich sofort.'
                  : 'If an attendee arrives without a QR code (email deleted, no phone reception, just forgot), you find them in the admin center via the search field above the participant table. The search matches first name, last name, email and attendee ID in parallel — one term suffices. Each table row has an "Einchecken" button in the right-hand action column → one click and the attendee is on status "Checked-in". The live counter and KPI tile update immediately.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Admin Center → Teilnehmerliste' : 'Admin Center → participants'} label2={isDe ? 'Einchecken' : 'Check in'} hint={isDe ? 'Name, E-Mail oder TeilnehmerID funktionieren als Suchbegriff.' : 'Name, email or attendee ID all work as search term.'} />,
          },
          {
            number: 11,
            title: isDe ? 'Check-In rückgängig machen' : 'Undo check-in',
            description: (
              <>
                {isDe
                  ? 'Wurde jemand versehentlich eingecheckt (falscher QR-Code, Verwechslung mit Namensvetter), kannst du das rückgängig machen: in der Admin-Center-Teilnehmerliste wechselt der "Einchecken"-Button bei bereits eingecheckten Zeilen auf "Auschecken". Ein Klick setzt den Status zurück auf "Angemeldet" (bzw. "QR versendet", je nachdem was vorher war), der Eingecheckt-Counter zählt um 1 runter. Keine Mail, kein Audit-Trail-Eintrag — ist eine reine Korrektur.'
                  : 'If someone was checked in by mistake (wrong QR code, name mixup), you can undo it: in the admin center participant list, the "Einchecken" button switches to "Auschecken" on already-checked-in rows. A click resets status to "Registered" (or "QR sent" depending on the previous state), the checked-in counter decrements by 1. No email, no audit trail — this is a plain correction.'}
              </>
            ),
            mockup: (
              <Callout variant="tip">
                {isDe
                  ? 'Aus dem Scanner heraus geht das Rückgängig-Machen nicht — dafür ist der Admin-Center-Weg nötig. Im Zweifel also zwischendurch aufs Admin-Center umschalten und wieder zurück in den Scanner.'
                  : 'Undo is not available directly from the scanner view — the admin-center route is needed. When in doubt, switch to the admin center briefly and then back to the scanner.'}
              </Callout>
            ),
          },
          {
            number: 12,
            title: isDe ? 'Mehrere Scanner-Geräte parallel' : 'Multiple scanner devices in parallel',
            description: (
              <>
                {isDe
                  ? 'Bei großen Events reicht ein einzelnes Handy oft nicht. Du kannst beliebig viele Geräte parallel als Scanner betreiben — jeder Organizer / Admin / QR-Scanner öffnet unabhängig seine eigene Scanner-Session. Da alle gegen dieselbe SharePoint-Teilnehmerliste arbeiten, sehen alle Geräte den aktuellen Zustand. Ein Doppel-Scan (z.B. zwei Geräte scannen denselben Teilnehmer gleichzeitig) wird erkannt — das zweite Gerät meldet orange "Bereits eingecheckt", der Counter springt nicht doppelt.'
                  : 'For large events a single phone often isn\'t enough. You can run as many devices in parallel as scanners as you like — each organizer / admin / QR scanner opens their own independent scanner session. Since they all work against the same SharePoint participant list, all devices see the current state. A double scan (e.g. two devices scan the same attendee at the same time) is detected — the second device reports orange "Already checked in", the counter does not double-count.'}
              </>
            ),
            tip: isDe
              ? 'Typisches Setup bei 200+ Teilnehmern: 3 Scanner-Handys am Einlass (je ein Organizer/Helfer) + 1 Tablet im Backoffice mit offenem Admin-Center, wo die KPI-Kacheln live mitlaufen.'
              : 'Typical setup for 200+ attendees: 3 scanner phones at the entrance (one organizer/helper each) + 1 tablet in the back office with the admin center open, showing live KPI tiles.',
          },
        ],
      },
    ],
  };
}
