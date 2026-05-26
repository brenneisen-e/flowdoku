import * as React from 'react';
import { ManualSection } from '../types';
import { DemoQRCode, DemoEmailPreview, Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import LandingPage from '../../LandingPage';
import CheckInPage from '../../CheckInPage';
import AdminPage from '../../AdminPage';
import EventCreationPage from '../../EventCreationPage';
import { DeloitteEvent } from '../../../types';

// v6.28: Zweit-Demo-Event für den Check-In-Event-Picker — mit zwei
// Events zu Auswahl wird die Picker-UI sichtbar (bei nur einem Event
// würde die App automatisch weiterleiten).
const demoSecondEvent: DeloitteEvent = {
  id: 'demo-event-2', eventNumber: 43, title: 'Summer Party Düsseldorf',
  type: 'Other', status: 'Active',
  organizers: ['Musterfrau, Maja'], organizerEmails: ['maja.musterfrau@deloitte.de'],
  qrScannerNames: [], qrScannerEmails: [],
  location: 'Rheinterrasse Düsseldorf', locationAudience: ['Düsseldorf'],
  audienceFilter: [], filterMode: 'OR',
  startDate: '2026-07-04T18:00:00', endDate: '2026-07-04T23:00:00',
  registrationDeadline: '2026-07-01T23:59:00', lastDeregisterDate: '2026-07-03T23:59:00',
  description: 'Sommerfest am Rhein.',
  maxParticipants: 200, currentParticipants: 134, waitlistCount: 0,
  imageUrl: '', outlookBody: '', emailLanguage: 'DE',
  agenda: [], transferTimes: [], documents: [], quiz: [],
  eventSpecificFields: [],
};

export function checkInSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'check-in',
    title: isDe ? 'Check-In & QR-Codes' : 'Check-in & QR codes',
    category: 'general',
    description: isDe
      ? 'Kompletter Ablauf für das Check-In am Veranstaltungstag — Standard ist die Namens-Suche im Admin Center (geht am Laptop, Tablet UND Handy); der Live-Kamera-Scan ist die schnellere Alternative bei großen Events, läuft aber ausschließlich auf dem Smartphone (inkl. der wichtigen Setup-Hinweise: SharePoint-App im Arbeitsprofil vorher deaktivieren, sonst kein Kamera-Zugriff!).'
      : 'Full guide for check-in on event day — by default search-by-name in the admin center (works on laptop, tablet AND phone); the live-camera scan is the faster alternative for large events, but runs on smartphone only (incl. the critical setup notes: disable the SharePoint app on the work profile beforehand, otherwise no camera access!).',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        title: isDe ? 'Als Teilnehmer:in' : 'As attendee',
        steps: [
          {
            number: 1,
            title: isDe ? 'QR-Code + Name aus der Bestätigungsmail' : 'QR code + name from the confirmation email',
            description: (
              <>
                {isDe
                  ? 'Sobald der Organizer die QR-Codes versendet (typischerweise 1-2 Tage vor dem Event), bekommst du eine Mail mit deinem persönlichen QR-Code. Direkt unter dem Code stehen dein Name und deine Teilnehmer-ID — falls der Scan am Event-Tag nicht klappt, reicht das Check-In-Team dich problemlos per Namen ein. Der QR-Code selbst enthält KEINE sensiblen Daten — nur deine Event- und Teilnehmer-ID.'
                  : 'As soon as the organizer sends out the QR codes (typically 1-2 days before the event), you receive an email with your personal QR code. Right below the code you find your name and attendee ID — if the scan fails on event day, the check-in team can check you in by name without any issue. The QR code itself contains NO sensitive data — only your event and attendee ID.'}
              </>
            ),
            mockup: (
              <div style={{ padding: '16px 8px', display: 'flex', justifyContent: 'center', background: 'var(--dex-gray-100)', borderRadius: 8 }}>
                <div style={{
                  width: 320, padding: '28px 10px 14px',
                  borderRadius: 34, background: '#1a1a1a',
                  boxShadow: '0 8px 24px rgba(0,0,0,0.24), inset 0 0 0 2px #333',
                  position: 'relative',
                }}>
                  <div style={{
                    position: 'absolute', top: 10, left: '50%', transform: 'translateX(-50%)',
                    width: 80, height: 14, background: '#000', borderRadius: 8,
                  }} />
                  <div style={{ background: '#fff', borderRadius: 22, overflow: 'hidden' }}>
                    <DemoEmailPreview
                      subject={isDe ? 'Dein Check-In-Code für Office Event Köln' : 'Your check-in code for Office Event Köln'}
                      heading={isDe ? 'Dein persönlicher QR-Code' : 'Your personal QR code'}
                      body={
                        <>
                          <p style={{ margin: '0 0 12px', fontSize: '0.85rem', lineHeight: 1.5 }}>
                            {isDe
                              ? 'Hallo Max, zeige diesen QR-Code am Event-Tag am Einlass — das Orga-Team scannt ihn und du bist eingecheckt. Wenn der Scan nicht klappt, nennst du einfach deinen Namen — beides funktioniert.'
                              : 'Hi Max, please show this QR code at the entrance on event day — the team will scan it and check you in. If the scan fails, just give them your name — both work.'}
                          </p>
                          <div style={{ textAlign: 'center', margin: '12px 0' }}>
                            <DemoQRCode size={160} label={isDe ? 'Dein Check-In-Code' : 'Your check-in code'} />
                          </div>
                          <div style={{ textAlign: 'center', margin: '4px 0 0' }}>
                            <div style={{ fontSize: '0.92rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>
                              Mustermann, Max
                            </div>
                            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {isDe ? 'Teilnehmer-ID' : 'Attendee ID'} 17
                            </div>
                          </div>
                        </>
                      }
                    />
                  </div>
                </div>
              </div>
            ),
            tip: isDe
              ? 'Speichere den QR-Code am besten gleich als Screenshot auf dem Handy ab, damit du ihn am Event-Tag offline parat hast — nützlich, falls am Veranstaltungsort schlechter Empfang herrscht.'
              : 'Save the QR code as a screenshot on your phone so you have it offline on event day — useful if reception at the venue is poor.',
          },
          {
            number: 2,
            title: isDe ? 'Beim Einlass: QR zeigen oder Namen nennen' : 'At the entrance: show QR or give your name',
            description: (
              <>
                {isDe
                  ? 'Vor Ort hast du zwei gleichwertige Wege: QR-Code dem Check-In-Team zeigen → Kamera drauf → „Pling", eingecheckt. Oder einfach deinen Namen nennen — das Check-In-Team findet dich in 1-2 Sekunden in der Liste und klickt auf „Einchecken". Beide Wege landen auf demselben Endpunkt; welcher schneller geht, hängt vom Veranstaltungs-Setup ab.'
                  : 'At the venue you have two equivalent paths: show the QR code to the team → camera on it → „beep", checked in. Or simply give your name — the team finds you in 1-2 seconds in the list and clicks „Check-in". Both routes lead to the same outcome; which is faster depends on the on-site setup.'}
              </>
            ),
          },
        ],
      },
      {
        perspective: 'organizer',
        title: isDe ? 'Als Organizer:in / QR-Scanner:in' : 'As organizer / QR scanner',
        intro: isDe ? (
          <>
            Das Check-In läuft über <strong>zwei Wege</strong>: <strong>Weg A</strong> ist die einfache Namens-Suche im Admin Center — funktioniert <strong>am Laptop, Tablet UND Handy</strong>, ohne irgendein Setup. <strong>Weg B</strong> ist der Live-Kamera-Scan und ist <strong>ausschließlich am Smartphone</strong> nutzbar (am Laptop gibt es keinen Scanner — also dort immer Weg A). Schneller bei großen Events, braucht aber einmaliges Setup (SharePoint-App im Arbeitsprofil deaktivieren, Kamera-Berechtigung erteilen). Schritt 1 (QR-Codes versenden) ist die gemeinsame Vorbereitung für beide Wege.
          </>
        ) : (
          <>
            Check-in runs along <strong>two paths</strong>: <strong>Path A</strong> is the simple search-by-name in the admin center — works on <strong>laptop, tablet AND phone</strong>, without any setup. <strong>Path B</strong> is the live camera scan and is <strong>smartphone-only</strong> (no scanner exists on laptop — so on a laptop always use Path A). Faster for large events but needs one-time setup (disable the SharePoint app on the work profile, grant camera permission). Step 1 (sending QR codes) is the shared preparation for both paths.
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Vorbereitung: QR-Codes an alle Teilnehmer versenden' : 'Preparation: send QR codes to all attendees',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center deines Events klickst du auf „QR-Codes versenden". Nur Teilnehmer mit Status „Angemeldet" werden angeschrieben — Wartelisten-Einträge und bereits Abgemeldete werden übersprungen. Vor dem Versand kommt eine Bestätigungs-Abfrage mit der Anzahl (z.B. „QR-Codes an 42 Teilnehmer versenden?"). Danach wird pro Teilnehmer eine Mail mit personalisiertem QR-Code + Name + Teilnehmer-ID verschickt, und der Status wird von „Angemeldet" auf „QR versendet" umgesetzt (sichtbar als lila KPI-Kachel). Diese Vorbereitung brauchen beide Wege — egal ob du nachher per Namens-Suche oder per Kamera-Scan eincheckst.'
                  : 'In the event\'s admin center click „Send QR codes". Only participants with status „Registered" are emailed — waitlist entries and cancelled ones are skipped. A confirmation dialog shows the count (e.g. „Send QR codes to 42 participants?") before sending. Then each attendee gets an email with their personalized QR code + name + attendee ID, and the status moves from „Registered" to „QR sent" (visible in the purple KPI tile). Both paths need this preparation — whether you later check in by name or by camera scan.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Admin Center → Event-Detailansicht (echte Ansicht)' : 'Admin center → event detail view (real view)'}
                role="Organizer"
                page="admin"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
              >
                <Header />
                <AdminPage />
              </AppPreview>
            ),
            tip: isDe
              ? 'Frühestens 1-2 Tage vor dem Event versenden, damit die Mail nicht in den älteren Posteingängen untergeht. Danach angemeldete Teilnehmer bekommen beim nächsten Klick auf den Button ebenfalls einen QR-Code — bereits verschickte werden übersprungen.'
              : 'Send no earlier than 1-2 days before the event, otherwise the email drowns in older inboxes. Participants who register afterwards get their QR code on the next click — already-sent ones are skipped.',
          },
          {
            number: 2,
            title: isDe ? 'Weg A (Standard): Namens-Suche im Admin Center' : 'Path A (default): search by name in the admin center',
            description: (
              <>
                {isDe
                  ? 'Der einfachste Weg — funktioniert auf jedem Gerät (Laptop, Tablet, Handy) ohne Setup. Im Admin Center deines Events tippst du oberhalb der Teilnehmertabelle in das Suchfeld den Namen, Vornamen, die E-Mail oder die Teilnehmer-ID — alle Felder werden parallel durchsucht, ein Suchbegriff reicht. Die Tabelle filtert sich sofort. Rechts in der Aktionen-Spalte klickst du bei der richtigen Zeile auf „Einchecken" — der Status springt auf „Eingecheckt", die grüne KPI-Kachel zählt um 1 hoch. Der Teilnehmer kann seinen Namen direkt aus der Bestätigungsmail ablesen (steht unter dem QR-Code), und du musst nichts installieren oder konfigurieren.'
                  : 'The simplest path — works on any device (laptop, tablet, phone) without any setup. In the admin center of your event, type the name, first name, email or attendee ID into the search box above the participant table — all fields are searched in parallel, one term is enough. The table filters instantly. In the right-hand action column click „Check-in" on the matching row — status flips to „Checked-in", the green KPI tile increments by 1. The attendee can read their name directly from the confirmation email (printed below the QR code), and you don\'t need to install or configure anything.'}
              </>
            ),
            mockup: (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                <Callout variant="tip" title={isDe ? 'Funktioniert auch hervorragend am Handy' : 'Works great on a phone too'}>
                  {isDe
                    ? 'Die Teilnehmertabelle ist responsive — auf dem Handy genauso bedienbar wie am Desktop. Bei sehr großen Listen ist Desktop entspannter (mehr Zeilen gleichzeitig sichtbar), aber kleines Event mit 30-50 Personen geht problemlos vom Handy.'
                    : 'The participant table is responsive — you can use it on a phone just as on desktop. For very large lists a desktop is more relaxed (more rows visible at once), but a small event with 30-50 people is fine from the phone.'}
                </Callout>
                <AppPreview
                  label={isDe ? 'Desktop-Ansicht: Admin Center → Teilnehmertabelle mit Suchfeld + Einchecken-Buttons' : 'Desktop view: admin center → participant table with search + check-in buttons'}
                  role="Organizer"
                  page="admin"
                  selectedEventId={DEMO_EVENT_ID}
                  width={1024}
                  device="laptop"
                >
                  <Header />
                  <AdminPage />
                </AppPreview>
                <AppPreview
                  label={isDe ? 'Mobile-Ansicht: gleiche Tabelle am Handy' : 'Mobile view: same table on phone'}
                  role="Organizer"
                  page="admin"
                  selectedEventId={DEMO_EVENT_ID}
                  width={430}
                  device="phone"
                >
                  <Header />
                  <AdminPage />
                </AppPreview>
              </div>
            ),
            tip: isDe
              ? 'Standard-Weg für kleine bis mittlere Events (bis ~100 Personen). Schnell genug, keine Kamera-Probleme, keine SharePoint-App-Hürden, kein Risiko dass am Eingang das Setup zickt.'
              : 'Default path for small to medium events (up to ~100 people). Fast enough, no camera issues, no SharePoint-app hurdles, no risk that on-site setup misbehaves.',
          },
          {
            number: 3,
            title: isDe ? 'Weg B Vorbereitung 1: SharePoint-App im Arbeitsprofil deaktivieren' : 'Path B prep 1: disable the SharePoint app in the work profile',
            description: (
              <>
                {isDe
                  ? 'Wenn du am Event-Tag den Live-Kamera-Scan nutzen willst (schneller bei großen Events), gibt es eine wichtige Hürde: Die SharePoint-Mobile-App im Arbeitsprofil fängt den Check-In-Link automatisch ab und öffnet ihn intern — und stellt aus Sicherheitsgründen KEINEN Kamera-Zugriff bereit. Der Scanner bleibt dann dunkel und meldet „Kamera-Zugriff nicht verfügbar". Lösung VOR dem Event-Tag: die SharePoint-App im Arbeitsprofil deaktivieren (Android: App-Einstellungen → Deaktivieren) oder ganz löschen. Danach öffnet der Link sauber in Edge oder Safari, und der Kamera-Zugriff funktioniert normal.'
                  : 'If you want to use the live camera scan on event day (faster for large events), there\'s an important hurdle: the SharePoint mobile app on the work profile intercepts the check-in link automatically and opens it internally — and for security reasons it provides NO camera access. The scanner stays dark and reports „Camera access not available". Fix BEFORE event day: disable the SharePoint app on the work profile (Android: app settings → Disable) or uninstall it entirely. Then the link opens cleanly in Edge or Safari, and camera access works normally.'}
              </>
            ),
            mockup: (
              <Callout variant="warning" title={isDe ? 'Vor dem Event-Tag erledigen — nicht erst am Eingang!' : 'Do this before event day — not at the entrance!'}>
                {isDe
                  ? 'Android: Einstellungen → Apps → Arbeitsprofil → SharePoint → Deaktivieren. iOS: SharePoint-App löschen (auf der App-Kachel lange drücken → „Entfernen"). Falls du SharePoint später wieder brauchst, kannst du sie aus dem Company Portal jederzeit neu installieren. Sende dir den Check-In-Link per Mail aufs Handy und klicke ihn ERST nach dem Deaktivieren der SharePoint-App an — sonst hat Microsoft die App schon als Standard-Handler registriert.'
                  : 'Android: Settings → Apps → Work profile → SharePoint → Disable. iOS: delete the SharePoint app (long-press the app tile → „Remove"). If you need SharePoint again later, reinstall it from Company Portal at any time. Email the check-in link to yourself on your phone and open it ONLY after disabling the SharePoint app — otherwise Microsoft has already registered the app as default handler.'}
              </Callout>
            ),
            tip: isDe
              ? 'Skip-able für Weg A — die Namens-Suche im Admin Center braucht keine Kamera, also auch keinen Edge/Safari-Workaround.'
              : 'Skippable for Path A — search-by-name in the admin center needs no camera, so no Edge/Safari workaround either.',
          },
          {
            number: 4,
            title: isDe ? 'Weg B Vorbereitung 2: Optional — QR-Scanner-Rolle pro Event vergeben' : 'Path B prep 2: optional — assign QR-scanner role per event',
            description: (
              <>
                {isDe
                  ? 'Seit v6.19 kannst du pro Event zusätzlich zu den Organizern explizite „QR-Code-Scanner" definieren. Das sind User, die am Event-Tag nur beim Einchecken helfen sollen — sie haben KEINE weiteren Rechte (kein Teilnehmer-Bearbeiten, kein Mail-Versand, kein Event-Edit). Scanner bekommen auch keine Organizer-Mails und tauchen nicht in der Organizer-Liste auf MyEvents/Registrierungsseite auf. Der Scanner-Pool besteht ausschließlich aus Usern, die bereits Organizer- oder Admin-Rolle in DEX_Roles haben — so wird verhindert, dass versehentlich ein einfacher User zu viele Rechte bekommt.'
                  : 'Since v6.19 you can define explicit „QR code scanners" per event — in addition to the organizers. These are users who only help with check-in on event day; they have NO further rights (no participant editing, no email sending, no event edit). Scanners also receive no organizer emails and do not appear in the organizer list on MyEvents / registration page. The scanner pool only contains users who already have Organizer or Admin role in DEX_Roles — preventing accidental escalation of a plain user.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Event-Editor → QR-Code-Scanner-Picker (echte Ansicht)' : 'Event editor → QR code scanner picker (real view)'}
                role="Organizer"
                page="edit-event"
                selectedEventId={DEMO_EVENT_ID}
                width={960}
              >
                <Header />
                <EventCreationPage />
              </AppPreview>
            ),
            tip: isDe
              ? 'Ideal, wenn du das Check-In an eine Assistentin oder einen Event-Helfer delegieren möchtest, ohne ihnen parallel Event-Bearbeitungsrechte geben zu müssen.'
              : 'Ideal if you want to delegate check-in to an assistant or event helper without granting event-editing rights at the same time.',
          },
          {
            number: 5,
            title: isDe ? 'Weg B: Scanner am Event-Tag öffnen (Mobile-Shortcut)' : 'Path B: open the scanner on event day (mobile shortcut)',
            description: (
              <>
                {isDe
                  ? 'Du öffnest die DEX-App im Mobil-Browser (nochmal: NICHT in der SharePoint-App, siehe Schritt 3). Seit v6.26 erscheint für User mit Check-In-Berechtigung (Admin / Organizer / QR-Scanner) direkt im Header neben dem QR-Code-Icon eine grüne Sprechblase „Jetzt einchecken". Ein einziger Tap führt dich — ohne Umweg über Start → Admin-Center — direkt zum Scanner. Wenn du genau EIN Event einchecken darfst, springt der Shortcut sofort in die Scanner-Ansicht dieses Events; bei Admins oder Usern mit mehreren Events kommt erst der Event-Picker (siehe Schritt 6).'
                  : 'You open the DEX app in the mobile browser (again: NOT in the SharePoint app, see step 3). Since v6.26, users with check-in permission (Admin / Organizer / QR scanner) see a green speech bubble „Check in now" right in the header next to the QR code icon. A single tap — without the detour via Start → Admin center — takes you directly to the scanner. If you\'re only allowed to check in for ONE event, the shortcut jumps straight to the scanner view; for admins or users with multiple events, the event picker appears first (see step 6).'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Landing-Page mit „Jetzt einchecken"-Bubble' : 'Landing page with „Check in now" bubble'}
                role="Organizer"
                width={390}
              >
                <Header />
                <LandingPage />
              </AppPreview>
            ),
            tip: isDe
              ? 'Der Shortcut erscheint ausschließlich auf Geräten mit Viewport ≤ 768px. Auf dem Desktop bleibt er versteckt — dort benutzt du weiter das Admin-Center (Button „Check-In").'
              : 'The shortcut only appears on devices with viewport ≤ 768px. On desktop it stays hidden — there you use the admin center as before (button „Check-in").',
          },
          {
            number: 6,
            title: isDe ? 'Weg B: Event auswählen (nur bei mehreren)' : 'Path B: pick event (only if multiple)',
            description: (
              <>
                {isDe
                  ? 'Darfst du mehrere Events einchecken (z.B. als Admin oder Organizer mehrerer Events), zeigt die Check-In-Seite erst eine Liste aller für dich relevanten Events — jeweils mit Titel, Datum und Ort als Karte. Per Tap wählst du das Event, für das du jetzt scannen möchtest. Genau dieses Event wird dann in der Scanner-Maske oben im Titel geführt („Check-In — <Event-Titel>"), und die KPI-Kacheln + der Live-Counter beziehen sich exakt auf die Teilnehmerliste dieses Events. Bei nur einem zugänglichen Event wird dieser Picker übersprungen.'
                  : 'If you can check in for multiple events (e.g. as admin or organizer of several events), the check-in page first shows a list of all events relevant to you — each as a card with title, date and location. Tap to pick the event you want to scan for now. That event is then shown in the scanner header („Check-in — <event title>"), and the KPI tiles + live counter refer exactly to the participant list of that event. With only one accessible event this picker is skipped.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Check-In: Event-Picker (echte Ansicht)' : 'Check-in: event picker (real view)'}
                role="Admin"
                page="check-in"
                extraEvents={[demoSecondEvent]}
                width={390}
              >
                <Header />
                <CheckInPage />
              </AppPreview>
            ),
          },
          {
            number: 7,
            title: isDe ? 'Weg B: Kamera-Berechtigung erteilen' : 'Path B: grant camera permission',
            description: (
              <>
                {isDe
                  ? 'Beim ersten Start fragt der Browser „Möchte diese Seite deine Kamera nutzen?" → bitte erlauben. Der Scanner nutzt bevorzugt die Rückkamera (environment facing). Wenn der Scanner dunkel bleibt, hast du entweder beim Permission-Prompt „Blockieren" angetippt oder eine andere App belegt die Kamera (z.B. laufender Teams-Anruf).'
                  : 'On first launch the browser asks „Does this site want to use your camera?" → please allow. The scanner prefers the rear camera (environment facing). If the scanner stays dark, you either tapped „Block" on the permission prompt or another app is holding the camera (e.g. an ongoing Teams call).'}
              </>
            ),
            mockup: (
              <Callout variant="tip" title={isDe ? 'Kamera nachträglich freigeben' : 'Re-enable camera afterwards'}>
                {isDe
                  ? 'iOS Safari: aA-Icon links in der Adresszeile → „Website-Einstellungen" → Kamera: „Erlauben". Android Chrome: Schloss-Icon in der Adresszeile → „Berechtigungen" → Kamera: „Zulassen". Danach Seite einmal neu laden. Die App zeigt beim Fehlschlag eine Fehlermeldung mit exakt diesen Anleitungen an.'
                  : 'iOS Safari: „aA" icon on the left of the address bar → „Website settings" → Camera: „Allow". Android Chrome: lock icon in the address bar → „Permissions" → Camera: „Allow". Then reload the page. The app shows the exact instructions as part of the error message if camera start fails.'}
              </Callout>
            ),
            tip: isDe
              ? 'Teams-Anruf, WhatsApp-Videoanruf oder eine andere Kamera-App geöffnet? Erst schließen, dann Scanner-Seite neu laden. Die App erkennt „NotReadableError" und zeigt dir den Grund dann auch explizit an.'
              : 'Teams call, WhatsApp video call, or another camera app open? Close it first, then reload the scanner page. The app detects „NotReadableError" and shows the reason explicitly.',
          },
          {
            number: 8,
            title: isDe ? 'Weg B: Scannen & Live-Ergebnis (grün / orange / rot)' : 'Path B: scan & live result (green / orange / red)',
            description: (
              <>
                {isDe
                  ? 'Du hältst die Kamera auf den QR-Code — die Erkennung passiert sofort, ohne auf einen Auslöser zu warten. Je nach Zustand des Teilnehmers gibt es drei mögliche Ergebnisse, die jeweils in einer farbigen Ergebnis-Karte oben in der Maske erscheinen: GRÜN = frisch eingecheckt (der Teilnehmer wird auf Status „Eingecheckt" gesetzt, CheckedIn-Counter erhöht sich um 1, Name + Foto werden zur Bestätigung angezeigt). ORANGE = war bereits eingecheckt (kein doppelter Counter-Sprung, nur Hinweis „Bereits eingecheckt am …"). ROT = Code nicht zugeordnet — z.B. QR eines anderen Events oder eines Abgemeldeten. Die Karte nennt dann den Grund. Der Live-Counter oben rechts („7 eingecheckt") zählt pro Scanner-Session lokal mit, damit du auf einen Blick siehst wie viele du schon durchgecheckt hast.'
                  : 'Point the camera at the QR code — recognition is instant, no shutter button. Depending on the attendee\'s state there are three possible outcomes, each shown as a colored result card at the top of the screen: GREEN = freshly checked in (status „Checked-in", CheckedIn counter +1, name + photo shown as confirmation). ORANGE = was already checked in (no double counter bump, just „Already checked in at …"). RED = code not recognized — e.g. QR from a different event or a cancelled registration. The card names the reason. The live counter top-right („7 checked in") tracks per scanner session locally so you see at a glance how many you\'ve processed.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Check-In-Scanner-Ansicht (echte Ansicht)' : 'Check-in scanner view (real view)'}
                role="Organizer"
                page="check-in"
                selectedEventId={DEMO_EVENT_ID}
                width={390}
              >
                <Header />
                <CheckInPage />
              </AppPreview>
            ),
            tip: isDe
              ? 'Neben dem lokalen Scanner-Counter aktualisieren sich im Admin Center in Echtzeit auch die KPI-Kacheln („Eingecheckt" wird grün). Auf einem Tablet am Stand kannst du parallel die KPIs zeigen, während auf dem Handy gescannt wird.'
              : 'Besides the local scanner counter, the admin-center KPI tiles update in real time as well („Checked-in" turns green). On a tablet at the booth you can show the KPIs while scanning on the phone.',
          },
          {
            number: 9,
            title: isDe ? 'Weg B Fallback: Foto-Upload statt Live-Kamera' : 'Path B fallback: photo upload instead of live camera',
            description: (
              <>
                {isDe
                  ? 'Wenn die Kamera partout nicht startet (eingebetteter Browser, restriktives Firmen-MDM), kannst du innerhalb von Weg B auf „Foto hochladen" wechseln: Foto vom QR-Code mit der Standard-Kamera-App machen, in der Upload-Box auswählen — die App liest den QR-Code aus dem Bild aus und verarbeitet das Scan-Ergebnis genau wie beim Live-Scan. Ergebnis-Karte grün/orange/rot identisch zu Schritt 8. Falls auch das nicht klappt, ist Weg A (Namens-Suche, Schritt 2) der zuverlässige Plan C.'
                  : 'If the camera absolutely refuses to start (embedded browser, restrictive corporate MDM), you can switch to „Upload photo" within Path B: take a photo of the QR code with the default camera app, pick it in the upload box — the app decodes the QR code from the image and processes the result like a live scan. Result card green/orange/red identical to step 8. If that also fails, Path A (search-by-name, step 2) is the reliable plan C.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Scanner-Maske mit Foto-Upload (echte Ansicht)' : 'Scanner page with photo upload (real view)'}
                role="Organizer"
                page="check-in"
                selectedEventId={DEMO_EVENT_ID}
                width={430}
              >
                <Header />
                <CheckInPage />
              </AppPreview>
            ),
          },
          {
            number: 10,
            title: isDe ? 'Allgemein: Check-In rückgängig machen' : 'General: undo check-in',
            description: (
              <>
                {isDe
                  ? 'Wurde jemand versehentlich eingecheckt (falscher QR-Code, Verwechslung mit Namensvetter), kannst du das rückgängig machen: in der Admin-Center-Teilnehmerliste wechselt der „Einchecken"-Button bei bereits eingecheckten Zeilen auf „Auschecken". Ein Klick setzt den Status zurück auf „Angemeldet" (bzw. „QR versendet", je nachdem was vorher war), der Eingecheckt-Counter zählt um 1 runter. Keine Mail, kein Audit-Trail-Eintrag — ist eine reine Korrektur.'
                  : 'If someone was checked in by mistake (wrong QR code, name mixup), you can undo it: in the admin center participant list, the „Check-in" button switches to „Check-out" on already-checked-in rows. A click resets status to „Registered" (or „QR sent" depending on the previous state), the checked-in counter decrements by 1. No email, no audit trail — this is a plain correction.'}
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
            number: 11,
            title: isDe ? 'Allgemein: Mehrere Scanner-Geräte parallel' : 'General: multiple scanner devices in parallel',
            description: (
              <>
                {isDe
                  ? 'Bei großen Events reicht ein einzelnes Handy oft nicht. Du kannst beliebig viele Geräte parallel als Scanner betreiben — jeder Organizer / Admin / QR-Scanner öffnet unabhängig seine eigene Scanner-Session. Da alle gegen dieselbe SharePoint-Teilnehmerliste arbeiten, sehen alle Geräte den aktuellen Zustand. Ein Doppel-Scan (z.B. zwei Geräte scannen denselben Teilnehmer gleichzeitig) wird erkannt — das zweite Gerät meldet orange „Bereits eingecheckt", der Counter springt nicht doppelt. Mischbetrieb Weg A + Weg B ist problemlos möglich: ein Helfer am Tablet macht Namens-Suche im Admin Center, gleichzeitig scannen zwei weitere mit dem Handy — alle landen in derselben Teilnehmerliste, ohne Konflikte.'
                  : 'For large events a single phone often isn\'t enough. You can run as many devices in parallel as scanners as you like — each organizer / admin / QR scanner opens their own independent scanner session. Since they all work against the same SharePoint participant list, all devices see the current state. A double scan (e.g. two devices scan the same attendee at the same time) is detected — the second device reports orange „Already checked in", the counter does not double-count. Mixed mode Path A + Path B is fine: one helper at a tablet does search-by-name in the admin center, two others scan with a phone — all land in the same participant list without conflicts.'}
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
