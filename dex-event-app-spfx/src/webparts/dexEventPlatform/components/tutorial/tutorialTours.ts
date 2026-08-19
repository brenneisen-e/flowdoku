/**
 * v22.21: Geführtes Tutorial — Tour-Definitionen.
 *
 * Jede Tour ist eine Liste von Schritten. Ein Schritt navigiert (falls nötig)
 * auf seine Seite und hebt optional ein Element per CSS-Selektor hervor
 * (Spotlight). Ohne Selektor — oder wenn das Element nicht gefunden wird —
 * erscheint die Schritt-Karte zentriert über einem abgedunkelten Backdrop.
 *
 * Sichtbarkeit: Die User-Tour gibt es für alle; die Organizer-Tour zusätzlich
 * für alle, die Events anlegen können oder Organizer/Co-Organizer mindestens
 * eines Events sind (gleiche Logik wie die Organizer-Kachel der Startseite).
 */

import { Page } from '../../context/NavigationContext';

export type TutorialTourId = 'user' | 'organizer';

export interface TutorialStep {
  /** Seite, auf der der Schritt spielt — wird vor dem Anzeigen angesteuert. */
  page: Page;
  /** CSS-Selektor des hervorzuhebenden Elements (Spotlight). Leer = zentriert. */
  selector?: string;
  /** v22.23: 0-basierter Wizard-Schritt — auf der create-event-Seite stellt
   *  die Tour den Wizard per CustomEvent auf diesen Schritt, bevor der
   *  Spotlight gesucht wird (Klick-Through durch alle Wizard-Schritte). */
  wizardStep?: number;
  titleDe: string;
  titleEn: string;
  bodyDe: string;
  bodyEn: string;
}

export interface TutorialTour {
  id: TutorialTourId;
  labelDe: string;
  labelEn: string;
  descDe: string;
  descEn: string;
  steps: TutorialStep[];
}

const USER_TOUR: TutorialTour = {
  id: 'user',
  labelDe: 'Teilnehmer-Tutorial',
  labelEn: 'Attendee tutorial',
  descDe: 'Events finden, anmelden, Meine Events und der QR-Code für den Check-in — in rund zwei Minuten.',
  descEn: 'Find events, register, My Events and your check-in QR code — in about two minutes.',
  steps: [
    {
      page: 'landing',
      titleDe: 'Willkommen bei DEX!',
      titleEn: 'Welcome to DEX!',
      bodyDe: 'Schön, dass du da bist! In rund zwei Minuten zeigen wir dir, wie du Events findest, dich anmeldest und am Eventtag entspannt eincheckst. Wir führen dich Schritt für Schritt durch die App — du klickst nur auf „Weiter".',
      bodyEn: 'Great to have you here! In about two minutes we will show you how to find events, register and check in smoothly on the event day. We guide you step by step — just click "Next".',
    },
    {
      page: 'landing',
      selector: '[data-tour="landing-start"]',
      titleDe: 'Hier geht es los',
      titleEn: 'This is where it starts',
      bodyDe: 'Mit diesem Button startest du in die App. Dahinter wartet die Startseite mit Kacheln für alles, was du brauchst — vom Event-Stöbern bis zu deinen eigenen Anmeldungen.',
      bodyEn: 'This button takes you into the app. Behind it is the start page with tiles for everything you need — from browsing events to your own registrations.',
    },
    {
      page: 'start',
      selector: '[data-tour="tile-register"]',
      titleDe: 'Events entdecken',
      titleEn: 'Discover events',
      bodyDe: 'Hinter dieser Kachel findest du alle Events, die für dich sichtbar sind — passend zu deinem Standort und deinen Verteilern. Stöbern lohnt sich!',
      bodyEn: 'Behind this tile you find all events that are visible to you — matching your location and distribution lists. Browsing pays off!',
    },
    {
      page: 'register',
      selector: '.event-grid',
      titleDe: 'Die Event-Liste',
      titleEn: 'The event list',
      bodyDe: 'Das ist deine Event-Übersicht. Ein Klick auf ein Event öffnet die Detailseite mit allen Infos — Datum, Ort, Agenda — und dem Anmeldeformular.',
      bodyEn: 'This is your event overview. Clicking an event opens the detail page with all the info — date, location, agenda — and the registration form.',
    },
    {
      page: 'register',
      selector: '.event-grid > *:first-child',
      titleDe: 'Anmelden in einer Minute',
      titleEn: 'Register in a minute',
      bodyDe: 'Ein Klick auf eine Event-Karte wie diese öffnet die Anmeldung. Deine persönlichen Daten sind schon vorausgefüllt — du beantwortest nur die Fragen des Organizers, wählst bei Bedarf Programmpunkte oder eine Gruppe und klickst auf „Anmelden". Bestätigungs-Mail und Outlook-Termin kommen automatisch.',
      bodyEn: 'Clicking an event card like this one opens the registration. Your personal data is already pre-filled — you just answer the organizer’s questions, pick sessions or a group if offered and click "Register". Confirmation email and Outlook invitation arrive automatically.',
    },
    {
      page: 'start',
      selector: '[data-tour="tile-myevents"]',
      titleDe: 'Deine Anmeldungen: Meine Events',
      titleEn: 'Your registrations: My Events',
      bodyDe: 'Hinter dieser Kachel verwaltest du alles, wofür du angemeldet bist. Schauen wir kurz rein!',
      bodyEn: 'Behind this tile you manage everything you are registered for. Let’s take a quick look!',
    },
    {
      page: 'my-events',
      selector: '.my-event-card',
      titleDe: 'Meine Events — deine Zentrale',
      titleEn: 'My Events — your home base',
      bodyDe: 'Hier siehst du alle deine Anmeldungen auf einen Blick: Status, Warteliste, Team. Du kannst deine Angaben nachträglich ändern, dich abmelden (dein Platz geht fair an die Warteliste) und deinen persönlichen QR-Code für den Check-in abrufen.',
      bodyEn: 'Here you see all your registrations at a glance: status, waitlist, team. You can update your answers later, cancel (your seat goes fairly to the waitlist) and open your personal QR code for check-in.',
    },
    {
      page: 'my-events',
      selector: '[data-tour="header-avatar"]',
      titleDe: 'Dein Profil & das Handbuch',
      titleEn: 'Your profile & the manual',
      bodyDe: 'Ein Klick auf dein Foto oben rechts öffnet dein Profil-Menü — dort findest du auch das ausführliche Handbuch mit Schritt-für-Schritt-Anleitungen zu jeder Funktion der App.',
      bodyEn: 'Clicking your photo in the top right opens your profile menu — there you also find the full manual with step-by-step guides for every feature of the app.',
    },
    {
      page: 'my-events',
      titleDe: 'Du bist startklar!',
      titleEn: 'You are ready to go!',
      bodyDe: 'Das war die Tour! Am Eventtag bekommst du deinen QR-Code per Mail — vorzeigen, scannen, drin. Viel Spaß bei deinem nächsten Event!',
      bodyEn: 'That was the tour! On the event day your QR code arrives by email — show it, scan it, you are in. Enjoy your next event!',
    },
  ],
};

const ORGANIZER_TOUR: TutorialTour = {
  id: 'organizer',
  labelDe: 'Organizer-Tutorial',
  labelEn: 'Organizer tutorial',
  descDe: 'Events anlegen mit dem Wizard (alle 9 Schritte), Teilnehmer verwalten im Organizer Center und der Check-in am Eventtag.',
  descEn: 'Create events with the wizard (all 9 steps), manage attendees in the Organizer Center and run check-in on the event day.',
  steps: [
    {
      page: 'start',
      titleDe: 'Willkommen, Organizer!',
      titleEn: 'Welcome, organizer!',
      bodyDe: 'Diese Tour zeigt dir, wie du Events anlegst (mit allen 9 Wizard-Schritten), Teilnehmer verwaltest und den Eventtag meisterst. Für die Dauer der Tour liegt ein Demo-Event in deiner Organizer-Liste — nur für dich sichtbar, zum gefahrlosen Anschauen.',
      bodyEn: 'This tour shows you how to create events (with all 9 wizard steps), manage attendees and run the event day. For the duration of the tour a demo event sits in your organizer list — visible only to you, safe to explore.',
    },
    {
      page: 'start',
      selector: '[data-tour="tile-admin"]',
      titleDe: 'Dein Organizer Center',
      titleEn: 'Your Organizer Center',
      bodyDe: 'Hinter dieser Kachel verwaltest du deine Events: Teilnehmerlisten, Statistiken, Warteliste, Mails und vieles mehr. Schauen wir rein!',
      bodyEn: 'Behind this tile you manage your events: attendee lists, statistics, waitlist, emails and much more. Let’s take a look!',
    },
    {
      page: 'admin',
      selector: '.page-container .card',
      titleDe: 'Deine Event-Liste — mit Demo-Event',
      titleEn: 'Your event list — with a demo event',
      bodyDe: 'Hier liegen alle Events, die du verwaltest. Während der Tour siehst du das Übungs-Event „Demo-Event — alle Funktionen" mit Beispiel-Teilnehmern — klick es nach der Tour gern an und probiere die Teilnehmer-Tabelle, die Statistiken und das Aktionen-Menü (QR-Versand, Massenmail, Excel-Export, Audit-Log) aus.',
      bodyEn: 'Here are all events you manage. During the tour you see the practice event “Demo event — all features” with sample attendees — after the tour, feel free to open it and try the attendee table, the statistics and the actions menu (QR sending, mass mail, Excel export, audit log).',
    },
    {
      page: 'admin',
      selector: '.page-container button.btn-primary',
      titleDe: 'Neues Event in Minuten',
      titleEn: 'A new event in minutes',
      bodyDe: 'Über diesen Button startest du den Event-Wizard. Wir gehen jetzt gemeinsam alle 9 Schritte durch — keine Sorge, es wird dabei nichts gespeichert.',
      bodyEn: 'This button starts the event wizard. We will now walk through all 9 steps together — don’t worry, nothing gets saved.',
    },
    {
      page: 'create-event',
      wizardStep: 0,
      selector: '[data-tour="wizard-step-0"]',
      titleDe: 'Schritt 1 — Grundlagen',
      titleEn: 'Step 1 — Basics',
      bodyDe: 'Titel, Zeitraum, Beschreibung und Event-Bild — und die Sub-Events: Hier legst du Programmpunkte oder Termine an und entscheidest, ob die Anmeldung nur über sie läuft. Auch der Entwurfs-Status wohnt hier — sichtbar wird das Event erst, wenn du es live schaltest.',
      bodyEn: 'Title, dates, description and event image — plus the sub-events: create program items or dates here and decide whether registration runs only through them. The draft status lives here too — the event only becomes visible once you publish it.',
    },
    {
      page: 'create-event',
      wizardStep: 1,
      selector: '[data-tour="wizard-step-1"]',
      titleDe: 'Schritt 2 — Organizer & Team',
      titleEn: 'Step 2 — Organizers & Team',
      bodyDe: 'Wer verantwortet das Event (Organizer) — sie sehen die Teilnehmerliste und stellen das Event ein. Optional: ein externer Ansprechpartner, ein Test-Team (sieht den Entwurf vorab) und das Check-in-Team für den Event-Tag.',
      bodyEn: 'Who runs the event (organizers) — they see the attendee list and configure the event. Optionally: an external contact, a test team (sees the draft in advance) and the check-in team for the event day.',
    },
    {
      page: 'create-event',
      wizardStep: 2,
      selector: '[data-tour="wizard-step-2"]',
      titleDe: 'Schritt 3 — Ort & Programm',
      titleEn: 'Step 3 — Location & programme',
      bodyDe: 'Veranstaltungsort und Adresse (landen automatisch im Outlook-Termin), dazu die Agenda für den Tagesablauf und optionale Transferzeiten (Bus, Bahn, Treffpunkt).',
      bodyEn: 'Venue and address (automatically included in the Outlook invite), plus the agenda for the day and optional transfer times (bus, train, meeting point).',
    },
    {
      page: 'create-event',
      wizardStep: 3,
      selector: '[data-tour="wizard-step-3"]',
      titleDe: 'Schritt 4 — Kapazität & Sichtbarkeit',
      titleEn: 'Step 4 — Capacity & visibility',
      bodyDe: 'Wie viele Plätze gibt es (mit fairer Warteliste und automatischem Nachrücken), bis wann kann man sich an- und abmelden — und wer sieht das Event überhaupt: Standortfilter, Mailverteiler oder einzelne Personen.',
      bodyEn: 'How many seats there are (with a fair waitlist and automatic promotion), the registration and cancellation deadlines — and who can see the event at all: location filter, distribution lists or individual people.',
    },
    {
      page: 'create-event',
      wizardStep: 4,
      selector: '[data-tour="wizard-step-4"]',
      titleDe: 'Schritt 5 — Felder',
      titleEn: 'Step 5 — Fields',
      bodyDe: 'Eigene Abfragen für das Anmeldeformular: Dropdowns, Checkboxen, Textfelder, Personen-Suche oder Dokument-Uploads — mit Pflichtfeld-Option und Sichtbarkeitsbedingungen. Die Antworten siehst du später pro Teilnehmer im Organizer Center.',
      bodyEn: 'Custom questions for the registration form: dropdowns, checkboxes, text fields, people pickers or document uploads — with required-field options and visibility conditions. You see the answers per attendee in the Organizer Center.',
    },
    {
      page: 'create-event',
      wizardStep: 5,
      selector: '[data-tour="wizard-step-5"]',
      titleDe: 'Schritt 6 — Kommunikation',
      titleEn: 'Step 6 — Communication',
      bodyDe: 'Alles Automatische: Bestätigungs- und Abmelde-Mails, der Outlook-Termin, Mail-Sprache und das Design (Texte, Logos, Kopfbild) — bei Bedarf pro Sub-Event unterschiedlich. Eine Übersichtsbox zeigt dir, was aktuell automatisch rausgeht.',
      bodyEn: 'Everything automatic: confirmation and cancellation emails, the Outlook invite, email language and the design (texts, logos, header image) — per sub-event if needed. A summary box shows what currently goes out automatically.',
    },
    {
      page: 'create-event',
      wizardStep: 6,
      selector: '[data-tour="wizard-step-6"]',
      titleDe: 'Schritt 7 — Team-Anmeldung',
      titleEn: 'Step 7 — Team registration',
      bodyDe: 'Für Events mit festen Teams (z.B. Quizabend, Sport): eine Person meldet ihr ganzes Team an. Du legst die Teamgröße fest und entscheidest, ob unvollständige Teams öffentlich beitretbar sind — auf Wunsch mit Bestätigung durch den Team-Kapitän.',
      bodyEn: 'For events with fixed teams (e.g. quiz night, sports): one person registers the whole team. You set the team size and decide whether incomplete teams are open to join — optionally with team captain approval.',
    },
    {
      page: 'create-event',
      wizardStep: 7,
      selector: '[data-tour="wizard-step-7"]',
      titleDe: 'Schritt 8 — Dokumente',
      titleEn: 'Step 8 — Documents',
      bodyDe: 'PDFs für deine Teilnehmer — etwa Anfahrtsbeschreibung, Programmheft oder FAQ. Sie erscheinen direkt auf der Anmeldeseite zum Download.',
      bodyEn: 'PDFs for your attendees — directions, programme booklet or an FAQ. They appear directly on the registration page for download.',
    },
    {
      page: 'create-event',
      wizardStep: 8,
      selector: '[data-tour="wizard-step-8"]',
      titleDe: 'Schritt 9 — Fun-Zone',
      titleEn: 'Step 9 — Fun zone',
      bodyDe: 'Das Sahnehäubchen: ein optionales Quiz rund um dein Event, das Teilnehmer vor Ort oder vorab spielen können — gut für Stimmung und Engagement.',
      bodyEn: 'The cherry on top: an optional quiz around your event that attendees can play on site or beforehand — great for mood and engagement.',
    },
    {
      page: 'create-event',
      wizardStep: 8,
      selector: '[data-tour="wizard-submit"]',
      titleDe: 'Und was passiert beim Anlegen?',
      titleEn: 'And what happens when you create it?',
      bodyDe: 'Mit diesem Button wird dein Event gespeichert: Die App legt automatisch die Teilnehmerliste und die Mail-Vorlagen an — das dauert einen Moment. Danach landet das Event in deinem Organizer Center, zunächst als Entwurf. Live schalten, testweise anmelden, Einladungsmail verschicken: All das machst du dort über die „Nächste Schritte"-Box.',
      bodyEn: 'This button saves your event: the app automatically creates the attendee list and the email templates — that takes a moment. The event then appears in your Organizer Center, initially as a draft. Publishing, test-registering, sending the invitation email: you do all that there via the “Next steps” box.',
    },
    {
      page: 'start',
      selector: '[data-tour="tile-checkin"]',
      titleDe: 'Der Eventtag: Check-In',
      titleEn: 'The event day: check-in',
      bodyDe: 'Hinter dieser Kachel checkst du am Eventtag deine Gäste ein — werfen wir noch einen kurzen Blick darauf.',
      bodyEn: 'Behind this tile you check in your guests on the event day — let’s take one quick look.',
    },
    {
      page: 'check-in',
      selector: '.page-container .card',
      titleDe: 'Einchecken: scannen oder selbst machen lassen',
      titleEn: 'Check-in: scan or let guests do it',
      bodyDe: 'Deine Teilnehmer bekommen ihren QR-Code automatisch per Mail. Hier scannst du ihn mit der Kamera oder checkst manuell per Name ein — oder du lässt die Gäste sich selbst einchecken: per ausgedrucktem QR-Plakat (PDF) oder rotierender Live-Anzeige am Eingang.',
      bodyEn: 'Your attendees receive their QR code automatically by email. Here you scan it with the camera or check people in manually by name — or let guests check themselves in: via a printed QR poster (PDF) or a rotating live display at the entrance.',
    },
    {
      page: 'start',
      titleDe: 'Bereit für dein erstes Event?',
      titleEn: 'Ready for your first event?',
      bodyDe: 'Das war die Organizer-Tour! Das Demo-Event verschwindet jetzt wieder aus deiner Liste — du kannst die Tour aber jederzeit neu starten. Details zu jedem Thema findest du im Handbuch (über dein Profilfoto oben rechts).',
      bodyEn: 'That was the organizer tour! The demo event now disappears from your list again — but you can restart the tour anytime. You find details on every topic in the manual (via your profile photo in the top right).',
    },
  ],
};

export const TUTORIAL_TOURS: Record<TutorialTourId, TutorialTour> = {
  user: USER_TOUR,
  organizer: ORGANIZER_TOUR,
};
