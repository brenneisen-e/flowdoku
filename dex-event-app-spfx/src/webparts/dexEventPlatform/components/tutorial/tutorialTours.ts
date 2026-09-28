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
  labelDe: 'Organizer-Rundgang',
  labelEn: 'Organizer tour',
  // v30.67: „alle Schritte" statt „alle 9" — Admins haben zehn.
  descDe: 'Rundgang: Organizer Center, Teilnehmer verwalten und der Check-in am Eventtag.',
  descEn: 'Tour: Organizer Center, managing attendees and check-in on the event day.',
  steps: [
    {
      page: 'start',
      titleDe: 'Willkommen, Organizer!',
      titleEn: 'Welcome, organizer!',
      bodyDe: 'Diese Tour zeigt dir, wo du deine Events verwaltest und wie du den Eventtag meisterst. Für die Dauer der Tour liegt ein Demo-Event in deiner Organizer-Liste — nur für dich sichtbar, zum gefahrlosen Anschauen.',
      bodyEn: 'This tour shows you where you manage your events and how to run the event day. For the duration of the tour a demo event sits in your organizer list — visible only to you, safe to explore.',
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
    // v32.1.0: Der Durchgang durch alle Assistenten-Schritte (elf Karten,
    // nur Blättern) ist ersetzt — Nutzer-Entscheidung 28.09.2026: EIN Weg
    // statt zwei. Das Anlegen lernt man im Mitmach-Tutorial (WizardCoach),
    // das ein echtes Test-Event mit der Person zusammen anlegt.
    {
      page: 'admin',
      selector: '.page-container button.btn-primary',
      titleDe: 'Neues Event — am besten einmal zum Üben',
      titleEn: 'A new event — best practised once',
      bodyDe: 'Über diesen Button startest du den Event-Assistenten. Für dein erstes Event gibt es ein eigenes Mitmach-Tutorial: Oben im Assistenten steht „Tutorial: Test-Event gemeinsam anlegen" — dort tippst und klickst du selbst, und am Ende steht ein echtes Test-Event, das nur du siehst und mit einem Klick wieder löschen kannst.',
      bodyEn: 'This button starts the event wizard. For your first event there is a hands-on tutorial: at the top of the wizard you find “Tutorial: create a test event together” — you type and click yourself, and at the end there is a real test event that only you can see and delete with one click.',
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
