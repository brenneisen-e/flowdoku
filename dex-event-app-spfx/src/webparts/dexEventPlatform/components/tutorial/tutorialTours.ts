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
      titleDe: 'Die Event-Liste',
      titleEn: 'The event list',
      bodyDe: 'Das ist deine Event-Übersicht. Ein Klick auf ein Event öffnet die Detailseite mit allen Infos — Datum, Ort, Agenda — und dem Anmeldeformular.',
      bodyEn: 'This is your event overview. Clicking an event opens the detail page with all the info — date, location, agenda — and the registration form.',
    },
    {
      page: 'register',
      titleDe: 'Anmelden in einer Minute',
      titleEn: 'Register in a minute',
      bodyDe: 'Deine persönlichen Daten sind schon vorausgefüllt. Du beantwortest nur die Fragen des Organizers, wählst bei Bedarf Programmpunkte oder eine Gruppe — und klickst auf „Anmelden". Bestätigungs-Mail und Outlook-Termin kommen automatisch.',
      bodyEn: 'Your personal data is already pre-filled. You just answer the organizer’s questions, pick sessions or a group if offered — and click "Register". Confirmation email and Outlook invitation arrive automatically.',
    },
    {
      page: 'my-events',
      titleDe: 'Meine Events — deine Zentrale',
      titleEn: 'My Events — your home base',
      bodyDe: 'Hier siehst du alle deine Anmeldungen auf einen Blick: Status, Warteliste, Team. Du kannst deine Angaben nachträglich ändern, dich abmelden (dein Platz geht fair an die Warteliste) und deinen persönlichen QR-Code für den Check-in abrufen.',
      bodyEn: 'Here you see all your registrations at a glance: status, waitlist, team. You can update your answers later, cancel (your seat goes fairly to the waitlist) and open your personal QR code for check-in.',
    },
    {
      page: 'my-events',
      selector: '.header-avatar',
      titleDe: 'Dein Profil & das Handbuch',
      titleEn: 'Your profile & the manual',
      bodyDe: 'Über dein Foto oben rechts erreichst du dein Profil — und das ausführliche Handbuch mit Schritt-für-Schritt-Anleitungen zu jeder Funktion der App.',
      bodyEn: 'Your photo in the top right opens your profile — and the full manual with step-by-step guides for every feature of the app.',
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
  descDe: 'Events anlegen mit dem Wizard, Teilnehmer verwalten im Organizer Center und der Check-in am Eventtag.',
  descEn: 'Create events with the wizard, manage attendees in the Organizer Center and run check-in on the event day.',
  steps: [
    {
      page: 'start',
      titleDe: 'Willkommen, Organizer!',
      titleEn: 'Welcome, organizer!',
      bodyDe: 'Diese Tour zeigt dir, wie du Events anlegst, Teilnehmer verwaltest und den Eventtag meisterst — alles in rund zwei Minuten. Los geht es auf der Startseite.',
      bodyEn: 'This tour shows you how to create events, manage attendees and run the event day — all in about two minutes. We start on the start page.',
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
      selector: '.page-container button.btn-primary',
      titleDe: 'Neues Event in Minuten',
      titleEn: 'A new event in minutes',
      bodyDe: 'Über diesen Button startest du den Event-Wizard. Er nimmt dich an die Hand — von den Grundlagen bis zum fertig eingerichteten Event mit eigener Teilnehmerliste.',
      bodyEn: 'This button starts the event wizard. It takes you by the hand — from the basics to a fully set up event with its own attendee list.',
    },
    {
      page: 'create-event',
      selector: '[data-tour="wizard-demo"]',
      titleDe: 'Der Event-Wizard',
      titleEn: 'The event wizard',
      bodyDe: 'Acht Schritte führen dich durch alles: Grundlagen, Ort & Programm, Kapazität & Sichtbarkeit, Team-Anmeldung, eigene Abfrage-Felder, Kommunikation (Mails & Outlook), Dokumente und die Fun-Zone. Tipp: Der „Demo"-Button füllt das Formular mit einer Beispiel-Vorlage — perfekt zum gefahrlosen Ausprobieren.',
      bodyEn: 'Eight steps walk you through everything: basics, location & program, capacity & visibility, team registration, custom fields, communication (emails & Outlook), documents and the fun zone. Tip: the "Demo" button fills the form with a sample template — perfect for safe experimenting.',
    },
    {
      page: 'admin',
      titleDe: 'Teilnehmer im Griff',
      titleEn: 'Attendees under control',
      bodyDe: 'Im Organizer Center wählst du dein Event und siehst alle Anmeldungen live: suchen, sortieren, ein- und auschecken, abmelden, Antworten bearbeiten. Das Aktionen-Menü bündelt alles Weitere — QR-Versand, Massenmails, Excel-Export, Audit-Log.',
      bodyEn: 'In the Organizer Center you pick your event and see all registrations live: search, sort, check in and out, cancel, edit answers. The actions menu bundles everything else — QR sending, mass emails, Excel export, audit log.',
    },
    {
      page: 'check-in',
      titleDe: 'Der Eventtag: Check-in',
      titleEn: 'The event day: check-in',
      bodyDe: 'Deine Teilnehmer bekommen ihren QR-Code per Mail. Hier scannst du ihn mit der Kamera — oder du lässt die Gäste sich selbst einchecken: per ausgedrucktem QR-Plakat (PDF) oder rotierender Live-Anzeige auf dem Bildschirm am Eingang.',
      bodyEn: 'Your attendees receive their QR code by email. Here you scan it with the camera — or let guests check themselves in: via a printed QR poster (PDF) or a rotating live display at the entrance.',
    },
    {
      page: 'start',
      titleDe: 'Bereit für dein erstes Event?',
      titleEn: 'Ready for your first event?',
      bodyDe: 'Das war die Organizer-Tour! Details zu jedem Thema findest du im Handbuch (über dein Profilfoto oben rechts). Und wenn du magst: Leg gleich ein Demo-Event an und spiel alles einmal durch.',
      bodyEn: 'That was the organizer tour! You find details on every topic in the manual (via your profile photo in the top right). And if you like: create a demo event right away and try everything out.',
    },
  ],
};

export const TUTORIAL_TOURS: Record<TutorialTourId, TutorialTour> = {
  user: USER_TOUR,
  organizer: ORGANIZER_TOUR,
};
