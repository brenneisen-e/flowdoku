/**
 * Language Context - Mehrsprachigkeit (DE/EN)
 *
 * Speichert die gewaehlte Sprache im localStorage.
 * Stellt eine t()-Funktion bereit die Uebersetzungen liefert.
 */

import * as React from 'react';

export type Locale = 'de' | 'en';

interface LanguageContextType {
  locale: Locale;
  setLocale: (locale: Locale) => void;
  t: (key: string) => string;
}

const LanguageContext = React.createContext<LanguageContextType | undefined>(undefined);

// ==================== Übersetzungen ====================

const translations: Record<Locale, Record<string, string>> = {
  de: {
    // Landing Page
    'landing.welcome': 'Willkommen auf der neuen',
    'landing.platform': 'Event Experience Platform',
    'landing.subtitle': 'Nutze die neue App für deine Registrierung bei Deloitte Events.',
    'landing.start': 'Start',
    'landing.about': 'Die Event Experience Platform ist eine Lösung zur Verwaltung von Teilnehmern bei Deloitte-Events. Entwickelt von Eike Brenneisen, Andreas Enk und Nils Felten. Aktuell in der Pilotphase.',

    // Start Page
    'start.title': 'Event Experience Platform',
    'start.register': 'Registrierung',
    'start.register.desc': 'Melde dich für ein Event an',
    'start.myevents': 'Meine Events',
    'start.myevents.desc': 'Deine Anmeldungen verwalten',
    'start.admin': 'Admin',
    'start.admin.desc': 'Events verwalten',

    // Header
    'header.registration': 'Registrierung | Verfügbare Events an deinem Standort',
    'header.myevents': 'Meine Events',
    'header.createevent': 'Deloitte Event erstellen',
    'header.editevent': 'Event bearbeiten',
    'header.settings': 'Einstellungen',
    'header.profile': 'Mein Profil',
    'header.admin': 'Admin',
    'header.rolematrix': 'Rollen-Matrix',
    'header.participants': 'Teilnehmer-Übersicht',
    'header.flowcharts': 'Prozess-Übersicht',
    'header.checkin': 'Check-in',

    // Event Card
    'event.freeplaces': 'freie Plätze',
    'event.full': 'Ausgebucht',
    'event.waitlist': 'Warteliste',

    // Registration Page
    'reg.selectedevent': 'Ausgewähltes Event',
    'reg.personalinfo': 'Persönliche Informationen',
    'reg.eventinfo': 'Event-spezifische Informationen',
    'reg.salutation': 'Anrede',
    'reg.firstname': 'Vorname',
    'reg.surname': 'Nachname',
    'reg.email': 'E-Mail',
    'reg.pleaseselect': 'Bitte wählen',
    'reg.required': 'Bitte ausfüllen',
    'reg.register': 'Registrieren',
    'reg.delete': 'Löschen',
    'reg.success': 'Registrierung erfolgreich!',
    'reg.error': 'Registrierung fehlgeschlagen. Möglicherweise bist du bereits angemeldet.',
    'reg.genericerror': 'Ein Fehler ist aufgetreten. Bitte versuche es erneut.',
    'reg.registerother': 'Für andere Person registrieren',
    'reg.registerself': 'Für mich selbst registrieren',
    'reg.requiredfield': 'Pflichtfeld',
    'reg.eventnotfound': 'Event nicht gefunden',
    'reg.backtoevents': 'Zurück zu Events',
    'reg.submitting': 'Wird gesendet...',
    'reg.requiredfields': 'Bitte alle Pflichtfelder ausfüllen.',
    'reg.requiredcustom': 'Bitte ausfüllen',
    'reg.waitlisttitle': 'Auf die Warteliste gesetzt',
    'reg.waitlistmsg': 'Du wurdest auf die Warteliste für "{title}" gesetzt. Du wirst benachrichtigt, sobald ein Platz frei wird.',
    'reg.successmsg': 'Du wurdest erfolgreich für "{title}" registriert. Eine Bestätigung wurde an {email} gesendet.',
    'reg.registeranother': 'Für ein weiteres Event registrieren',
    'reg.noadditional': 'Keine zusätzlichen Informationen erforderlich.',
    'reg.allplaces': 'Alle Plätze sind belegt. Aktuell stehen {count} Personen auf der Warteliste',
    'reg.locationnotice': 'Hinweis: Du würdest dieses Event als normaler User nicht sehen.',
    'reg.locationfilter': 'Standort-Filter',
    'reg.audience': 'Zielgruppe',
    'reg.andmode': 'UND-Verknüpfung',
    'reg.yourlocation': 'Dein Standort',
    'reg.unknown': 'unbekannt',
    'reg.until': 'bis',
    'reg.privacy': 'Im Rahmen des Events {title} werden von Deloitte Audio-, Bild- und Videoaufnahmen erstellt. Durch die eindeutige bestätigende Handlung, z.B. Posieren oder Lächeln in die Kamera, willigst du in die Erstellung von Audio-, Bild- und Videoaufnahmen ein. Die Aufnahmen können für Deloitte-interne Veröffentlichungszwecke (z.B. DeloitteNet) verwendet werden. Die Einwilligung kann jederzeit mit Wirkung für die Zukunft und ohne negative Konsequenzen per E-Mail an privacy@deloitte.de widerrufen werden.',

    // My Events
    'myevents.title': 'Meine Events',
    'myevents.loading': 'Lade deine Registrierungen...',
    'myevents.empty': 'Du bist noch für kein Event registriert.',
    'myevents.browse': 'Events durchsuchen',
    'myevents.cancel': 'Abmelden',
    'myevents.confirming': 'Wird abgemeldet...',
    'myevents.confirmcancel': 'Abmeldung bestätigen',
    'myevents.keepreg': 'Anmeldung behalten',
    'myevents.edit': 'Angaben bearbeiten',
    'myevents.save': 'Speichern',
    'myevents.saving': 'Wird gespeichert...',
    'myevents.cancelledevents': 'Abgemeldete Events',
    'myevents.cancelledon': 'Abgemeldet am',
    'myevents.location': 'Ort',
    'myevents.date': 'Datum',
    'myevents.registeredon': 'Angemeldet am',

    // Admin Page
    'admin.title': 'Admin / Organizer',
    'admin.splist': 'SharePoint-Liste',
    'admin.newevent': 'Neues Event erstellen',
    'admin.participants': 'Teilnehmer',
    'admin.processes': 'Prozesse',
    'admin.noevents': 'Keine Events vorhanden.',
    'admin.eventdetails': 'Event-Details',
    'admin.editbutton': 'Event bearbeiten',
    'admin.opensp': 'Teilnehmerliste in SharePoint öffnen',
    'admin.copyemails': 'E-Mail-Adressen kopieren',
    'admin.copied': 'Kopiert!',
    'admin.emailall': 'E-Mail an alle Teilnehmer',
    'admin.deleteevent': 'Event löschen',
    'admin.confirmdelete': 'Wirklich löschen?',
    'admin.deleting': 'Wird gelöscht...',
    'admin.checkin': 'QR-Code Check-in starten',
    'admin.sendqr': 'QR-Codes versenden',
    'admin.sendingqr': 'QR-Codes werden versendet...',
    'admin.loadingregs': 'Lade Teilnehmer...',
    'admin.noregs': 'Noch keine Teilnehmer registriert.',

    // Status
    'status.registered': 'Angemeldet',
    'status.qrsent': 'QR versendet',
    'status.waitlist': 'Warteliste',
    'status.checkedin': 'Eingecheckt',
    'status.cancelled': 'Abgemeldet',

    // Event Creation
    'create.title': 'Event erstellen',
    'create.eventtitle': 'Event-Titel',
    'create.eventtype': 'Event Typ',
    'create.description': 'Beschreibung',
    'create.location': 'Ort / Venue',
    'create.organizer': 'Organizer (Name)',
    'create.startdate': 'Start (Datum & Uhrzeit)',
    'create.enddate': 'Ende (Datum & Uhrzeit)',
    'create.deadline': 'Anmelde-Deadline',
    'create.lastcancel': 'Letzte Abmeldemöglichkeit',
    'create.maxparticipants': 'Max. Teilnehmer',
    'create.waitlist': 'Warteliste',
    'create.enabled': 'Aktiviert',
    'create.disabled': 'Deaktiviert',
    'create.outlookbody': 'Outlook-Termin Text',
    'create.customfields': 'Zusätzliche Registrierungsfelder',
    'create.addfield': 'Feld hinzufügen',
    'create.fieldname': 'Feldname',
    'create.addoption': '+ Option hinzufügen',
    'create.required': 'Pflicht',
    'create.cancel': 'Abbrechen',
    'create.preview': 'Vorschau anzeigen',
    'create.submit': 'Event erstellen',
    'create.save': 'Änderungen speichern',

    // Check-in
    'checkin.title': 'Check-in',
    'checkin.scan': 'QR-Code scannen',
    'checkin.manual': 'Oder Code manuell eingeben:',
    'checkin.check': 'Prüfen',
    'checkin.scanning': 'Scanne QR-Code...',
    'checkin.cancelscanning': 'Abbrechen',
    'checkin.invalidqr': 'Ungültiger QR-Code.',
    'checkin.notfound': 'nicht gefunden.',
    'checkin.notregistered': 'ist nicht für dieses Event registriert.',
    'checkin.checkinbtn': 'Einchecken',
    'checkin.checkingIn': 'Wird eingecheckt...',
    'checkin.success': 'Erfolgreich eingecheckt!',
    'checkin.alreadycheckedin': 'Bereits eingecheckt',
    'checkin.cancelled': 'Teilnehmer ist abgemeldet — Check-in nicht möglich',
    'checkin.next': 'Nächsten Teilnehmer scannen',
    'checkin.sessioncount': 'eingecheckt in dieser Session',

    // Profile
    'profile.details': 'Profil Details',
    'profile.firstname': 'Vorname',
    'profile.lastname': 'Nachname',
    'profile.jobtitle': 'Titel / Position',
    'profile.department': 'Abteilung',
    'profile.office': 'Standort',
    'profile.email': 'E-Mail',
    'profile.phone': 'Telefon',
    'profile.mobile': 'Mobiltelefon',
    'profile.manager': 'Manager',
    'profile.note': 'Diese Daten werden aus deinem SharePoint-Profil gelesen. Um sie zu aktualisieren, wende dich an deine IT oder aktualisiere dein Profil in Microsoft 365.',
    'profile.viewfull': 'Profil anzeigen',

    // Settings
    'settings.title': 'Einstellungen',
    'settings.rolemanagement': 'Rollenverwaltung',

    // Participants
    'participants.title': 'Teilnehmer-Übersicht',
    'participants.opensp': 'In SharePoint öffnen',
    'participants.search': 'Teilnehmer suchen (Name oder E-Mail)...',
    'participants.loading': 'Lade Teilnehmer...',
    'participants.empty': 'Noch keine Teilnehmer registriert.',
    'participants.noresults': 'Keine Teilnehmer gefunden.',
    'participants.registered': 'Registriert',
    'participants.waitlist': 'Warteliste',

    // General
    'general.event': 'Event',
    'general.back': 'Zurück',
    'general.cancel': 'Abbrechen',
    'general.save': 'Speichern',
    'general.delete': 'Löschen',
    'general.close': 'Schließen',
    'general.of': 'von',
    'general.filtered': 'gefiltert',
  },

  en: {
    // Landing Page
    'landing.welcome': 'Welcome to the new',
    'landing.platform': 'Event Experience Platform',
    'landing.subtitle': 'Enjoy the new app to handle your registration for your Deloitte Event.',
    'landing.start': 'Start',
    'landing.about': 'The Event Experience Platform is a solution for managing participants at Deloitte events. Developed by Eike Brenneisen, Andreas Enk and Nils Felten. Currently in pilot phase.',

    // Start Page
    'start.title': 'Event Experience Platform',
    'start.register': 'Registration',
    'start.register.desc': 'Register for an event',
    'start.myevents': 'My Events',
    'start.myevents.desc': 'Manage your registrations',
    'start.admin': 'Admin',
    'start.admin.desc': 'Manage events',

    // Header
    'header.registration': 'Registration | Available events at your location',
    'header.myevents': 'My Events',
    'header.createevent': 'Deloitte Event Creation',
    'header.editevent': 'Edit Event',
    'header.settings': 'Settings',
    'header.profile': 'My Profile',
    'header.admin': 'Admin',
    'header.rolematrix': 'Role Matrix',
    'header.participants': 'Participants Overview',
    'header.flowcharts': 'Process Overview',
    'header.checkin': 'Check-in',

    // Event Card
    'event.freeplaces': 'free places',
    'event.full': 'Fully booked',
    'event.waitlist': 'Waitlist',

    // Registration Page
    'reg.selectedevent': 'Selected Event',
    'reg.personalinfo': 'Personal Information',
    'reg.eventinfo': 'Event specific Information',
    'reg.salutation': 'Salutation',
    'reg.firstname': 'First Name',
    'reg.surname': 'Surname',
    'reg.email': 'E-Mail',
    'reg.pleaseselect': 'Please select',
    'reg.required': 'Please fill in',
    'reg.register': 'Register',
    'reg.delete': 'Delete',
    'reg.success': 'Registration successful!',
    'reg.error': 'Registration failed. You may already be registered.',
    'reg.genericerror': 'An error occurred. Please try again.',
    'reg.registerother': 'Register for another person',
    'reg.registerself': 'Register for myself',
    'reg.requiredfield': 'Required field',
    'reg.eventnotfound': 'Event not found',
    'reg.backtoevents': 'Back to Events',
    'reg.submitting': 'Submitting...',
    'reg.requiredfields': 'Please fill in all required fields.',
    'reg.requiredcustom': 'Please fill in',
    'reg.waitlisttitle': 'Added to waitlist',
    'reg.waitlistmsg': 'You have been added to the waitlist for "{title}". You will be notified when a spot becomes available.',
    'reg.successmsg': 'You have been successfully registered for "{title}". A confirmation has been sent to {email}.',
    'reg.registeranother': 'Register for another event',
    'reg.noadditional': 'No additional information required.',
    'reg.allplaces': 'All places are taken. There are currently {count} people on the waiting list',
    'reg.locationnotice': 'Note: You would not see this event as a regular user.',
    'reg.locationfilter': 'Location filter',
    'reg.audience': 'Audience',
    'reg.andmode': 'AND combination',
    'reg.yourlocation': 'Your location',
    'reg.unknown': 'unknown',
    'reg.until': 'until',
    'reg.privacy': 'As part of the event {title}, audio, image, and video recordings will be made for Deloitte. By the clear confirming action, e.g., posing or smiling at the camera, you consent to the creation of audio, image, and video recordings. The recordings may be used for Deloitte internal publication purposes (e.g., DeloitteNet). The consent can be revoked at any time with future effect and without any negative consequences by sending an email to privacy@deloitte.de.',

    // My Events
    'myevents.title': 'My Events',
    'myevents.loading': 'Loading your registrations...',
    'myevents.empty': 'You are not registered for any event yet.',
    'myevents.browse': 'Browse events',
    'myevents.cancel': 'Cancel registration',
    'myevents.confirming': 'Cancelling...',
    'myevents.confirmcancel': 'Confirm cancellation',
    'myevents.keepreg': 'Keep registration',
    'myevents.edit': 'Edit details',
    'myevents.save': 'Save',
    'myevents.saving': 'Saving...',
    'myevents.cancelledevents': 'Cancelled Events',
    'myevents.cancelledon': 'Cancelled on',
    'myevents.location': 'Location',
    'myevents.date': 'Date',
    'myevents.registeredon': 'Registered on',

    // Admin Page
    'admin.title': 'Admin / Organizer',
    'admin.splist': 'SharePoint List',
    'admin.newevent': 'Create new event',
    'admin.participants': 'Participants',
    'admin.processes': 'Processes',
    'admin.noevents': 'No events found.',
    'admin.eventdetails': 'Event Details',
    'admin.editbutton': 'Edit event',
    'admin.opensp': 'Open participant list in SharePoint',
    'admin.copyemails': 'Copy email addresses',
    'admin.copied': 'Copied!',
    'admin.emailall': 'Email all participants',
    'admin.deleteevent': 'Delete event',
    'admin.confirmdelete': 'Are you sure?',
    'admin.deleting': 'Deleting...',
    'admin.checkin': 'Start QR code check-in',
    'admin.sendqr': 'Send QR codes',
    'admin.sendingqr': 'Sending QR codes...',
    'admin.loadingregs': 'Loading participants...',
    'admin.noregs': 'No participants registered yet.',

    // Status
    'status.registered': 'Registered',
    'status.qrsent': 'QR sent',
    'status.waitlist': 'Waitlist',
    'status.checkedin': 'Checked in',
    'status.cancelled': 'Cancelled',

    // Event Creation
    'create.title': 'Create Event',
    'create.eventtitle': 'Event Title',
    'create.eventtype': 'Event Type',
    'create.description': 'Description',
    'create.location': 'Location / Venue',
    'create.organizer': 'Organizer (Name)',
    'create.startdate': 'Start (Date & Time)',
    'create.enddate': 'End (Date & Time)',
    'create.deadline': 'Registration Deadline',
    'create.lastcancel': 'Last Cancellation Date',
    'create.maxparticipants': 'Max. Participants',
    'create.waitlist': 'Waitlist',
    'create.enabled': 'Enabled',
    'create.disabled': 'Disabled',
    'create.outlookbody': 'Outlook Calendar Text',
    'create.customfields': 'Additional Registration Fields',
    'create.addfield': 'Add field',
    'create.fieldname': 'Field name',
    'create.addoption': '+ Add option',
    'create.required': 'Required',
    'create.cancel': 'Cancel',
    'create.preview': 'Show preview',
    'create.submit': 'Create event',
    'create.save': 'Save changes',

    // Check-in
    'checkin.title': 'Check-in',
    'checkin.scan': 'Scan QR code',
    'checkin.manual': 'Or enter code manually:',
    'checkin.check': 'Check',
    'checkin.scanning': 'Scanning QR code...',
    'checkin.cancelscanning': 'Cancel',
    'checkin.invalidqr': 'Invalid QR code.',
    'checkin.notfound': 'not found.',
    'checkin.notregistered': 'is not registered for this event.',
    'checkin.checkinbtn': 'Check in',
    'checkin.checkingIn': 'Checking in...',
    'checkin.success': 'Successfully checked in!',
    'checkin.alreadycheckedin': 'Already checked in',
    'checkin.cancelled': 'Participant is cancelled — check-in not possible',
    'checkin.next': 'Scan next participant',
    'checkin.sessioncount': 'checked in this session',

    // Profile
    'profile.details': 'Profile Details',
    'profile.firstname': 'First Name',
    'profile.lastname': 'Last Name',
    'profile.jobtitle': 'Job Title',
    'profile.department': 'Department',
    'profile.office': 'Office',
    'profile.email': 'E-Mail',
    'profile.phone': 'Phone',
    'profile.mobile': 'Mobile',
    'profile.manager': 'Manager',
    'profile.note': 'This data is read from your SharePoint profile. To update it, contact your IT or update your profile in Microsoft 365.',
    'profile.viewfull': 'View profile',

    // Settings
    'settings.title': 'Settings',
    'settings.rolemanagement': 'Role Management',

    // Participants
    'participants.title': 'Participants Overview',
    'participants.opensp': 'Open in SharePoint',
    'participants.search': 'Search participants (name or email)...',
    'participants.loading': 'Loading participants...',
    'participants.empty': 'No participants registered yet.',
    'participants.noresults': 'No participants found.',
    'participants.registered': 'Registered',
    'participants.waitlist': 'Waitlist',

    // General
    'general.event': 'Event',
    'general.back': 'Back',
    'general.cancel': 'Cancel',
    'general.save': 'Save',
    'general.delete': 'Delete',
    'general.close': 'Close',
    'general.of': 'of',
    'general.filtered': 'filtered',
  },
};

// ==================== Provider ====================

const STORAGE_KEY = 'dex-locale';

export function LanguageProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [locale, setLocaleState] = React.useState<Locale>(() => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored === 'de' || stored === 'en') return stored;
    } catch { /* */ }
    return 'en'; // Default: Englisch
  });

  const setLocale = (newLocale: Locale): void => {
    setLocaleState(newLocale);
    try {
      localStorage.setItem(STORAGE_KEY, newLocale);
    } catch { /* */ }
  };

  const t = (key: string): string => {
    return translations[locale][key] || translations['en'][key] || key;
  };

  return React.createElement(
    LanguageContext.Provider,
    { value: { locale, setLocale, t } },
    props.children
  );
}

export function useLanguage(): LanguageContextType {
  const ctx = React.useContext(LanguageContext);
  if (!ctx) throw new Error('useLanguage must be used within LanguageProvider');
  return ctx;
}
