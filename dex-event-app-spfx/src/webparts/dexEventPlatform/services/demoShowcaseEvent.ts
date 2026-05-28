// v17.25: Synthetisches Demo-Showcase-Event.
//
// Wird AUSSCHLIESSLICH im Demo-Impersonation-Modus client-seitig in die
// Event-Liste injiziert (siehe EventContext `eventsForConsumer`). Es
// existiert NICHT in SharePoint — Anmelden/Abmelden sind No-Ops. Zweck:
// einem Admin im Demo-Modus auf der Register-Seite alle Faehigkeiten eines
// Events in ein-/ausklappbaren Bereichen zeigen (Felder aller Typen,
// Agenda, Transferzeiten, geteilte Kapazitaet, Team-Anmeldung, Sub-Events,
// Zweisprachigkeit, Dokumente, Quiz).

import { DeloitteEvent, AgendaItem, TransferTime, EventSpecificField } from '../types';
import { SPRegistration } from './EventService';

export const DEMO_SHOWCASE_ID = '__demo_showcase__';
export const DEMO_SHOWCASE_CHILD_PREFIX = '__demo_showcase_child_';

export function isDemoShowcaseId(id: string | undefined | null): boolean {
  if (!id) return false;
  return id === DEMO_SHOWCASE_ID || id.indexOf(DEMO_SHOWCASE_CHILD_PREFIX) === 0;
}

// Relative Datumshilfen, damit das Demo-Event immer „in der Zukunft" liegt
// und die Deadlines plausibel sind — egal wann der Admin den Demo-Modus
// oeffnet.
function isoInDays(days: number, hour = 9, minute = 0): string {
  const d = new Date();
  d.setDate(d.getDate() + days);
  d.setHours(hour, minute, 0, 0);
  return d.toISOString();
}
function ymdInDays(days: number): string {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

// Alle Custom-Field-Typen einmal, zweisprachig (DE + EN), inkl. Sichtbar-
// keitsbedingung (showIf), Gruppen-Restriktion (onlyForGroup), Multi-Select.
function demoCustomFields(): EventSpecificField[] {
  return [
    {
      id: 'demo_text',
      label: 'Wunschname auf dem Badge',
      labelEn: 'Preferred name on the badge',
      type: 'text',
      required: true,
      helpText: 'Freitextfeld — z.B. fuer den Namen, der auf dem Namensschild erscheinen soll.',
      helpTextEn: 'Free-text field — e.g. the name to print on your badge.',
    },
    {
      id: 'demo_select',
      label: 'Essenspraeferenz',
      labelEn: 'Meal preference',
      type: 'select',
      required: true,
      options: ['Fleisch', 'Vegetarisch', 'Vegan'],
      optionsEn: ['Meat', 'Vegetarian', 'Vegan'],
      helpText: 'Einfachauswahl-Dropdown — genau eine Option waehlbar.',
      helpTextEn: 'Single-select dropdown — exactly one option.',
    },
    {
      id: 'demo_multi',
      label: 'Allergien',
      labelEn: 'Allergies',
      type: 'select',
      multi: true,
      required: false,
      options: ['Gluten', 'Laktose', 'Nuesse', 'Meeresfruechte'],
      optionsEn: ['Gluten', 'Lactose', 'Nuts', 'Seafood'],
      helpText: 'Mehrfachauswahl — beliebig viele Optionen ankreuzbar.',
      helpTextEn: 'Multi-select — tick any number of options.',
    },
    {
      id: 'demo_number',
      label: 'Anzahl Begleitpersonen',
      labelEn: 'Number of accompanying guests',
      type: 'number',
      required: false,
      helpText: 'Zahlenfeld.',
      helpTextEn: 'Numeric field.',
    },
    {
      id: 'demo_checkbox',
      label: 'Teilnahmebedingungen',
      labelEn: 'Terms of participation',
      type: 'checkbox',
      required: true,
      confirmLabel: 'Ich akzeptiere die Teilnahmebedingungen',
      confirmLabelEn: 'I accept the terms of participation',
      helpText: 'Pflicht-Checkbox mit individuellem Bestaetigungstext.',
      helpTextEn: 'Required checkbox with a custom confirmation label.',
    },
    {
      id: 'demo_hotel',
      label: 'Hotel benoetigt?',
      labelEn: 'Hotel needed?',
      type: 'checkbox',
      required: false,
      confirmLabel: 'Ja, ich brauche ein Hotelzimmer',
      confirmLabelEn: 'Yes, I need a hotel room',
      helpText: 'Steuerfeld fuer die folgende Sichtbarkeitsbedingung.',
      helpTextEn: 'Controls the visibility of the next field.',
    },
    {
      id: 'demo_roommate',
      label: 'Wunsch-Zimmerpartner:in',
      labelEn: 'Preferred roommate',
      type: 'roommate',
      required: false,
      helpText: 'Personen-Picker (Roommate) — erscheint nur, wenn „Hotel benoetigt?" angehakt ist. Die gewaehlte Person bekaeme im echten Event eine Info-Mail.',
      helpTextEn: 'Roommate picker — only shown when „Hotel needed?" is ticked. In a real event the selected person would get a notification email.',
      showIf: { fieldId: 'demo_hotel', values: ['true'] },
    },
    {
      id: 'demo_user',
      label: 'Ansprechpartner:in im Team',
      labelEn: 'Contact person in your team',
      type: 'user',
      required: false,
      helpText: 'Generischer Personen-Picker (ohne Mail-Versand).',
      helpTextEn: 'Generic people picker (no email sent).',
    },
    {
      id: 'demo_experience',
      label: 'Lauferfahrung',
      labelEn: 'Running experience',
      type: 'select',
      required: false,
      options: ['Anfaenger:in', 'Fortgeschritten', 'Profi'],
      optionsEn: ['Beginner', 'Intermediate', 'Pro'],
      helpText: 'Einfaches Dropdown.',
      helpTextEn: 'Simple dropdown.',
    },
  ];
}

function demoAgenda(): AgendaItem[] {
  const day = ymdInDays(14);
  return [
    { id: 'demo_ag_1', date: day, time: '09:00', endTime: '09:30', icon: 'Coffee', title: 'Ankunft & Registrierung', description: 'Check-in am Empfang, Badge abholen.' },
    { id: 'demo_ag_2', date: day, time: '09:30', endTime: '10:30', icon: 'Megaphone', title: 'Keynote', description: 'Begruessung und Ausblick.' },
    { id: 'demo_ag_3', date: day, time: '10:30', endTime: '12:00', icon: 'People', title: 'Workshops', description: 'Parallele Sessions in Gruppen.' },
    { id: 'demo_ag_4', date: day, time: '12:00', endTime: '13:00', icon: 'Food', title: 'Mittagessen', description: 'Buffet im Atrium.' },
    { id: 'demo_ag_5', date: day, time: '13:00', endTime: '17:00', icon: 'Running', title: 'Lauf / Aktivprogramm', description: 'B2Run-Strecke oder Spaziergang.' },
  ];
}

function demoTransfers(): TransferTime[] {
  const day = ymdInDays(14);
  return [
    { id: 'demo_tr_1', location: 'Hauptbahnhof', meetingPoint: 'Gleis 1, Treffpunkt-Schild', address: 'Bahnhofsplatz 1', date: day, departureTime: '08:00', arrivalTime: '08:45', description: 'Shuttle-Bus zum Veranstaltungsort.' },
    { id: 'demo_tr_2', location: 'Veranstaltungsort', meetingPoint: 'Haupteingang', address: 'Eventstrasse 5', date: day, departureTime: '18:00', arrivalTime: '18:45', description: 'Rueckfahrt zum Hauptbahnhof.' },
  ];
}

// Ein Sub-Event (Networking-Dinner) mit eigenen Feldern, damit die
// Sub-Event-Auswahl + die Pro-Sub-Event-Rueckfragen demonstriert werden.
function demoSubEvents(parentId: string): DeloitteEvent[] {
  return [
    {
      id: `${DEMO_SHOWCASE_CHILD_PREFIX}dinner`,
      eventNumber: 0,
      title: 'Networking-Dinner (Sub-Event)',
      type: 'Other',
      status: 'Active',
      organizers: ['Demo, Organizer'],
      organizerEmails: ['demo.organizer@deloitte.de'],
      qrScannerNames: [],
      qrScannerEmails: [],
      location: 'Restaurant „Zur Linde"',
      locationAudience: [],
      audienceFilter: [],
      filterMode: 'OR',
      startDate: isoInDays(14, 19, 0),
      endDate: isoInDays(14, 22, 0),
      registrationDeadline: isoInDays(10, 23, 59),
      lastDeregisterDate: isoInDays(12, 23, 59),
      description: 'Optionales Abendessen im Anschluss an das Hauptprogramm.',
      maxParticipants: 40,
      currentParticipants: 12,
      waitlistCount: 0,
      waitlistEnabled: true,
      outlookBody: '',
      emailLanguage: 'DE',
      bilingualFields: true,
      isDemoShowcase: true,
      parentEventId: parentId,
      eventSpecificFields: [
        {
          id: 'demo_dinner_drink',
          label: 'Getraenkewunsch',
          labelEn: 'Drink preference',
          type: 'select',
          required: false,
          options: ['Wein', 'Bier', 'Alkoholfrei'],
          optionsEn: ['Wine', 'Beer', 'Non-alcoholic'],
        },
      ],
      agenda: [],
      transferTimes: [],
      documents: [],
      quiz: [],
    } as DeloitteEvent,
  ];
}

// v17.26 / v18: Deterministische synthetische Teilnehmerliste fuer das
// Demo-Event, damit der Admin/Organizer unter „Organizer & Events" die
// Teilnehmer-Verwaltung mit echten Daten durchspielen kann (~20 Angemeldete,
// ein paar Abgemeldete, zwei Warteliste-Eintraege, ein Team). Rein
// client-seitig — kein SharePoint.
const DEMO_FIRST_NAMES = ['Anna', 'Ben', 'Clara', 'David', 'Eva', 'Felix', 'Greta', 'Hannes', 'Ida', 'Jonas', 'Klara', 'Lukas', 'Mia', 'Noah', 'Olivia', 'Paul', 'Quirin', 'Rosa', 'Sophie', 'Tom', 'Ulrike', 'Viktor', 'Wanda', 'Xaver', 'Yara', 'Zoe'];
const DEMO_LAST_NAMES = ['Müller', 'Schmidt', 'Schneider', 'Fischer', 'Weber', 'Meyer', 'Wagner', 'Becker', 'Schulz', 'Hoffmann', 'Koch', 'Bauer', 'Richter', 'Klein', 'Wolf', 'Schröder', 'Neumann', 'Braun', 'Werner', 'Krüger', 'Hofmann', 'Hartmann', 'Lange', 'Schmitt', 'Krause', 'Meier'];
const DEMO_LOCATIONS = ['Düsseldorf', 'München', 'Berlin', 'Hamburg', 'Frankfurt', 'Stuttgart', 'Köln'];

export function buildDemoRegistrations(): SPRegistration[] {
  const rows: SPRegistration[] = [];
  const baseDate = new Date();
  baseDate.setDate(baseDate.getDate() - 20);
  const mkDate = (offsetDays: number, hour = 10): string => {
    const d = new Date(baseDate);
    d.setDate(d.getDate() + offsetDays);
    d.setHours(hour, 0, 0, 0);
    return d.toISOString();
  };
  let tid = 1;
  // 20 Angemeldete, abwechselnd Gruppe A/B, mit Custom-Data. v18.6: keine
  // Team-Zuordnung mehr (Team-Anmeldung im Demo deaktiviert).
  for (let i = 0; i < 20; i++) {
    const first = DEMO_FIRST_NAMES[i % DEMO_FIRST_NAMES.length];
    const last = DEMO_LAST_NAMES[i % DEMO_LAST_NAMES.length];
    const email = `${first.toLowerCase()}.${last.toLowerCase().replace(/[äöü]/g, m => ({ ä: 'ae', ö: 'oe', ü: 'ue' }[m] || m))}@deloitte.de`;
    const starter = i % 2 === 0 ? 'Durchstarter' : 'Funstarter';
    rows.push({
      Id: 1000 + i,
      Title: email,
      TeilnehmerID: tid++,
      Anrede: i % 3 === 0 ? 'Frau' : i % 3 === 1 ? 'Herr' : 'Divers',
      Vorname: first,
      Nachname: last,
      StarterType: starter,
      ParticipantName: `${first} ${last}`,
      ParticipantEmail: email,
      Status: 'Angemeldet',
      RegistrationDate: mkDate(i),
      CancellationDate: '',
      Location: DEMO_LOCATIONS[i % DEMO_LOCATIONS.length],
      CustomData: JSON.stringify({
        demo_select: i % 3 === 0 ? 'Vegetarisch' : i % 3 === 1 ? 'Fleisch' : 'Vegan',
        demo_checkbox: 'true',
      }),
    });
  }
  // 2 Warteliste-Eintraege.
  for (let i = 0; i < 2; i++) {
    const first = DEMO_FIRST_NAMES[(20 + i) % DEMO_FIRST_NAMES.length];
    const last = DEMO_LAST_NAMES[(20 + i) % DEMO_LAST_NAMES.length];
    const email = `${first.toLowerCase()}.${last.toLowerCase()}@deloitte.de`;
    rows.push({
      Id: 2000 + i,
      Title: email,
      TeilnehmerID: tid++,
      Vorname: first,
      Nachname: last,
      ParticipantName: `${first} ${last}`,
      ParticipantEmail: email,
      Status: 'Warteliste',
      RegistrationDate: mkDate(20 + i),
      CancellationDate: '',
      Location: DEMO_LOCATIONS[i % DEMO_LOCATIONS.length],
      CustomData: '{}',
    });
  }
  // 3 Abgemeldete.
  for (let i = 0; i < 3; i++) {
    const first = DEMO_FIRST_NAMES[(22 + i) % DEMO_FIRST_NAMES.length];
    const last = DEMO_LAST_NAMES[(22 + i) % DEMO_LAST_NAMES.length];
    const email = `${first.toLowerCase()}.${last.toLowerCase()}@deloitte.de`;
    rows.push({
      Id: 3000 + i,
      Title: email,
      Vorname: first,
      Nachname: last,
      ParticipantName: `${first} ${last}`,
      ParticipantEmail: email,
      Status: 'Abgemeldet',
      RegistrationDate: mkDate(i),
      CancellationDate: mkDate(15 + i),
      Location: DEMO_LOCATIONS[i % DEMO_LOCATIONS.length],
      CustomData: '{}',
    });
  }
  return rows;
}

/**
 * Synthetische Eigen-Anmeldung des Demo-Users fuer die „Meine Events"-Ansicht.
 * v18: wird im Demo-Modus IMMER angezeigt (nicht an eine echte Anmeldung
 * gekoppelt — die Register-Anmeldung ist im Demo bewusst deaktiviert). Zeigt,
 * wie eine Anmeldung in „Meine Events" aussieht: Team-Badge, Custom-Field-
 * Antworten, Beschreibung etc.
 */
export function buildDemoMyRegistration(email: string, displayName: string): SPRegistration {
  const name = (displayName || 'Demo User').trim();
  const parts = name.split(' ');
  return {
    Id: 9999,
    Title: email || 'demo.user@deloitte.de',
    TeilnehmerID: 7,
    Anrede: 'Keine Angabe',
    Vorname: parts[0] || 'Demo',
    Nachname: parts.slice(1).join(' ') || 'User',
    StarterType: 'Durchstarter',
    ParticipantName: name,
    ParticipantEmail: email || 'demo.user@deloitte.de',
    Status: 'Angemeldet',
    RegistrationDate: new Date(Date.now() - 3 * 86400000).toISOString(),
    CancellationDate: '',
    Location: 'Düsseldorf',
    CustomData: JSON.stringify({ demo_text: 'Demo User', demo_select: 'Vegetarisch', demo_multi: 'Gluten | Laktose', demo_checkbox: 'true' }),
  };
}

/**
 * Baut das komplette Demo-Showcase-Set: das Hauptevent + Sub-Event(s).
 * Index 0 ist immer das Hauptevent (DEMO_SHOWCASE_ID).
 */
export function buildDemoShowcaseEvents(locale: 'de' | 'en' = 'de'): DeloitteEvent[] {
  const isDe = locale === 'de';
  const parent: DeloitteEvent = {
    id: DEMO_SHOWCASE_ID,
    eventNumber: 0,
    title: isDe ? 'Demo-Event — alle Funktionen' : 'Demo event — all features',
    type: 'Other',
    status: 'Active',
    organizers: ['Demo, Organizer', 'Beispiel, Anna'],
    organizerEmails: ['demo.organizer@deloitte.de', 'anna.beispiel@deloitte.de'],
    qrScannerNames: [],
    qrScannerEmails: [],
    contactName: isDe ? 'Demo-Ansprechpartner' : 'Demo contact',
    contactEmail: 'demo.kontakt@deloitte.de',
    contactInfo: isDe ? 'Erreichbar Mo–Fr 9–17 Uhr' : 'Available Mon–Fri 9am–5pm',
    location: isDe ? 'Deloitte Campus, Düsseldorf' : 'Deloitte Campus, Düsseldorf',
    locationAddress: { street: 'Schwannstrasse', houseNo: '6', zip: '40476', city: 'Düsseldorf' },
    locationAudience: [],
    audienceFilter: [],
    filterMode: 'OR',
    startDate: isoInDays(14, 9, 0),
    endDate: isoInDays(14, 17, 0),
    registrationDeadline: isoInDays(10, 23, 59),
    lastDeregisterDate: isoInDays(12, 23, 59),
    description: isDe
      ? '<p>Dies ist ein <strong>Demo-Event</strong>. Es zeigt alle Funktionen, die ein echtes Event haben kann: eigene Felder, Agenda, Transferzeiten, geteilte Kapazitäten, Team-Anmeldung, Sub-Events, Zweisprachigkeit und mehr.</p><p>Klapp die Bereiche unten auf, um die jeweiligen Funktionen auszuprobieren. <em>Es wird keine echte Anmeldung gespeichert.</em></p>'
      : '<p>This is a <strong>demo event</strong>. It showcases everything a real event can do: custom fields, agenda, transfer times, split capacity, team registration, sub-events, bilingual content and more.</p><p>Expand the sections below to try each capability. <em>No real registration is stored.</em></p>',
    maxParticipants: 50,
    currentParticipants: 25,
    waitlistCount: 2,
    waitlistEnabled: true,
    imageUrl: '',
    outlookBody: '',
    emailLanguage: isDe ? 'DE' : 'EN',
    isDemoShowcase: true,
    bilingualFields: true,
    askSalutation: true,
    allowAttendeeUpload: true,
    attendeeUploadLabel: isDe ? 'Dokument hochladen (Demo)' : 'Upload document (demo)',
    attendeeUploadHint: isDe ? 'Hier könntest du z.B. einen Nachweis hochladen.' : 'You could upload a proof document here.',
    // v18: Standardmaessig EINE Teilnehmergruppe (keine geteilte Kapazitaet).
    // Die „zwei Gruppen"-Funktion wird auf der Register-Seite als
    // einklappbarer Showcase-Bereich illustriert, nicht als Live-Split.
    // v18.6: Team-Anmeldung im Demo-Event deaktiviert (auf Wunsch — keine
    // Teams im Demo).
    teamRegistrationEnabled: false,
    // Sub-Event-Bezeichnung
    childEventTermSingular: isDe ? 'Sub-Event' : 'Sub-event',
    childEventTermPlural: isDe ? 'Sub-Events' : 'Sub-events',
    eventSpecificFields: demoCustomFields(),
    agenda: demoAgenda(),
    transferTimes: demoTransfers(),
    documents: [],
    quiz: [],
  };
  return [parent, ...demoSubEvents(DEMO_SHOWCASE_ID)];
}
