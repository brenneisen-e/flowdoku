// v17.25: Synthetisches Demo-Showcase-Event.
//
// Wird AUSSCHLIESSLICH im Demo-Impersonation-Modus client-seitig in die
// Event-Liste injiziert (siehe EventContext `eventsForConsumer`). Es
// existiert NICHT in SharePoint — Anmelden/Abmelden sind No-Ops. Zweck:
// einem Admin im Demo-Modus auf der Register-Seite alle Fähigkeiten eines
// Events in ein-/ausklappbaren Bereichen zeigen (Felder aller Typen,
// Agenda, Transferzeiten, geteilte Kapazität, Team-Anmeldung, Sub-Events,
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
      helpText: 'Freitextfeld — z.B. für den Namen, der auf dem Namensschild erscheinen soll.',
      helpTextEn: 'Free-text field — e.g. the name to print on your badge.',
    },
    {
      id: 'demo_select',
      label: 'Essenspräferenz',
      labelEn: 'Meal preference',
      type: 'select',
      required: true,
      options: ['Fleisch', 'Vegetarisch', 'Vegan'],
      optionsEn: ['Meat', 'Vegetarian', 'Vegan'],
      helpText: 'Einfachauswahl-Dropdown — genau eine Option wählbar.',
      helpTextEn: 'Single-select dropdown — exactly one option.',
    },
    {
      id: 'demo_multi',
      label: 'Allergien',
      labelEn: 'Allergies',
      type: 'select',
      multi: true,
      required: false,
      options: ['Gluten', 'Laktose', 'Nüsse', 'Meeresfrüchte'],
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
      helpText: 'Pflicht-Checkbox mit individuellem Bestätigungstext.',
      helpTextEn: 'Required checkbox with a custom confirmation label.',
    },
    {
      id: 'demo_hotel',
      label: 'Hotel benötigt?',
      labelEn: 'Hotel needed?',
      type: 'checkbox',
      required: false,
      confirmLabel: 'Ja, ich brauche ein Hotelzimmer',
      confirmLabelEn: 'Yes, I need a hotel room',
      helpText: 'Steuerfeld für die folgende Sichtbarkeitsbedingung.',
      helpTextEn: 'Controls the visibility of the next field.',
    },
    {
      id: 'demo_roommate',
      label: 'Wunsch-Zimmerpartner:in',
      labelEn: 'Preferred roommate',
      type: 'roommate',
      required: false,
      helpText: 'Personen-Picker (Roommate) — erscheint nur, wenn „Hotel benötigt?" angehakt ist. Die gewählte Person bekäme im echten Event eine Info-Mail.',
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
      options: ['Anfänger:in', 'Fortgeschritten', 'Profi'],
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
    { id: 'demo_ag_2', date: day, time: '09:30', endTime: '10:30', icon: 'Megaphone', title: 'Keynote', description: 'Begrüßung und Ausblick.' },
    { id: 'demo_ag_3', date: day, time: '10:30', endTime: '12:00', icon: 'People', title: 'Workshops', description: 'Parallele Sessions in Gruppen.' },
    { id: 'demo_ag_4', date: day, time: '12:00', endTime: '13:00', icon: 'Food', title: 'Mittagessen', description: 'Buffet im Atrium.' },
    { id: 'demo_ag_5', date: day, time: '13:00', endTime: '17:00', icon: 'Running', title: 'Lauf / Aktivprogramm', description: 'B2Run-Strecke oder Spaziergang.' },
  ];
}

function demoTransfers(): TransferTime[] {
  const day = ymdInDays(14);
  return [
    { id: 'demo_tr_1', location: 'Hauptbahnhof', meetingPoint: 'Gleis 1, Treffpunkt-Schild', address: 'Bahnhofsplatz 1', date: day, departureTime: '08:00', arrivalTime: '08:45', description: 'Shuttle-Bus zum Veranstaltungsort.' },
    { id: 'demo_tr_2', location: 'Veranstaltungsort', meetingPoint: 'Haupteingang', address: 'Eventstrasse 5', date: day, departureTime: '18:00', arrivalTime: '18:45', description: 'Rückfahrt zum Hauptbahnhof.' },
  ];
}

// Ein Sub-Event (Networking-Dinner) mit eigenen Feldern, damit die
// Sub-Event-Auswahl + die Pro-Sub-Event-Rückfragen demonstriert werden.
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
          label: 'Getränkewunsch',
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

// v17.26 / v18: Deterministische synthetische Teilnehmerliste für das
// Demo-Event, damit der Admin/Organizer unter „Organizer & Events" die
// Teilnehmer-Verwaltung mit echten Daten durchspielen kann (~20 Angemeldete,
// ein paar Abgemeldete, zwei Warteliste-Einträge, ein Team). Rein
// client-seitig — kein SharePoint.
const DEMO_FIRST_NAMES = ['Anna', 'Ben', 'Clara', 'David', 'Eva', 'Felix', 'Greta', 'Hannes', 'Ida', 'Jonas', 'Klara', 'Lukas', 'Mia', 'Noah', 'Olivia', 'Paul', 'Quirin', 'Rosa', 'Sophie', 'Tom', 'Ulrike', 'Viktor', 'Wanda', 'Xaver', 'Yara', 'Zoe'];
const DEMO_LAST_NAMES = ['Müller', 'Schmidt', 'Schneider', 'Fischer', 'Weber', 'Meyer', 'Wagner', 'Becker', 'Schulz', 'Hoffmann', 'Koch', 'Bauer', 'Richter', 'Klein', 'Wolf', 'Schröder', 'Neumann', 'Braun', 'Werner', 'Krüger', 'Hofmann', 'Hartmann', 'Lange', 'Schmitt', 'Krause', 'Meier'];
const DEMO_LOCATIONS = ['Düsseldorf', 'München', 'Berlin', 'Hamburg', 'Frankfurt', 'Stuttgart', 'Köln'];

const UMLAUT_MAP: Record<string, string> = { 'ä': 'ae', 'ö': 'oe', 'ü': 'ue', 'ß': 'ss' };
function demoFirst(idx: number): string { return DEMO_FIRST_NAMES[idx % DEMO_FIRST_NAMES.length]; }
function demoLast(idx: number): string { return DEMO_LAST_NAMES[idx % DEMO_LAST_NAMES.length]; }
function demoName(idx: number): string { return `${demoFirst(idx)} ${demoLast(idx)}`; }
function demoEmail(idx: number): string {
  return `${demoFirst(idx).toLowerCase()}.${demoLast(idx).toLowerCase().replace(/[äöüß]/g, m => UMLAUT_MAP[m] || m)}@deloitte.de`;
}
// People-Picker-Wert im Format „Name <email>" (so erwartet es die App +
// die Roommate-Match-Logik in AdminPage).
function pickerValue(idx: number): string { return `${demoName(idx)} <${demoEmail(idx)}>`; }

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
  // Roommate-Paare (gegenseitig → „Match"). Nur gerade Indizes (die haben
  // Hotel=true, daher ist das showIf-abhängige Roommate-Feld sichtbar):
  // 0↔2, 4↔6, 8↔10, 12↔14, 16↔18.
  const roommatePartner = (i: number): number => ((i / 2) % 2 === 0 ? i + 2 : i - 2);
  const MEALS = ['Vegetarisch', 'Fleisch', 'Vegan'];
  const ALLERGIES = [['Gluten', 'Laktose'], [], ['Nüsse'], ['Meeresfrüchte', 'Gluten']];
  const RUN_LEVELS = ['Anfänger:in', 'Fortgeschritten', 'Profi'];

  let tid = 1;
  // 20 Angemeldete (= voll bei maxParticipants 20), abwechselnd Gruppe A/B,
  // ALLE Custom-Fields mit Demo-Daten befüllt.
  for (let i = 0; i < 20; i++) {
    const starter = i % 2 === 0 ? 'Durchstarter' : 'Funstarter'; // gerade = Gruppe A
    const hotel = i % 2 === 0; // gerade Indizes brauchen Hotel → Roommate sichtbar
    const cd: Record<string, string> = {
      demo_text: demoFirst(i), // Wunschname auf dem Badge
      demo_select: MEALS[i % MEALS.length], // Essenspräferenz
      demo_multi: ALLERGIES[i % ALLERGIES.length].join(' | '), // Allergien (Multi)
      demo_number: String(i % 3), // Anzahl Begleitpersonen (0..2)
      demo_checkbox: 'true', // Teilnahmebedingungen
      demo_hotel: hotel ? 'true' : 'false', // Hotel benötigt?
      demo_user: pickerValue((i + 3) % 20), // Ansprechpartner:in im Team
    };
    // Roommate nur wenn Hotel benötigt (showIf demo_hotel=true) → gegenseitig.
    if (hotel) cd.demo_roommate = pickerValue(roommatePartner(i));
    // Gruppen-A-Feld (Lauferfahrung) nur für Durchstarter.
    if (starter === 'Durchstarter') cd.demo_group_a_field = RUN_LEVELS[i % RUN_LEVELS.length];

    const row: SPRegistration = {
      Id: 1000 + i,
      Title: demoEmail(i),
      TeilnehmerID: tid++,
      Anrede: i % 3 === 0 ? 'Frau' : i % 3 === 1 ? 'Herr' : 'Divers',
      Vorname: demoFirst(i),
      Nachname: demoLast(i),
      StarterType: starter,
      ParticipantName: demoName(i),
      ParticipantEmail: demoEmail(i),
      Status: 'Angemeldet',
      RegistrationDate: mkDate(i),
      CancellationDate: '',
      Location: DEMO_LOCATIONS[i % DEMO_LOCATIONS.length],
      CustomData: JSON.stringify(cd),
    };
    // „Registriert von" demonstrieren: eine Anmeldung stellvertretend durch
    // einen Organizer angelegt.
    if (i === 5) { row.RegisteredByName = 'Beispiel, Anna'; row.RegisteredByEmail = 'anna.beispiel@deloitte.de'; }
    // Nachrück-Audit demonstrieren: die letzten zwei Aktiven sind von der
    // Warteliste nachgerückt und haben je eine abgemeldete Person ersetzt.
    if (i === 18) { row.PromotedDate = mkDate(17); row.ReplacedParticipantEmail = demoEmail(22); }
    if (i === 19) { row.PromotedDate = mkDate(18); row.ReplacedParticipantEmail = demoEmail(23); }
    rows.push(row);
  }
  // 2 Warteliste-Einträge (konsistent: Event ist mit 20/20 voll), ebenfalls
  // mit Custom-Daten gefüllt.
  for (let k = 0; k < 2; k++) {
    const i = 20 + k;
    rows.push({
      Id: 2000 + k,
      Title: demoEmail(i),
      TeilnehmerID: tid++,
      Anrede: 'Keine Angabe',
      Vorname: demoFirst(i),
      Nachname: demoLast(i),
      PreferredStarterType: k % 2 === 0 ? 'Durchstarter' : 'Funstarter',
      ParticipantName: demoName(i),
      ParticipantEmail: demoEmail(i),
      Status: 'Warteliste',
      RegistrationDate: mkDate(20 + k),
      CancellationDate: '',
      Location: DEMO_LOCATIONS[i % DEMO_LOCATIONS.length],
      CustomData: JSON.stringify({
        demo_text: demoFirst(i),
        demo_select: MEALS[i % MEALS.length],
        demo_checkbox: 'true',
        demo_hotel: 'false',
        demo_user: pickerValue((i + 1) % 20),
      }),
    });
  }
  // 3 Abgemeldete — zwei davon wurden durch die Nachrücker (i=18/19) ersetzt.
  for (let k = 0; k < 3; k++) {
    const i = 22 + k;
    const row: SPRegistration = {
      Id: 3000 + k,
      Title: demoEmail(i),
      Vorname: demoFirst(i),
      Nachname: demoLast(i),
      ParticipantName: demoName(i),
      ParticipantEmail: demoEmail(i),
      Status: 'Abgemeldet',
      RegistrationDate: mkDate(k),
      CancellationDate: mkDate(15 + k),
      Location: DEMO_LOCATIONS[i % DEMO_LOCATIONS.length],
      CustomData: '{}',
    };
    if (k === 0) row.ReplacedByParticipantEmail = demoEmail(18);
    if (k === 1) row.ReplacedByParticipantEmail = demoEmail(19);
    rows.push(row);
  }
  return rows;
}

/**
 * Synthetische Eigen-Anmeldung des Demo-Users für die „Meine Events"-Ansicht.
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
    // v18.7: Kapazität 20 → 20 Angemeldete (voll) + 2 Warteliste sind damit
    // konsistent (vorher 50/20/2 = unstimmig).
    maxParticipants: 20,
    currentParticipants: 20,
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
    // v18: Standardmäßig EINE Teilnehmergruppe (keine geteilte Kapazität).
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
