/* Beispieldaten für die Teilnehmer-Seiten des Harness (v31.7).
 *
 * Warum eigene Daten statt eines einzigen Normalfall-Events: Die Seiten sehen
 * je Zustand anders aus, und genau die Zustände sind das, was ein Umbau
 * kaputtmachen kann — freie Plätze, volle Warteliste, Klammer-Event ohne
 * eigene Anmeldung, „QR versendet", Vergangenheit, Abmeldung. Wer nur den
 * Normalfall rendert, sieht die Hälfte der Oberfläche nie.
 *
 * Alles ist `any`: Die Seiten lesen viele optionale Felder von `DeloitteEvent`,
 * und ein vollständig getyptes Beispiel-Event wäre hundert Zeilen Pflichtfelder
 * ohne Erkenntnisgewinn. Der Harness ist ein Betrachter, kein Testfall.
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

export const me = {
  id: '1',
  firstName: 'Eike',
  surname: 'Brenneisen',
  email: 'eike.brenneisen@example.com',
  isAdmin: true,
  role: 'Admin',
  location: 'DE - Koeln',
  jobTitle: 'Manager',
};

/** ISO-Zeitstempel relativ zu heute — die Bilder sollen nie „abgelaufen" wirken. */
export const d = (offsetDays: number, hh: number, mm = 0): string => {
  const x = new Date();
  x.setDate(x.getDate() + offsetDays);
  x.setHours(hh, mm, 0, 0);
  return x.toISOString();
};
export const day = (offsetDays: number): string => d(offsetDays, 9).slice(0, 10);

/* Ein Platzhalter-Bild als Data-URI. SharePoint-Bilder gibt es hier nicht, und
 * eine leere Kachel sagt nichts über das Layout mit Bild. */
const heroSvg = (label: string, from: string, to: string): string =>
  'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="800" height="450">
      <defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="${from}"/><stop offset="100%" stop-color="${to}"/>
      </linearGradient></defs>
      <rect width="800" height="450" fill="url(#g)"/>
      <text x="400" y="240" font-family="Segoe UI, Arial" font-size="46" font-weight="700"
            fill="rgba(255,255,255,0.92)" text-anchor="middle">${label}</text>
    </svg>`)));

const IMG_SOMMER = heroSvg('Sommerfest', '#86bc25', '#2f6b0f');
const IMG_KOCH = heroSvg('Kochkurs', '#ed8b00', '#a34b00');
const IMG_OFFSITE = heroSvg('Offsite', '#0076a8', '#00344a');
const IMG_DTP = heroSvg('DTP Basics', '#62b5e5', '#005587');
const IMG_NEUJAHR = heroSvg('Neujahr', '#6a6a6a', '#2b2b2b');

/** Gemeinsamer Rumpf — nur die Felder, ohne die eine Karte gar nicht rendert. */
const base = (over: any): any => ({
  type: 'Other',
  status: 'Active',
  organizers: ['Rettinger, Carolin'],
  organizerEmails: ['carolin.rettinger@example.com'],
  coOrganizerNames: [],
  coOrganizerEmails: [],
  qrScannerNames: [],
  qrScannerEmails: [],
  locationAudience: [],
  audienceFilter: [],
  filterMode: 'OR',
  waitlistEnabled: false,
  waitlistCount: 0,
  currentParticipants: 0,
  maxParticipants: 0,
  emailLanguage: 'DE',
  outlookBody: '',
  imageUrl: '',
  agenda: [],
  transferTimes: [],
  documents: [],
  quiz: [],
  eventSpecificFields: [],
  ...over,
});

/* ---------------------------------------------------------------- Felder --- */

/* Eventspezifische Felder der Anmeldeseite: je Typ eins, ein Pflichtfeld
 * dabei, dazu ein abhängiges Feld (showIf) — das ist der Fall, der beim
 * Umbau der Formularsektion am ehesten verschwindet. */
const sommerFelder: any[] = [
  {
    id: 'anreise', label: 'Wie kommst du zum Sommerfest?', type: 'select', required: true,
    options: ['Mit der Bahn', 'Mit dem Auto', 'Zu Fuß / Rad', 'Ich weiß es noch nicht'],
    helpText: 'Wir planen damit die Shuttle-Zeiten ab dem Hauptbahnhof.',
    helpTextStyle: 'inline',
  },
  {
    id: 'essen', label: 'Was passt bei dir aufs Buffet?', type: 'select', required: true, multi: true,
    options: ['Alles', 'Vegetarisch', 'Vegan', 'Ohne Schwein', 'Glutenfrei', 'Laktosefrei'],
    defaultValue: '',
  },
  {
    id: 'allergien', label: 'Gibt es Allergien, von denen die Küche wissen sollte?', type: 'text',
    required: false, helpText: 'Freitext — steht nur den Organizern zur Verfügung.',
  },
  {
    id: 'begleitung', label: 'Bringst du eine Begleitung mit?', type: 'checkbox', required: false,
    confirmLabel: 'Ja, ich komme zu zweit',
  },
  {
    id: 'begleitname', label: 'Wie heißt deine Begleitung?', type: 'text', required: false,
    showIf: { fieldId: 'begleitung', values: ['true'] },
  },
  {
    id: 'ankunft', label: 'Wann willst du da sein?', type: 'date', required: false, withTime: true,
    helpText: 'Nur ein Richtwert für den Empfang.',
  },
  {
    id: 'shirt', label: 'Welche Shirt-Größe brauchst du?', type: 'select', required: false,
    options: ['S', 'M', 'L', 'XL', 'S', 'M', 'L', 'XL'],
    optionCategories: ['Damen', 'Damen', 'Damen', 'Damen', 'Herren', 'Herren', 'Herren', 'Herren'],
    prefilterLabel: 'Größentabelle',
  },
];

/* ---------------------------------------------------------------- Events --- */

/** Freie Plätze, keine eigene Anmeldung — das ist die Anmeldeseite im Normalfall. */
export const evOpen: any = base({
  id: 'ev-open', eventNumber: 5101, title: 'Sommerfest Köln 2026',
  startDate: d(24, 17), endDate: d(24, 23),
  registrationDeadline: d(18, 23, 59), lastDeregisterDate: d(21, 23, 59),
  location: 'Rheinterrassen, Köln-Deutz',
  locationAddress: { street: 'Kennedy-Ufer', houseNo: '2', zip: '50679', city: 'Köln' },
  description: '<p>Ein Abend am Rhein: Grill, Live-Musik und die Kolleginnen und Kollegen, '
    + 'die man sonst nur in Teams sieht. Getränke und Essen gehen aufs Haus.</p>'
    + '<p>Die Shuttles fahren ab 16:30 Uhr im Viertelstundentakt vom Hauptbahnhof.</p>',
  maxParticipants: 120, currentParticipants: 47,
  waitlistEnabled: true, waitlistCount: 0,
  imageUrl: IMG_SOMMER,
  eventSpecificFields: sommerFelder,
  agenda: [
    { id: 'so1', date: day(24), time: '17:00', endTime: '18:00', icon: 'Calendar', title: 'Empfang & Getränke', location: 'Terrasse', cluster: 'Abend' },
    { id: 'so2', date: day(24), time: '18:00', endTime: '19:30', icon: 'Calendar', title: 'Grill-Buffet', location: 'Zelt', cluster: 'Abend' },
    { id: 'so3', date: day(24), time: '20:00', endTime: '23:00', icon: 'Calendar', title: 'Live-Musik', location: 'Bühne', cluster: 'Abend' },
  ],
});

/** Voll, mit Warteliste — und die eigene Zeile steht auf der Warteliste. */
export const evWait: any = base({
  id: 'ev-wait', eventNumber: 5102, title: 'Kochkurs am Rhein',
  startDate: d(12, 18), endDate: d(12, 22),
  registrationDeadline: d(8, 23, 59), lastDeregisterDate: d(10, 23, 59),
  location: 'Kochwerkstatt Ehrenfeld',
  description: '<p>Vier Gänge, sechs Herde, ein Abend. Der Kurs ist ausgebucht — '
    + 'wer auf der Warteliste steht, rückt automatisch nach, sobald jemand absagt.</p>',
  maxParticipants: 12, currentParticipants: 12,
  waitlistEnabled: true, waitlistCount: 4,
  imageUrl: IMG_KOCH,
  eventSpecificFields: [
    { id: 'diaet', label: 'Isst du etwas nicht?', type: 'text', required: false },
  ],
});

/** Eigene Anmeldung mit Status „QR versendet" und Start übermorgen — damit
 *  zeigt die Landing Page ihren QR-Kasten und „Meine Events" den QR-Knopf. */
export const evQr: any = base({
  id: 'ev-qr', eventNumber: 5103, title: 'Team-Offsite Hamburg',
  startDate: d(2, 9), endDate: d(3, 16),
  registrationDeadline: d(-4, 23, 59), lastDeregisterDate: d(0, 23, 59),
  location: 'Speicherstadt, Hamburg',
  description: '<p>Zwei Tage Strategie, Retrospektive und Hafenrundfahrt. '
    + 'Die Zimmer sind gebucht, den Check-in-Code hast du per Mail bekommen.</p>',
  maxParticipants: 30, currentParticipants: 28,
  waitlistEnabled: true, waitlistCount: 1,
  imageUrl: IMG_OFFSITE,
  eventSpecificFields: [
    { id: 'hotel', label: 'Brauchst du eine Übernachtung?', type: 'select', required: true, options: ['Ja, bitte', 'Nein, danke'] },
  ],
});

/** Klammer-Event: nur die Termine sind buchbar (subEventsOnlyMode). */
export const evUmbrella: any = base({
  id: 'ev-umb', eventNumber: 5110, title: 'DTP Basics Training Oktober 2026',
  startDate: d(30, 9), endDate: d(32, 17),
  registrationDeadline: d(20, 23, 59), lastDeregisterDate: d(25, 23, 59),
  location: 'Campus Hackescher Markt (Haus A, Etage 4)',
  description: '<p>Drei Tage Grundlagen für neue Kolleginnen und Kollegen. '
    + 'Du buchst die Tage einzeln — jeder Tag hat eigene Plätze.</p>',
  subEventsOnlyMode: true,
  requireSubEventSelection: true,
  childEventTermSingular: 'Tag', childEventTermPlural: 'Tage',
  imageUrl: IMG_DTP,
  eventSpecificFields: [
    { id: 'vorwissen', label: 'Wie viel Vorwissen bringst du mit?', type: 'select', required: true, options: ['Keins', 'Etwas', 'Viel'] },
  ],
});

const kind = (id: string, nr: number, title: string, off: number, over: any = {}): any => base({
  id, eventNumber: nr, title, parentEventId: 'ev-umb',
  startDate: d(off, 9), endDate: d(off, 17),
  registrationDeadline: d(20, 23, 59), lastDeregisterDate: d(25, 23, 59),
  location: 'Campus Hackescher Markt',
  description: '<p>Fachthema am Vormittag, Fallarbeit am Nachmittag.</p>',
  maxParticipants: 20, currentParticipants: 8,
  organizers: ['Rettinger, Carolin'], organizerEmails: ['carolin.rettinger@example.com'],
  ...over,
});

export const evUmbrellaKids: any[] = [
  kind('ev-umb-1', 5111, 'Tag 1 · Business Chemistry', 30, { currentParticipants: 8 }),
  kind('ev-umb-2', 5112, 'Tag 2 · Joint Venture', 31, {
    currentParticipants: 20, waitlistEnabled: true, waitlistCount: 3,
  }),
  kind('ev-umb-3', 5113, 'Tag 3 · Post Merger Integration', 32, { currentParticipants: 4 }),
];

/** Vergangen — in „Meine Events" landet es im Rückblick-Block. */
export const evPast: any = base({
  id: 'ev-past', eventNumber: 5120, title: 'Neujahrsempfang 2026',
  startDate: d(-96, 18), endDate: d(-96, 23),
  registrationDeadline: d(-104, 23, 59), lastDeregisterDate: d(-100, 23, 59),
  location: 'Foyer, Köln',
  description: '<p>Der Jahresauftakt mit Rückblick und Ausblick.</p>',
  maxParticipants: 200, currentParticipants: 173,
  imageUrl: IMG_NEUJAHR,
});

/** Abgemeldet — die Zeile steht nicht mehr in DEX_Participants, nur die
 *  Teilnehmerliste kennt sie noch mit Status „Abgemeldet". */
export const evCancelled: any = base({
  id: 'ev-can', eventNumber: 5130, title: 'Legal Update Q4',
  startDate: d(40, 10), endDate: d(40, 12),
  registrationDeadline: d(35, 23, 59), lastDeregisterDate: d(38, 23, 59),
  location: 'Online (Teams)',
  description: '<p>Was sich zum Jahreswechsel ändert — kompakt in zwei Stunden.</p>',
  maxParticipants: 500, currentParticipants: 96,
});

export const allEvents: any[] = [
  evOpen, evWait, evQr, evUmbrella, ...evUmbrellaKids, evPast, evCancelled,
];
export const topLevelEvents: any[] = allEvents.filter(e => !e.parentEventId);

/* --------------------------------------------------------- Anmeldungen --- */

const reg = (over: any): any => ({
  Id: 1, Title: me.email, TeilnehmerID: 1,
  Anrede: '', Vorname: me.firstName, Nachname: me.surname,
  ParticipantName: `${me.surname}, ${me.firstName}`,
  ParticipantEmail: me.email,
  Status: 'Angemeldet',
  RegistrationDate: d(-6, 10), CancellationDate: '',
  CustomData: '{}',
  ...over,
});

/** Eigene Anmeldung je Event-Id. `undefined` heißt: keine Zeile. */
export const myRegistrations: Record<string, any> = {
  'ev-qr': reg({
    Id: 41, TeilnehmerID: 17, QrSentId: 17, Status: 'QR versendet',
    RegistrationDate: d(-18, 9),
    CustomData: JSON.stringify({ hotel: 'Ja, bitte' }),
  }),
  // Die Wartelisten-Position ist der Rang nach TeilnehmerID, gerechnet als
  // `TeilnehmerID - maxParticipants` (MyEventCard). Eine kleine Nummer ergäbe
  // hier „Warteliste #-8" — die Nummer muss ÜBER der Kapazität liegen.
  'ev-wait': reg({
    Id: 42, TeilnehmerID: 14, Status: 'Warteliste',
    RegistrationDate: d(-2, 14),
    CustomData: JSON.stringify({ diaet: 'kein Fisch' }),
  }),
  // Klammer: Schattenzeile mit den übergreifenden Antworten, kein Platz.
  'ev-umb': reg({
    Id: 43, TeilnehmerID: 6, Status: 'Angemeldet',
    RegistrationDate: d(-9, 11),
    CustomData: JSON.stringify({ vorwissen: 'Etwas' }),
  }),
  'ev-umb-1': reg({ Id: 44, TeilnehmerID: 6, Status: 'Angemeldet', RegistrationDate: d(-9, 11) }),
  'ev-umb-2': reg({ Id: 45, TeilnehmerID: 22, Status: 'Warteliste', RegistrationDate: d(-9, 11) }),
  'ev-past': reg({
    Id: 46, TeilnehmerID: 88, QrSentId: 88, Status: 'Eingecheckt',
    RegistrationDate: d(-120, 8),
    CheckedInDate: d(-96, 18), CheckedInByName: 'Rettinger, Carolin',
  }),
  'ev-can': reg({
    Id: 47, TeilnehmerID: 12, Status: 'Abgemeldet',
    RegistrationDate: d(-30, 9), CancellationDate: d(-11, 16),
  }),
};

/** Was DEX_Participants über mich weiß. Die abgemeldete Zeile steht hier
 *  bewusst NICHT — genau so ist es live, und nur deshalb läuft in
 *  „Meine Events" die Abgemeldet-Nachsuche über alle übrigen Events. */
export const myEventNumbers = {
  registered: [5103, 5110, 5111, 5120],
  waitlisted: [5102, 5112],
};

/** Teilnehmerliste eines Events — für Ansichten, die zählen statt zu lesen. */
export const registrationsOf = (eventId: string): any[] => {
  const mine = myRegistrations[eventId];
  return mine ? [mine] : [];
};
