// DEX Event Platform - Type Definitions
// Author: Eike Brenneisen

export type EventType = 'B2Run' | 'JPMorgan' | 'Other';
export type EventStatus = 'Under Construction' | 'Active' | 'Completed' | 'Cancelled';
export type RegistrationStatus = 'Angemeldet' | 'QR versendet' | 'Warteliste' | 'Eingecheckt' | 'Abgemeldet';
export type Salutation = 'Herr' | 'Frau' | 'Divers';

export interface DeloitteEvent {
  id: string;
  eventNumber: number;
  title: string;
  type: EventType;
  status: EventStatus;
  organizers: string[];
  organizerEmails: string[]; // ';'-separiert in SharePoint (OrganizerEmail), hier als Array fuer Benachrichtigungen
  location: string;
  locationAddress?: { street: string; houseNo: string; zip: string; city: string };
  locationAudience: string[];
  audienceFilter: string[];
  filterMode: 'AND' | 'OR';
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  lastDeregisterDate: string;
  description: string;
  maxParticipants: number;
  currentParticipants: number;
  waitlistCount: number;
  imageUrl?: string;
  subsiteUrl?: string;
  outlookBody: string;
  emailLanguage: string; // 'DE' | 'EN'
  emailTemplateOverrides?: string; // JSON mit Event-spezifischen Template-Anpassungen
  disableEmails?: boolean;   // Wenn true: keine E-Mails bei An-/Abmeldung
  disableOutlook?: boolean;  // Wenn true: keine Outlook-Kalendereintraege
  isFictive?: boolean;       // Wenn true: Event nur fuer Admins + eigene Organizer sichtbar (Test-Event)
  durchstarterCapacity?: number; // B2Run: getrennte Kapazitaet fuer Durchstarter
  funstarterCapacity?: number;   // B2Run: getrennte Kapazitaet fuer Funstarter
  agenda: AgendaItem[];
  transferTimes: TransferTime[];
  documents: EventDocument[];
  quiz: QuizQuestion[];
  /** Anzahl Fragen pro Ansicht im Quiz-Player (1..4). Default 1 = klassisch einzeln. */
  quizClusterSize?: number;
  eventSpecificFields: EventSpecificField[];
  /** Optionale Sub-Events (z.B. Trainingssessions fuer B2Run). Eigenes Datum/Ort/Kapazitaet/Outlook-Termin.
   * Kein separater Teilnehmerlisten-Eintrag — die Registrierung wird als JSON-Array
   * (SubEventIds) im Teilnehmer-Eintrag mitgefuehrt. */
  subEvents?: SubEvent[];
}

export interface SubEvent {
  id: string;                       // uuid
  title: string;
  description?: string;
  location?: string;
  startDate: string;                // ISO-DateTime
  endDate: string;                  // ISO-DateTime
  maxParticipants?: number;         // optional — 0 oder undefined = keine Begrenzung
  registrationDeadline?: string;    // ISO-DateTime, optional
  lastDeregisterDate?: string;      // ISO-DateTime, optional
  /** HTML-Body fuer den Outlook-Termin dieses Sub-Events (bereits im Deloitte-Layout gewrappt). */
  outlookBody?: string;
  /** Wenn true: Mails fuer An-/Abmeldung dieses Sub-Events unterdruecken. */
  disableEmails?: boolean;
  /** Wenn true: keinen separaten Outlook-Kalendereintrag fuer dieses Sub-Event. */
  disableOutlook?: boolean;
  /** Icon-Name aus Fluent UI (optional, z.B. 'Running', 'Education'). */
  icon?: string;
}

export interface AgendaItem {
  id: string;
  date: string;        // YYYY-MM-DD
  time: string;         // HH:mm
  endTime?: string;     // HH:mm
  icon: string;         // FluentUI icon name (z.B. 'Calendar', 'People', 'Food')
  title: string;
  description?: string;
}

export interface TransferTime {
  id: string;
  location: string;    // Stadt/Standort (z.B. "Düsseldorf", "München")
  meetingPoint: string; // Treffpunkt (z.B. "Flughafen Terminal 1", "Hauptbahnhof Gleis 5")
  address?: string;     // Genaue Adresse (z.B. "Terminalstraße 1, 40474 Düsseldorf")
  date: string;        // YYYY-MM-DD
  departureTime: string; // HH:mm
  arrivalTime?: string;  // HH:mm
  description?: string;  // Zusatzinfo
}

export interface QuizQuestion {
  id: string;
  question: string;
  options: string[];
  correctIndices: number[]; // Indices der richtigen Antworten (0-basiert, mehrere moeglich)
  /** Optionales Bild zur Frage (data-URL mit base64-jpeg oder png). Wird inline im
   * FunZone-JSON gespeichert. Max empfohlen: ~80KB pro Bild (komprimiert). */
  imageBase64?: string;
  /** Optionaler Bereichs-Name. Fragen mit gleichem Bereich werden zusammen auf
   * einer Quiz-Seite angezeigt (statt nach clusterSize gepaginiert). */
  section?: string;
}

export interface EventDocument {
  name: string;
  url: string;
  size?: number;
}

export interface EventSpecificField {
  id: string;
  label: string;
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user';
  required: boolean;
  options?: string[];
  helpText?: string;
  spInternalName?: string;
  /** Optionale externe Links (AGB, Datenschutz etc.) unter dem Feld */
  externalLinks?: Array<{ label: string; url: string }>;
}

export interface Registration {
  id: string;
  eventId: string;
  eventTitle: string;
  salutation: Salutation;
  firstName: string;
  surname: string;
  email: string;
  status: RegistrationStatus;
  registrationDate: string;
  cancellationDate?: string;
  waitlistPosition?: number;
  eventSpecificData: Record<string, string>;
  /** IDs der Sub-Events, fuer die sich der Teilnehmer zusaetzlich angemeldet hat. */
  subEventIds?: string[];
}

export type UserRole = 'Admin' | 'Organizer' | 'User';

export interface User {
  id: string;
  firstName: string;
  surname: string;
  email: string;
  isAdmin: boolean;
  role: UserRole;
  location: string;
  jobTitle?: string; // z.B. "Assistant", "Senior Assistant", "Director", "Partner"
}

export interface RoleAssignment {
  id: number;
  userEmail: string;
  userName: string;
  role: UserRole;
  location: string;
  assignedBy: string;
  assignedDate: string;
}
