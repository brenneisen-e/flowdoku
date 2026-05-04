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
  /** Seit v6.19: QR-Code-Scanner pro Event. Diese User haben eingeschränkten Admin-Zugriff:
   *  - Dürfen NICHT Teilnehmer bearbeiten, Event editieren, Mails versenden etc.
   *  - Dürfen das QR-Code-Scanner-Tool nutzen + Check-In-KPIs ansehen (wie viele eingecheckt, wie viele ausstehen).
   *  - Erscheinen NICHT in der Organizer-Liste auf MyEvents/RegistrationPage.
   *  - Bekommen KEINE Organizer-Mails.
   *  Persistenz: EmailTemplateOverrides._qrScanners (JSON). qrScannerNames + qrScannerEmails sind index-synchron. */
  qrScannerNames: string[];
  qrScannerEmails: string[];
  /** v9.18: Co-Organizer pro Event — beliebiger Deloitte-User, kein Admin/Organizer-Status nötig.
   *  Diese User haben für DIESES eine Event die gleichen Rechte wie der Hauptorganizer.
   *  Persistenz: EmailTemplateOverrides._coOrganizers (JSON, analog _qrScanners).
   *  Ab v9.20 wird das Co-Organizer-Konzept faktisch durch den Graph-Search-Organizer-
   *  Picker ersetzt — der Hauptorganizer selbst kann jeden beliebigen Deloitte-User
   *  als Organizer hinterlegen. Felder bleiben fuer Backward-Compat. */
  coOrganizerNames?: string[];
  coOrganizerEmails?: string[];
  /** v9.21: Test-Team pro Event — Personen die diesen Event bereits im Entwurfs-
   *  Modus sehen + sich anmelden duerfen. Sie testen das Event durch, bevor der
   *  Organizer das "Entwurf"-Haekchen wegnimmt. Persistenz:
   *  EmailTemplateOverrides._testTeam (JSON, analog _qrScanners). */
  testTeamNames?: string[];
  testTeamEmails?: string[];
  /** v9.21: Datum, ab dem das Event auto-aktiv wird (sonst bleibt es im
   *  Entwurfs-/Under-Construction-Modus). Wenn leer = sofort aktiv (sobald
   *  isFictive=false und EventStatus='Active'). */
  activeFrom?: string;
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
  waitlistEnabled?: boolean;
  autoSendQRCode?: boolean; // v9.15 — automatisch QR-Code-Mail nach Anmeldung versenden
  imageUrl?: string;
  subsiteUrl?: string;
  outlookBody: string;
  emailLanguage: string; // 'DE' | 'EN'
  emailTemplateOverrides?: string; // JSON mit Event-spezifischen Template-Anpassungen
  disableEmails?: boolean;   // Wenn true: keine E-Mails bei An-/Abmeldung
  disableOutlook?: boolean;  // Wenn true: keine Outlook-Kalendereintraege
  /** v8.5: Granulare Organizer-Benachrichtigung bei Anmeldungen.
   *  - 'never' (Default): Organizer bekommt nichts mit
   *  - 'always': Organizer wird bei jeder Anmeldung als BCC dazugesetzt
   *  - 'fromDate': erst ab notifyOrgRegisterFromDate wird BCC'd
   *  Hintergrund: bei grossen Events ist Organizer-Spam waehrend der
   *  Hauptanmeldephase nervig — kurz vor dem Event will man aber sehen,
   *  wer kurzfristig noch dazustoesst. */
  notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
  notifyOrgRegisterFromDate?: string; // ISO; nur relevant wenn mode='fromDate'

  /** v8.5: Granulare Organizer-Benachrichtigung bei Abmeldungen.
   *  - 'never': Organizer bekommt keine Storno-Bestaetigungs-Mails
   *  - 'always': BCC bei jedem Storno
   *  - 'afterDeadline': BCC erst wenn die Abmeldefrist (lastDeregisterDate)
   *    bereits ueberschritten ist — also nur bei "spaeten" Stornos, die
   *    fuer den Organizer planungsrelevant sind. */
  notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';

  /** v8.6: Liste explizit ausgeschlossener User-E-Mails (semikolon-separiert
   *  in SP, Array hier). Greift NACH Standortfilter + Mailverteiler/User —
   *  egal warum jemand sonst Sichtbarkeit haette, wenn seine Mail hier
   *  drin ist, sieht er das Event nicht. Genutzt z.B. wenn aus einem
   *  100er-Verteiler 5 Personen gezielt rausgenommen werden sollen. */
  excludedUsers?: string[];
  isFictive?: boolean;       // Wenn true: Event nur fuer Admins + eigene Organizer sichtbar (Test-Event)
  durchstarterCapacity?: number; // B2Run: getrennte Kapazitaet fuer Durchstarter
  funstarterCapacity?: number;   // B2Run: getrennte Kapazitaet fuer Funstarter
  /** Seit v6.15: optionale Verknüpfung Starter-Typ → Startblock. Wenn gesetzt,
   *  wird bei der Registrierung automatisch das passende Startblock-Custom-Field
   *  gesetzt (keine Einzel-Auswahl durch den User nötig). */
  durchstarterStartblock?: string;
  funstarterStartblock?: string;
  /** Seit v6.15: optionale Admin-Option. Wenn true, muss der User beim Wählen
   *  von Durchstarter eine Checkbox "Leistungsnachweis vorhanden" bestätigen. */
  durchstarterRequiresProof?: boolean;
  agenda: AgendaItem[];
  transferTimes: TransferTime[];
  documents: EventDocument[];
  quiz: QuizQuestion[];
  /** Anzahl Fragen pro Ansicht im Quiz-Player (1..4). Default 1 = klassisch einzeln. */
  quizClusterSize?: number;
  eventSpecificFields: EventSpecificField[];
  /** Seit v6.4: wenn gesetzt, ist dieses Event ein Sub-Event (z.B. Trainingssession) und zeigt auf die id des Parent-Events. Top-Level-Events haben parentEventId=undefined/''. */
  parentEventId?: string;
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
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate';
  required: boolean;
  options?: string[];
  helpText?: string;
  spInternalName?: string;
  /** Optionale externe Links (AGB, Datenschutz etc.) unter dem Feld */
  externalLinks?: Array<{ label: string; url: string }>;
  /** v7.11: Wenn `type === 'select'` und `multi === true`, wird das Feld als
   *  Mehrfachauswahl gerendert (Checkbox-Liste). Mehrere gewählte Werte werden
   *  in `eventSpecificData[id]` als " | "-getrennter String gespeichert (Pipe
   *  mit Spaces, damit Optionen mit Komma im Label nicht zerrissen werden). */
  multi?: boolean;
  /** v7.21: Sichtbarkeitsbedingung. Wenn gesetzt, wird das Feld nur dann
   *  angezeigt, wenn das Quell-Feld (`fieldId`) einen der `values` als
   *  Antwort hat. Beispiel: Roommate-Feld nur sichtbar wenn Zimmerart =
   *  "Doppelzimmer". Bei Multi-Select-Quelle reicht es, wenn EINE der
   *  Quell-Antworten in `values` enthalten ist. Bei Checkbox-Quelle ist
   *  values = ['true'] oder ['false']. */
  showIf?: { fieldId: string; values: string[] };
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
