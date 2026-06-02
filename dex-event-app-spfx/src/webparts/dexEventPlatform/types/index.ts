// DEX Event Platform - Type Definitions
// Author: Eike Brenneisen

export type EventType = 'B2Run' | 'JPMorgan' | 'Other';
// v11.89: 'Under Construction' wurde konsolidiert in das IsFictive-Flag.
// Der „Entwurf"-Zustand eines Events lebt jetzt ausschließlich auf
// IsFictive — EventStatus beschreibt nur noch den Lebenszyklus.
export type EventStatus = 'Active' | 'Completed' | 'Cancelled';
export type RegistrationStatus = 'Angemeldet' | 'QR versendet' | 'Warteliste' | 'Eingecheckt' | 'Abgemeldet';
export type Salutation = 'Herr' | 'Frau' | 'Divers' | 'Keine Angabe';

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
  /** v10.16: Optionaler Ansprechpartner pro Event — z.B. die Person vor Ort
   *  für Nachfragen, die NICHT die App-Organizer-Rechte haben muss. Anzeige
   *  auf der Registration-Page als zusätzlicher Kontakt unter den Organizern,
   *  in den Bestätigungs-Mails optional. Reines Anzeige-Feld (kein App-Login,
   *  keine SP-Berechtigungen), daher als Freitext. */
  contactName?: string;
  contactEmail?: string;
  contactInfo?: string;
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
  /** v16.4: Vor-aufgeloeste E-Mails der DLs aus audienceFilter (lowercase),
   *  beim Event-Save via Graph auf transitive Members aufgeloest. matchesAudience
   *  in EventListPage checkt zusaetzlich gegen diese Liste — damit funktioniert
   *  die Sichtbarkeit auch fuer verschachtelte DLs, die /me/memberOf nicht
   *  zurueckliefert. Bei DL-Member-Aenderungen muss der Organizer das Event
   *  einmal neu speichern, damit der Cache aktualisiert wird. */
  audienceResolvedEmails?: string[];
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
  /** v18.34/v18.40: Ort für das Location-Feld des Outlook-Termins. Standard =
   *  Veranstaltungsort + Adresse (automatisch gebaut), im Wizard überschreibbar. */
  outlookLocation?: string;
  /** OutlookEventId aus DEX_Events. Wird vom DEX_CreateOutlookEvent-Flow nach
   *  erfolgreichem Anlegen des Kalendertermins geschrieben (Wert='FAILED' bei
   *  Fehler). Leer = noch kein Outlook-Termin angelegt; wichtig u.a. für die
   *  Erkennung „nachträglich Outlook aktiviert" auf Sub-Events
   *  (`persistSubEventsForParent` in `EventCreationPage.tsx`). */
  outlookEventId?: string;
  /** CalendarLink (iCalUId) aus DEX_Events. Wird vom DEX_CreateOutlookEvent-Flow
   *  nach erfolgreichem Anlegen geschrieben — und ist tatsaechlich das einzige
   *  Feld, das auf Erfolgsbasis gefuellt wird (OutlookEventId nur bei Fehler
   *  ='FAILED'). Wird in der App genutzt, um zu erkennen, ob ein Event einen
   *  Outlook-Termin besitzt — wichtig fuer die v11.57-Update-Confirm-Logik. */
  calendarLink?: string;
  /** v11.57: Marker, dass beim letzten Edit Outlook-relevante Felder (Title,
   *  Start, End, OutlookBody) geändert wurden, ohne dass der Organizer ein
   *  Update an die Teilnehmer-Outlook-Termine ausgelöst hat. Wird im
   *  Wizard-Schritt 1 (Grundlagen) als Hinweis-Box angezeigt, damit beim
   *  nächsten Save bewusst entschieden werden kann, ob die Teilnehmer eine
   *  „Aktualisierter Termin"-Benachrichtigung bekommen. */
  outlookDirty?: boolean;
  emailLanguage: string; // 'DE' | 'EN'
  /** v17.25: Markiert das synthetische Demo-Showcase-Event, das nur im
   *  Demo-Impersonation-Modus client-seitig in die Event-Liste injiziert
   *  wird. Treibt die „DEMO"-Markierung in der Liste, den Showcase-Banner
   *  + die einklappbaren Bereiche auf der Register-Seite und blockt einen
   *  echten Submit. Existiert nie in SharePoint. */
  isDemoShowcase?: boolean;
  /** v18.9: Wenn true, werden die Organizer-Chips (Name + Foto) auf der
   *  Anmelde-Seite und in „Meine Events" NICHT angezeigt. Rein visuell —
   *  Organizer-Rechte, Mails und BCC bleiben unberührt. Persistiert als
   *  Piggyback `_hideOrganizer` in EmailTemplateOverrides. */
  hideOrganizer?: boolean;
  /** v17.20: Wenn true UND der Organizer hat pro Custom-Field die
   *  EN-Variante hinterlegt (`labelEn` etc.), zeigt die Anmeldeseite die
   *  Felder in der **Locale des Teilnehmers** (App-Spracheinstellung) statt
   *  in der Event-Mail-Sprache. Zusätzlich wechselt das Form-Chrome
   *  (Placeholder, Hinweise, Sub-Event-Sektion-Labels) ebenfalls auf die
   *  Teilnehmer-Locale. Default false = altes Verhalten (Form-Chrome folgt
   *  `emailLanguage`, Custom-Field-Labels einsprachig). */
  bilingualFields?: boolean;
  emailTemplateOverrides?: string; // JSON mit Event-spezifischen Template-Anpassungen
  disableEmails?: boolean;   // Wenn true: keine E-Mails bei An-/Abmeldung
  disableOutlook?: boolean;  // Wenn true: keine Outlook-Kalendereintraege
  /** v14.5: Wenn true UND es existieren Sub-Events, muss der Teilnehmer
   *  beim Anmelden mindestens ein Sub-Event auswählen. Typischer Use-Case:
   *  Hauptevent-Kommunikation ist abgestellt (disableEmails/disableOutlook),
   *  Kommunikation läuft nur über die Sub-Events — dann darf niemand „nur
   *  Hauptevent" anmelden, sonst landet er ohne Bestätigung im System. */
  requireSubEventSelection?: boolean;
  /** v14.8: „Nur Sub-Events"-Modus. Wenn true, gibt es kein
   *  Hauptevent-Anmelde-Angebot mehr — die Anmelde-Checkbox für das
   *  Hauptevent ist im RegistrationForm ausgeblendet, der Teilnehmer
   *  meldet sich ausschließlich für einzelne Sub-Events an. Die
   *  Hauptevent-Kommunikation (Mails, Outlook) ist dadurch implizit
   *  irrelevant — der entsprechende Tab in Schritt 6 wird ausgegraut.
   *  Konsistenzregel: subEventsOnlyMode=true impliziert
   *  requireSubEventSelection=true. */
  subEventsOnlyMode?: boolean;
  /** v14.8: Organizer-konfigurierbarer Begriff für die untergeordneten
   *  Events. Default „Sub-Event(s)" (DE/EN). Per Wizard-Dropdown auch
   *  „Workshop", „Session", „Programmpunkt", „Event-Section" oder
   *  Freitext (Singular + Plural). Wird überall im User-facing Text
   *  verwendet. Persistiert als Piggyback `_childEventTerm` in
   *  EmailTemplateOverrides. */
  childEventTermSingular?: string;
  childEventTermPlural?: string;
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
  /** Geteilte Kapazitäten — bis v10.19 nur fuer B2Run gedacht (Durchstarter /
   *  Funstarter), seit v10.20 generisch fuer beliebige Events nutzbar. Wenn
   *  beide > 0 zeigt die Registration-Seite zwei nebeneinanderstehende Boxen
   *  zur Auswahl. Die Beschriftungen kommen aus splitLabelA/splitLabelB; ist
   *  dort nichts hinterlegt, wird der historische Default
   *  "Durchstarter"/"Funstarter" verwendet (Backward-Compat fuer B2Run-Events
   *  vor v10.20). */
  durchstarterCapacity?: number;
  funstarterCapacity?: number;
  /** v10.20: frei wählbare Beschriftungen fuer die zwei Kapazitäts-Gruppen.
   *  Beispiel: "Vormittag" / "Nachmittag", "Lauf 5 km" / "Lauf 10 km",
   *  "VIP" / "Standard". Bei B2Run-Legacy-Events leer = Default. */
  splitLabelA?: string;
  splitLabelB?: string;
  /** v10.20: Warteliste-Verhalten bei aktiver Split-Capacity.
   *  - false / undefined (Default): zwei getrennte Wartelisten — eine pro
   *    Gruppe, Nachrueck-Logik bleibt typ-bewusst (alter B2Run-Stil).
   *  - true: eine gemeinsame Warteliste — wer zuerst auf der Warteliste war,
   *    rueckt nach, unabhaengig von der Gruppe. Sinnvoll z.B. wenn die zwei
   *    Gruppen nur ein UI-Konstrukt sind aber organisatorisch fluide
   *    behandelt werden. */
  splitSharedWaitlist?: boolean;
  /** v11.25: Reine Display-Reihenfolge der zwei Gruppen-Karten in der
   *  Registrierungs-UI umkehren. Default false — Karte A links / zuerst,
   *  Karte B rechts / zweitens. Mit true wird Karte B zuerst gezeigt
   *  (Group 2 prominent). Aenderung beeinflusst NUR die Anzeige-Reihen-
   *  folge in RegistrationPage und der Kapazitaets-Uebersicht im
   *  AdminCenter — keinerlei Daten-Migration, splitLabelA/B,
   *  durchstarter-/funstarterCapacity und die internen StarterType-
   *  IDs auf den Anmeldungen bleiben unangetastet. */
  splitDisplayOrderReversed?: boolean;
  /** v11.0: Wenn true, koennen sich Teilnehmer ueber „Meine Events" eine
   *  PDF-Datei pro Anmeldung hochladen (z.B. Reisekostenbeleg, Foto-
   *  Einverstaendnis, Zertifikat). Die Datei wird als Item-Attachment an
   *  die Teilnehmerlisten-Zeile in der Subsite gehaengt und ist im Admin-
   *  Center sichtbar / herunterladbar. Default: false. */
  allowAttendeeUpload?: boolean;
  /** v11.0: Optionaler Hinweistext, der dem Teilnehmer im Upload-Bereich
   *  ueber dem File-Input angezeigt wird — z.B. „Bitte lade hier deinen
   *  unterschriebenen Datenschutzbogen hoch". Wenn leer, erscheint nur
   *  ein generischer Default-Text. */
  attendeeUploadHint?: string;
  /** v11.0: Anzeige-Name des Upload-Feldes in „Meine Events" — z.B.
   *  „Reisekostenbeleg" oder „Datenschutz-Erklärung". Default-Fallback
   *  wenn leer: „Dokumenten-Upload" / „Document upload". */
  attendeeUploadLabel?: string;
  /** Seit v6.15: optionale Verknüpfung Starter-Typ → Startblock. Wenn gesetzt,
   *  wird bei der Registrierung automatisch das passende Startblock-Custom-Field
   *  gesetzt (keine Einzel-Auswahl durch den User nötig). */
  durchstarterStartblock?: string;
  funstarterStartblock?: string;
  /** Seit v6.15: optionale Admin-Option. Wenn true, muss der User beim Wählen
   *  von Durchstarter eine Checkbox "Leistungsnachweis vorhanden" bestätigen. */
  durchstarterRequiresProof?: boolean;
  /** v11.80: Wenn true, sehen Teilnehmer im Registrierungsformular ein
   *  Anrede-Dropdown (Frau / Herr / Divers / Keine Angabe). Wenn false /
   *  undefined (Default), wird die Anrede-Auswahl ausgeblendet — viele
   *  Events brauchen keine Anrede. Gespeicherte Anrede ist in dem Fall
   *  ein leerer String. */
  askSalutation?: boolean;
  /** v18.35: Anmeldesprache vorgeben. Wenn gesetzt ('de' | 'en'), wird die
   *  komplette Registrierungsseite (inkl. Disclaimer) in dieser Sprache
   *  angezeigt — unabhängig von der App-Sprache des Users. Leer/undefined =
   *  Anmeldeseite folgt der App-Sprache (Default). */
  registrationLanguage?: 'de' | 'en';
  /** v18.33: Self-Check-in. Wenn true, kann sich jeder angemeldete
   *  Teilnehmer selbst einchecken, indem er den event-spezifischen
   *  QR-Code scannt (statischer QR im PDF ODER rotierender Live-QR am
   *  Eingang). Default false. */
  selfCheckInEnabled?: boolean;
  /** v18.33: Geheimer Token pro Event — dient als Schlüssel für den
   *  statischen Check-in-Link (?action=selfcheckin&token=…) UND als
   *  HMAC-Schlüssel für den rotierenden Live-QR-Code. Wird beim Aktivieren
   *  einmalig generiert und bleibt danach stabil. */
  selfCheckInToken?: string;
  /** v18.33: Optionaler Beginn des Self-Check-in-Zeitfensters (ISO).
   *  Leer = Default „nur am Event-Tag" (Start- bis End-Datum). */
  selfCheckInFrom?: string;
  /** v18.33: Optionales Ende des Self-Check-in-Zeitfensters (ISO).
   *  Leer = Default „nur am Event-Tag". */
  selfCheckInTo?: string;
  /** v11.80: Team-Anmeldung — eine Person meldet ein ganzes Team an.
   *  Default false. Die tatsächliche Multi-Person-Anmelde-Logik folgt
   *  mit v11.82+; aktuell wird die Konfiguration nur persistiert. */
  teamRegistrationEnabled?: boolean;
  /** v11.80: Maximale Teamgröße (Default 0 = nicht gesetzt; UI-Default
   *  beim Aktivieren = 4). Nur relevant wenn teamRegistrationEnabled. */
  teamSize?: number;
  /** v11.80: Wenn true, fragt das Anmeldeformular bei einer Team-
   *  Anmeldung zusaetzlich nach einem Team-Namen (frei waehlbar). */
  askTeamName?: boolean;
  /** v11.81: Beitritts-Modus. Wenn true, kann der Team-Lead unvollständige
   *  Teams anmelden (z.B. 2 von 4 Slots). Wenn false (Default), müssen
   *  immer komplette Teams angemeldet werden. */
  teamPartialAllowed?: boolean;
  /** v11.81: Wenn true, sehen andere Teilnehmer offene Slots in der
   *  Registrierungsseite als „Team mit X freien Plätzen" und können dem
   *  Team beitreten. Die Namen der bereits angemeldeten Mitglieder bleiben
   *  privat. Default false. */
  teamOpenSlotsVisible?: boolean;
  /** v11.81: Wenn true, geht jeder Beitritt zu einem offenen Team erst in
   *  eine Approve-Queue — der Team-Lead bekommt eine Mail mit Bestätigen/
   *  Ablehnen-Buttons. Default false (Beitritt sofort gültig). Nur sinnvoll
   *  wenn teamOpenSlotsVisible aktiv ist. */
  teamJoinRequiresApproval?: boolean;
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
  /** v18.18: Darstellung der Beschreibung (`helpText`). `'tooltip'` (Default,
   *  Backward-Compat): erscheint als „i"-Hover-Box neben dem Label.
   *  `'inline'`: erscheint als nicht-fetter Erklär-Text direkt unter dem
   *  (fetten) Feld-Label. Pro Feld vom Organizer wählbar. */
  helpTextStyle?: 'tooltip' | 'inline';
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
  /** v10.24: Bei aktiver Split-Capacity (DurchstarterCapacity > 0 AND
   *  FunstarterCapacity > 0) kann der Organizer ein Feld auf eine der zwei
   *  Gruppen einschränken. 'A' = nur Gruppe A (intern: Durchstarter), 'B' =
   *  nur Gruppe B (intern: Funstarter). Nicht gesetzt oder 'all' = beide
   *  Gruppen. Beispiel: Pflicht-Checkbox „Leistungsnachweis vorhanden" nur
   *  für Gruppe A einblenden. Bei Events ohne Split-Capacity wirkungslos. */
  onlyForGroup?: 'all' | 'A' | 'B';
  /** v11.94: Nur bei `type === 'checkbox'` — Text, der NEBEN der Checkbox
   *  als klickbarer Bestätigungs-Hinweis angezeigt wird. Default-Fallback:
   *  „Ja, bestätigen" (DE) bzw. „Yes, confirm" (EN). Wird leer = Default. */
  confirmLabel?: string;
  /** v17.20: Englische Übersetzungen — werden nur dann verwendet, wenn auf
   *  Event-Ebene `bilingualFields=true` UND die Locale des Teilnehmers
   *  englisch ist. Pro Feld vier separate Texte: Label, Help-Text (Tooltip),
   *  Confirm-Label (Checkbox-Text) und die Optionen eines Dropdowns. Die
   *  Options-EN-Liste wird positional gemappt — Index i in `options` ist
   *  der DE-Text, derselbe Index in `optionsEn` ist der EN-Text. Fehlt ein
   *  EN-Eintrag, fällt die App still auf den DE-Wert zurück. */
  labelEn?: string;
  helpTextEn?: string;
  confirmLabelEn?: string;
  optionsEn?: string[];
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
  // v11.94: Personal-Info-Card auf der Registration-Page zeigt zusätzlich
  // Department + Mobile (statt Festnetz) read-only aus dem SP-User-Profil.
  department?: string;
  mobilePhone?: string;
}

export interface RoleAssignment {
  id: number;
  userEmail: string;
  userName: string;
  role: UserRole;
  location: string;
  assignedBy: string;
  assignedDate: string;
  /** v18.5: „Power User" ist KEINE eigene Rolle, sondern ein Zusatz-Flag auf
   *  einem Organizer (oder Admin). Power User kennen sich besonders gut aus
   *  und werden auf der Event-Erstellungs-Seite als Hilfe-Ansprechpartner mit
   *  Name + Foto angezeigt. Eine Person hat also genau EINEN DEX_Roles-Eintrag
   *  und kann gleichzeitig Organizer UND Power User sein. */
  isPowerUser?: boolean;
}
