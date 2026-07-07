// DEX Event Platform - Type Definitions
// Author: Eike Brenneisen

export type EventType = 'B2Run' | 'JPMorgan' | 'Other';
// v11.89: 'Under Construction' wurde konsolidiert in das IsFictive-Flag.
// Der „Entwurf"-Zustand eines Events lebt jetzt ausschließlich auf
// IsFictive — EventStatus beschreibt nur noch den Lebenszyklus.
export type EventStatus = 'Active' | 'Completed' | 'Cancelled';
// v23.28: 'No-Show' = war angemeldet/eingeladen, ist aber nicht erschienen
// (vom Check-in-Team beim Event vergeben).
export type RegistrationStatus = 'Angemeldet' | 'QR versendet' | 'Warteliste' | 'Eingecheckt' | 'No-Show' | 'Abgemeldet';
export type Salutation = 'Herr' | 'Frau' | 'Divers' | 'Keine Angabe';

export interface DeloitteEvent {
  id: string;
  eventNumber: number;
  title: string;
  type: EventType;
  status: EventStatus;
  organizers: string[];
  organizerEmails: string[]; // ';'-separiert in SharePoint (OrganizerEmail), hier als Array für Benachrichtigungen
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
   *  als Organizer hinterlegen. Felder bleiben für Backward-Compat. */
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
  /** v26.18: E-Mail eines der Organizer, der als Hauptkontakt markiert ist —
   *  wird auf der Anmeldeseite grün hervorgehoben (ersetzt die separate
   *  Kontakt-Angabe). Leer = kein Organizer hervorgehoben. */
  contactOrganizerEmail?: string;
  /** v26.22: SP-Erstell-Zeitstempel (ISO) — u.a. für die Duplikat-Anzeige. */
  created?: string;
  /** v9.21: Test-Team pro Event — Personen die diesen Event bereits im Entwurfs-
   *  Modus sehen + sich anmelden dürfen. Sie testen das Event durch, bevor der
   *  Organizer das "Entwurf"-Häkchen wegnimmt. Persistenz:
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
  /** v16.4: Vor-aufgelöste E-Mails der DLs aus audienceFilter (lowercase),
   *  beim Event-Save via Graph auf transitive Members aufgelöst. matchesAudience
   *  in EventListPage checkt zusätzlich gegen diese Liste — damit funktioniert
   *  die Sichtbarkeit auch für verschachtelte DLs, die /me/memberOf nicht
   *  zurückliefert. Bei DL-Member-Aenderungen muss der Organizer das Event
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
  /** v24.64: Pflicht-Sub-Event — wenn true, MUSS dieses (Sub-)Event bei der
   *  Anmeldung ausgewählt werden. Pro Sub-Event im Wizard-Schritt „Sub-Events"
   *  einstellbar. Ersetzt das alte globale requireSubEventSelection. */
  mandatoryRegistration?: boolean;
  autoSendQRCode?: boolean; // v9.15 — automatisch QR-Code-Mail nach Anmeldung versenden
  imageUrl?: string;
  subsiteUrl?: string;
  outlookBody: string;
  /** v18.42: Betreff des Outlook-Termins. Leer = Event-Titel (Flow-Fallback). */
  outlookSubject?: string;
  /** v18.44: abweichende Start-/End-Zeit des Outlook-Termins (ISO). Leer =
   *  Event-Start/-Ende (Flow-Fallback). */
  outlookStart?: string;
  outlookEnd?: string;
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
   *  nach erfolgreichem Anlegen geschrieben — und ist tatsächlich das einzige
   *  Feld, das auf Erfolgsbasis gefüllt wird (OutlookEventId nur bei Fehler
   *  ='FAILED'). Wird in der App genutzt, um zu erkennen, ob ein Event einen
   *  Outlook-Termin besitzt — wichtig für die v11.57-Update-Confirm-Logik. */
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
  /** v24.15: Wenn true (nur relevant wenn hideOrganizer=true): NICHT alle
   *  Organizer ausblenden, sondern nur die in hiddenOrganizerEmails. Piggyback
   *  `_hideOrgIndividual`. */
  hideOrganizerIndividualOnly?: boolean;
  /** v24.8 (J): EINZELNE Organizer (E-Mails, lowercase), die auf der
   *  Anmelde-Seite nicht als Ansprechpartner gezeigt werden — Rechte bleiben.
   *  Piggyback `_hiddenOrganizers` in EmailTemplateOverrides. */
  hiddenOrganizerEmails?: string[];
  /** v23.6: Wenn true, sehen Personen mit dem Job-Title „Assistenz" dieses
   *  Event generell — auch wenn der Standort-/Verteiler-Filter sie sonst
   *  ausschließen würde (damit sie stellvertretend anmelden können).
   *  Persistiert als Piggyback `_assistantsCanSee` in EmailTemplateOverrides. */
  assistantsCanSee?: boolean;
  /** v23.14: Nur relevant, wenn `activeFrom` in der Zukunft liegt. true = das
   *  Event ist für reguläre Teilnehmer schon VOR dem Aktivierungszeitpunkt in
   *  der Event-Liste als Vorschau sichtbar (mit Hinweis „Anmeldung ab …", aber
   *  noch nicht buchbar). false/undefined = vor `activeFrom` für reguläre User
   *  komplett unsichtbar (bisheriges Verhalten). Piggyback `_previewBeforeActive`. */
  previewBeforeActive?: boolean;
  /** v23.25: Organizer auf der Anmeldeseite groß darstellen (Foto + Name +
   *  E-Mail + Rolle direkt sichtbar, wie das Hover-Popup) statt klein als Chip
   *  mit Hover. false/undefined = kleiner Chip (bisheriges Verhalten).
   *  Piggyback `_organizerDisplayLarge` in EmailTemplateOverrides. */
  organizerDisplayLarge?: boolean;
  /** v23.19: Optionale, pro Ansicht unterschiedliche Darstellung des
   *  Event-Bildes (Zoom + vertikale Position), damit es in jeder Ansicht gut
   *  sitzt. Leer/undefined = Standard (cover, zentriert) — der Default-Flow ist
   *  „einfach Foto hochladen", ohne hier etwas einzustellen. Piggyback
   *  `_imageDisplay` in EmailTemplateOverrides. zoom: 1..3, posY: 0..100 (%). */
  imageDisplay?: { card?: { zoom: number; posY: number; height?: number }; hero?: { zoom: number; posY: number; height?: number } };
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
  disableRegistrationEmail?: boolean; // v19.21: Wenn true: keine Anmelde-Bestätigung (Master disableEmails sticht weiterhin)
  disableCancellationEmail?: boolean; // v19.21: Wenn true: keine Abmelde-Bestätigung
  autoDeregisterOnDecline?: boolean; // v19.23: Wenn true: Outlook-Absage meldet automatisch vom Event ab (Flow-getrieben)
  /** v26.40: Verhalten, wenn erkannt wird, dass eine angemeldete Person nicht
   *  mehr bei Deloitte ist. 'notify' (Default) = Organizer per Mail informieren
   *  (Organizer entscheidet). 'autoderegister' = beim Öffnen der App durch einen
   *  Organizer automatisch abmelden (mit Modal-Hinweis). */
  inactiveHandling?: 'notify' | 'autoderegister';
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
   *  Hintergrund: bei grossen Events ist Organizer-Spam während der
   *  Hauptanmeldephase nervig — kurz vor dem Event will man aber sehen,
   *  wer kurzfristig noch dazustößt. */
  notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
  notifyOrgRegisterFromDate?: string; // ISO; nur relevant wenn mode='fromDate'

  /** v8.5: Granulare Organizer-Benachrichtigung bei Abmeldungen.
   *  - 'never': Organizer bekommt keine Storno-Bestätigungs-Mails
   *  - 'always': BCC bei jedem Storno
   *  - 'afterDeadline': BCC erst wenn die Abmeldefrist (lastDeregisterDate)
   *    bereits ueberschritten ist — also nur bei "späten" Stornos, die
   *    für den Organizer planungsrelevant sind. */
  notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';

  /** v8.6: Liste explizit ausgeschlossener User-E-Mails (semikolon-separiert
   *  in SP, Array hier). Greift NACH Standortfilter + Mailverteiler/User —
   *  egal warum jemand sonst Sichtbarkeit hätte, wenn seine Mail hier
   *  drin ist, sieht er das Event nicht. Genutzt z.B. wenn aus einem
   *  100er-Verteiler 5 Personen gezielt rausgenommen werden sollen. */
  excludedUsers?: string[];
  isFictive?: boolean;       // Wenn true: Event nur für Admins + eigene Organizer sichtbar (Test-Event)
  /** Geteilte Kapazitäten — bis v10.19 nur für B2Run gedacht (Durchstarter /
   *  Funstarter), seit v10.20 generisch für beliebige Events nutzbar. Wenn
   *  beide > 0 zeigt die Registration-Seite zwei nebeneinanderstehende Boxen
   *  zur Auswahl. Die Beschriftungen kommen aus splitLabelA/splitLabelB; ist
   *  dort nichts hinterlegt, wird der historische Default
   *  "Durchstarter"/"Funstarter" verwendet (Backward-Compat für B2Run-Events
   *  vor v10.20). */
  durchstarterCapacity?: number;
  funstarterCapacity?: number;
  /** v10.20: frei wählbare Beschriftungen für die zwei Kapazitäts-Gruppen.
   *  Beispiel: "Vormittag" / "Nachmittag", "Lauf 5 km" / "Lauf 10 km",
   *  "VIP" / "Standard". Bei B2Run-Legacy-Events leer = Default. */
  splitLabelA?: string;
  splitLabelB?: string;
  /** v26.72: frei konfigurierbare Beschreibung pro Gruppe — erscheint unter
   *  dem Gruppen-Namen in der Auswahl-Karte auf der Anmeldeseite (mehrzeilig). */
  splitDescA?: string;
  splitDescB?: string;
  /** v10.20: Warteliste-Verhalten bei aktiver Split-Capacity.
   *  - false / undefined (Default): zwei getrennte Wartelisten — eine pro
   *    Gruppe, Nachrück-Logik bleibt typ-bewusst (alter B2Run-Stil).
   *  - true: eine gemeinsame Warteliste — wer zuerst auf der Warteliste war,
   *    rückt nach, unabhängig von der Gruppe. Sinnvoll z.B. wenn die zwei
   *    Gruppen nur ein UI-Konstrukt sind aber organisatorisch fluide
   *    behandelt werden. */
  splitSharedWaitlist?: boolean;
  /** v11.25: Reine Display-Reihenfolge der zwei Gruppen-Karten in der
   *  Registrierungs-UI umkehren. Default false — Karte A links / zuerst,
   *  Karte B rechts / zweitens. Mit true wird Karte B zuerst gezeigt
   *  (Group 2 prominent). Aenderung beeinflusst NUR die Anzeige-Reihen-
   *  folge in RegistrationPage und der Kapazitäts-Uebersicht im
   *  AdminCenter — keinerlei Daten-Migration, splitLabelA/B,
   *  durchstarter-/funstarterCapacity und die internen StarterType-
   *  IDs auf den Anmeldungen bleiben unangetastet. */
  splitDisplayOrderReversed?: boolean;
  /** v11.0: Wenn true, können sich Teilnehmer ueber „Meine Events" eine
   *  PDF-Datei pro Anmeldung hochladen (z.B. Reisekostenbeleg, Foto-
   *  Einverständnis, Zertifikat). Die Datei wird als Item-Attachment an
   *  die Teilnehmerlisten-Zeile in der Subsite gehängt und ist im Admin-
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
  /** v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung. Default aus.
   *  confirmDialogMode: 'summary' = Auswahl-Übersicht (Haupt-/Sub-Events mit
   *  De-/Selektieren), 'freetext' = eigener Hinweis-Text. */
  confirmDialogEnabled?: boolean;
  confirmDialogMode?: string;
  confirmDialogText?: string;
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
   *  Leer = Default (v20.3): 2 Stunden vor Event-Start. */
  selfCheckInFrom?: string;
  /** v18.33: Optionales Ende des Self-Check-in-Zeitfensters (ISO).
   *  Leer = Default (v20.3): Event-Ende. */
  selfCheckInTo?: string;
  /** v11.80: Team-Anmeldung — eine Person meldet ein ganzes Team an.
   *  Default false. Die tatsächliche Multi-Person-Anmelde-Logik folgt
   *  mit v11.82+; aktuell wird die Konfiguration nur persistiert. */
  teamRegistrationEnabled?: boolean;
  /** v11.80: Maximale Teamgröße (Default 0 = nicht gesetzt; UI-Default
   *  beim Aktivieren = 4). Nur relevant wenn teamRegistrationEnabled. */
  teamSize?: number;
  /** v11.80: Wenn true, fragt das Anmeldeformular bei einer Team-
   *  Anmeldung zusätzlich nach einem Team-Namen (frei wählbar). */
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
  /** v22.78: Frei benennbarer Team-Begriff (wie Event-Sections), z.B.
   *  „Break-Out Session". Leer = Default „Team". */
  teamTermSingular?: string;
  teamTermPlural?: string;
  /** v22.78: Wenn true, dürfen Teilnehmer KEINE neuen Teams selbst erstellen —
   *  die Zuordnung übernimmt der Organizer (z.B. Break-Out-Sessions). */
  teamMembersCannotCreate?: boolean;
  /** v24.58: Anzeige-Bezeichnung des Haupt-Events in der Sub-Event-Auswahl
   *  (Anmeldeseite). 'default' = „Haupt-Event", 'custom' = freier Text in
   *  mainEventLabel, 'none' = kein Präfix (nur der Event-Titel). Piggyback im
   *  EmailTemplateOverrides-JSON (_mainEventLabel). */
  mainEventLabelMode?: 'default' | 'custom' | 'none';
  mainEventLabel?: string;
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
  correctIndices: number[]; // Indices der richtigen Antworten (0-basiert, mehrere möglich)
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
  // v19.0: 'document' = PDF/Bild-Upload, der als Attachment an die Teilnehmer-
  // Zeile gehängt wird (kein Spaltenwert).
  // v24.25: 'date' = Datums-/Kalender-Auswahl (optional mit Uhrzeit via withTime).
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate' | 'document' | 'date';
  required: boolean;
  options?: string[];
  /** v24.25: Nur für `type === 'date'` — zusätzlich die Uhrzeit abfragen
   *  (datetime-local statt date). */
  withTime?: boolean;
  helpText?: string;
  /** v18.18: Darstellung der Beschreibung (`helpText`). `'tooltip'` (Default,
   *  Backward-Compat): erscheint als „i"-Hover-Box neben dem Label.
   *  `'inline'`: erscheint als nicht-fetter Erklär-Text direkt unter dem
   *  (fetten) Feld-Label. Pro Feld vom Organizer wählbar. */
  helpTextStyle?: 'tooltip' | 'inline';
  spInternalName?: string;
  /** v18.41: Nur für People-Picker-Felder (`type === 'user'` / `'roommate'`).
   *  Wenn true, wird die im Feld ausgewählte Person bei der An-/Abmelde-Mail
   *  des Teilnehmers auf CC gesetzt (z.B. die Assistenz). Betrifft NUR die
   *  E-Mails — NICHT den Outlook-Termin. */
  ccOnEmails?: boolean;
  /** v26.60: Nur für `type === 'roommate'`. Steuert die separate
   *  „Zimmerpartner-Anfrage"-Mail an die ausgewählte Person direkt nach der
   *  Anmeldung. Default (undefined) = true (Bestandsverhalten); explizit
   *  false = keine Benachrichtigungs-Mail an den Roommate. */
  notifyRoommate?: boolean;
  /** Optionale externe Links (AGB, Datenschutz etc.) unter dem Feld */
  externalLinks?: Array<{ label: string; url: string }>;
  /** v7.11: Wenn `type === 'select'` und `multi === true`, wird das Feld als
   *  Mehrfachauswahl gerendert (Checkbox-Liste). Mehrere gewählte Werte werden
   *  in `eventSpecificData[id]` als " | "-getrennter String gespeichert (Pipe
   *  mit Spaces, damit Optionen mit Komma im Label nicht zerrissen werden). */
  multi?: boolean;
  /** v26.74: Nur für `type === 'select'` ohne `multi`. Optionale Vorauswahl —
   *  eine der `options`, die im Anmeldeformular vorausgewählt ist (leer/
   *  undefined = keine Vorauswahl, „Bitte wählen"). */
  defaultValue?: string;
  /** v26.75: Nur für `type === 'select'` ohne `multi`. Vorfilter/Kategorie pro
   *  Option — POSITIONAL zu `options` (gleicher Index). Ist mindestens eine
   *  Kategorie gesetzt, zeigt das Anmeldeformular zuerst ein Kategorie-Dropdown
   *  und filtert die eigentliche Optionsliste darauf (z.B. „Herren"/„Damen" →
   *  nur die passenden Größen). Optionen mit leerer Kategorie sind in jeder
   *  Kategorie sichtbar. */
  optionCategories?: string[];
  /** v26.75: Beschriftung des Vorfilter-Dropdowns (z.B. „Größentabelle"). Leer =
   *  generisches „Kategorie". */
  prefilterLabel?: string;
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

// v26.33: 'IT-Admin' = technischer Admin mit den GLEICHEN Rechten wie 'Admin',
// aber bewusst KEIN Empfänger der Benachrichtigungs-Mails (Ticket-Fragen,
// Wochenbericht, Organizer-Anträge, Inaktiv-Hinweise). Da alle Empfänger-Listen
// exakt auf Role='Admin' filtern, wird ein IT-Admin dort automatisch NICHT
// aufgenommen — er behält aber vollen App-Zugriff (isAdmin = true).
export type UserRole = 'Admin' | 'IT-Admin' | 'Organizer' | 'User';

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
  // v24.29: Unternehmenszugehörigkeit / Rechtsträger („Company name" aus dem
  // M365-Profil, z.B. „Deloitte GmbH" / „Deloitte Consulting").
  company?: string;
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

// ==================== Ticketsystem (v26.0.0) ====================
// Fragen der Nutzer (grüner „Hast du Fragen?"-Header-Button) + Antworten der
// Power-User/Organizer. Persistiert in der globalen Liste DEX_Tickets.
export type TicketStatus = 'Open' | 'InProgress' | 'Closed';
// PowerUser = Frage eines Organizers/Admins (geht an Power-User, Admins in CC).
// Organizer  = Frage eines normalen Users zu einem Event (geht an dessen Organizer).
export type TicketAudience = 'PowerUser' | 'Organizer';

export interface TicketAttachment {
  fileName: string;
  /** Server-relative URL des Anhangs (für <img src>). */
  url: string;
  /** ask = vom Fragesteller angehängt, ans = von der Antwort, other = sonstiges. */
  kind: 'ask' | 'ans' | 'other';
}

// v26.8: Rückfragen-Verlauf NACH der ersten Antwort. Der Fragesteller kann auf
// die Antwort erneut antworten (geht nur an die Person, die geantwortet hat),
// die Antwortende kann erneut antworten oder „keine Antwort nötig" klicken.
export interface TicketFollowUp {
  byEmail: string;
  byName: string;
  /** asker = Rückfrage des Fragestellers, answerer = Folge-Antwort der/des Beantwortenden. */
  byRole: 'asker' | 'answerer';
  text: string;
  at: string;
}

export interface DexTicket {
  id: number;
  title: string;
  /** Eine oder mehrere Fragen (per „+" im Modal). */
  questions: string[];
  status: TicketStatus;
  askerEmail: string;
  askerName: string;
  askerRole: UserRole;
  /** v26.8: Standort + Position des Fragestellers (für Foto-Kontaktkarte). */
  askerLocation: string;
  askerJobTitle: string;
  audience: TicketAudience;
  /** Event-Kontext beim Stellen der Frage (leer = kein Event). */
  eventId: string;
  eventTitle: string;
  /** Bei Audience='Organizer': E-Mails der zuständigen Organizer. */
  assignedOrganizers: string[];
  /** Page-ID / Bildschirm, von dem die Frage gestellt wurde. */
  pageContext: string;
  /** v26.30: 1-basierter Event-Wizard-Schritt, in dem die Frage gestellt wurde
   *  (Organizer im Wizard) — null = nicht im Wizard. */
  askWizardStep: number | null;
  /** v26.60: 'bug' = Bug-Report (Benachrichtigung geht an die DEX-Maintainer
   *  statt an alle Power-User) · 'question' = inhaltliche Frage (Default,
   *  auch für Bestandstickets ohne Category-Spalte). */
  category: 'question' | 'bug';
  answerText: string;
  /** IDs der in der Antwort verlinkten Handbuch-Artikel. */
  answerArticleIds: string[];
  /** 1-basierter Event-Wizard-Schritt, der in der Antwort eingebunden ist (null = keiner). */
  answerWizardStep: number | null;
  /** v26.52: Markierungsbox auf der Live-Wizard-Vorschau des Antwort-Schritts —
   *  „hier klicken". Prozent-Koordinaten relativ zum Vorschau-Container
   *  (x/y = linke obere Ecke, w/h = Größe). null = keine Markierung. */
  answerWizardMarker: { x: number; y: number; w: number; h: number } | null;
  answeredByEmail: string;
  answeredByName: string;
  /** v26.8: Standort + Position der/des Beantwortenden (für Foto-Kontaktkarte). */
  answeredByLocation: string;
  answeredByJobTitle: string;
  answeredAt: string;
  /** Wer das Ticket gerade bearbeitet (Echtzeit „in Bearbeitung"). */
  claimedByEmail: string;
  claimedByName: string;
  claimedAt: string;
  created: string;
  attachments: TicketAttachment[];
  /** v26.8: Rückfragen-Verlauf nach der ersten Antwort. */
  followUps: TicketFollowUp[];
}
