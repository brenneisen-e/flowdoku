/**
 * Event Service - SharePoint-Operationen für Events und Teilnehmerlisten
 *
 * Erstellt DEX_Events-Liste automatisch beim ersten Start.
 * Erstellt pro Event eine Subsite mit einer "Teilnehmer"-Liste.
 *
 * Struktur auf SharePoint:
 *   DOL-c-DE-B2Run (Hauptsite)
 *   ├── DEX_Events (zentrale Event-Liste)
 *   ├── DEX_Roles (Rollenverwaltung)
 *   ├── b2run-frankfurt-2026 (Subsite)
 *   │   └── Teilnehmer (Registrierungsliste)
 *   ├── jpmorgan-muenchen-2026 (Subsite)
 *   │   └── Teilnehmer (Registrierungsliste)
 *   └── ...
 *
 * - Eike, Maerz 2026
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, SPHttpClientConfiguration } from '@microsoft/sp-http';
// v29.48: Alle SharePoint-Requests dieser Klasse laufen über _sp (s.u.) und
// damit durch den 429-Retry.
import { withThrottleRetry } from '../utils/spThrottle';
import * as emailQueue from './events/emailQueue';
import * as profileData from './events/profileData';
import * as outlookQueue from './events/outlookQueue';
import * as hotelPlanning from './events/hotelPlanning';
import * as idReorder from './events/idReorder';
import * as changeLog from './events/changeLog';
import * as emailTemplatesList from './events/emailTemplatesList';
import * as archive from './events/archive';
import * as weeklyReport from './events/weeklyReport';
import * as organizer from './events/organizer';
import * as eventsCrud from './events/eventsCrud';
import * as subsiteProvisioning from './events/subsiteProvisioning';
import * as registrationEdit from './events/registrationEdit';
import * as registration from './events/registration';
import * as registrationStatus from './events/registrationStatus';
import * as eventsListSchema from './events/eventsListSchema';
import * as participantsRegistry from './events/participantsRegistry';
import * as teamJoinRequests from './events/teamJoinRequests';
import * as seats from './events/seats';
import * as overbooking from './events/overbooking';
import * as waitlist from './events/waitlist';
import * as regListRepair from './events/regListRepair';
import * as quiz from './events/quiz';
import * as teilnehmerIdCounter from './events/teilnehmerIdCounter';
import * as eventAssets from './events/eventAssets';
import * as branding from './events/branding';
import * as eventStats from './events/eventStats';
import * as registrationAttachments from './events/registrationAttachments';
import * as accessQueues from './events/accessQueues';
import * as notificationLogs from './events/notificationLogs';
import * as permissionsAudit from './events/permissionsAudit';
import { DexTicket, TicketFollowUp } from '../types';
// v28.95: Erstes nach Thema herausgeloestes Fach-Modul (siehe CLAUDE.md).
import * as tickets from './tickets';
export const REG_LIST_NAME = 'Teilnehmer';

/** v28.61: Je Teilnehmerliste nur einmal pro Sitzung die Hotel-Spalten
 *  anlegen — der Aufruf ist idempotent, aber nicht kostenlos (drei POSTs). */
export const HOTEL_COLS_READY = new Set<string>();
// v30.66: exportiert — die Themen-Module schreiben auf dieselbe Liste.
export const REG_LIST_ITEM_TYPE = 'SP.Data.TeilnehmerListItem';
// v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe (ETag-basiert).
// Pro Subsite eine Liste mit genau einem Item, dessen NextValue beim
// Anmelden via If-Match-Header inkrementiert wird. So können mehrere
// User parallel registrieren ohne dass IDs doppelt vergeben werden.
//
// Listenname mit Unterstrich → SharePoint kodiert das in der Item-Type-
// Bezeichnung als `_x005f_`. Genauso wie wir das schon bei DEX_Events
// machen ('SP.Data.DEX_x005f_EventsListItem'). Ohne dieses Encoding
// schlägt der POST stillschweigend mit HTTP 400 fehl, weil SP den
// Typ-Namen nicht auflösen kann.
// v30.66: exportiert, weil das Zaehler-Thema in services/events/teilnehmerIdCounter.ts
// liegt, die Namen aber auch im Rest des Service vorkommen.
export const COUNTER_LIST_NAME = 'DEX_TeilnehmerCounter';
export const COUNTER_LIST_ITEM_TYPE = 'SP.Data.DEX_x005f_TeilnehmerCounterListItem';

// v11.36: Die Status, die einen Platz belegen (Überbuchungs-Schutz und
// Platz-Zähler rechnen darüber).
// v30.66: war `private static readonly` an der Klasse — jetzt exportiert,
// weil seats.ts und overbooking.ts dieselbe Liste brauchen.
export const ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export interface DeclinedAttendee {
  email: string;
  name: string;
}

// v18.66: Ergebnis-Zusammenfassung des Template-Reseeds, damit die UI dem
// Admin konkret zurückmeldet, was passiert ist (statt nur "erfolgreich").
/** v26.41: Eine Zeile aus dem Kommunikations-Log (DEX_EventComms) — eine
 *  Event-Rundmail (Einladung/Massenmail/Ankündigung) mit Zeitstempel. */
export interface EventCommRow {
  id: number;
  eventId: string;
  subject: string;
  bodyHtml: string;
  emailType: string;
  sentByName: string;
  sentByEmail: string;
  created: string;
}

export interface ReseedSummary {
  created: number;
  updated: number;
  skipped: number;
  failed: number;
  errors: string[];
}

export type DeclineCheckReason = 'no-pointer' | 'not-found' | 'forbidden' | 'error';

// Flaches Ergebnis-Shape, weil der TS-4.7-Compiler bei Discriminated Unions
// über eine async-Funktion + React-Callback teils nicht korrekt narrowed.
// Erfolg: ok=true, attendees gefüllt, reason/message leer.
// Fehler: ok=false, reason gesetzt, attendees leeres Array.
export interface DeclineCheckResult {
  ok: boolean;
  attendees: DeclinedAttendee[];
  reason?: DeclineCheckReason;
  message?: string;
}

export interface SPEvent {
  Id: number;
  Title: string;
  EventStatus: string;
  EventType?: string; // @deprecated seit v5.2 — wird aus CustomFields abgeleitet; Spalte kann aus DEX_Events entfernt werden.
  EventNumber: number;
  Description: string;
  Location: string;
  LocationAddress: string; // JSON-String: { street, houseNo, zip, city }
  LocationFilter: string;
  Audience: string; // Zielgruppen-Filter (Gruppen + Emails, kommasepariert)
  FilterMode: string; // 'AND' | 'OR' - Verknüpfung Standort+Zielgruppe
  StartDate: string;
  EndDate: string;
  RegistrationDeadline: string;
  LastDeregisterDate: string;
  MaxParticipants: number;
  CurrentParticipants?: number; // v26.63: persistierte aktuelle Teilnehmerzahl (von Organizer/Admin gepflegt)
  WaitlistEnabled: boolean;
  MandatoryRegistration?: boolean; // v24.64: Pflicht-Sub-Event (pro Sub-Event)
  EventImageUrl: string;
  Modified?: string; // SP-System-Feld; v26.17 als Cache-Buster für das Event-Bild
  Created?: string;  // SP-System-Feld; v26.22 Erstell-Zeitstempel (Duplikat-Anzeige)
  EmailImageBase64: string; // Base64 Event-Bild für E-Mails/Outlook
  Organizer: string;
  OrganizerEmail: string;
  /** v10.16: optionaler Ansprechpartner (Anzeige-Feld). */
  ContactName?: string;
  ContactEmail?: string;
  ContactOrganizerEmail?: string; // v26.18: als Hauptkontakt markierter Organizer
  ContactInfo?: string;
  OutlookEventId: string;
  CalendarLink: string;
  OutlookBody: string; // Text für den Outlook-Kalendereintrag
  OutlookSubject?: string; // v18.42: Betreff des Outlook-Termins (leer = Event-Titel)
  OutlookStart?: string; // v18.44: abweichende Start-Zeit des Outlook-Termins (ISO, leer = Event-Start)
  OutlookEnd?: string;   // v18.44: abweichende End-Zeit des Outlook-Termins (ISO, leer = Event-Ende)
  OutlookLocation?: string; // v18.34: lesbarer Ort für das Location-Feld des Outlook-Termins
  AllDay?: boolean; // v29.52: ganztägiger Termin — der Outlook-Flow setzt daraus isAllDay
  ShowAsFree?: boolean; // v29.54: Termin als „Frei" zeigen (Flow: showAs). Negativ gespeichert, s. types/index.ts
  SkipOrganizerInvite?: boolean; // v29.55: Organizer nicht als Teilnehmer eintragen (Flow: requiredAttendees)
  OutlookIsOnlineMeeting?: boolean; // v30.26: Flow soll für DIESEN Termin einen Teams-Link erzeugen (isOnlineMeeting)
  EmailLanguage: string; // DE oder EN
  RegistrationLanguage?: string; // v18.35: erzwungene Sprache der Anmeldeseite ('de' | 'en' | '')
  EmailTemplateOverrides: string; // JSON mit Event-spezifischen Template-Anpassungen
  DisableEmails: boolean; // true = keine E-Mails bei An-/Abmeldung
  DisableRegistrationEmail?: boolean; // v19.21: true = keine Anmelde-Bestätigung (Master DisableEmails sticht weiterhin)
  DisableCancellationEmail?: boolean; // v19.21: true = keine Abmelde-Bestätigung
  AutoDeregisterOnDecline?: boolean; // v19.23: true = Outlook-Absage meldet automatisch vom Event ab
  InactiveHandling?: string; // v26.40: 'notify' (Default, Organizer informieren) | 'autoderegister' (automatisch abmelden), wenn eine Person Deloitte verlassen hat
  DisableOutlook: boolean; // true = keine Outlook-Kalendereinträge
  OutlookDirty?: boolean; // v11.57: true = Outlook-relevante Felder geändert, Update an Teilnehmer-Termine noch nicht angestoßen
  AutoSendQRCode?: boolean; // v9.15: true = nach Anmeldung automatisch QR-Code-Mail versenden
  ActiveFrom?: string; // v9.21: ISO-Datum, ab dem ein "Active"-Event tatsächlich sichtbar wird
  // v8.5: Granulare Organizer-BCC-Modi
  NotifyOrgRegisterMode?: string; // 'Never' | 'Always' | 'FromDate'
  NotifyOrgRegisterFromDate?: string;
  NotifyOrgCancelMode?: string; // 'Never' | 'Always' | 'AfterDeadline'
  ExcludedUsers?: string; // v8.6: semikolon-separierte User-Mails die das Event NICHT sehen sollen
  IsFictive?: boolean; // true = Test-Event (nur Admin + eigene Organizer sichtbar)
  DurchstarterCapacity?: number; // Split-Capacity Gruppe A (historisch B2Run-Durchstarter)
  FunstarterCapacity?: number;   // Split-Capacity Gruppe B (historisch B2Run-Funstarter)
  SplitLabelA?: string; // v10.20: frei wählbare Bezeichnung Gruppe A
  SplitLabelB?: string; // v10.20: frei wählbare Bezeichnung Gruppe B
  SplitDescA?: string; // v26.72: Beschreibung Gruppe A (mehrzeilig)
  SplitDescB?: string; // v26.72: Beschreibung Gruppe B (mehrzeilig)
  SplitHelpText?: string; // v26.83: Hinweistext über der Gruppen-Auswahl (mehrzeilig)
  SplitSectionTitle?: string; // v26.83: frei wählbare Überschrift der Gruppen-Auswahl
  SplitSharedWaitlist?: boolean; // v10.20: true = gemeinsame Warteliste, false = getrennt (Default)
  AllowAttendeeUpload?: boolean; // v11.0: Teilnehmer können PDF an ihre Anmeldung hängen
  AttendeeUploadHint?: string;   // v11.0: optionaler Hinweistext über dem Upload-Input
  AttendeeUploadLabel?: string;  // v11.0: Anzeige-Name des Upload-Felds in MyEvents
  AskSalutation?: boolean;       // v11.80: Anrede im Registrierungsformular abfragen
  ConfirmDialogEnabled?: boolean; // v18.75: Sicherheitshinweis vor dem Absenden
  ConfirmDialogMode?: string;     // v18.75: 'summary' | 'freetext'
  ConfirmDialogText?: string;     // v18.75: Freitext-Hinweis
  SelfCheckInEnabled?: boolean;  // v18.33: Self-Check-in per QR-Code erlauben
  SelfCheckInToken?: string;     // v18.33: Geheimer Token (statischer Link + HMAC-Schlüssel rotierender QR)
  SelfCheckInFrom?: string;      // v18.33: optionaler Start des Check-in-Fensters (ISO), leer = Default 2 Std. vor Start (v20.3)
  SelfCheckInTo?: string;        // v18.33: optionales Ende des Check-in-Fensters (ISO)
  TeamRegistrationEnabled?: boolean; // v11.80: Team-Anmeldung erlauben
  TeamSize?: number;             // v11.80: Maximale Teamgröße
  AskTeamName?: boolean;         // v11.80: Team-Namen abfragen
  TeamPartialAllowed?: boolean;       // v11.81: Auch Teil-Teams erlauben (statt nur komplette)
  TeamOpenSlotsVisible?: boolean;     // v11.81: Offene Slots öffentlich sichtbar für Beitritt
  TeamJoinRequiresApproval?: boolean; // v11.81: Beitritt erfordert Bestätigung durch Team-Kapitän
  BilingualFields?: boolean; // v17.20: Custom-Fields zweisprachig (DE + EN) anbieten
  CustomFields: string; // JSON-String mit konfigurierbaren Feldern
  Agenda: string; // JSON-Array mit Agenda-Einträgen
  Transfers: string; // JSON-Array mit Transferzeiten
  Documents: string; // JSON-Array mit Dokumenten
  FunZone: string; // JSON-Array mit Quiz-Fragen
  QuizClusterSize?: number; // 1..4 - wie viele Fragen pro Quiz-Ansicht. Optional, Default 1.
  ParentEventId?: string; // Seit v6.4: wenn gesetzt, ist dies ein Sub-Event und zeigt auf das Parent-Event. Leer = Top-Level-Event.
  RegistrationListName: string;
  SubsiteUrl: string; // Absolute URL der Event-Subsite
}

export interface CustomField {
  id: string;
  label: string;
  // v19.0: document = Datei-Upload (Attachment); v24.25: date = Kalender-Auswahl;
  // v28.63: daterange = Übernachtungs-Zeitraum (Anreise + Abreise, Nächte berechnet)
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate' | 'document' | 'date' | 'daterange';
  required: boolean;
  options?: string[]; // für select-Felder
  /** v24.25: Nur für `type === 'date'` — zusätzlich die Uhrzeit abfragen. */
  withTime?: boolean;
  /** v28.63: Nur für `type === 'daterange'` — buchbares Fenster + Nächte-Limit. */
  rangeStart?: string;
  rangeEnd?: string;
  maxNights?: number;
  visible: boolean;
  /** v7.20: Optionaler Hilfe-/Beschreibungstext, der im Registrierungs-
   *  Formular als "i"-Tooltip neben dem Feld-Label sichtbar ist. */
  helpText?: string;
  /** v18.18: 'tooltip' (Default) oder 'inline' (Erklär-Text unter dem Label). */
  helpTextStyle?: 'tooltip' | 'inline';
  /** v18.41: People-Picker (user/roommate): ausgewählte Person bei An-/Abmelde-
   *  Mail auf CC (nicht im Outlook-Termin). */
  ccOnEmails?: boolean;
  /** v26.60: roommate-Felder — false schaltet die separate
   *  „Zimmerpartner-Anfrage"-Mail ab (undefined = an, Bestandsverhalten). */
  notifyRoommate?: boolean;
  /** v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn das Quell-Feld
   *  einen der `values` als Antwort hat. */
  showIf?: { fieldId: string; values: string[] };
  /** v10.24: Bei Split-Capacity-Events kann ein Feld auf eine der zwei
   *  Gruppen eingeschränkt werden ('A' = Durchstarter, 'B' = Funstarter).
   *  'all'/undefined = für beide Gruppen sichtbar. */
  onlyForGroup?: 'all' | 'A' | 'B';
  /** Optionale externe Links, die unter dem Feld als klickbare Links erscheinen.
   * Aktuell vor allem für B2Run-Zustimmung (AGB + Datenschutz von b2run.de). */
  externalLinks?: Array<{ label: string; url: string }>;
  /** v7.11: Bei type=select erlaubt true Mehrfachauswahl. Wert wird in der
   *  Teilnehmerliste " | "-getrennt gespeichert. */
  multi?: boolean;
  /** v11.94: Nur für type='checkbox' — Text neben der Checkbox (Default
   *  „Ja, bestätigen" / „Yes, confirm"). */
  confirmLabel?: string;
  /** v17.20: Englische Varianten — nur relevant wenn auf Event-Ebene
   *  `bilingualFields=true`. optionsEn ist positional zu options gemappt. */
  labelEn?: string;
  helpTextEn?: string;
  confirmLabelEn?: string;
  optionsEn?: string[];
}

/** v24.41: Eine Zeile aus DEX_AssistantAccess — Verknüpfung Person ↔ Assistenz
 *  + optionale Änderungs-/Abmelde-Anforderung. Enthält KEINE Anmelde-Antworten. */
export interface AssistantLink {
  id: number;
  subsiteUrl: string;
  itemId: number;
  eventId: string;
  eventTitle: string;
  participantEmail: string;
  participantName: string;
  assistantEmail: string;
  assistantName: string;
  ownerEmail: string;
  linkType: string; // 'delegation' | 'proxy'
  status: string;   // 'Active' | 'Cancelled'
  requestType: string;   // '' | 'change' | 'cancel'
  requestNote: string;
  requestedByEmail: string;
  requestedByName: string;
  requestStatus: string; // '' | 'Open' | 'Done' | 'Rejected'
  created: string;
}

export interface SPRegistration {
  Id: number;
  Title: string; // Email
  TeilnehmerID?: number;
  Anrede?: string;
  Vorname?: string;
  Nachname?: string;
  StarterType?: string; // B2Run: 'Durchstarter' | 'Funstarter'
  PreferredStarterType?: string; // B2Run: Wunsch (bei Fallback/Warteliste)
  QuizScore?: number; // Anzahl richtige Antworten
  QuizAnswers?: string; // JSON-Array der gegebenen Antworten
  QuizCompletedAt?: string; // ISO-DateTime
  ParticipantName: string;
  ParticipantEmail: string;
  Status: string;
  /** v11.36: '' = normal, 'Pending' = vom „Überbuchung prüfen"-Lauf als
   *  über Gruppen-/Event-Kapazität markiert; wartet auf Admin-Entscheidung. */
  OverbookReview?: string;
  /** v26.47: Externe stellvertretende Anmeldung — '' = normal/bestätigt,
   *  'Pending' = Datenschutz-Rückmeldung der externen Person steht noch aus
   *  (Anmelder verschickt die Einladung selbst; Mail-Flow kann externe
   *  Adressen nicht erreichen). Anzeige: „Angemeldet (Datenschutzrückmeldung
   *  offen)"; Bestätigung per Button in der Teilnehmerliste. */
  ConsentReview?: string;
  RegistrationDate: string;
  /** v28.38: Hotel-Zuordnung (Name des Hotels aus den Event-Stammdaten). */
  Hotel?: string;
  /** v28.38: An-/Abreise als ISO-DateTime. Die Naechte ergeben sich daraus. */
  HotelFrom?: string;
  HotelTo?: string;
  RegisteredByName?: string;   // Audit: Name des Users der die Anmeldung durchführte
  RegisteredByEmail?: string;  // Audit: E-Mail des Users der die Anmeldung durchführte
  /** v27.12: SP-Item-Metadaten als Fallback für „Registriert am/von", wenn die
   *  Zeile NICHT über die App angelegt wurde (z.B. direkt in der SharePoint-
   *  Listenansicht) und RegistrationDate/RegisteredBy* deshalb leer sind.
   *  Created kommt immer mit; Author nur bei $expand=Author (getAllRegistrations). */
  Created?: string;
  Author?: { Title?: string; EMail?: string };
  CancellationDate: string;
  CancelledByName?: string;    // Audit: Name des Users der die Abmeldung ausgelöst hat
  CancelledByEmail?: string;   // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
  /** v7.16: Check-In-Audit — gesetzt sobald checkInParticipant() aufgerufen wird. */
  CheckedInDate?: string;      // ISO-DateTime, wann der Teilnehmer eingecheckt wurde
  CheckedInByName?: string;    // Name des Helfers, der den Check-In ausgelöst hat
  CheckedInByEmail?: string;   // E-Mail des Helfers, der den Check-In ausgelöst hat
  /** v11.82: Team-Anmeldung — TeamId ist die UUID, die alle Mitglieder eines
   *  gemeinsam angemeldeten Teams gruppiert. TeamLead=true nur für die
   *  anmeldende Person. TeamName ist optional (nur wenn AskTeamName aktiv). */
  TeamId?: string;
  TeamLead?: boolean;
  TeamName?: string;
  /** v11.86: Standort des Teilnehmers (aus Anmeldeformular ableiten oder vom
   *  Profil übernommen). Wird im Team-Badge in „Meine Events" zur
   *  Mitglieder-Identifikation angezeigt. Auf der SP-Liste seit jeher
   *  vorhanden — hier nur als TypeScript-Property nachgezogen. */
  Location?: string;
  /** v24.29: Unternehmenszugehörigkeit / Rechtsträger („Company name" aus dem
   *  Profil, z.B. „Deloitte GmbH" / „Deloitte Consulting"). */
  Company?: string;
  /** v17.15: Nachrück-Audit. PromotedDate ist die ISO-Zeit des Promote
   *  (Warteliste → Angemeldet). ReplacedParticipantEmail ist die E-Mail
   *  der Person, deren Cancel diesen Promote ausgelöst hat („Ersetzt
   *  wen"). ReplacedByParticipantEmail ist auf der cancelnden Person
   *  gesetzt und zeigt die E-Mail der nachrückenden Person („Ersetzt
   *  durch"). Beide Felder Single-Line-Text. */
  PromotedDate?: string;
  ReplacedParticipantEmail?: string;
  ReplacedByParticipantEmail?: string;
  CustomData: string; // JSON mit Custom Field Werten
}

export interface SPParticipant {
  Id: number;
  Title: string; // Email
  Vorname: string;
  Nachname: string;
  Email: string;
  EventRegistered: string; // Kommaseparierte EventNumbers
  EventOnWaitlist: string; // Kommaseparierte EventNumbers
}

// v26.79: Ergebnis der Berechtigungs-Aufräumung (Audit / Cleanup über die
// gesamte Site-Collection). Jede „finding"-Zeile ist eine Abweichung vom
// Soll-Konzept oder eine durchgeführte Korrektur.
export interface PermCleanupFinding {
  /** Menschliche Beschreibung des betroffenen Ortes, z.B. „Liste DEX_Events" oder „Subsite B2Run Köln – Web". */
  scope: string;
  /** Art des Befunds. */
  kind: 'stray-write' | 'ils' | 'inheritance' | 'error';
  /** Betroffene Person (E-Mail/Anzeigename), falls es um eine Einzel-Freigabe geht. */
  principal?: string;
  /** Detailtext (welches Recht, was korrigiert wird/wurde). */
  detail: string;
  /** Im Apply-Lauf: wurde die Korrektur durchgeführt? (im Dry-Run immer false) */
  fixed: boolean;
}

export interface PermCleanupReport {
  apply: boolean;
  websScanned: number;
  listsScanned: number;
  strayWriteFound: number;
  strayWriteRemoved: number;
  ilsIssues: number;
  ilsFixed: number;
  errors: number;
  findings: PermCleanupFinding[];
}

// v26.81: Verwaiste Subsite (existiert als Web, wird aber von KEINEM Event in
// DEX_Events referenziert — z.B. Test-Subsite, deren Event gelöscht wurde).
export interface OrphanSubsite {
  url: string;
  serverRel: string;
  title: string;
  created: string;            // ISO, falls lesbar
  hasParticipantList: boolean; // 'Teilnehmer'-Liste vorhanden → sehr wahrscheinlich Event-Rest
  participantCount: number;    // Anzahl Zeilen in der Teilnehmerliste (0 = leer)
}

export interface OrphanScanResult {
  websScanned: number;
  eventSubsites: number;   // Anzahl von Events referenzierter Subsites
  orphans: OrphanSubsite[];
}

export class EventService {
  public context: WebPartContext;
  public siteUrl: string;
  // v26.81: Pro-Web-Request-Digest-Cache. SPFx spHttpClient injiziert den
  // Digest nur für das AKTUELLE Web — Schreib-Requests (MERGE/DELETE) an ANDERE
  // Webs (Subsites) werden sonst mit 403 „security validation" abgelehnt. Für
  // solche Cross-Web-Schreibzugriffe holen wir den Digest des Ziel-Webs.
  private _digestByWeb: Map<string, string> = new Map();

    /**
   * v29.48/v29.50: Ersatz für `this.context.spHttpClient` — gleiche Signatur.
   *
   * **Nur `post` wiederholt.** Die Drosselung, um die es ging, trifft das
   * SPEICHERN (21 Sub-Events in einem Save). Lesezugriffe hier ebenfalls durch
   * die Schranke zu schicken, war ein Fehler: Der Start der App besteht fast
   * nur aus GETs, und ein einziges 429 legte damit den gesamten Bootvorgang
   * still — die Seite stand bei 8 %, der Browser meldete „reagiert nicht".
   * Ein abgelehnter Lesezugriff war vorher ein fehlendes Stück Anzeige; das
   * ist unschön, aber die App lebt. Deshalb geht `get` wieder direkt raus.
   *
   * Wer neu dazuschreibt, nimmt `this._sp`.
   */
  // v30.6: public — die Themen-Module (services/events/*) nutzen dieselbe
  // Infrastruktur; der Unterstrich bleibt als "intern"-Signal.
  public _sp = {
    get: (url: string, cfg: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse> =>
      this.context.spHttpClient.get(url, cfg, options),
    post: (url: string, cfg: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse> =>
      withThrottleRetry(() => this.context.spHttpClient.post(url, cfg, options), url),
  };

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  /** Web-Basis-URL aus einer Securable-/Listen-API-URL ableiten
   *  (…/_api/web… → Teil vor „/_api/web"). */
  public _webOf(apiBase: string): string {
    const idx = apiBase.indexOf('/_api/web');
    return idx > 0 ? apiBase.slice(0, idx) : this.siteUrl;
  }

  /** Request-Digest (FormDigestValue) des angegebenen Webs holen (gecacht).
   *  Leerer String, wenn nicht ermittelbar. */
  public async _webDigest(webUrl: string): Promise<string> {
    const key = (webUrl || this.siteUrl).toLowerCase().replace(/\/+$/, '');
    const cached = this._digestByWeb.get(key);
    if (cached) return cached;
    try {
      const r = await this._sp.post(
        `${webUrl}/_api/contextinfo`, SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (r.ok) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const d: any = await r.json();
        const val: string = d.FormDigestValue || d.GetContextWebInformation?.FormDigestValue || d.d?.GetContextWebInformation?.FormDigestValue || '';
        if (val) { this._digestByWeb.set(key, val); return val; }
      } else {
        console.warn('[DEX PermFix] contextinfo HTTP', r.status, webUrl);
      }
    } catch (e) { console.warn('[DEX PermFix] contextinfo ERROR', webUrl, e); }
    return '';
  }

  /**
   * v26.87: Setzt ReadSecurity/WriteSecurity einer Liste ZUVERLÄSSIG.
   * Der frühere `odata=verbose`+`__metadata`-MERGE schlug FLÄCHENDECKEND mit
   * HTTP 400 („Bad Request") fehl — auch auf dem aktuellen Web, nicht nur
   * cross-web: SPFx `spHttpClient` erzwingt den Header `odata-version: 3.0`,
   * unter dem die Verbose-`__metadata`-Annotation im Body ungültig ist. Daher
   * jetzt `nometadata` OHNE `__metadata` (wie PnPjs) + der Request-Digest des
   * Ziel-Webs (nötig für Schreibzugriff auf Subsites). Dadurch griff die
   * Element-Sicherheit („nur eigene Elemente") bislang NIE — weder bei der
   * Listen-Erstellung noch beim Admin-Aufräumen. Gibt den HTTP-Status zurück
   * (-1 bei Exception).
   */
  public async _setListSecurity(
    listBase: string,
    values: { ReadSecurity?: number; WriteSecurity?: number },
  ): Promise<number> {
    try {
      const digest = await this._webDigest(this._webOf(listBase));
      const resp = await this._sp.post(
        listBase, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE',
            ...(digest ? { 'X-RequestDigest': digest } : {}),
          },
          body: JSON.stringify(values),
        }
      );
      return resp.status;
    } catch (e) {
      console.warn('[DEX] _setListSecurity ERROR', listBase, e);
      return -1;
    }
  }

  // ==================== DEX_Emails Liste ====================
  // v30.6 (Modularisierung Stufe 2, CLAUDE.md): Implementierung liegt in
  // services/events/emailQueue.ts — hier stehen nur noch Delegations-Stubs
  // mit unveraenderter Signatur, damit KEINE Aufrufstelle angepasst werden
  // musste. Neue Queue-Logik gehoert ins Modul, nicht hierher.

  public async ensureEmailsList(): Promise<void> {
    return emailQueue.ensureEmailsList(this);
  }

  public async setQueueListPermissions(listName: string): Promise<void> {
    return emailQueue.setQueueListPermissions(this, listName);
  }

  public async setEmailsListPermissions(listName: string): Promise<void> {
    return emailQueue.setEmailsListPermissions(this, listName);
  }

  public async queueEmail(
    subject: string,
    recipient: string,
    recipientName: string,
    body: string,
    emailType: string,
    eventTitle: string,
    eventId: string,
    cc?: string,
    bcc?: string,
    importance?: 'High' | 'Normal',
    attachment?: { fileName: string; content: string }
  ): Promise<boolean> {
    return emailQueue.queueEmail(this, subject, recipient, recipientName, body, emailType, eventTitle, eventId, cc, bcc, importance, attachment);
  }

  public async hasQueuedEmail(emailType: string, eventId: string): Promise<boolean> {
    return emailQueue.hasQueuedEmail(this, emailType, eventId);
  }

  public async hasQueuedEmailSince(emailType: string, sinceIso: string): Promise<boolean> {
    return emailQueue.hasQueuedEmailSince(this, emailType, sinceIso);
  }

  public async getParticipantDeletionWarningDates(): Promise<Record<string, string>> {
    return emailQueue.getParticipantDeletionWarningDates(this);
  }


  public async getEventCustomFieldsVersions(itemId: number): Promise<Array<{ created: string; customFields: string }>> {
    return eventsCrud.getEventCustomFieldsVersions(this, itemId);
  }

  // ==================== DEX_Outlook Liste ====================
  // v30.6 (Modularisierung Stufe 2, Tranche 2): Implementierung in
  // services/events/outlookQueue.ts — hier nur Delegations-Stubs.

  public async ensureOutlookList(): Promise<void> {
    return outlookQueue.ensureOutlookList(this);
  }

  // v18.34: Session-Merker des OutlookLocation-Backfills — Instanz-Zustand,
  // das Modul liest/setzt ihn ueber svc.
  public _outlookLocationBackfilled = new Set<string>();

  public async getInvitedRecipients(eventId: string | number, onHttpError?: (_status: number) => void): Promise<string[]> {
    return outlookQueue.getInvitedRecipients(this, eventId, onHttpError);
  }

  public async queueOutlookEvent(attendee: string, eventId: string, eventTitle: string, actionType: 'Einladen' | 'Ausladen' | 'UpdateEvent'): Promise<boolean> {
    return outlookQueue.queueOutlookEvent(this, attendee, eventId, eventTitle, actionType);
  }

  public async queueOutlookDeleteEvent(eventId: string, eventTitle: string, calendarLink: string): Promise<boolean> {
    return outlookQueue.queueOutlookDeleteEvent(this, eventId, eventTitle, calendarLink);
  }

  public async getDeclinedAttendees(eventId: number | string): Promise<DeclineCheckResult> {
    return outlookQueue.getDeclinedAttendees(this, eventId);
  }

  public async checkAccountsActive(emails: string[]): Promise<{ ok: boolean; inactive: string[] }> {
    return outlookQueue.checkAccountsActive(this, emails);
  }

  public async getMyCompanyViaGraph(): Promise<string> {
    return outlookQueue.getMyCompanyViaGraph(this);
  }

  public async getCompaniesByEmails(emails: string[]): Promise<Record<string, string>> {
    return outlookQueue.getCompaniesByEmails(this, emails);
  }

  public async backfillCompanyForList(subsiteUrl: string): Promise<{ updated: number; checked: number }> {
    return outlookQueue.backfillCompanyForList(this, subsiteUrl);
  }


  // ==================== Hotel-Planung / IDReorder / ChangeLog ====================
  // v30.6 (Modularisierung Stufe 2): Implementierungen liegen in
  // services/events/{hotelPlanning,idReorder,changeLog}.ts — hier stehen nur
  // noch Delegations-Stubs mit unveraenderter Signatur.

  public async ensureHotelColumns(subsiteUrl: string): Promise<void> {
    return hotelPlanning.ensureHotelColumns(this, subsiteUrl);
  }

  public async setHotelAssignment(subsiteUrl: string, itemId: number, hotel: string, fromIso: string, toIso: string): Promise<boolean> {
    return hotelPlanning.setHotelAssignment(this, subsiteUrl, itemId, hotel, fromIso, toIso);
  }

  public async ensureIDReorderList(): Promise<void> {
    return idReorder.ensureIDReorderList(this);
  }

  // v18.65: Session-Merker der CancelledName-Spalten-Nachruestung — bleibt
  // als Instanz-Zustand an der Klasse, das Modul liest/setzt ihn ueber svc.
  public _idReorderCancelledFieldEnsured = false;

  public async queueIDReorder(
    eventId: string,
    eventNumber: number,
    subsiteUrl: string,
    eventTitle: string,
    cancelledName?: string,
    cancelledEmail?: string
  ): Promise<boolean> {
    return idReorder.queueIDReorder(this, eventId, eventNumber, subsiteUrl, eventTitle, cancelledName, cancelledEmail);
  }

  public async ensureChangeLogList(): Promise<void> {
    return changeLog.ensureChangeLogList(this);
  }

  public async writeChangeLog(entry: changeLog.ChangeLogEntryInput): Promise<void> {
    return changeLog.writeChangeLog(this, entry);
  }

  public async readChangeLog(opts?: { eventId?: string; top?: number }): Promise<changeLog.ChangeLogRow[]> {
    return changeLog.readChangeLog(this, opts);
  }


  // ==================== DEX_EmailTemplates Liste ====================
  // v30.11: Thema ausgelagert nach services/events/emailTemplatesList.ts —
  // hier stehen nur noch Delegations-Stubs mit unveränderter Signatur.

  public async ensureEmailTemplatesList(): Promise<void> {
    return emailTemplatesList.ensureEmailTemplatesList(this);
  }

  public async reseedDefaultEmailTemplates(): Promise<ReseedSummary> {
    return emailTemplatesList.reseedDefaultEmailTemplates(this);
  }

  public async getKpiCache(): Promise<{ participants: number; events: number } | null> {
    return emailTemplatesList.getKpiCache(this);
  }

  public async getFAConfig(): Promise<{ infoRecipients: string[]; listRecipients: string[]; log: Array<{ ts: string; by: string; action: string; old?: string; neu?: string }> }> {
    return emailTemplatesList.getFAConfig(this);
  }

  public async saveFAConfig(cfg: { infoRecipients: string[]; listRecipients: string[]; log: Array<{ ts: string; by: string; action: string; old?: string; neu?: string }> }): Promise<boolean> {
    return emailTemplatesList.saveFAConfig(this, cfg);
  }

  public async bumpKpiParticipants(delta: number): Promise<number | null> {
    return emailTemplatesList.bumpKpiParticipants(this, delta);
  }

  public async bumpKpiEvents(delta: number): Promise<number | null> {
    return emailTemplatesList.bumpKpiEvents(this, delta);
  }

  public async recomputeEventKpiOnly(): Promise<number | null> {
    return emailTemplatesList.recomputeEventKpiOnly(this);
  }

  public async persistCurrentParticipants(eventId: number, count: number): Promise<boolean> {
    return emailTemplatesList.persistCurrentParticipants(this, eventId, count);
  }

  public async updateKpiCache(values: { participants: number; events: number }): Promise<boolean> {
    return emailTemplatesList.updateKpiCache(this, values);
  }

  public async getAllEventsForKpi(): Promise<Array<{ id: number; parentEventId: string; status: string; subsiteUrl: string; isFictive: boolean }>> {
    return emailTemplatesList.getAllEventsForKpi(this);
  }

  public async getKpiTotals(): Promise<{ participants: number; events: number } | null> {
    return emailTemplatesList.getKpiTotals(this);
  }

  public async getParticipantsListCount(): Promise<number | null> {
    return emailTemplatesList.getParticipantsListCount(this);
  }

  public async getAppViewCount(): Promise<number | null> {
    return emailTemplatesList.getAppViewCount(this);
  }

  public async incrementAppViewCount(): Promise<number | null> {
    return emailTemplatesList.incrementAppViewCount(this);
  }

  public async ensureLogosInConfig(): Promise<void> {
    return emailTemplatesList.ensureLogosInConfig(this);
  }

  public async getTestTeamEmails(): Promise<string[]> {
    return emailTemplatesList.getTestTeamEmails(this);
  }

  public async setTestTeamEmails(emails: string[]): Promise<boolean> {
    return emailTemplatesList.setTestTeamEmails(this, emails);
  }

  public async getEmailTemplate(templateType: string, language: string = 'EN'): Promise<{ subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string } | null> {
    return emailTemplatesList.getEmailTemplate(this, templateType, language);
  }

  public async getAllEmailTemplates(): Promise<Array<{ id: number; templateType: string; language: string; subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string }>> {
    return emailTemplatesList.getAllEmailTemplates(this);
  }

  public async updateEmailTemplate(id: number, fields: { subject?: string; heading?: string; subheading?: string; headingColor?: string; bodyHtml?: string }): Promise<boolean> {
    return emailTemplatesList.updateEmailTemplate(this, id, fields);
  }

  // v30.11: war privater Helfer der Sektion, wird aber auch vom
  // Logo-&-Branding-Block (v26.50) gebraucht — bleibt deshalb als Stub public.
  public async loadFileAsBase64(path: string): Promise<string> {
    return emailTemplatesList.loadFileAsBase64(this, path);
  }

  // ==================== DEX_Participants (Personen-Register) ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/participantsRegistry.ts — hier nur Delegations-Stubs.

  public async ensureParticipantsList(): Promise<void> {
    return participantsRegistry.ensureParticipantsList(this);
  }

  public async getParticipantByEmail(email: string, onHttpError?: (_status: number) => void): Promise<SPParticipant | null> {
    return participantsRegistry.getParticipantByEmail(this, email, onHttpError);
  }

  public async upsertParticipant(
    vorname: string,
    nachname: string,
    email: string,
    eventNumber: number,
    status: string // 'Angemeldet' | 'Warteliste'
  ): Promise<boolean> {
    return participantsRegistry.upsertParticipant(this, vorname, nachname, email, eventNumber, status);
  }

  public async analyzeRegistryAgainstLists(
    events: Array<{ eventNumber?: number; title: string; subsiteUrl?: string }>,
    onProgress?: (_done: number, _total: number, _title: string) => void,
    onRead?: (_loaded: number) => void,
  ): Promise<{
    checkedEvents: number; skippedEvents: number;
    stale: Array<{ email: string; eventNumber: number; title: string }>;
    /** Je geprüftem Event: wie viele Verweise, wie viele davon ohne Zeile,
     *  wie viele aktive Zeilen die Liste überhaupt hat. Das ist die Grundlage
     *  für die Plausibilitäts-Prüfung unten — und für die Frage, WARUM etwas
     *  auseinanderläuft. */
    perEvent: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number; suspicious: boolean; listGone: boolean }>;
    /** Events, bei denen (fast) ALLE Verweise ins Leere zeigen, OBWOHL ihre
     *  Teilnehmerliste lesbar ist. Ihre Verweise stehen NICHT in `stale` —
     *  siehe Begründung unten. */
    suspiciousEvents: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number }>;
    /** v29.3: Events, deren Teilnehmerliste NICHT MEHR EXISTIERT (HTTP 404 —
     *  in aller Regel das 3-Monats-Löschkonzept, das die Subsite recycelt und
     *  das Event-Item stehen lässt). Ihre Verweise sind eindeutig Rückstand
     *  und stehen in `stale`. */
    deletedListEvents: Array<{ title: string; eventNumber: number; referenced: number }>;
  }> {
    return participantsRegistry.analyzeRegistryAgainstLists(this, events, onProgress, onRead);
  }

  public async collectOrphanRegistryNumbers(
    onRead?: (_loaded: number) => void,
  ): Promise<Array<{ email: string; eventNumber: number }>> {
    return participantsRegistry.collectOrphanRegistryNumbers(this, onRead);
  }

  public async pruneStaleRegistryNumbers(
    stale: Array<{ email: string; eventNumber: number }>,
    onProgress?: (_done: number, _total: number) => void,
  ): Promise<{ updated: number; removed: number; failed: number }> {
    return participantsRegistry.pruneStaleRegistryNumbers(this, stale, onProgress);
  }

  public async removeParticipantEvent(email: string, eventNumber: number): Promise<boolean> {
    return participantsRegistry.removeParticipantEvent(this, email, eventNumber);
  }

  public async getAllParticipants(): Promise<SPParticipant[]> {
    return participantsRegistry.getAllParticipants(this);
  }

  public async analyzeParticipantRegistry(
    validEventNumbers: number[],
    onRead?: (loaded: number) => void,
  ): Promise<{
    total: number; duplicateGroups: number; surplusRecords: number; orphanNumbers: number; noEmail: number;
  }> {
    return participantsRegistry.analyzeParticipantRegistry(this, validEventNumbers, onRead);
  }

  public async mergeDuplicateParticipants(
    onProgress?: (done: number, total: number) => void,
    onRead?: (loaded: number) => void,
  ): Promise<{ groups: number; deleted: number; failed: number }> {
    return participantsRegistry.mergeDuplicateParticipants(this, onProgress, onRead);
  }

  public async fetchAllParticipantsOrThrow(onPage?: (loaded: number) => void): Promise<SPParticipant[]> {
    return participantsRegistry.fetchAllParticipantsOrThrow(this, onPage);
  }

  public async backfillParticipantRegistry(
    subsiteUrl: string,
    eventNumber: number,
    onProgress?: (done: number, total: number) => void,
  ): Promise<{ active: number; fixed: number; failed: number }> {
    return participantsRegistry.backfillParticipantRegistry(this, subsiteUrl, eventNumber, onProgress);
  }

  // ==================== DEX_Events: Schema, Migrationen, Ansicht ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/eventsListSchema.ts — hier nur Delegations-Stubs.
  // `getEventsFieldDefinitions` und `_upgradeTextFieldToNote` sind public,
  // weil sie auch ausserhalb des Themas gebraucht werden.

  public async ensureEventsList(): Promise<void> {
    return eventsListSchema.ensureEventsList(this);
  }

  public getEventsFieldDefinitions(): Array<{ title: string; type: number; choices?: string[]; metaType?: string; richText?: boolean; numberOfLines?: number }> {
    return eventsListSchema.getEventsFieldDefinitions();
  }

  public async upgradeAudienceFieldToNote(): Promise<void> {
    return eventsListSchema.upgradeAudienceFieldToNote(this);
  }

  public async upgradeOrganizerFieldsToNote(): Promise<void> {
    return eventsListSchema.upgradeOrganizerFieldsToNote(this);
  }

  public async upgradeOverflowTextFieldsToNote(): Promise<void> {
    return eventsListSchema.upgradeOverflowTextFieldsToNote(this);
  }

  public async userHasSiteAccess(email: string): Promise<boolean | null> {
    return eventsListSchema.userHasSiteAccess(this, email);
  }

  public async _upgradeTextFieldToNote(listName: string, fieldName: string): Promise<void> {
    return eventsListSchema._upgradeTextFieldToNote(this, listName, fieldName);
  }

  public async configureDefaultView(listName: string, fieldNames: string[], baseUrl?: string, opts?: { rebuild?: boolean }): Promise<void> {
    return eventsListSchema.configureDefaultView(this, listName, fieldNames, baseUrl, opts);
  }

  /**
   * Strip SharePoint-Note-Field-Wrapper.
   *
   * Seit der Note-Migration wickelt SP die Werte für `Organizer` + `OrganizerEmail`
   * beim REST-Read in `<div class="ExternalClassXXXX">…</div>`. Vor dem Splitten
   * via `;` muss der Wrapper raus, sonst landen die Tag-Reste in den
   * Email-Listen → falsche Match-Vergleiche, kaputte Permissions.
   *
   * Idempotent: Werte ohne Wrapper bleiben unverändert.
   */
  // v30.7: public — profileData-Modul nutzt den Wrapper-Strip ebenfalls.
  public static stripNoteWrapper(value: string | null | undefined): string {
    if (!value) return '';
    let v = value.trim();
    v = v.replace(/^<div\b[^>]*>/i, '');
    v = v.replace(/<\/div>\s*$/i, '');
    return v.trim();
  }

  /** v26.51: Klartext-Grund des letzten fehlgeschlagenen updateEvent-Aufrufs —
   *  wird dem Organizer in der Fehlermeldung angezeigt (vorher nur Konsole).
   *  v30.66: Instanz-Zustand bleibt an der Klasse, das Thema liegt im Modul. */
  public lastUpdateEventError = '';

  // ==================== Events lesen, anlegen, ändern, löschen ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/eventsCrud.ts — hier nur Delegations-Stubs.

  public async seedEvents(): Promise<void> {
    return eventsCrud.seedEvents(this);
  }

  public async getEvents(): Promise<SPEvent[]> {
    return eventsCrud.getEvents(this);
  }

  public async getEvent(eventId: number): Promise<SPEvent | null> {
    return eventsCrud.getEvent(this, eventId);
  }

  public async getChildEventIds(parentId: number): Promise<number[] | null> {
    return eventsCrud.getChildEventIds(this, parentId);
  }

  public async getEventBySelfCheckInToken(token: string): Promise<SPEvent | null> {
    return eventsCrud.getEventBySelfCheckInToken(this, token);
  }

  public async getEventByEventNumber(eventNumber: number): Promise<SPEvent | null> {
    return eventsCrud.getEventByEventNumber(this, eventNumber);
  }

  public async createEvent(event: {
    title: string;
    status: string;
    type: string;
    description: string;
    location: string;
    locationAddress?: string; // JSON-String: { street, houseNo, zip, city }
    outlookSubject?: string; // v18.42: Betreff des Outlook-Termins (leer = Titel)
    outlookStart?: string; // v18.44: abweichende Start-Zeit (ISO, leer = Event-Start)
    outlookEnd?: string;   // v18.44: abweichende End-Zeit (ISO, leer = Event-Ende)
    outlookLocation?: string; // v18.40: manueller Outlook-Ort (leer = Auto aus Ort + Adresse)
    allDay?: boolean; // v29.52: ganztägiger Termin (Flow setzt daraus isAllDay)
    showAsFree?: boolean; // v29.54: Termin als „Frei" anzeigen (Flow: showAs)
    skipOrganizerInvite?: boolean; // v29.55: Organizer nicht einladen (Flow: requiredAttendees)
    outlookIsOnlineMeeting?: boolean; // v30.26: Termin als Teams-Besprechung anlegen (Flow: isOnlineMeeting)
    locationFilter: string;
    audience: string;
    /** v16.4: Vor-aufgelöste E-Mails der Audience-DLs, ';'-separiert, lowercase. */
    audienceResolvedEmails?: string;
    filterMode: string;
    startDate: string;
    endDate: string;
    registrationDeadline: string;
    lastDeregisterDate: string;
    /** v29.19: Auto-Aktivierungszeitpunkt (UTC-ISO). Der Wizard bot das Feld
     *  auch beim ANLEGEN an, persistiert wurde es aber nur im Edit-Pfad —
     *  ein als Entwurf angelegtes Event mit „Aktiv ab" ging nie von allein
     *  live. */
    activeFrom?: string;
    maxParticipants: number;
    waitlistEnabled: boolean;
    mandatoryRegistration?: boolean; // v24.64: Pflicht-Sub-Event

    eventImageUrl: string;
    organizer: string;
    organizerEmail: string;
    /** v10.16: optionaler Ansprechpartner (Anzeige-Feld). */
    contactName?: string;
    contactEmail?: string;
    contactOrganizerEmail?: string;
    contactInfo?: string;
    outlookEventId: string;
    outlookBody: string;
    agenda?: string; // JSON-Array mit Agenda-Einträgen
    transfers?: string; // JSON-Array mit Transferzeiten
    documents?: string; // JSON-Array mit Dokumenten
    funZone?: string; // JSON-Array mit Quiz-Fragen
    quizClusterSize?: number; // 1..4 - Fragen pro Quiz-Ansicht
    /** Seit v6.4: wenn gesetzt, wird dieses Event als Sub-Event angelegt und zeigt auf das angegebene Parent-Event. */
    parentEventId?: string;
    emailLanguage?: string;
    registrationLanguage?: 'de' | 'en';
    emailTemplateOverrides?: string;
    disableEmails?: boolean;
    disableRegistrationEmail?: boolean;
    disableCancellationEmail?: boolean;
    autoDeregisterOnDecline?: boolean;
    inactiveHandling?: string;
    disableOutlook?: boolean;
    notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
    notifyOrgRegisterFromDate?: string;
    notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';
    excludedUsers?: string[];
    isFictive?: boolean;
    durchstarterCapacity?: number;
    funstarterCapacity?: number;
    splitLabelA?: string;
    splitLabelB?: string;
    splitDescA?: string;
    splitDescB?: string;
    splitHelpText?: string;
    splitSectionTitle?: string;
    splitSharedWaitlist?: boolean;
    allowAttendeeUpload?: boolean;
    attendeeUploadHint?: string;
    attendeeUploadLabel?: string;
    /** v11.80: Anrede im Registrierungsformular abfragen (Default false). */
    askSalutation?: boolean;
    /** v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung. */
    confirmDialogEnabled?: boolean;
    confirmDialogMode?: string; // 'summary' | 'freetext'
    confirmDialogText?: string;
    /** v18.33: Self-Check-in per QR-Code erlauben (Default false). */
    selfCheckInEnabled?: boolean;
    /** v18.33: Geheimer Token (statischer Link + HMAC-Schlüssel rotierender QR). */
    selfCheckInToken?: string;
    /** v18.33: optionaler Start des Check-in-Fensters (ISO). */
    selfCheckInFrom?: string;
    /** v18.33: optionales Ende des Check-in-Fensters (ISO). */
    selfCheckInTo?: string;
    /** v11.80: Team-Anmeldung erlauben (Default false). */
    teamRegistrationEnabled?: boolean;
    /** v11.80: Maximale Teamgröße (0 = nicht gesetzt). */
    teamSize?: number;
    /** v11.80: Team-Name abfragen (Default false). */
    askTeamName?: boolean;
    /** v11.81: Auch Teil-Teams zulassen (Default false = nur komplette Teams). */
    teamPartialAllowed?: boolean;
    /** v11.81: Offene Slots öffentlich für Beitritt sichtbar (Default false). */
    teamOpenSlotsVisible?: boolean;
    /** v11.81: Beitritt erfordert Bestätigung durch Team-Kapitän (Default false). */
    teamJoinRequiresApproval?: boolean;
    /** v17.20: Custom-Fields zweisprachig anbieten (DE+EN). */
    bilingualFields?: boolean;
    customFields: CustomField[];
    /** v11.69: Wenn `existingSubsiteUrl` UND `existingRegistrationListName`
     *  gesetzt sind, wird KEINE neue Subsite und KEINE neue Teilnehmer-
     *  liste angelegt — stattdessen werden die mitgegebenen Werte direkt in
     *  das neue DEX_Events-Item geschrieben. Hintergrund: Outlook-Termin
     *  nachträglich aktivieren ohne Verlust bestehender Anmeldungen — das
     *  Sub-Event wird mit `deleteEventItemOnly()` aus DEX_Events entfernt
     *  und hier neu angelegt, wobei die alte Subsite + Teilnehmerliste
     *  unangetastet bleiben und an die neue Event-Zeile angehängt werden.
     *  Damit triggert der `DEX_CreateOutlookEvent`-Flow (GetOnNewItems) auf
     *  dem neuen Item und legt den Outlook-Termin an. */
    existingSubsiteUrl?: string;
    existingRegistrationListName?: string;
    /** v11.87: Optionaler Progress-Callback. Wird zu Beginn jeder Teil-
     *  Operation aufgerufen — die UI kann darauf den Fortschrittsbalken
     *  und die Unter-Caption sichtbar bewegen, statt minutenlang auf
     *  „Event wird vorbereitet..." stehen zu bleiben. Stages decken
     *  die langsamen SP-Operationen ab (Subsite-Create, Listen-Create,
     *  Permissions, Counter, Views). */
    onProgress?: (stage:
      | 'start'
      | 'subsite-creating'
      | 'subsite-done'
      | 'permissions'
      | 'list-creating'
      | 'list-done'
      | 'item-insert'
      | 'done'
    ) => void;
  }): Promise<number | null> {
    return eventsCrud.createEvent(this, event);
  }

  public async markExpiredEventsAsCompleted(): Promise<number> {
    return eventsCrud.markExpiredEventsAsCompleted(this);
  }

  public async getEventCustomFieldsHistory(eventId: number): Promise<Array<{
    versionLabel: string;
    modified: string;
    customFields: Array<Record<string, unknown>>;
  }>> {
    return eventsCrud.getEventCustomFieldsHistory(this, eventId);
  }

  public async updateEvent(eventId: number, updates: Record<string, unknown>, retried?: boolean): Promise<boolean> {
    return eventsCrud.updateEvent(this, eventId, updates, retried);
  }

  public async deleteEvent(eventId: number): Promise<boolean> {
    return eventsCrud.deleteEvent(this, eventId);
  }

  // ==================== Event-Statistik & Datenlöschung ====================
  // v26.32: Löschkonzept — Teilnehmerliste 3 Monate nach Event-Ende löschen;
  // Event + KPIs bleiben im Statistik-Archiv (DEX_EventStats) erhalten.
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/eventStats.ts — hier nur Delegations-Stubs.

  public async ensureEventStatsList(): Promise<void> {
    return eventStats.ensureEventStatsList(this);
  }

  public async getArchivedStatsEventNumbers(): Promise<Set<number>> {
    return eventStats.getArchivedStatsEventNumbers(this);
  }

  public async getEventStats(): Promise<Array<{
    id: number; eventNumber: number; eventTitle: string; eventType: string;
    location: string; startDate: string; endDate: string; maxParticipants: number | null;
    registeredCount: number; qrSentCount: number; checkedInCount: number;
    noShowCount: number; waitlistCount: number; deregisteredCount: number;
    organizer: string; archivedByEmail: string; archivedDate: string;
  }>> {
    return eventStats.getEventStats(this);
  }

  public async archiveEventStats(meta: {
    eventNumber: number; eventTitle: string; eventType?: string; location?: string;
    startDate?: string; endDate?: string; maxParticipants?: number; subsiteUrl?: string;
    organizer?: string;
  }): Promise<boolean> {
    return eventStats.archiveEventStats(this, meta);
  }

  public async deleteEventStatsRow(eventNumber: number): Promise<boolean> {
    return eventStats.deleteEventStatsRow(this, eventNumber);
  }

  public async deleteParticipantData(eventId: number): Promise<boolean> {
    return eventStats.deleteParticipantData(this, eventId);
  }

  public async deleteEventItemOnly(eventId: string | number): Promise<boolean> {
    return eventStats.deleteEventItemOnly(this, eventId);
  }

  // ==================== Event-Subsite anlegen und berechtigen ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/subsiteProvisioning.ts — hier nur Delegations-Stubs.
  // Die vier bisher privaten Helfer sind public, weil sie ausserhalb des
  // Themas (Event-CRUD, Anmeldung) über `svc` aufgerufen werden.

  public async createEventSubsite(title: string, description: string): Promise<string | null> {
    return subsiteProvisioning.createEventSubsite(this, title, description);
  }

  public async createRegistrationList(
    subsiteUrl: string,
    customFields: CustomField[],
    organizerEmail: string
  ): Promise<Record<string, string>> {
    return subsiteProvisioning.createRegistrationList(this, subsiteUrl, customFields, organizerEmail);
  }

  public async ensureOrganizerPermissions(subsiteUrl: string, organizerEmails: string): Promise<void> {
    return subsiteProvisioning.ensureOrganizerPermissions(this, subsiteUrl, organizerEmails);
  }

  // v30.67: Rueckgabetyp aus dem Modul durchreichen — der alte Inline-Typ blendete
  // `failed` aus, und die Aufrufer konnten fehlgeschlagene Rechtevergaben nicht
  // anzeigen (der Fund: "meldet Erfolg, ohne eine Antwort zu pruefen").
  public async ensureOrganizerPermissionsMulti(
    subsiteUrls: string[],
    organizerEmails: string
  ): Promise<subsiteProvisioning.OrganizerPermissionsResult> {
    return subsiteProvisioning.ensureOrganizerPermissionsMulti(this, subsiteUrls, organizerEmails);
  }

  public async setSubsitePermissions(subsiteUrl: string, organizerEmail: string): Promise<void> {
    return subsiteProvisioning.setSubsitePermissions(this, subsiteUrl, organizerEmail);
  }

  public async assignRegistrationToAssistant(
    subsiteUrl: string,
    itemId: number,
    assistantEmail: string,
    assistantName: string
  ): Promise<boolean> {
    return subsiteProvisioning.assignRegistrationToAssistant(this, subsiteUrl, itemId, assistantEmail, assistantName);
  }

  public async trySetItemAuthor(subsiteUrl: string, listName: string, itemId: number, participantEmail: string): Promise<void> {
    return subsiteProvisioning.trySetItemAuthor(this, subsiteUrl, listName, itemId, participantEmail);
  }

  public async repairProxyRegistrationAccess(
    subsiteUrl: string,
    onProgress?: (done: number, total: number) => void
  ): Promise<{ ilsWasWrong: boolean; ilsFixed: boolean; itemsTotal: number; proxyFound: number; authorFixed: number; authorFailed: number }> {
    return subsiteProvisioning.repairProxyRegistrationAccess(this, subsiteUrl, onProgress);
  }

  // ==================== Anmelden (inkl. Team-Anmeldung) ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/registration.ts — hier nur Delegations-Stubs.

  public async registerForEvent(
    subsiteUrl: string,
    firstName: string,
    surname: string,
    participantEmail: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string>, // cf.id -> SP InternalName
    starterType?: string, // B2Run: effektiver Typ (nach Fallback)
    preferredStarterType?: string, // B2Run: Wunsch-Typ (was der User eigentlich wollte)
    registeredByName?: string, // Audit: Name des Users der die Anmeldung auslöst
    registeredByEmail?: string, // Audit: E-Mail des Users der die Anmeldung auslöst
    // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung — wird in
    // die SP-Spalte ProxyConsent geschrieben (leer bei Selbst-Anmeldung).
    proxyConsent?: string,
    // v19.9: Der CLIENT hat bereits zuverlässig festgestellt, dass der/die
    // Anmeldende Haupt- oder Co-Organizer dieses Events ist (aus dem geladenen
    // Event-Objekt: organizerEmails/coOrganizerEmails ⊇ aktueller User). Diese
    // Information wird hier durchgereicht und hat Vorrang vor der fragilen
    // serverseitigen Ableitung (SubsiteUrl-Filter + Note-Feld-Parsing +
    // pageContext-Identität), die im Tenant gelegentlich fehlschlug und
    // legitime Organizer mit „bereits angemeldet" ablehnte.
    actorIsEventOrganizer: boolean = false,
    // v23.10: Der CLIENT hat bereits validiert, dass der/die Anmeldende eine
    // Assistenz ist UND das Ziel Partner/Director (RegistrationPage prüft das
    // beim Submit gegen den Picker-JobTitle). Diesem Flag vertrauen wir — die
    // serverseitige Ableitung in canRegisterForOthers hängt an
    // Profil-Lookups (Title/SPS-JobTitle), die im Tenant unzuverlässig leer
    // zurückkamen und legitime Assistenzen mit „nicht berechtigt" ablehnten.
    // Gilt NUR für Check A (Berechtigung) — NICHT für die Deadline (Assistenz
    // darf wie ein normaler User nicht nach Frist anmelden).
    clientAssistantAllowed: boolean = false
  // v23.9: Statt nacktem boolean ein konkreter Grund bei Misserfolg, damit die
  // UI nicht mehr pauschal „bereits registriert" anzeigt (irreführend, wenn der
  // echte Grund Berechtigung/Deadline/Insert-Fehler war).
    // v30.58: `detail` trägt die Klartext-Antwort von SharePoint bei einem
    // abgelehnten Insert (z.B. „The field or property 'X' does not exist") —
    // der `reason` bleibt maschinenlesbar, die Ursache geht nicht verloren.
  ): Promise<{ ok: boolean; reason?: 'not-allowed' | 'deadline' | 'insert-failed' | 'error'; detail?: string }> {
    return registration.registerForEvent(this, subsiteUrl, firstName, surname, participantEmail, customData, status, customFieldMap, starterType, preferredStarterType, registeredByName, registeredByEmail, proxyConsent, actorIsEventOrganizer, clientAssistantAllowed);
  }

  public async registerTeamMember(
    subsiteUrl: string,
    args: {
      firstName: string;
      lastName: string;
      email: string;
      profile: { department: string; location: string; jobTitle: string; phone: string; company?: string };
      status: 'Angemeldet' | 'Warteliste';
      teamId: string;
      teamLead: boolean;
      teamName?: string;
      customData?: Record<string, string>;
      customFieldMap?: Record<string, string>;
      starterType?: string;
      preferredStarterType?: string;
      registeredByName?: string;
      registeredByEmail?: string;
      salutation?: string;
    }
  ): Promise<{ ok: boolean; teilnehmerId?: number; itemId?: number }> {
    return registration.registerTeamMember(this, subsiteUrl, args);
  }

  public async getTeamMembers(subsiteUrl: string, teamId: string): Promise<SPRegistration[]> {
    return registration.getTeamMembers(this, subsiteUrl, teamId);
  }

  public async promoteToTeamLead(subsiteUrl: string, itemId: number): Promise<boolean> {
    return registration.promoteToTeamLead(this, subsiteUrl, itemId);
  }

  public async transferTeamLead(
    subsiteUrl: string,
    fromLeadItemId: number,
    toNewLeadItemId: number
  ): Promise<boolean> {
    return registration.transferTeamLead(this, subsiteUrl, fromLeadItemId, toNewLeadItemId);
  }

  public async isUserAlreadyOnEvent(subsiteUrl: string, email: string): Promise<boolean | null> {
    return registration.isUserAlreadyOnEvent(this, subsiteUrl, email);
  }

  // ==================== DEX_TeamJoinRequests (v11.83) ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/teamJoinRequests.ts — hier nur Delegations-Stubs.

  public async ensureTeamJoinRequestsList(): Promise<void> {
    return teamJoinRequests.ensureTeamJoinRequestsList(this);
  }

  // ==================== Zugriffs-Queues (AccessFix / AssistantAccess) ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/accessQueues.ts — hier nur Delegations-Stubs.

  public async ensureAccessFixList(): Promise<void> {
    return accessQueues.ensureAccessFixList(this);
  }

  public async queueAccessFix(subsiteUrl: string, itemId: number, participantEmail: string): Promise<void> {
    return accessQueues.queueAccessFix(this, subsiteUrl, itemId, participantEmail);
  }

  public async ensureAssistantAccessList(): Promise<void> {
    return accessQueues.ensureAssistantAccessList(this);
  }

  public async queueAssistantAccess(args: {
    subsiteUrl: string; itemId: number; eventId: string; eventTitle: string;
    participantEmail: string; participantName: string; assistantEmail: string; assistantName: string;
    ownerEmail: string; linkType: 'delegation' | 'proxy';
  }): Promise<void> {
    return accessQueues.queueAssistantAccess(this, args);
  }

  public async getAssistantLinksForUser(myEmail: string): Promise<AssistantLink[]> {
    return accessQueues.getAssistantLinksForUser(this, myEmail);
  }

  public async setAssistantLinkRequest(linkId: number, args: {
    requestType: 'change' | 'cancel'; note: string; requestedByEmail: string; requestedByName: string;
  }): Promise<boolean> {
    return accessQueues.setAssistantLinkRequest(this, linkId, args);
  }

  public async resolveAssistantLinkRequest(linkId: number, decision: 'Done' | 'Rejected'): Promise<boolean> {
    return accessQueues.resolveAssistantLinkRequest(this, linkId, decision);
  }

  public async setAssistantLinkStatusForRegistration(itemId: number, subsiteUrl: string, status: 'Cancelled'): Promise<void> {
    return accessQueues.setAssistantLinkStatusForRegistration(this, itemId, subsiteUrl, status);
  }

  // ==================== Merk-/Protokoll-Listen rund um Mails ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/notificationLogs.ts — hier nur Delegations-Stubs.

  public async ensureInactiveNoticesList(): Promise<void> {
    return notificationLogs.ensureInactiveNoticesList(this);
  }

  public async getSentInactiveNotices(eventId: string): Promise<Set<string>> {
    return notificationLogs.getSentInactiveNotices(this, eventId);
  }

  public async recordInactiveNotice(eventId: string, participantEmail: string): Promise<void> {
    return notificationLogs.recordInactiveNotice(this, eventId, participantEmail);
  }

  public async ensurePostEventMailsList(): Promise<boolean> {
    return notificationLogs.ensurePostEventMailsList(this);
  }

  public async getPostEventMailSentEventIds(): Promise<Set<string>> {
    return notificationLogs.getPostEventMailSentEventIds(this);
  }

  public async recordPostEventMail(eventId: string | number): Promise<void> {
    return notificationLogs.recordPostEventMail(this, eventId);
  }

  public async ensureEventCommsList(): Promise<void> {
    return notificationLogs.ensureEventCommsList(this);
  }

  public async logEventComm(meta: { eventId: string | number; eventTitle: string; subject: string; bodyHtml: string; emailType: string }): Promise<void> {
    return notificationLogs.logEventComm(this, meta);
  }

  public async getEventComms(eventId: string | number): Promise<EventCommRow[]> {
    return notificationLogs.getEventComms(this, eventId);
  }

  public async deleteEventComm(id: number): Promise<boolean> {
    return notificationLogs.deleteEventComm(this, id);
  }

  public async hasEventComms(
    eventId: string | number,
    excludeTypes?: string[],
  ): Promise<boolean> {
    return notificationLogs.hasEventComms(this, eventId, excludeTypes);
  }

  public async ensureOutlookLocksList(): Promise<void> {
    return notificationLogs.ensureOutlookLocksList(this);
  }

  // ==================== Team-Beitritts-Anfragen: anlegen / lesen / entscheiden ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/teamJoinRequests.ts — hier nur Delegations-Stubs.

  public async createTeamJoinRequest(args: {
    eventId: string;
    eventTitle: string;
    teamId: string;
    requesterEmail: string;
    requesterDisplayName: string;
    // v18.73: event-spezifische Antworten als JSON (optional).
    customData?: string;
  }): Promise<{ ok: boolean; itemId?: number }> {
    return teamJoinRequests.createTeamJoinRequest(this, args);
  }

  public async listTeamJoinRequests(args: {
    eventId?: string;
    teamId?: string;
    status?: 'Pending' | 'Approved' | 'Rejected';
  }): Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string; DecidedDate?: string; DecidedByEmail?: string; CustomData?: string }>> {
    return teamJoinRequests.listTeamJoinRequests(this, args);
  }

  public async decideTeamJoinRequest(
    requestId: number,
    decision: 'Approved' | 'Rejected',
    decidedByEmail: string
  ): Promise<boolean> {
    return teamJoinRequests.decideTeamJoinRequest(this, requestId, decision, decidedByEmail);
  }

  // ==================== Anmeldung lesen und ändern ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/registrationEdit.ts — hier nur Delegations-Stubs.

  public async reactivateRegistration(
    subsiteUrl: string,
    itemId: number,
    firstName: string,
    surname: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string>,
    registeredByName?: string, // Audit: Name des Users der die Re-Anmeldung auslöst
    registeredByEmail?: string, // Audit: E-Mail des Users der die Re-Anmeldung auslöst
    proxyConsent?: string, // v18.74: Zustimmungs-Nachweis bei stellvertretender Re-Anmeldung
    starterType?: string, // v30.67: Gruppe bei geteilten Kapazitäten (wie registerForEvent)
    preferredStarterType?: string // v30.67: gewünschte Gruppe
  ): Promise<boolean> {
    return registrationEdit.reactivateRegistration(this, subsiteUrl, itemId, firstName, surname, customData, status, customFieldMap, registeredByName, registeredByEmail, proxyConsent, starterType, preferredStarterType);
  }

  public async mergeRegistrationFields(
    subsiteUrl: string,
    itemId: number,
    body: Record<string, unknown>
  ): Promise<boolean> {
    return registrationEdit.mergeRegistrationFields(this, subsiteUrl, itemId, body);
  }

  public async assignRegistrationToTeam(
    subsiteUrl: string,
    itemId: number,
    teamId: string,
    teamName: string | undefined,
    isLead: boolean,
  ): Promise<boolean> {
    return registrationEdit.assignRegistrationToTeam(this, subsiteUrl, itemId, teamId, teamName, isLead);
  }

  public async updateRegistrationData(
    subsiteUrl: string,
    itemId: number,
    customData: Record<string, string>,
    customFieldMap?: Record<string, string>,
    oldCustomData?: Record<string, string>,
    fieldLabelMap?: Record<string, string> // cf.id -> label
  ): Promise<boolean> {
    return registrationEdit.updateRegistrationData(this, subsiteUrl, itemId, customData, customFieldMap, oldCustomData, fieldLabelMap);
  }

  public async adminUpdateRegistration(
    subsiteUrl: string,
    itemId: number,
    patch: Record<string, unknown>,
    actor: { name: string; email: string },
    oldValues?: Record<string, unknown>,
    fieldLabelMap?: Record<string, string>
  ): Promise<boolean> {
    return registrationEdit.adminUpdateRegistration(this, subsiteUrl, itemId, patch, actor, oldValues, fieldLabelMap);
  }

  public async getMyRegistration(
    subsiteUrl: string,
    email: string,
    onHttpError?: (_status: number) => void // v30.67: „konnte nicht lesen" von „keine Zeile" trennen
  ): Promise<SPRegistration | null> {
    return registrationEdit.getMyRegistration(this, subsiteUrl, email, onHttpError);
  }

  public async getProxyRegistrationsByActor(
    subsiteUrl: string,
    actorEmail: string
  ): Promise<SPRegistration[]> {
    return registrationEdit.getProxyRegistrationsByActor(this, subsiteUrl, actorEmail);
  }

  public async markConsentPendingByEmail(subsiteUrl: string, participantEmail: string): Promise<boolean> {
    return registrationEdit.markConsentPendingByEmail(this, subsiteUrl, participantEmail);
  }

  public async storeInviteEmlByEmail(subsiteUrl: string, participantEmail: string, emlContent: string): Promise<number> {
    return registrationEdit.storeInviteEmlByEmail(this, subsiteUrl, participantEmail, emlContent);
  }

  public async getInviteEmlByItem(subsiteUrl: string, itemId: number): Promise<{ fileName: string; content: string } | null> {
    return registrationEdit.getInviteEmlByItem(this, subsiteUrl, itemId);
  }

  public async confirmConsentReview(subsiteUrl: string, itemId: number, meta?: { eventId?: string; eventTitle?: string; participantName?: string }): Promise<boolean> {
    return registrationEdit.confirmConsentReview(this, subsiteUrl, itemId, meta);
  }

  public async ensureStartNumberColumn(subsiteUrl: string): Promise<boolean> {
    return registrationEdit.ensureStartNumberColumn(this, subsiteUrl);
  }

  public async getAllRegistrations(subsiteUrl: string, onHttpError?: (_status: number) => void): Promise<SPRegistration[]> {
    return registrationEdit.getAllRegistrations(this, subsiteUrl, onHttpError);
  }

  // ==================== Platz-Zähler / Wartelisten-Reihenfolge ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/seats.ts — hier nur Delegations-Stubs. `seatFieldFor` und
  // `adjustSeatCounterField` sind public, weil das Überbuchungs-Modul sie
  // über `svc` aufruft.

  public async setWaitlistPosition(
    subsiteUrl: string,
    itemId: number,
    targetPosition: number,
    group?: string // v30.67: Rang innerhalb dieser Gruppe (PreferredStarterType) bei getrennten Wartelisten
  ): Promise<{ ok: boolean; from: number; to: number; changed: number; error?: string }> {
    return seats.setWaitlistPosition(this, subsiteUrl, itemId, targetPosition, group);
  }

  public async reorderParticipantIDs(
    subsiteUrl: string,
    onProgress?: (pct: number) => void
  ): Promise<{ success: number; errors: number }> {
    return seats.reorderParticipantIDs(this, subsiteUrl, onProgress);
  }

  public seatFieldFor(group: string): 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun' {
    return seats.seatFieldFor(this, group);
  }

  public async reserveSeat(
    subsiteUrl: string,
    group: '' | 'Durchstarter' | 'Funstarter',
    cap: number,
    count: number = 1
  ): Promise<'reserved' | 'full' | 'error'> {
    return seats.reserveSeat(this, subsiteUrl, group, cap, count);
  }

  public async syncSeatsToActiveCount(
    subsiteUrl: string,
    opts: { isSplit: boolean }
  ): Promise<boolean> {
    return seats.syncSeatsToActiveCount(this, subsiteUrl, opts);
  }

  public async adjustSeatCounterField(
    subsiteUrl: string,
    field: 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun',
    delta: number
  ): Promise<void> {
    return seats.adjustSeatCounterField(this, subsiteUrl, field, delta);
  }

  public async releaseSeatAfterCancel(
    subsiteUrl: string,
    opts: { isSplit: boolean; previousStatus: string; starterType?: string; waitlistDisabled?: boolean }
  ): Promise<void> {
    return seats.releaseSeatAfterCancel(this, subsiteUrl, opts);
  }

  public async adjustWaitlistCounter(subsiteUrl: string, delta: number): Promise<void> {
    return seats.adjustWaitlistCounter(this, subsiteUrl, delta);
  }

  public async getCounterStats(subsiteUrl: string, isSplit: boolean): Promise<seats.CounterStats | null> {
    return seats.getCounterStats(this, subsiteUrl, isSplit);
  }

  public async subscribeListRealtime(
    subsiteUrl: string,
    kind: 'counter' | 'participants',
    onChange: () => void
  ): Promise<() => void> {
    return seats.subscribeListRealtime(this, subsiteUrl, kind, onChange);
  }

  // ==================== Überbuchung erkennen und auflösen ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/overbooking.ts — hier nur Delegations-Stubs.

  public async detectOverbooking(
    subsiteUrl: string,
    opts: { isSplit: boolean; maxParticipants?: number; durchstarterCapacity?: number; funstarterCapacity?: number }
  ): Promise<{ groups: Array<{ group: string; cap: number; activeBefore: number; marked: number }>; total: number; errors: number }> {
    return overbooking.detectOverbooking(this, subsiteUrl, opts);
  }

  public async clearOverbookMark(subsiteUrl: string, itemId: number): Promise<boolean> {
    return overbooking.clearOverbookMark(this, subsiteUrl, itemId);
  }

  public async resolveOverbookToWaitlist(
    subsiteUrl: string,
    itemId: number,
    group: string
  ): Promise<boolean> {
    return overbooking.resolveOverbookToWaitlist(this, subsiteUrl, itemId, group);
  }

  public async resolveOverbookKeepActive(subsiteUrl: string, itemId: number): Promise<boolean> {
    return overbooking.resolveOverbookKeepActive(this, subsiteUrl, itemId);
  }

  public async resolveOverbookKeepAsFirstWaitlist(
    subsiteUrl: string,
    itemId: number,
    group: string
  ): Promise<boolean> {
    return overbooking.resolveOverbookKeepAsFirstWaitlist(this, subsiteUrl, itemId, group);
  }

  public async buildOverbookApologyEmail(
    name: string,
    eventTitle: string,
    lang: string,
    waitlistPos?: number
  ): Promise<{ subject: string; body: string }> {
    return overbooking.buildOverbookApologyEmail(this, name, eventTitle, lang, waitlistPos);
  }

  // ==================== Teilnehmerliste prüfen / reparieren ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/regListRepair.ts — hier nur Delegations-Stubs.

  public async diagnoseRegistrationList(
    subsiteUrl: string,
    customFields: Array<{ id: string; label: string; spInternalName?: string }>
  ): Promise<{ ok: boolean; listMissing: boolean; missingColumns: Array<{ id: string; label: string; column: string }>; error?: string }> {
    return regListRepair.diagnoseRegistrationList(this, subsiteUrl, customFields);
  }

  public async fixRegistrationListColumns(
    subsiteUrl: string,
    eventContext?: {
      isB2Run?: boolean;  // Event hat Durchstarter/Funstarter Kapazität
      hasQuiz?: boolean;  // Event hat Quizfragen
      customFields?: CustomField[]; // Event-spezifische Custom-Fields — fehlende SP-Spalten werden angelegt
    },
    // v11.56: Optionaler Confirm-Callback. Wird aufgerufen, wenn Duplikat-Spalten
    // erkannt wurden, BEVOR irgendetwas gelöscht wird. Liefert der Callback false,
    // werden die Duplikate übersprungen (die Hauptfix-Logik läuft trotzdem).
    confirmDeleteDuplicates?: (count: number, titles: string[]) => boolean | Promise<boolean>
  ): Promise<{ added: string[]; removed: string[]; viewFixed: boolean; customFieldMap?: Record<string, string>; duplicatesRemoved?: string[]; duplicatesWithData?: string[] }> {
    return regListRepair.fixRegistrationListColumns(this, subsiteUrl, eventContext, confirmDeleteDuplicates);
  }


  // ==================== Fun-Zone / Quiz ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/quiz.ts — hier nur ein Delegations-Stub.

  public async saveQuizProgress(
    subsiteUrl: string,
    itemId: number,
    score: number,
    answers: number[][],
    isComplete: boolean
  ): Promise<boolean> {
    return quiz.saveQuizProgress(this, subsiteUrl, itemId, score, answers, isComplete);
  }

  // ==================== Nachrücken von der Warteliste ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/waitlist.ts — hier nur ein Delegations-Stub.

  public async promoteFirstWaitlistItem(
    subsiteUrl: string,
    inheritStarterType?: string,
    maxParticipants?: number,
    /** Seit v6.5: Bei B2Run-Events mit getrennten Durchstarter-/Funstarter-Wartelisten
     * hier den freigewordenen Starter-Typ mitgeben — dann wird NUR der erste
     * Warteliste-Teilnehmer mit passendem PreferredStarterType nachgerückt.
     * Wenn leer: Default-Verhalten (beliebiger Warteliste-Teilnehmer). */
    onlyWithPreferredType?: string,
    /** v17.15: Audit-Tracking — wenn der Promote durch das Cancel einer
     *  konkreten Person ausgelöst wurde (in der App-Pfad), die E-Mail
     *  und Item-Id dieser Person mitgeben. Wird auf der nachrückenden
     *  Person als ReplacedParticipantEmail + PromotedDate gespeichert,
     *  und zusätzlich auf der cancelnden Person als
     *  ReplacedByParticipantEmail (zweite MERGE-PATCH). */
    replacedByCancel?: { itemId: number; participantEmail: string },
  ): Promise<{ success: boolean; email?: string; name?: string; itemId?: number; skippedOverbooked?: boolean }> {
    return waitlist.promoteFirstWaitlistItem(this, subsiteUrl, inheritStarterType, maxParticipants, onlyWithPreferredType, replacedByCancel);
  }

  // ==================== Anhänge an Anmeldezeilen ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/registrationAttachments.ts — hier nur Delegations-Stubs.

  public async listRegistrationAttachments(
    subsiteUrl: string,
    itemId: number,
  ): Promise<Array<{ fileName: string; serverRelativeUrl: string }>> {
    return registrationAttachments.listRegistrationAttachments(this, subsiteUrl, itemId);
  }

  public async addRegistrationAttachment(
    subsiteUrl: string,
    itemId: number,
    file: File,
    // v19.0: optionaler Präfix, um ein Attachment einem Dokument-Custom-Field
    // zuzuordnen (z.B. 'dxf-<fieldId>--'). Leer = generischer Attendee-Upload.
    fieldPrefix: string = '',
  ): Promise<boolean> {
    return registrationAttachments.addRegistrationAttachment(this, subsiteUrl, itemId, file, fieldPrefix);
  }

  public async deleteRegistrationAttachment(
    subsiteUrl: string,
    itemId: number,
    fileName: string,
  ): Promise<boolean> {
    return registrationAttachments.deleteRegistrationAttachment(this, subsiteUrl, itemId, fileName);
  }

  // ==================== Status einer Anmeldung ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/registrationStatus.ts — hier nur Delegations-Stubs.

  public async flipAllStarterTypes(subsiteUrl: string): Promise<{ ok: boolean; updated: number; failed: number }> {
    return registrationStatus.flipAllStarterTypes(this, subsiteUrl);
  }

  public async switchSplitGroup(
    subsiteUrl: string,
    itemId: number,
    newType: 'Durchstarter' | 'Funstarter',
    durchstarterCapacity: number,
    funstarterCapacity: number,
  ): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }> {
    return registrationStatus.switchSplitGroup(this, subsiteUrl, itemId, newType, durchstarterCapacity, funstarterCapacity);
  }

  public async cancelRegistration(
    subsiteUrl: string,
    itemId: number,
    cancelledByName?: string,
    cancelledByEmail?: string
  ): Promise<boolean> {
    return registrationStatus.cancelRegistration(this, subsiteUrl, itemId, cancelledByName, cancelledByEmail);
  }

  public async declineRegistration(
    subsiteUrl: string,
    firstName: string,
    surname: string,
    participantEmail: string,
    actorName?: string,
    actorEmail?: string
  ): Promise<boolean> {
    return registrationStatus.declineRegistration(this, subsiteUrl, firstName, surname, participantEmail, actorName, actorEmail);
  }

  public async checkInParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    return registrationStatus.checkInParticipant(this, subsiteUrl, itemId);
  }

  public async markNoShowParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    return registrationStatus.markNoShowParticipant(this, subsiteUrl, itemId);
  }

  public async checkOutParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    return registrationStatus.checkOutParticipant(this, subsiteUrl, itemId);
  }

  public async setQRSentStatus(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    return registrationStatus.setQRSentStatus(this, subsiteUrl, itemId);
  }

  public async getRegistrationByEmail(
    subsiteUrl: string,
    email: string
  ): Promise<SPRegistration | null> {
    return registrationStatus.getRegistrationByEmail(this, subsiteUrl, email);
  }

  public async getRegistrationCount(subsiteUrl: string): Promise<{ registered: number; waitlist: number }> {
    return registrationStatus.getRegistrationCount(this, subsiteUrl);
  }

  public async getParticipantEmailsByStatus(subsiteUrl: string): Promise<{ active: string[]; waitlist: string[] }> {
    return registrationStatus.getParticipantEmailsByStatus(this, subsiteUrl);
  }

  public async updateRegistrationTitle(
    subsiteUrl: string,
    itemId: number,
    newTitle: string
  ): Promise<boolean> {
    return registrationStatus.updateRegistrationTitle(this, subsiteUrl, itemId, newTitle);
  }

  // ==================== Branding-Assets (Logo, Orb, Video) ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/branding.ts — hier nur Delegations-Stubs.

  public async ensureAssetsFolders(): Promise<void> {
    return branding.ensureAssetsFolders(this);
  }

  public async getBranding(): Promise<{ logoBase64: string; orbBase64: string; videoUrl: string; videoFileName: string }> {
    return branding.getBranding(this);
  }

  public async saveBrandingOrb(orbDataUri: string): Promise<boolean> {
    return branding.saveBrandingOrb(this, orbDataUri);
  }

  public async saveBrandingLogo(logoDataUri: string): Promise<boolean> {
    return branding.saveBrandingLogo(this, logoDataUri);
  }

  public async uploadBrandingVideo(file: File): Promise<string> {
    return branding.uploadBrandingVideo(this, file);
  }

  /** Serialisiert die Overrides-Schreibvorgänge (siehe services/events/eventAssets.ts).
   *  v30.66: public — das Thema liegt im Modul, der Instanz-Zustand bleibt hier. */
  public _ovQueue: Promise<void> = Promise.resolve();

  // ==================== Event-Bilder, -Dokumente, EmailTemplateOverrides ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/eventAssets.ts — hier nur Delegations-Stubs. Die
  // Serialisierungs-Warteschlange `_ovQueue` bleibt als Instanz-Zustand hier.

  public async uploadEventImageAsAttachment(eventId: number, file: File): Promise<string> {
    return eventAssets.uploadEventImageAsAttachment(this, eventId, file);
  }

  public async deleteEventOrigImageAttachment(eventId: number): Promise<void> {
    return eventAssets.deleteEventOrigImageAttachment(this, eventId);
  }

  public async uploadEventOrigImageAsAttachment(eventId: number, file: File): Promise<string> {
    return eventAssets.uploadEventOrigImageAsAttachment(this, eventId, file);
  }

  public async patchEventOverridesKey(eventId: number, key: string, value: string): Promise<void> {
    return eventAssets.patchEventOverridesKey(this, eventId, key, value);
  }

  public async patchEventOverridesValue(eventId: number, key: string, value: unknown): Promise<boolean> {
    return eventAssets.patchEventOverridesValue(this, eventId, key, value);
  }

  public async patchEventOverridesValueEx(
    eventId: number, key: string, value: unknown,
  ): Promise<{ ok: boolean; status: number; detail: string }> {
    return eventAssets.patchEventOverridesValueEx(this, eventId, key, value);
  }

  public async updateEventImageUrl(eventId: number, url: string): Promise<boolean> {
    return eventAssets.updateEventImageUrl(this, eventId, url);
  }

  public async uploadEventDocument(eventId: number, file: File): Promise<string> {
    return eventAssets.uploadEventDocument(this, eventId, file);
  }

  public async deleteEventDocument(eventId: number, fileName: string): Promise<boolean> {
    return eventAssets.deleteEventDocument(this, eventId, fileName);
  }

  public async getEventAttachments(eventId: number): Promise<Array<{ name: string; url: string; size: number }>> {
    return eventAssets.getEventAttachments(this, eventId);
  }

  // ==================== Profil-Daten ====================
  // v30.7 (Modularisierung Stufe 2, Tranche 3): Implementierung in
  // services/events/profileData.ts — hier nur Delegations-Stubs.

  public async canRegisterForOthers(subsiteUrl: string, targetParticipantEmail: string): Promise<boolean> {
    return profileData.canRegisterForOthers(this, subsiteUrl, targetParticipantEmail);
  }

  public async getCurrentUserProfile(): Promise<{ department: string; location: string; jobTitle: string; phone: string; firstName: string; lastName: string; displayName: string; company: string }> {
    return profileData.getCurrentUserProfile(this);
  }

  public async fixEventParticipantsProfileData(subsiteUrl: string, n: number = 1000): Promise<{ scanned: number; updated: number; failedLookups: number }> {
    return profileData.fixEventParticipantsProfileData(this, subsiteUrl, n);
  }

  public async fixRecentParticipantsProfileData(n: number): Promise<{ scanned: number; updated: number; failedLookups: number }> {
    return profileData.fixRecentParticipantsProfileData(this, n);
  }

  public async getUserProfileByEmail(email: string): Promise<{ department: string; location: string; jobTitle: string; phone: string; firstName: string; lastName: string; displayName: string; company: string }> {
    return profileData.getUserProfileByEmail(this, email);
  }

  // v30.66: `onProgress` fehlte hier seit dem Auszug in v30.7. Das Modul ruft den
  // Rueckruf an zwei Stellen auf (je Zeile und am Schluss), der Stub nahm ihn aber
  // gar nicht erst entgegen — ueber die Klasse war der Fortschritt also nicht
  // erreichbar. Der einzige Aufrufer uebergibt bisher keinen, deshalb ist nie
  // etwas aufgefallen; wer einen uebergeben haette, waere am Compiler gescheitert.
  public async repairClaimNamesInRegistrations(
    subsiteUrl: string,
    onProgress?: (done: number, total: number) => void,
  ): Promise<{ scanned: number; hits: number; fixed: number; failed: number }> {
    return profileData.repairClaimNamesInRegistrations(this, subsiteUrl, onProgress);
  }

  public async displayNameForEmail(email: string): Promise<string> {
    return profileData.displayNameForEmail(this, email);
  }


  // ==================== Hilfsmethoden ====================

  public async listExists(listName: string): Promise<boolean> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')`,
        SPHttpClient.configurations.v1
      );
      // 200 = OK, 403 = existiert aber kein Zugriff (User hat nur Read)
      return response.ok || response.status === 403;
    } catch {
      return false;
    }
  }

  /**
   * ID der SharePoint-Gruppe "DEALL" (alle Deloitte-Mitarbeiter) ermitteln.
   */
  /**
   * ID der Visitors-Gruppe ermitteln (dort ist DEALL / alle Deloitte-Mitarbeiter hinterlegt).
   */
  public async getVisitorsGroupId(): Promise<number | null> {
    try {
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/associatedvisitorgroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const d = await resp.json();
        return d.d?.Id || d.Id || null;
      }
    } catch { /* */ }
    return null;
  }

  public async _post(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
      },
      body: JSON.stringify(body),
    };
    return this._sp.post(url, SPHttpClient.configurations.v1, options);
  }

  /**
   * MERGE-Request für Item-Updates.
   * SharePoint verarbeitet den MERGE korrekt, antwortet aber auf manchen
   * Subsite-Listen mit 406 (Accept-Format nicht unterstützt). Da bei MERGE
   * kein Response-Body benötigt wird, behandeln wir 406 als Erfolg.
   */
  public async _merge(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    const response = await this._sp.post(url, SPHttpClient.configurations.v1, options);
    // 406: SharePoint hat den MERGE ausgeführt, kann aber nicht im
    // gewünschten Format antworten. Daten sind trotzdem gespeichert.
    if (response.status === 406) {
      return { ok: true, status: 204, statusText: 'No Content' } as unknown as SPHttpClientResponse;
    }
    return response;
  }

  // v7.28: Variante von _merge, die den übergebenen ETag im IF-MATCH-Header
  // mitsendet (statt '*'). SharePoint vergleicht den ETag mit dem aktuellen
  // Stand des Items und antwortet mit HTTP 412 (Precondition Failed), wenn
  // ein anderer Client zwischenzeitlich geschrieben hat. So können wir
  // optimistic-concurrency-Pattern für den TeilnehmerID-Counter umsetzen.
  public async _mergeIfMatch(url: string, body: object, etag: string): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': etag,
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    const response = await this._sp.post(url, SPHttpClient.configurations.v1, options);
    if (response.status === 406) {
      return { ok: true, status: 204, statusText: 'No Content' } as unknown as SPHttpClientResponse;
    }
    return response;
  }

  // v24.76: Counter-Felder (inkl. des neuen WaitlistTaken) EINMAL pro Subsite
  // pro Session sicherstellen, bevor darauf geschrieben wird — sonst liefert der
  // MERGE auf Bestands-Events ein HTTP 400 (Feld existiert noch nicht). Gecacht,
  // damit nicht bei jedem Reconcile/Bump erneut geprobt wird.
  // v30.66: public — der Merker ist Instanz-Zustand, das Thema liegt im Modul.
  public _counterFieldsEnsured: Set<string> = new Set<string>();

  // ==================== TeilnehmerID-Zähler ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/teilnehmerIdCounter.ts — hier nur Delegations-Stubs.
  // Die bisher privaten Helfer sind public, weil das Modul sie über `svc`
  // aufruft (Unterstrich-Konvention: „intern").

  public async getNextTeilnehmerId(subsiteUrl: string): Promise<number | undefined> {
    return teilnehmerIdCounter.getNextTeilnehmerId(this, subsiteUrl);
  }

  public async ensureCounterFieldsOnce(subsiteUrl: string): Promise<void> {
    return teilnehmerIdCounter.ensureCounterFieldsOnce(this, subsiteUrl);
  }

  public async ensureCounterListField(subsiteUrl: string): Promise<void> {
    return teilnehmerIdCounter.ensureCounterListField(this, subsiteUrl);
  }

  public async ensureCounterList(subsiteUrl: string): Promise<{ created: boolean; seededValue?: number }> {
    return teilnehmerIdCounter.ensureCounterList(this, subsiteUrl);
  }

  public async setCounterListPermissions(subsiteUrl: string, organizerEmail?: string): Promise<void> {
    return teilnehmerIdCounter.setCounterListPermissions(this, subsiteUrl, organizerEmail);
  }

  public async getCurrentMaxTeilnehmerId(subsiteUrl: string): Promise<number> {
    return teilnehmerIdCounter.getCurrentMaxTeilnehmerId(this, subsiteUrl);
  }

  public async resetCounterToMax(subsiteUrl: string): Promise<{ counter: number; max: number }> {
    return teilnehmerIdCounter.resetCounterToMax(this, subsiteUrl);
  }

  public async syncCounterToMax(subsiteUrl: string): Promise<void> {
    return teilnehmerIdCounter.syncCounterToMax(this, subsiteUrl);
  }

  public async deleteRegistration(subsiteUrl: string, itemId: number): Promise<boolean> {
    return registrationEdit.deleteRegistration(this, subsiteUrl, itemId);
  }

  // ==================== Berechtigungs-Audit / verwaiste Subsites ====================
  // v30.66 (Modularisierung Stufe 2): Implementierung in
  // services/events/permissionsAudit.ts — hier nur Delegations-Stubs.

  public async hardenQueueListsIls(): Promise<{ fixed: string[]; failed: string[] }> {
    return permissionsAudit.hardenQueueListsIls(this);
  }

  public async auditOrCleanupPermissions(
    apply: boolean,
    ctx: { adminEmails: string[]; organizerEmails: string[]; subsiteOrganizers: Record<string, string>; selfEmail: string },
    onProgress?: (msg: string, done: number, total: number) => void
  ): Promise<PermCleanupReport> {
    return permissionsAudit.auditOrCleanupPermissions(this, apply, ctx, onProgress);
  }

  public async findOrphanSubsites(
    onProgress?: (msg: string, done: number, total: number) => void
  ): Promise<OrphanScanResult> {
    return permissionsAudit.findOrphanSubsites(this, onProgress);
  }

  public async deleteSubsiteWeb(webUrl: string): Promise<boolean> {
    return permissionsAudit.deleteSubsiteWeb(this, webUrl);
  }

  // ==================== v21: Archivierung ====================
  // v30.11: Thema ausgelagert nach services/events/archive.ts (DEX_Archive-
  // Liste, Archiv-Lauf über ARCHIVE_SOURCES, Löschkonzept) — hier stehen nur
  // noch Delegations-Stubs mit unveränderter Signatur.

  public async ensureArchiveList(): Promise<void> {
    return archive.ensureArchiveList(this);
  }

  // ==================== Wochenbericht (v23.8) ====================
  // v30.11: Thema ausgelagert nach services/events/weeklyReport.ts —
  // Delegations-Stubs mit unveränderter Signatur.

  public async ensureWeeklyReportsList(): Promise<void> {
    return weeklyReport.ensureWeeklyReportsList(this);
  }

  public async getLastWeeklyReport(): Promise<{ created: string; periodTo: string; draftEventIds: string[] } | null> {
    return weeklyReport.getLastWeeklyReport(this);
  }

  public async recordWeeklyReport(fromIso: string, toIso: string, draftEventIds?: string[]): Promise<void> {
    return weeklyReport.recordWeeklyReport(this, fromIso, toIso, draftEventIds);
  }

  // ==================== Ticketsystem (v26.0.0) ====================
  // v28.95: Der Inhalt liegt in services/tickets.ts. Die Klasse behält ihre
  // Methoden — so ändert sich an keiner der Aufrufstellen etwas, und der
  // Compiler prueft den Weg von hier bis in das Modul.

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async ensureTicketsList(): Promise<void> {
    return tickets.ensureTicketsList(this);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async createTicket(t: {
    questions: string[];
    askerEmail: string; askerName: string; askerRole: string;
    askerLocation?: string; askerJobTitle?: string;
    audience: string; eventId: string; eventTitle: string;
    assignedOrganizers: string[]; pageContext: string;
    askWizardStep?: number | null;
    /** v26.60: 'bug' = Bug-Report (Benachrichtigung an die DEX-Maintainer
     *  statt an alle Power-User); sonst inhaltliche Frage. */
    category?: 'question' | 'bug';
  }): Promise<number | null> {
    return tickets.createTicket(this, t);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async addTicketAttachment(itemId: number, file: File, kind: 'ask' | 'ans'): Promise<boolean> {
    return tickets.addTicketAttachment(this, itemId, file, kind);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async getTickets(): Promise<DexTicket[]> {
    return tickets.getTickets(this);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async getMyTickets(email: string): Promise<DexTicket[]> {
    return tickets.getMyTickets(this, email);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async claimTicket(
    itemId: number, email: string, name: string, opts?: { onlyIfOpen?: boolean }
  ): Promise<{ ok: boolean; conflict?: boolean; claimedByName?: string; status?: string }> {
    return tickets.claimTicket(this, itemId, email, name, opts);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async releaseTicket(itemId: number): Promise<boolean> {
    return tickets.releaseTicket(this, itemId);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async answerTicket(itemId: number, a: {
    answerText: string; articleIds: string[]; wizardStep: number | null;
    wizardMarker?: { x: number; y: number; w: number; h: number } | null;
    answeredByEmail: string; answeredByName: string;
    answeredByLocation?: string; answeredByJobTitle?: string;
  }): Promise<boolean> {
    return tickets.answerTicket(this, itemId, a);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async setTicketFollowUps(itemId: number, followUps: TicketFollowUp[], extra: {
    status?: string; claimedByEmail?: string; claimedByName?: string; claimedAt?: string | null;
  }): Promise<boolean> {
    return tickets.setTicketFollowUps(this, itemId, followUps, extra);
  }

  /** v28.95: delegiert an services/tickets.ts (siehe dort). */
  public async closeTicketNoAnswer(itemId: number): Promise<boolean> {
    return tickets.closeTicketNoAnswer(this, itemId);
  }

  // ==================== Organizer-Verwaltung ====================
  // v30.12: Organizer-Archiv (v24.6), Organizer-Anträge (v23.37) und
  // Rollen-/Bericht-Abfragen ausgelagert nach services/events/organizer.ts —
  // hier stehen nur noch Delegations-Stubs mit unveränderter Signatur.

  public async ensureOrganizerRequestsList(): Promise<void> {
    return organizer.ensureOrganizerRequestsList(this);
  }

  public async ensureOrganizerArchivedList(): Promise<void> {
    return organizer.ensureOrganizerArchivedList(this);
  }

  public async getOrganizerArchivedEventIds(email: string): Promise<Set<string>> {
    return organizer.getOrganizerArchivedEventIds(this, email);
  }

  public async archiveEventForOrganizer(eventId: string, email: string): Promise<boolean> {
    return organizer.archiveEventForOrganizer(this, eventId, email);
  }

  public async unarchiveEventForOrganizer(eventId: string, email: string): Promise<boolean> {
    return organizer.unarchiveEventForOrganizer(this, eventId, email);
  }

  public async createOrganizerRequest(email: string, name: string, location: string, message: string): Promise<{ ok: boolean; itemId?: number }> {
    return organizer.createOrganizerRequest(this, email, name, location, message);
  }

  public async getOrganizerRequests(onlyPending: boolean = true): Promise<Array<{ id: number; email: string; name: string; location: string; message: string; status: string; created: string }>> {
    return organizer.getOrganizerRequests(this, onlyPending);
  }

  public async getOrganizerRequestDetails(id: number): Promise<{ id: number; email: string; name: string; status: string; decidedByEmail: string; decidedDate: string } | null> {
    return organizer.getOrganizerRequestDetails(this, id);
  }

  public async grantSiteReadAccess(email: string): Promise<boolean> {
    return organizer.grantSiteReadAccess(this, email);
  }

  public async updateOrganizerRequestStatus(id: number, status: 'Approved' | 'Rejected', decidedByEmail: string): Promise<boolean> {
    return organizer.updateOrganizerRequestStatus(this, id, status, decidedByEmail);
  }

  public async getRoleRecipients(role: string): Promise<Array<{ email: string; name: string }>> {
    return organizer.getRoleRecipients(this, role);
  }

  public async getRoleEmails(role: string): Promise<string[]> {
    return organizer.getRoleEmails(this, role);
  }

  public async getRoleItemsCreatedSince(role: string, fromIso: string): Promise<Array<{ email: string; created: string }>> {
    return organizer.getRoleItemsCreatedSince(this, role, fromIso);
  }

  public async getEventsCreatedSince(fromIso: string): Promise<Array<{ title: string; author: string; created: string; isDraft: boolean }>> {
    return organizer.getEventsCreatedSince(this, fromIso);
  }

  public async countRegistrations(subsiteUrl: string, fromIso: string): Promise<{ total: number; since: number }> {
    return organizer.countRegistrations(this, subsiteUrl, fromIso);
  }

  public async countArchivableRows(
    expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
    allEventIds: Set<string> = new Set(), allSubsiteUrls: Set<string> = new Set()
  ): Promise<{ total: number; perList: Record<string, number> }> {
    return archive.countArchivableRows(this, expiredEventIds, expiredSubsiteUrls, allEventIds, allSubsiteUrls);
  }

  public async archiveExpiredRows(
    expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
    eventTitleById: Record<string, string>,
    onProgress?: (listIdx: number, listTotal: number, listName: string, done: number, total: number) => void,
    shouldCancel?: () => boolean,
    allEventIds: Set<string> = new Set(), allSubsiteUrls: Set<string> = new Set()
  ): Promise<{ archived: number; failed: number; cancelled: boolean; perList: Record<string, number> }> {
    return archive.archiveExpiredRows(this, expiredEventIds, expiredSubsiteUrls, eventTitleById, onProgress, shouldCancel, allEventIds, allSubsiteUrls);
  }

  // ==================== Archiv-Löschkonzept (v23.40) ====================
  // v30.11: ausgelagert nach services/events/archive.ts — Stubs s.o.

  public async countDeletableArchiveRows(olderThanIso: string): Promise<number> {
    return archive.countDeletableArchiveRows(this, olderThanIso);
  }

  public async deleteOldArchiveRows(
    olderThanIso: string,
    onProgress?: (done: number, total: number) => void,
    shouldCancel?: () => boolean
  ): Promise<{ deleted: number; failed: number; cancelled: boolean }> {
    return archive.deleteOldArchiveRows(this, olderThanIso, onProgress, shouldCancel);
  }

  // v30.11: public — der Archiv-Lauf (services/events/archive.ts) und weitere
  // Aufrufstellen brauchen den DELETE-Helfer von außen (Unterstrich = intern).
  public async _delete(url: string): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE',
      },
    };
    return this._sp.post(url, SPHttpClient.configurations.v1, options);
  }
}
