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
import { wrapTemplateForStorage, buildEmailFromTemplate, normalizeMadeWithLink } from './EmailTemplates';
// v28.95: Die Mail-Koerper liegen jetzt in ./mailBodies — die Datei begann
// sonst mit 400 Zeilen HTML, bevor die erste Methode kam.
import { buildOutlookLocation } from '../utils/eventFormat';
import { subscribeListChanges } from '../utils/spListRealtime';
import { DexTicket, TicketFollowUp } from '../types';
// v28.95: Erstes nach Thema herausgeloestes Fach-Modul (siehe CLAUDE.md).
import * as tickets from './tickets';
export const REG_LIST_NAME = 'Teilnehmer';

/** v28.61: Je Teilnehmerliste nur einmal pro Sitzung die Hotel-Spalten
 *  anlegen — der Aufruf ist idempotent, aber nicht kostenlos (drei POSTs). */
export const HOTEL_COLS_READY = new Set<string>();
const REG_LIST_ITEM_TYPE = 'SP.Data.TeilnehmerListItem';
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
const COUNTER_LIST_NAME = 'DEX_TeilnehmerCounter';
const COUNTER_LIST_ITEM_TYPE = 'SP.Data.DEX_x005f_TeilnehmerCounterListItem';

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
  private _webOf(apiBase: string): string {
    const idx = apiBase.indexOf('/_api/web');
    return idx > 0 ? apiBase.slice(0, idx) : this.siteUrl;
  }

  /** Request-Digest (FormDigestValue) des angegebenen Webs holen (gecacht).
   *  Leerer String, wenn nicht ermittelbar. */
  private async _webDigest(webUrl: string): Promise<string> {
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


  /** v26.13: Versions-Historie der CustomFields-Spalte eines DEX_Events-Items
   *  (neueste zuerst). Grundlage für die Wiederherstellung versehentlich
   *  überschriebener Custom-Field-Beschreibungen (helpText etc.) aus der
   *  SharePoint-Versionshistorie. */
  public async getEventCustomFieldsVersions(itemId: number): Promise<Array<{ created: string; customFields: string }>> {
    // WICHTIG (v26.15): $select=Created,CustomFields ist PFLICHT — sonst liefert
    // der versions-Endpunkt ALLE Felder pro Version (inkl. der riesigen
    // OutlookBody-/EmailTemplateOverrides-Base64-Logos). Bei stark bearbeiteten
    // Events (z.B. 188 Versionen) sprengt das die Antwortgröße und SharePoint
    // bricht nach ~51 Versionen ab → die Version MIT der Beschreibung fehlte und
    // es kam fälschlich „helpText in Historie: false". KEIN $orderby (das löst
    // auf dem versions-Endpunkt 400 aus) — wir sortieren clientseitig nach
    // Created absteigend (neueste zuerst). Folgeseiten via nextLink einsammeln.
    const out: Array<{ created: string; customFields: string }> = [];
    let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${itemId})/versions?$select=Created,CustomFields&$top=500`;
    let guard = 0;
    try {
      while (url && guard < 25) {
        guard++;
        const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!resp.ok) { console.warn('[DEX restore] versions HTTP', resp.status, 'für Item', itemId); break; }
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const v of (items as any[])) {
          out.push({ created: v.Created || '', customFields: typeof v.CustomFields === 'string' ? v.CustomFields : '' });
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        url = (data as any)['@odata.nextLink'] || (data.d && (data.d as any).__next) || null;
      }
    } catch (e) {
      console.warn('[DEX restore] versions fetch failed für Item', itemId, e);
    }
    out.sort((a, b) => (b.created || '').localeCompare(a.created || ''));
    return out;
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

  public async getInvitedRecipients(eventId: string | number): Promise<string[]> {
    return outlookQueue.getInvitedRecipients(this, eventId);
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

  // ==================== DEX_Participants Liste ====================

  /**
   * Zentrale Teilnehmer-Liste erstellen falls nicht vorhanden.
   * Speichert pro Person die EventNumbers für Registrierung und Warteliste.
   */
  public async ensureParticipantsList(): Promise<void> {
    const listName = 'DEX_Participants';
    const exists = await this.listExists(listName);
    if (exists) {
      await this.ensureMissingParticipantsFields(listName);
      await this.ensureParticipantsIndexes(listName);
      await this.configureDefaultView(listName, [
        'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
      ]);
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Zentrale Teilnehmerliste der DEX Event Experience Platform',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields = [
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'Email', type: 2 },
      { title: 'EventRegistered', type: 3 }, // Note für beliebig viele EventNumbers
      { title: 'EventOnWaitlist', type: 3 }, // Note für beliebig viele EventNumbers
    ];

    for (const f of fields) {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    }

    await this.configureDefaultView(listName, [
      'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
    ]);
    await this.ensureParticipantsIndexes(listName);

    await this.setEmailsListPermissions(listName);
  }

  /**
   * v28.25: Index auf der Spalte `Email` sicherstellen.
   *
   * DEX_Participants wächst mit jeder je angemeldeten Person. Überschreitet die
   * Liste die SharePoint-Schwelle von 5000 Elementen, scheitert JEDE Abfrage,
   * die auf einer NICHT indizierten Spalte filtert oder sortiert — und zwar mit
   * HTTP 500 („exceeds the list view threshold"), nicht mit einer sprechenden
   * Meldung. Genau das legte im Tenant `getParticipantByEmail` (Filter auf
   * Email) und damit die gesamte Schattenbuchhaltung lahm: „Meine Events" blieb
   * leer, jede An-/Abmeldung konnte das Register nicht mehr fortschreiben und
   * die Reparatur-Aktion lief in eine Fehlerwand.
   *
   * Ein Index auf `Email` hebt die Sperre für genau diese Filter-Abfragen auf.
   * Best-effort: Das Setzen braucht „Listen verwalten" (Admin/Organizer der
   * Site); fehlt das Recht, bleibt alles wie bisher.
   */
  private async ensureParticipantsIndexes(listName: string): Promise<void> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Email')`;
      const resp = await this._sp.get(`${url}?$select=Indexed`, SPHttpClient.configurations.v1);
      if (!resp.ok) return;
      const data = await resp.json();
      if (data && data.Indexed === true) return;
      const m = await this._merge(url, { 'Indexed': true });
      if (!m.ok) console.warn('[DEX] Index auf DEX_Participants.Email konnte nicht gesetzt werden (HTTP ' + m.status + ') — bei >5000 Einträgen scheitern gefilterte Abfragen.');
    } catch (err) {
      console.warn('[DEX] ensureParticipantsIndexes fehlgeschlagen (best-effort):', err);
    }
  }

  private async ensureMissingParticipantsFields(listName: string): Promise<void> {
    const requiredFields = [
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'Email', type: 2 },
      { title: 'EventRegistered', type: 3 },
      { title: 'EventOnWaitlist', type: 3 },
    ];

    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title&$filter=Hidden eq false&$top=200`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;
      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const existingNames = new Set((data.value || []).flatMap((f: any) => [f.InternalName, f.Title]));

      for (const f of requiredFields) {
        if (!existingNames.has(f.title)) {
          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
            '__metadata': { 'type': 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          });
        }
      }
    } catch { /* optional */ }
  }

  /**
   * Teilnehmer-Eintrag per Email suchen
   */
  public async getParticipantByEmail(email: string): Promise<SPParticipant | null> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$filter=Email eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      if (data.value && data.value.length > 0) return data.value[0];
      return null;
    } catch {
      return null;
    }
  }

  /**
   * Teilnehmer anlegen oder aktualisieren bei Registrierung.
   * Fügt eventNumber zu EventRegistered oder EventOnWaitlist hinzu.
   */
  public async upsertParticipant(
    vorname: string,
    nachname: string,
    email: string,
    eventNumber: number,
    status: string // 'Angemeldet' | 'Warteliste'
  ): Promise<boolean> {
    try {
      const existing = await this.getParticipantByEmail(email);
      return this.applyParticipantEvent(existing, vorname, nachname, email, eventNumber, status);
    } catch {
      return false;
    }
  }

  /**
   * v28.25: Schreib-Teil von `upsertParticipant`, aber mit BEREITS bekanntem
   * Register-Eintrag. Der Massen-Abgleich (backfillParticipantRegistry) lädt
   * das Register einmal komplett und spart sich damit die Einzelabfrage pro
   * Person — bei mehreren hundert Teilnehmern hunderte Requests weniger, und
   * es funktioniert auch dann, wenn die Einzelabfrage an der 5000-Element-
   * Schwelle scheitern würde.
   */
  private async applyParticipantEvent(
    existing: SPParticipant | null,
    vorname: string,
    nachname: string,
    email: string,
    eventNumber: number,
    status: string
  ): Promise<boolean> {
    try {
      if (existing) {
        // EventNumber zu richtigem Feld hinzufügen
        let registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s) : [];
        let waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s) : [];
        const en = eventNumber.toString();

        // Erst aus beiden entfernen
        registered = registered.filter(n => n !== en);
        waitlist = waitlist.filter(n => n !== en);

        if (status === 'Warteliste') {
          waitlist.push(en);
        } else {
          registered.push(en);
        }

        await this._merge(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
          {
            'Vorname': vorname,
            'Nachname': nachname,
            'EventRegistered': registered.join(','),
            'EventOnWaitlist': waitlist.join(','),
          }
        );
      } else {
        // Neuen Eintrag erstellen
        const isWaitlist = status === 'Warteliste';
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items`, {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_ParticipantsListItem' },
          'Title': email,
          'Vorname': vorname,
          'Nachname': nachname,
          'Email': email,
          'EventRegistered': isWaitlist ? '' : eventNumber.toString(),
          'EventOnWaitlist': isWaitlist ? eventNumber.toString() : '',
        });
      }
      return true;
    } catch {
      return false;
    }
  }

  /**
   * EventNumber aus den Feldern eines Teilnehmers entfernen (bei Abmeldung).
   */
  /**
   * v29.0: Register GEGEN die Teilnehmerlisten prüfen.
   *
   * `analyzeParticipantRegistry` findet nur zwei Dinge: mehrere Einträge zur
   * selben E-Mail und Verweise auf GELÖSCHTE Events. Ob eine Event-Nummer, die
   * auf ein existierendes Event zeigt, dort auch eine Zeile hat, prüft sie
   * nicht — genau daraus entsteht der Fall „Meine Events sagt angemeldet, die
   * Teilnehmerliste kennt die Person nicht" (v28.99).
   *
   * Vorgehen: einmal das Register lesen, daraus je Event-Nummer die Menge der
   * E-Mails bilden, und dann JE EVENT dessen Teilnehmerliste EINMAL laden.
   * Verglichen wird gegen aktive Zeilen — „Abgemeldet" zählt nicht, denn eine
   * abgemeldete Person gehört nicht mehr ins Register.
   *
   * WICHTIG: Schlägt das Lesen einer Liste fehl, wird das Event ÜBERSPRUNGEN
   * und als solches gezählt. Aus einem Netzwerkfehler „keine Zeile gefunden"
   * abzuleiten, würde gültige Anmeldungen aus dem Register werfen.
   */
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
    const all = await this.fetchAllParticipantsOrThrow(onRead);
    // Event-Nummer → E-Mails, die laut Register dort angemeldet sind.
    const byNumber: Record<number, Set<string>> = {};
    for (const p of all) {
      const em = (p.Email || '').trim().toLowerCase();
      if (!em) continue;
      `${p.EventRegistered || ''},${p.EventOnWaitlist || ''}`
        .split(',').map(x => parseInt(x.trim(), 10)).filter(n => !isNaN(n) && n > 0)
        .forEach(n => { (byNumber[n] = byNumber[n] || new Set<string>()).add(em); });
    }
    const relevant = events.filter(e =>
      typeof e.eventNumber === 'number' && e.eventNumber > 0
      && !!e.subsiteUrl && !!byNumber[e.eventNumber]);
    const stale: Array<{ email: string; eventNumber: number; title: string }> = [];
    const perEvent: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number; suspicious: boolean; listGone: boolean }> = [];
    const suspiciousEvents: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number }> = [];
    const deletedListEvents: Array<{ title: string; eventNumber: number; referenced: number }> = [];
    let checkedEvents = 0;
    let skippedEvents = 0;
    for (let i = 0; i < relevant.length; i++) {
      const ev = relevant[i];
      const num = ev.eventNumber as number;
      if (onProgress) onProgress(i + 1, relevant.length, ev.title || '');
      let rows: SPRegistration[] | null = null;
      /**
       * v29.3: Der Lesefehler wird jetzt AUSGEWERTET statt verschluckt.
       * `getAllRegistrations` bricht bei einem HTTP-Fehler ab und liefert eine
       * leere Liste — bis v29.2 war deshalb „Subsite existiert nicht mehr"
       * (404) von „niemand ist angemeldet" nicht zu unterscheiden. Genau das
       * ist hier der Regelfall und nicht die Ausnahme: Das 3-Monats-
       * Löschkonzept (`deleteParticipantData`) recycelt die Subsite und LÄSST
       * das Event-Item stehen. Das Event ist danach weiter in der Liste, seine
       * Teilnehmerliste aber weg — und alle Register-Verweise darauf sind
       * Rückstand, den genau dieses Löschkonzept hinterlassen hat.
       */
      let httpStatus = 0;
      let hadHttpError = false;
      try {
        // eslint-disable-next-line no-await-in-loop
        rows = await this.getAllRegistrations(ev.subsiteUrl as string, s => { hadHttpError = true; httpStatus = s; });
      } catch { rows = null; }
      // 404/410 = Liste/Subsite gibt es nicht mehr → eindeutige Aussage.
      const listGone = hadHttpError && (httpStatus === 404 || httpStatus === 410);
      // Jeder ANDERE Fehler (403 fehlende Rechte, 429 Drosselung, 500, Netz)
      // sagt nichts über den Inhalt aus — solche Events werden übersprungen,
      // nicht geraten.
      if (!rows || (hadHttpError && !listGone)) { skippedEvents += 1; continue; }
      checkedEvents += 1;
      if (listGone) {
        const refsGone = Array.from(byNumber[num]);
        perEvent.push({ title: ev.title || '', eventNumber: num, referenced: refsGone.length, missing: refsGone.length, rows: 0, suspicious: false, listGone: true });
        deletedListEvents.push({ title: ev.title || '', eventNumber: num, referenced: refsGone.length });
        refsGone.forEach(em => { stale.push({ email: em, eventNumber: num, title: ev.title || '' }); });
        continue;
      }
      const active = new Set<string>();
      for (const r of rows) {
        const st = (r.Status || '').trim();
        if (st === 'Abgemeldet') continue;
        const em = (r.ParticipantEmail || '').trim().toLowerCase();
        if (em) active.add(em);
      }
      const refs = Array.from(byNumber[num]);
      const missing = refs.filter(em => !active.has(em));
      /**
       * Plausibilitäts-Riegel — gilt seit v29.3 nur noch für Events mit
       * LESBARER Teilnehmerliste. Wenn dort praktisch alle Verweise ins Leere
       * zeigen, ist die naheliegende Erklärung nicht, dass hunderte
       * Abmeldungen einzeln schiefgingen — sondern dass Register und Liste bei
       * diesem Event gar nicht vergleichbar sind. Denkbare Gründe: eine andere
       * Teilnehmerliste als die Standard-Liste, eine neu angelegte/geleerte
       * Liste bei erhaltenem Register, oder ein Event, das seine Anmeldungen
       * woanders führt.
       *
       * In dem Fall wäre ein Entfernen der Verweise ein Datenverlust, kein
       * Aufräumen. Solche Events werden deshalb ausgewiesen, aber NICHT
       * bereinigt — die Entscheidung darüber braucht einen Blick in die
       * betroffene Liste, nicht einen Knopfdruck.
       */
      const suspicious = refs.length >= 5 && missing.length >= Math.ceil(refs.length * 0.9);
      perEvent.push({ title: ev.title || '', eventNumber: num, referenced: refs.length, missing: missing.length, rows: active.size, suspicious, listGone: false });
      if (suspicious) {
        suspiciousEvents.push({ title: ev.title || '', eventNumber: num, referenced: refs.length, missing: missing.length, rows: active.size });
        continue;
      }
      missing.forEach(em => { stale.push({ email: em, eventNumber: num, title: ev.title || '' }); });
    }
    perEvent.sort((a, b) => b.missing - a.missing);
    deletedListEvents.sort((a, b) => b.referenced - a.referenced);
    return { checkedEvents, skippedEvents, stale, perEvent, suspiciousEvents, deletedListEvents };
  }

  /**
   * v29.4: Alle vergebenen `EventNumber` direkt aus `DEX_Events` lesen —
   * strikt, also mit Fehler statt stiller Teilliste.
   *
   * Bewusst NICHT die im Client geladene Event-Liste: `loadEvents` lässt
   * einzelne Events aus, wenn ihr Mapping scheitert (v9.41, damit ein kaputtes
   * Event nicht die ganze Liste kippt). Für eine Anzeige ist das richtig — für
   * die Frage „gibt es dieses Event noch?" wäre es fatal, weil ein
   * ausgelassenes Event wie ein gelöschtes aussähe und seine Verweise entfernt
   * würden.
   *
   * Fensterung nach `Id` wie in `readAllParticipants`: schwellenfest und
   * unabhängig vom nextLink-Format.
   */
  private async readAllEventNumbersOrThrow(): Promise<Set<number>> {
    const out = new Set<number>();
    const PAGE = 2000;
    const MAX_PAGES = 100;
    let lastId = 0;
    for (let page = 0; page < MAX_PAGES; page++) {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`
        + `?$select=Id,EventNumber&$filter=Id gt ${lastId}&$orderby=Id asc&$top=${PAGE}`;
      // eslint-disable-next-line no-await-in-loop
      const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) throw new Error(`DEX_Events nicht vollständig lesbar (HTTP ${resp.status}).`);
      // eslint-disable-next-line no-await-in-loop
      const data = await resp.json();
      const items = (data.value || data.d?.results || []) as Array<{ Id: number; EventNumber?: number }>;
      if (items.length === 0) break;
      items.forEach(it => {
        if (typeof it.EventNumber === 'number' && it.EventNumber > 0) out.add(it.EventNumber);
        if (typeof it.Id === 'number' && it.Id > lastId) lastId = it.Id;
      });
      if (items.length < PAGE) break;
    }
    return out;
  }

  /**
   * v29.4: Verweise im Register auf Event-Nummern, die es in `DEX_Events`
   * NICHT MEHR GIBT. Bis v29.3 wurden die nur gezählt („wirkungslos, aber
   * harmlos") — sie sind aber personenbezogener Rückstand gelöschter Events
   * und gehören weg.
   *
   * Zwei Riegel, weil ein Fehlurteil hier das ganze Register leeren würde:
   *  - Event-Nummern und Register werden BEIDE strikt gelesen; ein Lesefehler
   *    wirft, statt eine Teilmenge als „alles" zu behandeln.
   *  - Eine leere Nummern-Menge wird als Fehler gewertet, nicht als „es gibt
   *    keine Events mehr".
   */
  public async collectOrphanRegistryNumbers(
    onRead?: (_loaded: number) => void,
  ): Promise<Array<{ email: string; eventNumber: number }>> {
    const valid = await this.readAllEventNumbersOrThrow();
    if (valid.size === 0) throw new Error('Keine Event-Nummern gefunden — Abbruch, statt alle Verweise als verwaist zu werten.');
    const all = await this.fetchAllParticipantsOrThrow(onRead);
    const out: Array<{ email: string; eventNumber: number }> = [];
    for (const p of all) {
      const em = (p.Email || '').trim().toLowerCase();
      if (!em) continue;
      const nums = `${p.EventRegistered || ''},${p.EventOnWaitlist || ''}`
        .split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n) && n > 0);
      const seen = new Set<number>();
      nums.forEach(n => { if (!valid.has(n) && !seen.has(n)) { seen.add(n); out.push({ email: em, eventNumber: n }); } });
    }
    return out;
  }

  /**
   * v29.0: Die von `analyzeRegistryAgainstLists` gefundenen Verweise aus dem
   * Register nehmen. Je E-Mail EIN Schreibvorgang, auch wenn mehrere Nummern
   * betroffen sind. Der Eintrag selbst bleibt stehen — er kann weitere,
   * gültige Events tragen.
   */
  public async pruneStaleRegistryNumbers(
    stale: Array<{ email: string; eventNumber: number }>,
    onProgress?: (_done: number, _total: number) => void,
  ): Promise<{ updated: number; removed: number; failed: number }> {
    const byEmail: Record<string, number[]> = {};
    stale.forEach(s => { (byEmail[s.email] = byEmail[s.email] || []).push(s.eventNumber); });
    const emails = Object.keys(byEmail);
    let updated = 0; let removed = 0; let failed = 0;
    for (let i = 0; i < emails.length; i++) {
      const em = emails[i];
      if (onProgress) onProgress(i + 1, emails.length);
      try {
        // eslint-disable-next-line no-await-in-loop
        const rec = await this.getParticipantByEmail(em);
        if (!rec) { failed += 1; continue; }
        const drop = new Set(byEmail[em].map(n => String(n)));
        const keep = (v?: string): string => (v || '').split(',').map(x => x.trim())
          .filter(x => x && !drop.has(x)).join(',');
        const nextReg = keep(rec.EventRegistered);
        const nextWait = keep(rec.EventOnWaitlist);
        // eslint-disable-next-line no-await-in-loop
        const resp = await this._merge(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${rec.Id})`,
          { 'EventRegistered': nextReg, 'EventOnWaitlist': nextWait },
        );
        if (resp.ok) { updated += 1; removed += byEmail[em].length; }
        else failed += 1;
      } catch { failed += 1; }
    }
    return { updated, removed, failed };
  }

  public async removeParticipantEvent(email: string, eventNumber: number): Promise<boolean> {
    try {
      const existing = await this.getParticipantByEmail(email);
      if (!existing) return false;

      const en = eventNumber.toString();
      const registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];
      const waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];

      await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
        {
          'EventRegistered': registered.join(','),
          'EventOnWaitlist': waitlist.join(','),
        }
      );
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Alle Teilnehmer laden (für Admin-Seite).
   */
  /**
   * v28.27: Das Register vollständig lesen — per ID-Fenster statt über den
   * „nextLink".
   *
   * Zwei Fallen stecken hier drin, und v28.25/26 sind in beide getreten:
   *  1. **Schwelle:** `$orderby` über eine nicht indizierte Spalte scheitert ab
   *     5000 Listenelementen mit HTTP 500. Nach `Id` zu sortieren ist dagegen
   *     immer erlaubt — die ID-Spalte ist von Haus aus indiziert.
   *  2. **nextLink:** SharePoint benennt den Folgeseiten-Link je nach
   *     ausgehandeltem OData-Format unterschiedlich (`odata.nextLink`,
   *     `@odata.nextLink`, `d.__next`). Wer nur eine Variante prüft, hält nach
   *     der ERSTEN Seite an und meldet fröhlich Vollständigkeit — genau das
   *     ließ die Dubletten-Prüfung „2000 Einträge geprüft" melden, obwohl die
   *     Liste ein Vielfaches davon enthält.
   *
   * Deshalb hier gar kein nextLink mehr: Wir holen aufsteigend nach `Id` und
   * setzen als Fenster `Id gt <letzte gelesene Id>`. Das ist deterministisch,
   * schwellenfest und formatunabhängig.
   *
   * @param strict wirft bei einem HTTP-Fehler, statt still eine unvollständige
   *   Liste zu liefern (für Abläufe, die aus dem Ergebnis auf „unbekannt"
   *   schließen).
   */
  private async readAllParticipants(strict: boolean, onPage?: (loaded: number) => void): Promise<SPParticipant[]> {
    const out: SPParticipant[] = [];
    const PAGE = 2000;
    const MAX_PAGES = 100; // Reißleine (200k Einträge) gegen Endlosschleifen
    let lastId = 0;
    for (let page = 0; page < MAX_PAGES; page++) {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items`
        + `?$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist`
        + `&$orderby=Id&$top=${PAGE}&$filter=${encodeURIComponent(`Id gt ${lastId}`)}`;
      let items: SPParticipant[] = [];
      try {
        const response = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) {
          if (strict) {
            throw new Error(`DEX_Participants nicht lesbar (HTTP ${response.status}). Bei mehr als 5000 Einträgen braucht die Spalte „Email" einen Index — die App versucht ihn beim Start automatisch zu setzen (erfordert „Listen verwalten").`);
          }
          break;
        }
        const data = await response.json();
        items = (data.value || data.d?.results || []) as SPParticipant[];
      } catch (err) {
        if (strict) throw err;
        break;
      }
      out.push(...items);
      // v28.29: Nach jeder Seite melden — das Register hat inzwischen mehrere
      // tausend Zeilen, der Lesevorgang dauert spürbar. Ohne Rückmeldung sah
      // die Wartungs-Kachel aus, als würde nichts passieren.
      if (onPage) { try { onPage(out.length); } catch { /* UI-Fehler nie durchreichen */ } }
      if (items.length < PAGE) break;
      const last = items[items.length - 1];
      if (!last || typeof last.Id !== 'number' || last.Id <= lastId) break; // Schutz vor Stillstand
      lastId = last.Id;
    }
    return out;
  }

  public async getAllParticipants(): Promise<SPParticipant[]> {
    const allItems = await this.readAllParticipants(false);
    allItems.sort((a, b) =>
      (a.Nachname || '').localeCompare(b.Nachname || '', 'de')
      || (a.Vorname || '').localeCompare(b.Vorname || '', 'de'));
    return allItems;
  }

  /**
   * v28.23: Teilnehmer-Register (DEX_Participants) für EIN Event nachziehen.
   *
   * DEX_Participants ist die zentrale „Schattenbuchhaltung": Pro Person stehen
   * dort die Event-Nummern, für die sie angemeldet ist bzw. auf der Warteliste
   * steht. Sie liegt auf der Haupt-Site und unterliegt NICHT der
   * Item-Level-Security der Teilnehmerlisten — deshalb ist sie die einzige
   * Quelle, die auch stellvertretend angelegte Anmeldungen zuverlässig kennt.
   * „Meine Events" startet von hier, und seit v28.22 hängt auch die
   * Doppel-Anmelde-Vorwarnung daran.
   *
   * Der Dual-Write bei jeder Anmeldung ist best-effort (`.catch(warn)`) —
   * schlägt er fehl (Netzwerk, Rechte, Timeout), fehlt der Eintrag dauerhaft.
   * Diese Methode gleicht ihn für die übergebene Teilnehmerliste ab: Sie
   * ergänzt fehlende Event-Nummern und korrigiert Einträge, die im falschen
   * Feld stehen (Warteliste ↔ angemeldet). Es wird NICHTS entfernt — für
   * abgemeldete Personen räumt der normale Abmelde-Pfad auf.
   */
  /**
   * v28.26: Zustand des zentralen Teilnehmer-Registers analysieren.
   *
   * Über die Jahre sammeln sich dort zwei Sorten Altlasten:
   *  - **Dubletten:** mehrere Einträge zur selben E-Mail. Sie entstehen, wenn
   *    der Lookup vor dem Schreiben scheitert (z.B. der HTTP-500-Fall aus
   *    v28.25) — dann legt die App einen zweiten Eintrag an, und ab da landen
   *    Anmeldungen mal im einen, mal im anderen. „Meine Events" zeigt dann je
   *    nach Treffer nur einen Teil der Events.
   *  - **Verwaiste Event-Nummern:** Anmeldungen zu Events, die es nicht mehr
   *    gibt. Beim Löschen eines Events räumt die App das Register NICHT mit
   *    auf. Harmlos (die Nummer läuft ins Leere), aber Ballast.
   */
  public async analyzeParticipantRegistry(
    validEventNumbers: number[],
    onRead?: (loaded: number) => void,
  ): Promise<{
    total: number; duplicateGroups: number; surplusRecords: number; orphanNumbers: number; noEmail: number;
  }> {
    const all = await this.fetchAllParticipantsOrThrow(onRead);
    const valid = new Set(validEventNumbers.filter(n => typeof n === 'number' && n > 0));
    const byEmail: Record<string, SPParticipant[]> = {};
    let noEmail = 0;
    let orphanNumbers = 0;
    for (const p of all) {
      const em = (p.Email || '').trim().toLowerCase();
      if (!em) { noEmail += 1; continue; }
      (byEmail[em] = byEmail[em] || []).push(p);
      if (valid.size > 0) {
        const nums = `${p.EventRegistered || ''},${p.EventOnWaitlist || ''}`
          .split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n) && n > 0);
        orphanNumbers += nums.filter(n => !valid.has(n)).length;
      }
    }
    let duplicateGroups = 0;
    let surplusRecords = 0;
    Object.keys(byEmail).forEach(em => {
      const n = byEmail[em].length;
      if (n > 1) { duplicateGroups += 1; surplusRecords += n - 1; }
    });
    return { total: all.length, duplicateGroups, surplusRecords, orphanNumbers, noEmail };
  }

  /**
   * v28.26: Dubletten im Teilnehmer-Register zusammenführen.
   *
   * Je E-Mail bleibt der ÄLTESTE Eintrag (kleinste Id) stehen und erhält die
   * VEREINIGUNG aller Event-Nummern; steht dieselbe Nummer bei einem Eintrag
   * als „angemeldet" und beim anderen als „Warteliste", gewinnt „angemeldet".
   * Name-Felder werden aus dem ersten nicht-leeren Wert aufgefüllt. Die
   * überzähligen Einträge werden danach gelöscht. Es gehen also KEINE
   * Anmelde-Informationen verloren — im Gegenteil, die zusammengeführte Zeile
   * kennt danach alle Events der Person.
   */
  public async mergeDuplicateParticipants(
    onProgress?: (done: number, total: number) => void,
    onRead?: (loaded: number) => void,
  ): Promise<{ groups: number; deleted: number; failed: number }> {
    const all = await this.fetchAllParticipantsOrThrow(onRead);
    const byEmail: Record<string, SPParticipant[]> = {};
    for (const p of all) {
      const em = (p.Email || '').trim().toLowerCase();
      if (!em) continue;
      (byEmail[em] = byEmail[em] || []).push(p);
    }
    const groups = Object.keys(byEmail).filter(em => byEmail[em].length > 1);
    let deleted = 0;
    let failed = 0;
    let done = 0;
    const parseNums = (s?: string): string[] => (s || '').split(',').map(x => x.trim()).filter(Boolean);
    for (const em of groups) {
      const recs = byEmail[em].slice().sort((a, b) => a.Id - b.Id);
      const keeper = recs[0];
      const registered = new Set<string>();
      const waitlist = new Set<string>();
      let vorname = '';
      let nachname = '';
      for (const r of recs) {
        parseNums(r.EventRegistered).forEach(n => registered.add(n));
        parseNums(r.EventOnWaitlist).forEach(n => waitlist.add(n));
        if (!vorname && (r.Vorname || '').trim()) vorname = (r.Vorname || '').trim();
        if (!nachname && (r.Nachname || '').trim()) nachname = (r.Nachname || '').trim();
      }
      // „Angemeldet" sticht „Warteliste" — dieselbe Nummer nie in beiden Feldern.
      registered.forEach(n => waitlist.delete(n));
      try {
        const m = await this._merge(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${keeper.Id})`,
          {
            'Vorname': vorname,
            'Nachname': nachname,
            'EventRegistered': Array.from(registered).join(','),
            'EventOnWaitlist': Array.from(waitlist).join(','),
          },
        );
        if (!m.ok) { failed += 1; done += 1; if (onProgress) onProgress(done, groups.length); continue; }
        for (const r of recs.slice(1)) {
          try {
            const resp = await this._sp.post(
              `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${r.Id})`,
              SPHttpClient.configurations.v1,
              { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'Accept': 'application/json;odata=nometadata' } },
            );
            if (resp.ok) deleted += 1; else failed += 1;
          } catch { failed += 1; }
        }
      } catch { failed += 1; }
      done += 1;
      if (onProgress) onProgress(done, groups.length);
    }
    return { groups: groups.length, deleted, failed };
  }

  /**
   * v28.25: Wie `getAllParticipants`, wirft aber bei einem HTTP-Fehler, statt
   * still eine unvollständige Liste zu liefern. Für Abläufe, die aus dem
   * Ergebnis auf „Person ist unbekannt" schließen (Register-Abgleich).
   */
  private async fetchAllParticipantsOrThrow(onPage?: (loaded: number) => void): Promise<SPParticipant[]> {
    return this.readAllParticipants(true, onPage);
  }

  public async backfillParticipantRegistry(
    subsiteUrl: string,
    eventNumber: number,
    onProgress?: (done: number, total: number) => void,
  ): Promise<{ active: number; fixed: number; failed: number }> {
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const regs = await this.getAllRegistrations(subsiteUrl);
    const active = regs.filter(r => ACTIVE.indexOf(r.Status || '') >= 0
      && (r.ParticipantEmail || '').indexOf('@') > 0);
    if (active.length === 0 || !eventNumber) return { active: 0, fixed: 0, failed: 0 };
    // v28.25: Register EINMAL laden — und einen Lesefehler NICHT verschlucken.
    // Wäre die Liste nicht lesbar (z.B. 5000-Element-Schwelle) und wir liefen
    // trotzdem weiter, hielte der Abgleich jede Person für unbekannt und legte
    // reihenweise Dubletten an. Lieber sauber abbrechen.
    const all = await this.fetchAllParticipantsOrThrow();
    const byEmail: Record<string, SPParticipant> = {};
    for (const p of all) {
      const e = (p.Email || '').trim().toLowerCase();
      if (e) byEmail[e] = p;
    }
    const en = String(eventNumber);
    const has = (field: string | undefined): boolean =>
      (field || '').split(',').map(s => s.trim()).indexOf(en) >= 0;
    let fixed = 0;
    let failed = 0;
    let done = 0;
    for (const r of active) {
      const em = (r.ParticipantEmail || '').trim().toLowerCase();
      const rec = byEmail[em];
      const wantWaitlist = r.Status === 'Warteliste';
      const alreadyRight = !!rec && (wantWaitlist ? has(rec.EventOnWaitlist) : has(rec.EventRegistered));
      if (!alreadyRight) {
        // Bekannten Register-Eintrag direkt mitgeben — spart die Einzelabfrage
        // pro Person (die an der Schwelle ohnehin scheitern könnte).
        const ok = await this.applyParticipantEvent(
          rec || null, r.Vorname || '', r.Nachname || '', r.ParticipantEmail,
          eventNumber, wantWaitlist ? 'Warteliste' : 'Angemeldet',
        );
        if (ok) fixed += 1; else failed += 1;
      }
      done += 1;
      if (onProgress && (done % 10 === 0 || done === active.length)) onProgress(done, active.length);
    }
    return { active: active.length, fixed, failed };
  }

  // ==================== DEX_Events Liste ====================

  /**
   * Events-Liste erstellen falls nicht vorhanden
   */
  public async ensureEventsList(): Promise<void> {
    const listName = 'DEX_Events';
    const exists = await this.listExists(listName);
    if (exists) {
      await this.ensureMissingFields(listName);

      // Default-View komplett neu aufbauen: ID, Title, EventImageUrl, dann Rest
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
          {}
        );
      } catch { /* ignore */ }
      await this.configureDefaultView(listName, [
        'ID', 'LinkTitle', 'EventImageUrl',
        'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
        'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
        'WaitlistEnabled', 'Organizer', 'DisableEmails', 'DisableOutlook',
        'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
      ], undefined, { rebuild: true });
      await this.setColumnFormatting(listName, 'EventImageUrl', {
        '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        'elmType': 'img',
        'attributes': { 'src': '@currentField' },
        'style': { 'max-height': '60px', 'max-width': '120px', 'border-radius': '6px', 'box-shadow': '0 1px 3px rgba(0,0,0,0.15)' },
      });
      try {
        const listInfo = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
          SPHttpClient.configurations.v1
        );
        if (listInfo.ok) {
          const data = await listInfo.json();
          if (!data.HasUniqueRoleAssignments) {
            await this.setEventsListPermissions(listName);
          }
        }
      } catch { /* ignore */ }
      return;
    }

    // Liste erstellen
    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Events der DEX Event Experience Platform',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // Spalten hinzufügen
    const fields = this.getEventsFieldDefinitions();
    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        payload['Choices'] = { 'results': f.choices };
      }
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    await this.configureDefaultView(listName, [
      'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
      'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
      'WaitlistEnabled', 'Organizer', 'EventImageUrl', 'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
    ], undefined, { rebuild: true });
    await this.setColumnFormatting(listName, 'EventImageUrl', {
      '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
      'elmType': 'img',
      'attributes': { 'src': '@currentField' },
      'style': { 'max-height': '50px', 'max-width': '120px', 'border-radius': '4px' },
    });

    await this.setEventsListPermissions(listName);
  }

  /**
   * Feld-Definitionen für DEX_Events Liste
   */
  private getEventsFieldDefinitions(): Array<{ title: string; type: number; choices?: string[]; metaType?: string; richText?: boolean; numberOfLines?: number }> {
    return [
      { title: 'EventStatus', type: 6, choices: ['Under Construction', 'Active', 'Completed', 'Cancelled'], metaType: 'SP.FieldChoice' },
      // EventType-Spalte ab v5.2 deprecated (Feld wird nicht mehr angelegt/aktualisiert).
      // Typ wird beim Laden aus CustomFields abgeleitet. Bestehende Spalte in DEX_Events
      // kann manuell entfernt werden.
      { title: 'Description', type: 3 },
      { title: 'Location', type: 2 },
      { title: 'LocationAddress', type: 2 }, // JSON-String: { street, houseNo, zip, city }
      { title: 'LocationFilter', type: 2 },
      // Audience ist Multi-Line-Text (Note) damit es auch bei 100+ E-Mail-Adressen
      // nicht abgeschnitten wird (Single-Line-Text ist auf 255 Zeichen limitiert).
      // Für bestehende Events siehe upgradeAudienceFieldToNote() — migriert die
      // alte Text-Spalte zu Note ohne Datenverlust.
      { title: 'Audience', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      // v16.4: Vor-aufgelöste E-Mails der Audience-DLs (Multi-Line, ';'-
      // separiert). Wird beim Event-Save vom EventCreationPage-Flow
      // gesetzt; matchesAudience im EventListPage checkt zusätzlich
      // gegen diese Liste. Damit funktioniert die Sichtbarkeit auch für
      // verschachtelte DLs, die /me/memberOf nicht zurückliefert.
      { title: 'AudienceResolvedEmails', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 8 },
      { title: 'FilterMode', type: 6, choices: ['AND', 'OR'], metaType: 'SP.FieldChoice' },
      { title: 'StartDate', type: 4 },
      { title: 'EndDate', type: 4 },
      { title: 'RegistrationDeadline', type: 4 },
      { title: 'LastDeregisterDate', type: 4 },
      { title: 'MaxParticipants', type: 9 },
      // v26.63: Denormalisierte aktuelle Teilnehmerzahl am Event-Item. Wird von
      // Organizern/Admins gepflegt (nur die haben Schreibrechte auf DEX_Events)
      // — beim Laden der echten Zahl best-effort persistiert. So ist die
      // Teilnehmerzahl pro Event ohne Subsite-Scan aus DEX_Events lesbar.
      { title: 'CurrentParticipants', type: 9 },
      { title: 'WaitlistEnabled', type: 8 },
      { title: 'MandatoryRegistration', type: 8 },
      // v26.55: EventImageUrl ist Note (mehrzeilig) statt Single-Line-Text —
      // SharePoint-Asset-URLs (SiteAssets-Pfad + langer Original-Dateiname)
      // überschreiten real das 255-Zeichen-Limit (MD Academy: 261 Zeichen) und
      // ließen den kompletten Event-Save mit „Invalid text value" abbrechen.
      // Bestands-Listen migriert upgradeOverflowTextFieldsToNote() beim Boot
      // bzw. die Selbstheilung in updateEvent() beim ersten fehlschlagenden Save.
      { title: 'EventImageUrl', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 2 },
      { title: 'EmailImageBase64', type: 3 }, // Base64 Event-Bild für E-Mails/Outlook (Flow ersetzt {{ORB_URL}})
      // Organizer + OrganizerEmail sind Multi-Line-Text (Note) damit sie auch bei
      // 10+ Co-Organizern nicht abgeschnitten werden (Single-Line-Text ist auf 255
      // Zeichen limitiert — bei ~17 Personen mit Format `vorname.nachname@deloitte.de;`
      // wird das überschritten und SP antwortet mit „Invalid text value" beim Update).
      // Für bestehende Events siehe upgradeOrganizerFieldsToNote() — migriert die
      // alten Text-Spalten ohne Datenverlust.
      { title: 'Organizer', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      { title: 'OrganizerEmail', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      // v10.16: Optionaler Ansprechpartner pro Event (Anzeige-Feld, kein App-Login).
      { title: 'ContactName', type: 2 },
      { title: 'ContactEmail', type: 2 },
      { title: 'ContactInfo', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      { title: 'EventNumber', type: 9 },
      { title: 'OutlookEventId', type: 2 },
      { title: 'CalendarLink', type: 2 },
      { title: 'OutlookBody', type: 3 }, // Multiline - Text für Outlook-Termin
      { title: 'OutlookSubject', type: 2 }, // v18.42: Single line - Betreff des Outlook-Termins (leer = Titel)
      { title: 'OutlookStart', type: 4 }, // v18.44: DateTime - abweichende Start-Zeit (leer = Event-Start)
      { title: 'OutlookEnd', type: 4 },   // v18.44: DateTime - abweichende End-Zeit (leer = Event-Ende)
      { title: 'OutlookLocation', type: 2 }, // v18.34: Single line - lesbarer Ort für den Outlook-Termin
      // v29.52: Boolean - ganztägiger Termin. Echte Spalte (kein Piggyback in
      // EmailTemplateOverrides), weil der Outlook-Flow den Wert direkt über
      // triggerBody()?['AllDay'] lesen muss — JSON parsen kann er dort nicht.
      { title: 'AllDay', type: 8, metaType: 'SP.Field' },
      // v29.54: Boolean - Termin als „Frei" statt „Beschäftigt" anzeigen.
      // NEGATIV benannt, damit bestehende Einträge (leer/false) weiter
      // beschäftigt bleiben — siehe Kommentar an DeloitteEvent.showAsFree.
      { title: 'ShowAsFree', type: 8, metaType: 'SP.Field' },
      // v30.26: Boolean — Outlook-Termin als echte Teams-Besprechung anlegen
      // (Flow „Create event (V4)": Is online meeting + Provider Teams).
      // Leer/false = wie bisher KEIN automatisches Meeting; der Organizer
      // entscheidet das pro Event im Wizard (Ort → Online-Meeting).
      { title: 'OutlookIsOnlineMeeting', type: 8, metaType: 'SP.Field' },
      // v29.55: Boolean - Organizer nicht in requiredAttendees des Outlook-
      // Termins. Ebenfalls negativ, damit Bestandsevents unverändert bleiben.
      { title: 'SkipOrganizerInvite', type: 8, metaType: 'SP.Field' },
      { title: 'EmailLanguage', type: 2 }, // DE oder EN
      { title: 'RegistrationLanguage', type: 2 }, // v18.35: erzwungene Anmeldeseiten-Sprache ('de'|'en'|'')
      { title: 'EmailTemplateOverrides', type: 3 }, // JSON mit Event-spezifischen Template-Anpassungen
      { title: 'DisableEmails', type: 8, metaType: 'SP.Field' }, // Boolean - keine E-Mails versenden
      { title: 'DisableRegistrationEmail', type: 8, metaType: 'SP.Field' }, // v19.21 Boolean - keine Anmelde-Bestätigung
      { title: 'DisableCancellationEmail', type: 8, metaType: 'SP.Field' }, // v19.21 Boolean - keine Abmelde-Bestätigung
      { title: 'AutoDeregisterOnDecline', type: 8, metaType: 'SP.Field' }, // v19.23 Boolean - Outlook-Absage = Auto-Abmeldung
      { title: 'InactiveHandling', type: 2, metaType: 'SP.Field' }, // v26.40 Text - 'notify' | 'autoderegister' bei Ex-Deloitte-Konten
      { title: 'DisableOutlook', type: 8, metaType: 'SP.Field' }, // Boolean - keine Outlook-Kalendereinträge
      { title: 'OutlookDirty', type: 8, metaType: 'SP.Field' }, // v11.57 Boolean - Outlook-Update ausstehend nach Bearbeitung
      { title: 'AutoSendQRCode', type: 8, metaType: 'SP.Field' }, // v9.15 Boolean - QR-Code automatisch nach Anmeldung versenden
      { title: 'ActiveFrom', type: 4, metaType: 'SP.Field' }, // v9.21 DateTime - Auto-Aktivierungs-Datum
      { title: 'NotifyOrgRegisterMode', type: 6, choices: ['Never', 'Always', 'FromDate'], metaType: 'SP.FieldChoice' }, // v8.5
      { title: 'NotifyOrgRegisterFromDate', type: 4 }, // v8.5: ISO-Date, nur für Mode='FromDate'
      { title: 'NotifyOrgCancelMode', type: 6, choices: ['Never', 'Always', 'AfterDeadline'], metaType: 'SP.FieldChoice' }, // v8.5
      { title: 'ExcludedUsers', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 }, // v8.6: explizit ausgeschlossene User
      { title: 'IsFictive', type: 8, metaType: 'SP.Field' }, // Boolean - Test-Event (nur Admin + eigene Organizer sichtbar)
      { title: 'DurchstarterCapacity', type: 9 }, // Split-Capacity Gruppe A (historisch B2Run-Durchstarter)
      { title: 'FunstarterCapacity', type: 9 }, // Split-Capacity Gruppe B (historisch B2Run-Funstarter)
      { title: 'SplitLabelA', type: 2 }, // v10.20: frei wählbare Bezeichnung Gruppe A (Single line text)
      { title: 'SplitLabelB', type: 2 }, // v10.20: frei wählbare Bezeichnung Gruppe B (Single line text)
      { title: 'SplitDescA', type: 3 }, // v26.72: Beschreibung Gruppe A (Note/mehrzeilig)
      { title: 'SplitDescB', type: 3 }, // v26.72: Beschreibung Gruppe B (Note/mehrzeilig)
      { title: 'SplitHelpText', type: 3 }, // v26.83: Hinweistext über der Gruppen-Auswahl (Note/mehrzeilig)
      { title: 'SplitSectionTitle', type: 2 }, // v26.83: frei wählbare Überschrift der Gruppen-Auswahl (Single line)
      { title: 'SplitSharedWaitlist', type: 8, metaType: 'SP.Field' }, // v10.20: Boolean - true = gemeinsame Warteliste
      { title: 'AllowAttendeeUpload', type: 8, metaType: 'SP.Field' }, // v11.0: Boolean - Teilnehmer-PDF-Upload erlauben
      { title: 'AttendeeUploadHint', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 }, // v11.0: Hinweistext
      { title: 'AttendeeUploadLabel', type: 2 }, // v11.0: Single-line Label für den Upload-Block in MyEvents
      { title: 'AskSalutation', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Anrede im Registrierungsformular abfragen
      { title: 'ContactOrganizerEmail', type: 2 }, // v26.18: E-Mail des als Hauptkontakt markierten Organizers (grün hervorgehoben auf der Anmeldeseite)
      { title: 'ConfirmDialogEnabled', type: 8, metaType: 'SP.Field' }, // v18.75: Boolean - Sicherheitshinweis vor dem Absenden anzeigen
      { title: 'ConfirmDialogMode', type: 2 }, // v18.75: Single line text - 'summary' (Auswahl-Übersicht) | 'freetext'
      { title: 'ConfirmDialogText', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 }, // v18.75: Note - Freitext-Hinweis
      { title: 'SelfCheckInEnabled', type: 8, metaType: 'SP.Field' }, // v18.33: Boolean - Self-Check-in per QR-Code erlauben
      { title: 'SelfCheckInToken', type: 2 }, // v18.33: Single line text - geheimer Token (statischer Link + HMAC-Schlüssel)
      { title: 'SelfCheckInFrom', type: 4 }, // v18.33: DateTime - optionaler Start des Check-in-Fensters
      { title: 'SelfCheckInTo', type: 4 }, // v18.33: DateTime - optionales Ende des Check-in-Fensters
      { title: 'TeamRegistrationEnabled', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Team-Anmeldung erlauben
      { title: 'TeamSize', type: 9 }, // v11.80: Number - Maximale Teamgröße (0 = nicht gesetzt)
      { title: 'AskTeamName', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Team-Name abfragen
      { title: 'TeamPartialAllowed', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - Auch Teil-Teams erlauben
      { title: 'TeamOpenSlotsVisible', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - offene Slots öffentlich sichtbar
      { title: 'TeamJoinRequiresApproval', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - Lead muss Beitritt bestätigen
      { title: 'BilingualFields', type: 8, metaType: 'SP.Field' }, // v17.20: Boolean - Custom-Fields zweisprachig (DE + EN) anbieten
      { title: 'CustomFields', type: 3 },
      { title: 'Agenda', type: 3 }, // JSON-Array mit Agenda-Einträgen
      { title: 'Transfers', type: 3 }, // JSON-Array mit Transferzeiten
      { title: 'Documents', type: 3 }, // JSON-Array mit Dokumenten
      { title: 'FunZone', type: 3 }, // JSON-Array mit Quiz-Fragen
      { title: 'QuizClusterSize', type: 9 }, // Number - 1..4 Fragen pro Quiz-Ansicht
      { title: 'ParentEventId', type: 2 }, // Seit v6.4: ID des Parent-Events (wenn dies ein Sub-Event ist)
      { title: 'RegistrationListName', type: 2 },
      { title: 'RegistrationListUrl', type: 2 },
      { title: 'SubsiteUrl', type: 2 },
    ];
  }

  /**
   * Fehlende Spalten auf einer bestehenden DEX_Events-Liste nachträglich hinzufügen.
   */
  private async ensureMissingFields(listName: string): Promise<void> {
    const requiredFields = this.getEventsFieldDefinitions();

    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$filter=Hidden eq false&$top=200`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;

      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const existingFields = new Set((data.value || []).map((f: any) => f.InternalName));

      for (const f of requiredFields) {
        if (!existingFields.has(f.title)) {
          // Fehlende Spalte nachträglich hinzufügen
          const payload: Record<string, unknown> = {
            '__metadata': { 'type': f.metaType || 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          };
          if (f.choices) {
            payload['Choices'] = { 'results': f.choices };
          }
          if (f.metaType === 'SP.FieldMultiLineText') {
            payload['RichText'] = !!f.richText;
            if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
          }
          try {
            await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
          } catch {
            console.warn('[DEX] Konnte Spalte nicht hinzufügen:', f.title);
          }
        }
      }
    } catch (e) {
      console.warn('[DEX] ensureMissingFields Error:', e);
    }
  }

  /**
   * Migration: alte Audience-Spalte (Type 2, Single-Line-Text, 255 Zeichen Limit)
   * auf Multi-Line-Text (Type 3, Plain-Text, 63.999 Zeichen) umstellen.
   *
   * Nötig weil bei Zielgruppen mit vielen Email-Adressen (~10+) der 255-Zeichen-
   * Cutoff schon griff und Adressen stumm abgeschnitten wurden.
   *
   * Ablauf (idempotent):
   *   1. Check TypeAsString der Audience-Spalte. Wenn schon 'Note' -> skip.
   *   2. Backup aller Event-Werte (id -> audience) im Speicher.
   *   3. Alte Spalte löschen.
   *   4. Neue Spalte als SP.FieldMultiLineText (RichText=false, NumberOfLines=4) anlegen.
   *   5. Werte aus dem Backup zurückschreiben (MERGE pro Event).
   *
   * Läuft beim App-Start (nur für Admins, weil wir Write-Rechte auf DEX_Events brauchen).
   */
  public async upgradeAudienceFieldToNote(): Promise<void> {
    const listName = 'DEX_Events';
    try {
      // 1. TypeAsString abfragen
      const fieldResp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')?$select=TypeAsString,FieldTypeKind`,
        SPHttpClient.configurations.v1
      );
      if (!fieldResp.ok) return;
      const fieldData = await fieldResp.json();
      const typeAsString: string = fieldData.TypeAsString || fieldData.d?.TypeAsString || '';
      const fieldTypeKind: number = fieldData.FieldTypeKind ?? fieldData.d?.FieldTypeKind ?? 0;
      if (typeAsString === 'Note' || fieldTypeKind === 3) {
        // Schon migriert
        return;
      }
      if (typeAsString !== 'Text' && fieldTypeKind !== 2) {
        // Unerwarteter Typ - nicht anfassen
        console.warn(`[DEX] upgradeAudienceFieldToNote: Audience hat unerwarteten Typ '${typeAsString}' (kind=${fieldTypeKind}) - skip.`);
        return;
      }

      // 2. Alle Event-Werte laden und backuppen
      const itemsResp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Audience&$top=2000`,
        SPHttpClient.configurations.v1
      );
      if (!itemsResp.ok) return;
      const itemsData = await itemsResp.json();
      const items: Array<{ Id: number; Audience: string }> = itemsData.value || itemsData.d?.results || [];
      const backup: Record<number, string> = {};
      for (const it of items) {
        if (it.Audience) backup[it.Id] = it.Audience;
      }
      console.warn(`[DEX] upgradeAudienceFieldToNote: Backup ${Object.keys(backup).length} von ${items.length} Event-Audience-Werten.`);

      // 3. Alte Spalte löschen
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')/deleteObject`,
          {}
        );
      } catch (e) {
        console.warn('[DEX] upgradeAudienceFieldToNote: Delete alte Audience-Spalte fehlgeschlagen, Migration abgebrochen:', e);
        return;
      }

      // 4. Neue Spalte als Multi-Line-Text anlegen
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          {
            '__metadata': { 'type': 'SP.FieldMultiLineText' },
            'Title': 'Audience',
            'FieldTypeKind': 3,
            'Required': false,
            'RichText': false,
            'NumberOfLines': 4,
          }
        );
      } catch (e) {
        console.error('[DEX] upgradeAudienceFieldToNote: Konnte neue Audience-Note-Spalte nicht anlegen - Daten könnten verloren gehen:', e, backup);
        return;
      }

      // 5. Werte zurückschreiben per _merge (odata=nometadata, daher kein __metadata im Body nötig)
      let restored = 0;
      let failed = 0;
      for (const idStr of Object.keys(backup)) {
        const id = Number(idStr);
        try {
          const resp = await this._merge(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
            { 'Audience': backup[id] }
          );
          if (resp.ok) restored += 1;
          else failed += 1;
        } catch { failed += 1; }
      }
      console.warn(`[DEX] upgradeAudienceFieldToNote: Migration fertig — ${restored} Werte zurückgeschrieben, ${failed} Fehler.`);
    } catch (e) {
      console.warn('[DEX] upgradeAudienceFieldToNote Error:', e);
    }
  }

  /**
   * Migration: alte `Organizer` + `OrganizerEmail`-Spalten (Type 2, Single-Line-Text,
   * 255 Zeichen Limit) auf Multi-Line-Text (Type 3, Plain-Text, 63.999 Zeichen) umstellen.
   *
   * Nötig weil bei Events mit 10+ Co-Organizern der 255-Zeichen-Cutoff greift und
   * SharePoint beim Update mit „Invalid text value. A text field contains invalid data."
   * (HTTP 500, Microsoft.SharePoint.SPException) antwortet — der Save bricht komplett ab.
   *
   * Beispiel-Overflow: 17 × `vorname.nachname@deloitte.de;` ≈ 425 Zeichen.
   *
   * Ablauf pro Feld (idempotent, parallel für beide Felder):
   *   1. Check TypeAsString. Wenn schon 'Note' -> skip.
   *   2. Backup aller Event-Werte (id -> wert) im Speicher.
   *   3. Alte Spalte löschen.
   *   4. Neue Spalte als SP.FieldMultiLineText (RichText=false, NumberOfLines=4) anlegen.
   *   5. Werte aus dem Backup zurückschreiben (MERGE pro Event).
   *
   * Läuft beim App-Start (nur für Admins, weil wir Write-Rechte auf DEX_Events brauchen).
   */
  public async upgradeOrganizerFieldsToNote(): Promise<void> {
    await this._upgradeTextFieldToNote('DEX_Events', 'Organizer');
    await this._upgradeTextFieldToNote('DEX_Events', 'OrganizerEmail');
  }

  /**
   * v26.53/v26.55: Migration überlauf-gefährdeter Single-Line-Text-Spalten auf
   * Multi-Line-Text (Note). Auf Bestands-Listen liegen diese Spalten als
   * einzeiliger Text (255-Zeichen-Limit) — längere Werte ließen den kompletten
   * Event-Save mit „Invalid text value. A text field contains invalid data."
   * (HTTP 500) abbrechen. Konkrete Fälle:
   *  - ConfirmDialogText (v26.53): Freitext des Bestätigungs-Dialogs, z. B.
   *    Stornoregeln mit ~450 Zeichen.
   *  - EventImageUrl (v26.55): SiteAssets-Bild-URL mit langem Original-
   *    Dateinamen, real 261 Zeichen (MD Academy 2026).
   * Idempotent: ist eine Spalte schon Note, passiert nichts. Zusätzlich heilt
   * updateEvent() solche Spalten seit v26.54 auch beim fehlschlagenden Save
   * selbst (Boot-Ensure läuft nur einmal pro Version — s. ENSURE_FLAG_KEY).
   */
  public async upgradeOverflowTextFieldsToNote(): Promise<void> {
    await this._upgradeTextFieldToNote('DEX_Events', 'ConfirmDialogText');
    await this._upgradeTextFieldToNote('DEX_Events', 'EventImageUrl');
  }

  /**
   * v26.57: Best-effort-Check, ob eine Person bereits Zugriff auf die Site
   * hat (mindestens Seiten ansehen). Genutzt vor der „SharePoint-Zugriff
   * benötigt"-Admin-Mail bei internationalen Zielgruppen-Personen — wer schon
   * berechtigt ist, taucht in der Mail nicht mehr auf.
   *
   * Ablauf: LoginName über siteusers auflösen (für Gäste weicht er mit
   * #EXT#-Format von der Mail ab), dann getusereffectivepermissions und das
   * ViewPages-Bit prüfen. Gruppen-basierte Rechte (z. B. „Deloitte DE ALL")
   * löst SharePoint dabei serverseitig mit auf.
   *
   * Rückgabe: true = hat Zugriff · false = sicher kein Zugriff · null = nicht
   * prüfbar (User unbekannt, keine Enumerate-Permissions-Rechte des Aufrufers,
   * Netzwerkfehler). Aufrufer behandeln null wie „kein Zugriff" — lieber
   * einmal zu viel benachrichtigen als eine Freigabe verpassen.
   */
  public async userHasSiteAccess(email: string): Promise<boolean | null> {
    const mail = (email || '').trim();
    if (!mail) return null;
    try {
      let login = `i:0#.f|membership|${mail.toLowerCase()}`;
      try {
        const resp = await this._sp.get(
          `${this.siteUrl}/_api/web/siteusers?$filter=Email eq '${encodeURIComponent(mail.replace(/'/g, "''"))}'&$select=LoginName&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (resp.ok) {
          const data = await resp.json();
          const item = (data.value || data.d?.results || [])[0];
          if (item && item.LoginName) login = item.LoginName;
        }
      } catch { /* Fallback auf membership-Claim */ }
      const permResp = await this._sp.get(
        `${this.siteUrl}/_api/web/getusereffectivepermissions(@u)?@u='${encodeURIComponent(login.replace(/'/g, "''"))}'`,
        SPHttpClient.configurations.v1
      );
      if (!permResp.ok) return null;
      const perm = await permResp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const raw: any = (perm && (perm as any).GetUserEffectivePermissions) || (perm as any).d?.GetUserEffectivePermissions || perm;
      const low = Number(raw?.Low || 0);
      // ViewPages = 0x20000 — reicht, um die App-Seite zu öffnen.
      return (low & 0x20000) !== 0;
    } catch {
      return null;
    }
  }

  /**
   * Generischer Helper: migriert ein einzelnes Single-Line-Text-Feld einer Liste auf
   * Multi-Line-Text (Note). Idempotent — wenn das Feld schon Note ist, no-op. Wenn das
   * Feld einen anderen Typ hat (Choice/Number/etc.), no-op mit Warnung.
   *
   * Wird von `upgradeAudienceFieldToNote()` (existierendes Audience-Feld) und
   * `upgradeOrganizerFieldsToNote()` (Organizer + OrganizerEmail) genutzt. Die alte
   * `upgradeAudienceFieldToNote()`-Implementierung ist aus Kompatibilitätsgründen
   * unberührt geblieben — neue Migrationen sollten diesen Helper nutzen.
   */
  private async _upgradeTextFieldToNote(listName: string, fieldName: string): Promise<void> {
    const tag = `_upgradeTextFieldToNote(${listName}.${fieldName})`;
    try {
      const fieldResp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${fieldName}')?$select=TypeAsString,FieldTypeKind`,
        SPHttpClient.configurations.v1
      );
      if (!fieldResp.ok) return;
      const fieldData = await fieldResp.json();
      const typeAsString: string = fieldData.TypeAsString || fieldData.d?.TypeAsString || '';
      const fieldTypeKind: number = fieldData.FieldTypeKind ?? fieldData.d?.FieldTypeKind ?? 0;
      if (typeAsString === 'Note' || fieldTypeKind === 3) return;
      if (typeAsString !== 'Text' && fieldTypeKind !== 2) {
        console.warn(`[DEX] ${tag}: unerwarteter Typ '${typeAsString}' (kind=${fieldTypeKind}) — skip.`);
        return;
      }

      const itemsResp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,${fieldName}&$top=2000`,
        SPHttpClient.configurations.v1
      );
      if (!itemsResp.ok) return;
      const itemsData = await itemsResp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const items: Array<any> = itemsData.value || itemsData.d?.results || [];
      const backup: Record<number, string> = {};
      for (const it of items) {
        const v = it[fieldName];
        if (v) backup[it.Id] = v;
      }
      console.warn(`[DEX] ${tag}: Backup ${Object.keys(backup).length} von ${items.length} Werten.`);

      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${fieldName}')/deleteObject`,
          {}
        );
      } catch (e) {
        console.warn(`[DEX] ${tag}: Delete alte Spalte fehlgeschlagen, Migration abgebrochen:`, e);
        return;
      }

      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          {
            '__metadata': { 'type': 'SP.FieldMultiLineText' },
            'Title': fieldName,
            'FieldTypeKind': 3,
            'Required': false,
            'RichText': false,
            'NumberOfLines': 4,
          }
        );
      } catch (e) {
        console.error(`[DEX] ${tag}: Konnte neue Note-Spalte nicht anlegen — Daten könnten verloren gehen:`, e, backup);
        return;
      }

      let restored = 0;
      let failed = 0;
      for (const idStr of Object.keys(backup)) {
        const id = Number(idStr);
        try {
          const resp = await this._merge(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
            { [fieldName]: backup[id] }
          );
          if (resp.ok) restored += 1;
          else failed += 1;
        } catch { failed += 1; }
      }
      console.warn(`[DEX] ${tag}: Migration fertig — ${restored} Werte zurückgeschrieben, ${failed} Fehler.`);
    } catch (e) {
      console.warn(`[DEX] ${tag} Error:`, e);
    }
  }

  /**
   * Berechtigungen für DEX_Events setzen
   */
  private async setEventsListPermissions(listName: string): Promise<void> {
    try {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      const ownersResponse = await this._sp.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
          {}
        );
      }

      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
          {}
        );
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden
    }
  }

  /**
   * Default View einer Liste konfigurieren
   */
  public async configureDefaultView(listName: string, fieldNames: string[], baseUrl?: string, opts?: { rebuild?: boolean }): Promise<void> {
    const url = baseUrl || this.siteUrl;
    try {
      let existingFields: string[] = [];
      if (opts?.rebuild) {
        // Komplett neu aufbauen — SP-Defaults (Modified, Created, ID, Type,
        // Compliance-Tag, App Created By, ...) werden rausgeworfen.
        try {
          await this._post(
            `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
            {}
          );
        } catch { /* ignore */ }
      }
      // IMMER die tatsächlich noch vorhandenen View-Felder lesen — auch im
      // Rebuild-Pfad. Hintergrund: SharePoints `addviewfield` ist NICHT
      // idempotent und antwortet mit HTTP 500, wenn das Feld bereits in der
      // View liegt. Greift `removeallviewfields` nicht (Permission/Throttle),
      // blieben die alten Felder drin → 500-Rauschen bei jedem Boot (z.B.
      // `addviewfield('DisableOutlook')`). Mit dem Re-Read überspringen wir
      // bereits vorhandene Felder und senden den fehlschlagenden Request gar
      // nicht erst ab.
      try {
        const existingResponse = await this._sp.get(
          `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields`,
          SPHttpClient.configurations.v1
        );
        if (existingResponse.ok) {
          const existingData = await existingResponse.json();
          if (existingData.Items) existingFields = existingData.Items;
          else if (existingData.d?.Items) existingFields = existingData.d.Items;
          else if (existingData.value) existingFields = existingData.value;
        }
      } catch { /* ignore — dann werden ggf. alle Felder versucht */ }

      for (const fieldName of fieldNames) {
        // Nur hinzufügen wenn noch nicht in der View
        if (existingFields.indexOf(fieldName) < 0) {
          await this._post(
            `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/addviewfield('${fieldName}')`,
            {}
          );
        }
      }
    } catch {
      // View-Konfiguration ist optional
    }
  }

  /**
   * Column Formatting auf ein Feld setzen (z.B. Bild-Vorschau für URL-Spalten)
   */
  private async setColumnFormatting(listName: string, fieldName: string, formatJson: object): Promise<void> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq '${fieldName}'&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;
      const data = await response.json();
      const field = data.value?.[0];
      if (!field) return;

      await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields('${field.Id}')`,
        { CustomFormatter: JSON.stringify(formatJson) }
      );
      // Column Formatting gesetzt
    } catch {
      // Column Formatting ist optional
    }
  }

  // ==================== Events CRUD ====================

  private static readonly EVENT_SELECT = 'Id,Title,EventStatus,EventNumber,Description,Location,LocationAddress,LocationFilter,Audience,AudienceResolvedEmails,FilterMode,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,CurrentParticipants,WaitlistEnabled,MandatoryRegistration,EventImageUrl,EmailImageBase64,Organizer,OrganizerEmail,ContactName,ContactEmail,ContactOrganizerEmail,ContactInfo,OutlookEventId,CalendarLink,OutlookBody,OutlookSubject,OutlookStart,OutlookEnd,OutlookLocation,AllDay,ShowAsFree,SkipOrganizerInvite,EmailLanguage,RegistrationLanguage,EmailTemplateOverrides,DisableEmails,DisableRegistrationEmail,DisableCancellationEmail,AutoDeregisterOnDecline,InactiveHandling,DisableOutlook,OutlookDirty,AutoSendQRCode,ActiveFrom,NotifyOrgRegisterMode,NotifyOrgRegisterFromDate,NotifyOrgCancelMode,ExcludedUsers,IsFictive,DurchstarterCapacity,FunstarterCapacity,SplitLabelA,SplitLabelB,SplitDescA,SplitDescB,SplitHelpText,SplitSectionTitle,SplitSharedWaitlist,AllowAttendeeUpload,AttendeeUploadHint,AttendeeUploadLabel,AskSalutation,ConfirmDialogEnabled,ConfirmDialogMode,ConfirmDialogText,SelfCheckInEnabled,SelfCheckInToken,SelfCheckInFrom,SelfCheckInTo,TeamRegistrationEnabled,TeamSize,AskTeamName,TeamPartialAllowed,TeamOpenSlotsVisible,TeamJoinRequiresApproval,BilingualFields,CustomFields,Agenda,Transfers,Documents,FunZone,QuizClusterSize,ParentEventId,RegistrationListName,SubsiteUrl,Modified,Created';

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

  /**
   * Seed-Events anlegen falls sie nicht existieren (einmalig beim ersten Start).
   */
  public async seedEvents(): Promise<void> {
    try {
      // Prüfen ob "Assistenz Meeting 2026" schon existiert
      const check = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=Title eq 'Assistenz Meeting 2026'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (check.ok) {
        const data = await check.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0) return; // Existiert bereits
      }

      // Event anlegen
      await this.createEvent({
        title: 'Assistenz Meeting 2026',
        type: 'Other',
        status: 'Active',
        description: 'Assistenz Meeting Mai 2026 - Frankfurt am Main',
        location: 'Frankfurt am Main',
        locationFilter: '',
        audience: 'All',
        filterMode: 'OR',
        startDate: '2026-05-07T11:00:00.000Z',
        endDate: '2026-05-08T15:00:00.000Z',
        registrationDeadline: '2026-04-09T00:00:00.000Z',
        lastDeregisterDate: '',
        maxParticipants: 130,
        waitlistEnabled: true,
        eventImageUrl: '',
        organizer: 'Maerzluft, Petra; Schwartz, Eva',
        organizerEmail: 'pmaerzluft@deloitte.de',
        outlookEventId: '',
        outlookBody: '',
        emailLanguage: 'EN',
        emailTemplateOverrides: '',
        customFields: [
          { id: 'travel', label: 'You will travel with?', type: 'select', required: false, visible: true, options: ['Train', 'Car', 'Public Transport'] },
          { id: 'deutschlandticket', label: 'Do you own a Deutschlandticket?', type: 'select', required: false, visible: true, options: ['Yes', 'No'] },
          { id: 'expenses', label: 'Please insert the total amount of your travel expenses!', type: 'text', required: false, visible: true },
        ],
        agenda: '[]',
        transfers: '[]',
        documents: '[]',
      });
    } catch { /* Seed fehlgeschlagen - nicht kritisch */ }
  }

  /**
   * Alle Events laden
   */
  public async getEvents(): Promise<SPEvent[]> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EventService.EVENT_SELECT}&$orderby=StartDate desc&$top=100`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return [];
      // v29.51 (Messpunkt): Das ist die EINZIGE blockierende Datenabfrage des
      // Starts — und EVENT_SELECT holt 79 Spalten, darunter EmailImageBase64
      // und EmailTemplateOverrides mit eingebetteten Bildern. Ob das ein paar
      // Kilobyte oder mehrere Megabyte sind, entscheidet über den nächsten
      // Optimierungsschritt; bisher wurde darüber geraten. `.text()` +
      // JSON.parse ist genau das, was `.json()` intern auch tut — der Umweg
      // kostet nichts und liefert die exakte Byte-Zahl.
      const raw = await response.text();
      const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : 0;
      const data = JSON.parse(raw);
      const parseMs = t0 ? Math.round(performance.now() - t0) : -1;
      const rows = data.value || [];
      // eslint-disable-next-line no-console
      console.log(
        `[DEX][perf][getEvents] ${rows.length} Events · ${Math.round(raw.length / 1024)} KB JSON · parse ${parseMs} ms`
      );
      return rows;
    } catch {
      return [];
    }
  }

  /**
   * Einzelnes Event laden
   */
  public async getEvent(eventId: number): Promise<SPEvent | null> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=${EventService.EVENT_SELECT}`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      return await response.json();
    } catch {
      return null;
    }
  }

  /**
   * v18.33: Event anhand des Self-Check-in-Tokens finden (für den statischen
   * Check-in-Link ?action=selfcheckin&token=…). Liefert das erste Event mit
   * passendem Token. Alle eingeloggten User dürfen DEX_Events lesen.
   */
  public async getEventBySelfCheckInToken(token: string): Promise<SPEvent | null> {
    try {
      const safe = token.replace(/'/g, "''");
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EventService.EVENT_SELECT}&$filter=SelfCheckInToken eq '${safe}'&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      const arr = data.value || (data.d && data.d.results) || [];
      return arr.length > 0 ? arr[0] : null;
    } catch {
      return null;
    }
  }

  /**
   * v18.33: Event anhand der Event-Nummer finden (für den rotierenden Live-QR
   * ?action=selfcheckin&event=<Nr>&code=…&t=…).
   */
  public async getEventByEventNumber(eventNumber: number): Promise<SPEvent | null> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EventService.EVENT_SELECT}&$filter=EventNumber eq ${eventNumber}&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      const arr = data.value || (data.d && data.d.results) || [];
      return arr.length > 0 ? arr[0] : null;
    } catch {
      return null;
    }
  }

  /**
   * Neues Event erstellen + Subsite mit Teilnehmerliste anlegen
   */
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
    const reportProgress = (stage:
      | 'start'
      | 'subsite-creating'
      | 'subsite-done'
      | 'permissions'
      | 'list-creating'
      | 'list-done'
      | 'item-insert'
      | 'done'
    ): void => {
      try { event.onProgress?.(stage); } catch { /* */ }
    };
    try {
      reportProgress('start');
      // 0. Nächste EventNumber ermitteln
      let nextEventNumber = 1;
      try {
        const enResp = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=EventNumber&$orderby=EventNumber desc&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (enResp.ok) {
          const enData = await enResp.json();
          if (enData.value && enData.value.length > 0 && enData.value[0].EventNumber) {
            nextEventNumber = enData.value[0].EventNumber + 1;
          }
        }
      } catch { /* Fallback: 1 */ }

      // v11.69: Reuse-Pfad — wenn `existingSubsiteUrl` UND
      // `existingRegistrationListName` mitgegeben wurden, überspringen wir
      // 1) Subsite-Anlegen, 2) Subsite-Permissions, 3) Teilnehmerliste
      // anlegen. Die mitgegebene Subsite bleibt unangetastet inkl. aller
      // Teilnehmer-Anmeldungen. Custom-Fields werden ohne spInternalName-
      // Anreicherung übernommen — die Felder existieren bereits auf der
      // alten Teilnehmerliste mit den korrekten Internal-Names.
      const reuseSubsite = !!(event.existingSubsiteUrl && event.existingRegistrationListName);
      let subsiteUrl: string;
      let enrichedCustomFields: CustomField[];
      const coOrgEmailsForPerm: string[] = (() => {
        try {
          const o = JSON.parse(event.emailTemplateOverrides || '{}');
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const list = (o as any)._coOrganizers;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          if (Array.isArray(list)) return list.map((x: any) => String(x?.email || '')).filter(Boolean);
        } catch { /* */ }
        return [];
      })();
      const allOrgEmails = [event.organizerEmail || '', ...coOrgEmailsForPerm].filter(Boolean).join(';');
      const regListName = reuseSubsite ? (event.existingRegistrationListName as string) : REG_LIST_NAME;

      if (reuseSubsite) {
        // v11.69: bestehende Subsite + Teilnehmerliste wiederverwenden.
        subsiteUrl = event.existingSubsiteUrl as string;
        // Custom-Fields unverändert übernehmen — die Liste existiert
        // bereits, kein neues Schema nötig.
        enrichedCustomFields = event.customFields.map(cf => ({ ...cf }));
      } else {
        // 1. Subsite für das Event erstellen
        reportProgress('subsite-creating');
        const createdSubsite = await this.createEventSubsite(event.title, event.description);
        if (!createdSubsite) {
          console.error('[DEX] Subsite konnte nicht erstellt werden');
          throw new Error('Subsite konnte nicht erstellt werden. Fehlende Berechtigung? Bitte wende dich an einen Site-Administrator.');
        }
        subsiteUrl = createdSubsite;
        reportProgress('subsite-done');

        // 2. Subsite-Berechtigungen: Members der Parent-Site auf der Subsite berechtigen.
        // v9.18: Co-Organizer-Emails aus emailTemplateOverrides._coOrganizers extrahieren
        // und mit dem Hauptorganizer zusammen Full Control erteilen.
        reportProgress('permissions');
        await this.setSubsitePermissions(subsiteUrl, allOrgEmails);

        // 3. Teilnehmerliste auf der Subsite erstellen
        reportProgress('list-creating');
        const fieldMap: Record<string, string> = await this.createRegistrationList(subsiteUrl, event.customFields, allOrgEmails);
        reportProgress('list-done');

        // Custom Fields mit SP InternalName anreichern
        enrichedCustomFields = event.customFields.map(cf => ({
          ...cf,
          spInternalName: fieldMap[cf.id] || '',
        }));
      }

      // 3. Event in DEX_Events eintragen
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
        'Title': event.title,
        'EventNumber': nextEventNumber,
        'EventStatus': event.status,
        'Description': event.description,
        'Location': event.location,
        'LocationAddress': event.locationAddress || '',
        // v18.42: Outlook-Betreff (leer = Flow fällt auf Titel zurück via coalesce).
        'OutlookSubject': (event.outlookSubject && event.outlookSubject.trim()) ? event.outlookSubject.trim() : '',
        // v18.44: abweichendes Outlook-Datum (leer = Flow nutzt StartDate/EndDate).
        'OutlookStart': event.outlookStart || null,
        'OutlookEnd': event.outlookEnd || null,
        // v29.52: ganztägig. Start/Ende bleiben bewusst wie gesetzt (00:00/23:59)
        // — die Umrechnung auf die Ganztags-Grenzen macht der Flow. Solange der
        // Flow das Feld noch nicht liest, verhält sich alles wie bisher.
        'AllDay': !!event.allDay,
        'ShowAsFree': !!event.showAsFree, // v29.54
        'OutlookIsOnlineMeeting': !!event.outlookIsOnlineMeeting, // v30.26
        'SkipOrganizerInvite': !!event.skipOrganizerInvite, // v29.55
        // v18.34/v18.40: Outlook-Ort = manuelle Überschreibung, sonst
        // automatisch aus Veranstaltungsort + Adresse. Flow mappt OutlookLocation 1:1.
        // v26.54: hart auf 255 kappen (einzeilige Text-Spalte — s. updateEvent).
        'OutlookLocation': ((event.outlookLocation && event.outlookLocation.trim())
          ? event.outlookLocation.trim()
          : buildOutlookLocation(event.location, event.locationAddress)).slice(0, 255),
        'LocationFilter': event.locationFilter,
        'Audience': event.audience,
        'AudienceResolvedEmails': event.audienceResolvedEmails || '',
        'FilterMode': event.filterMode || 'OR',
        'StartDate': event.startDate || null,
        // v22.17/v28.66: EndDate darf NIE leer in DEX_Events landen — der
        // DEX_CreateOutlookEvent-Flow rechnet convertFromUtc(coalesce(
        // OutlookEnd, EndDate)); bei null stürzt „Create event (V4)" ab und es
        // entsteht kein Outlook-Termin. Die Aufrufer setzen den Fallback zwar
        // schon, hier wird er zentral erzwungen (letzte Instanz vor dem
        // Schreiben — s. auch updateEvent).
        'EndDate': event.endDate || event.startDate || null,
        'RegistrationDeadline': event.registrationDeadline || null,
        // v29.19: s. Interface — vorher nur im Edit-Pfad geschrieben.
        'ActiveFrom': event.activeFrom || null,
        'LastDeregisterDate': event.lastDeregisterDate || null,
        'MaxParticipants': event.maxParticipants,
        'WaitlistEnabled': event.waitlistEnabled,
        'MandatoryRegistration': !!event.mandatoryRegistration,
        'EventImageUrl': event.eventImageUrl,
        // Custom-Event-Logo aus emailTemplateOverrides._eventLogo extrahieren (falls
        // vorhanden) und als EmailImageBase64 persistieren — damit der Power-Automate-Flow
        // es als {{ORB_URL}} in Mail + Outlook-Termin einsetzt.
        'EmailImageBase64': (() => {
          try {
            const o = JSON.parse(event.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })(),
        'Organizer': event.organizer,
        'OrganizerEmail': event.organizerEmail,
        // v10.16: optionaler Ansprechpartner (Anzeige-Feld). Strings können
        // leer sein — leer = kein Ansprechpartner gepflegt.
        'ContactName': event.contactName || '',
        'ContactEmail': event.contactEmail || '',
        'ContactOrganizerEmail': event.contactOrganizerEmail || '',
        'ContactInfo': event.contactInfo || '',
        'OutlookEventId': event.outlookEventId,
        // outlookBody kommt bereits vollständig gewickelt + mit aufgelösten Variablen
        // aus EventCreationPage — hier nur durchreichen.
        // v29.42: auch im Termin-Text die Fußzeile auf die kanonische Adresse.
        'OutlookBody': normalizeMadeWithLink(event.outlookBody || ''),
        'EmailLanguage': event.emailLanguage || 'EN',
        'RegistrationLanguage': event.registrationLanguage || '',
        'EmailTemplateOverrides': event.emailTemplateOverrides || '',
        'DisableEmails': !!event.disableEmails,
        'DisableRegistrationEmail': !!event.disableRegistrationEmail,
        'DisableCancellationEmail': !!event.disableCancellationEmail,
        'AutoDeregisterOnDecline': !!event.autoDeregisterOnDecline,
        'InactiveHandling': event.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify',
        'DisableOutlook': !!event.disableOutlook,
        'NotifyOrgRegisterMode': (() => {
          const m = event.notifyOrgRegisterMode || 'never';
          return m === 'always' ? 'Always' : m === 'fromDate' ? 'FromDate' : 'Never';
        })(),
        'NotifyOrgRegisterFromDate': event.notifyOrgRegisterFromDate || null,
        'NotifyOrgCancelMode': (() => {
          const m = event.notifyOrgCancelMode || 'never';
          return m === 'always' ? 'Always' : m === 'afterDeadline' ? 'AfterDeadline' : 'Never';
        })(),
        'ExcludedUsers': (event.excludedUsers || []).filter(Boolean).join(';'),
        'IsFictive': !!event.isFictive,
        'DurchstarterCapacity': typeof event.durchstarterCapacity === 'number' ? event.durchstarterCapacity : null,
        'FunstarterCapacity': typeof event.funstarterCapacity === 'number' ? event.funstarterCapacity : null,
        'SplitLabelA': event.splitLabelA || '',
        'SplitLabelB': event.splitLabelB || '',
        'SplitDescA': event.splitDescA || '',
        'SplitDescB': event.splitDescB || '',
        'SplitHelpText': event.splitHelpText || '',
        'SplitSectionTitle': event.splitSectionTitle || '',
        'SplitSharedWaitlist': !!event.splitSharedWaitlist,
        'AllowAttendeeUpload': !!event.allowAttendeeUpload,
        'AttendeeUploadHint': event.attendeeUploadHint || '',
        'AttendeeUploadLabel': event.attendeeUploadLabel || '',
        'AskSalutation': !!event.askSalutation,
        'ConfirmDialogEnabled': !!event.confirmDialogEnabled,
        'ConfirmDialogMode': event.confirmDialogMode || '',
        'ConfirmDialogText': event.confirmDialogText || '',
        'SelfCheckInEnabled': !!event.selfCheckInEnabled,
        'SelfCheckInToken': event.selfCheckInToken || '',
        'SelfCheckInFrom': event.selfCheckInFrom || null,
        'SelfCheckInTo': event.selfCheckInTo || null,
        'TeamRegistrationEnabled': !!event.teamRegistrationEnabled,
        'TeamSize': typeof event.teamSize === 'number' && event.teamSize > 0 ? event.teamSize : null,
        'AskTeamName': !!event.askTeamName,
        'TeamPartialAllowed': !!event.teamPartialAllowed,
        'TeamOpenSlotsVisible': !!event.teamOpenSlotsVisible,
        'TeamJoinRequiresApproval': !!event.teamJoinRequiresApproval,
        'BilingualFields': !!event.bilingualFields,
        'CustomFields': JSON.stringify(enrichedCustomFields),
        'Agenda': event.agenda || '[]',
        'Transfers': event.transfers || '[]',
        'Documents': event.documents || '[]',
        'FunZone': event.funZone || '[]',
        'QuizClusterSize': typeof event.quizClusterSize === 'number' ? event.quizClusterSize : null,
        'ParentEventId': event.parentEventId || '',
        'RegistrationListName': regListName,
        'RegistrationListUrl': `${subsiteUrl}/Lists/${regListName}/AllItems.aspx`,
        'SubsiteUrl': subsiteUrl,
      };

      reportProgress('item-insert');
      // v28.10: gleicher 2-MB-Schutz wie in updateEvent — zu große Payloads
      // (eingebettete Logos/Bilder) sauber abfangen statt kryptischem 400.
      if (JSON.stringify(payload).length > 1_900_000) {
        throw new Error('Die Event-Daten überschreiten das SharePoint-Limit von 2 MB. Ursache ist fast immer ein zu großes eingebettetes Bild (Mail-Logo, Outlook-Kopfbild oder ein Bild im Mail-/Termin-Text). Bitte das Bild entfernen oder neu (kleiner) hochladen.');
      }
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`,
        payload
      );

      if (!response.ok) return null;
      const result = await response.json();
      reportProgress('done');
      return result.d?.Id || result.Id;
    } catch (err) {
      if (err instanceof Error) throw err;
      return null;
    }
  }


  /**
   * Admin-Cleanup beim App-Start: alle Events mit EventStatus='Active' und EndDate < jetzt
   * werden automatisch auf 'Completed' gesetzt. Liefert die Anzahl der aktualisierten Events.
   */
  public async markExpiredEventsAsCompleted(): Promise<number> {
    try {
      // SharePoint OData Filter: Active + EndDate < jetzt
      const nowIso = new Date().toISOString();
      const filter = `EventStatus eq 'Active' and EndDate lt datetime'${nowIso}'`;
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=${encodeURIComponent(filter)}&$select=Id,Title,EndDate&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return 0;
      const data = await resp.json();
      const items: Array<{ Id: number; Title: string }> = data.value || data.d?.results || [];
      if (items.length === 0) return 0;

      let updated = 0;
      for (const it of items) {
        try {
          const ok = await this.updateEvent(it.Id, { 'EventStatus': 'Completed' });
          if (ok) updated += 1;
        } catch { /* einzelnes Update überspringen */ }
      }
      return updated;
    } catch (err) {
      console.warn('[DEX] markExpiredEventsAsCompleted failed:', err);
      return 0;
    }
  }

  /**
   * Event aktualisieren
   */
  /**
   * v11.11: Versionsverlauf des Event-Items aus DEX_Events lesen, um
   * versehentlich gelöschte Custom-Fields (z.B. b2run_*-Felder nach
   * der zu aggressiven v11.9-Migration) wieder zurückzuholen.
   *
   * Liefert eine Liste der Versionen, jeweils mit dem geparsten
   * `CustomFields`-Array (sortiert: neueste zuerst). Werte ohne
   * CustomFields oder mit leerem Array fallen einfach mit raus, sind
   * aber nicht gefiltert — der Caller entscheidet, welche Version
   * relevant ist.
   */
  public async getEventCustomFieldsHistory(eventId: number): Promise<Array<{
    versionLabel: string;
    modified: string;
    customFields: Array<Record<string, unknown>>;
  }>> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/versions?$select=VersionLabel,Modified,CustomFields`;
      const response = await this._sp.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
      });
      if (!response.ok) {
        console.warn('[DEX] getEventCustomFieldsHistory failed:', response.status);
        return [];
      }
      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const versions: any[] = data.value || [];
      return versions.map(v => {
        let parsed: Array<Record<string, unknown>> = [];
        try {
          const raw = (v.CustomFields || '').toString();
          if (raw.trim()) {
            const obj = JSON.parse(raw);
            if (Array.isArray(obj)) parsed = obj as Array<Record<string, unknown>>;
          }
        } catch { /* invalid JSON in old version → leer */ }
        return {
          versionLabel: String(v.VersionLabel || ''),
          modified: String(v.Modified || ''),
          customFields: parsed,
        };
      });
    } catch (err) {
      console.warn('[DEX] getEventCustomFieldsHistory error:', err);
      return [];
    }
  }

  /** v26.51: Klartext-Grund des letzten fehlgeschlagenen updateEvent-Aufrufs —
   *  wird dem Organizer in der Fehlermeldung angezeigt (vorher nur Konsole). */
  public lastUpdateEventError = '';

  public async updateEvent(eventId: number, updates: Record<string, unknown>, retried?: boolean): Promise<boolean> {
    this.lastUpdateEventError = '';
    try {
      // v28.66: zentraler Schutz für EndDate — analog zu createEvent. Ein
      // leeres EndDate in DEX_Events lässt den DEX_CreateOutlookEvent-Flow in
      // „Create event (V4)" mit convertFromUtc(null) abstürzen. Deshalb hier,
      // am gemeinsamen Nadelöhr aller Update-Pfade, aufräumen:
      //  - leeres EndDate + StartDate im selben Update -> Start als Ende,
      //  - sonst das Feld weglassen, statt einen gespeicherten Wert mit null
      //    zu überschreiben (leer war ohnehin nie ein gültiger Zustand).
      const safeUpdates: Record<string, unknown> = { ...updates };
      if ('EndDate' in safeUpdates && !safeUpdates.EndDate) {
        if (safeUpdates.StartDate) {
          safeUpdates.EndDate = safeUpdates.StartDate;
        } else {
          delete safeUpdates.EndDate;
        }
      }
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
        ...safeUpdates,
      };

      // v28.10: SharePoint lehnt REST-Bodies > 2 MB mit einem kryptischen
      // HTTP 400 ab („The request message is too big"). Vorab prüfen und
      // eine verständliche Meldung liefern — Verursacher ist praktisch
      // immer ein zu großes eingebettetes Bild (Mail-/Outlook-Logo oder
      // ein ins Mail-/Termin-Template eingefügtes Bild).
      const LIMIT = 1_900_000;
      const payloadStr = JSON.stringify(payload);
      if (payloadStr.length > LIMIT) {
        // v28.31: Statt aufzugeben in MEHREREN Requests nacheinander schreiben.
        // Das 2-MB-Limit gilt pro REST-Aufruf, nicht pro Item — ein Event mit
        // eingebetteten Bildern (OutlookBody + EmailTemplateOverrides +
        // EmailImageBase64 tragen dasselbe Bild je einmal) passt problemlos,
        // wenn man die Felder auf mehrere MERGEs verteilt. Vorher brach der
        // Save hier still ab: In der Konsole stand nur eine Warnung, im Wizard
        // passierte auf „Speichern" schlicht nichts.
        const FIELD_OVERHEAD = 160; // __metadata + Klammern/Kommas
        const entries = Object.keys(safeUpdates)
          .map(k => ({ k, size: JSON.stringify({ [k]: safeUpdates[k] }).length }))
          .sort((a, b) => b.size - a.size);
        // Ein EINZELNES Feld über dem Limit lässt sich nicht aufteilen — hier
        // hilft nur ein kleineres Bild. Feldname mitgeben, damit der Organizer
        // weiß, wo er suchen muss.
        const tooBig = entries.filter(e => e.size + FIELD_OVERHEAD > LIMIT);
        if (tooBig.length > 0) {
          this.lastUpdateEventError = `Ein einzelnes Feld ist zu groß für SharePoint (${tooBig.map(e => `${e.k}: ${Math.round(e.size / 1024)} KB`).join(', ')}). Ursache ist praktisch immer ein zu großes eingebettetes Bild (Mail-Logo, Outlook-Kopfbild oder ein Bild im Mail-/Termin-Text). Bitte das Bild entfernen oder kleiner erneut hochladen.`;
          console.warn('[DEX] updateEvent: einzelnes Feld über dem Limit —', tooBig);
          return false;
        }
        const groups: Array<Record<string, unknown>> = [];
        let cur: Record<string, unknown> = {};
        let curSize = FIELD_OVERHEAD;
        for (const e of entries) {
          if (curSize + e.size > LIMIT && Object.keys(cur).length > 0) {
            groups.push(cur); cur = {}; curSize = FIELD_OVERHEAD;
          }
          cur[e.k] = safeUpdates[e.k];
          curSize += e.size;
        }
        if (Object.keys(cur).length > 0) groups.push(cur);
        console.warn(`[DEX] updateEvent: Payload ${payloadStr.length} Bytes > Limit — wird in ${groups.length} aufeinanderfolgende Schreibvorgänge aufgeteilt.`);
        for (let i = 0; i < groups.length; i++) {
          const ok = await this.updateEvent(eventId, groups[i], retried);
          if (!ok) {
            // lastUpdateEventError kommt aus dem fehlgeschlagenen Teil-Request.
            this.lastUpdateEventError = `Teil ${i + 1} von ${groups.length} konnte nicht gespeichert werden. ${this.lastUpdateEventError}`.trim();
            return false;
          }
        }
        return true;
      }

      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            'odata-version': '',
          },
          body: payloadStr,
        }
      );
      if (!response.ok) {
        const errText = await response.text().catch(() => '');
        console.warn('[DEX] updateEvent failed:', response.status, errText.substring(0, 400));
        // SharePoint-Fehlertext extrahieren (verbose: error.message.value).
        let spMsg = '';
        try {
          const parsed = JSON.parse(errText);
          spMsg = parsed?.error?.message?.value || parsed?.['odata.error']?.message?.value || '';
        } catch { /* kein JSON */ }
        const statusHint = response.status === 403
          ? 'Keine Berechtigung — du brauchst Schreibrechte auf der Event-Liste (Organizer/Admin).'
          : response.status === 404
            ? 'Das Event wurde in der Liste nicht gefunden — womöglich wurde es zwischenzeitlich gelöscht.'
            : response.status === 409 || response.status === 412
              ? 'Das Event wurde zeitgleich von jemand anderem geändert — bitte neu laden und erneut speichern.'
              : '';
        this.lastUpdateEventError = [`HTTP ${response.status}`, statusHint, spMsg && spMsg !== statusHint ? spMsg.slice(0, 300) : '']
          .filter(Boolean).join(' — ');

        // v26.54: „Invalid text value" = ein String-Wert passt nicht in eine
        // EINZEILIGE Text-Spalte (255-Zeichen-Limit). SharePoint nennt das
        // betroffene Feld nicht — wir diagnostizieren selbst: Payload-Werte
        // gegen die Live-Feldtypen der Liste halten. Spalten, die laut
        // Schema-Definition ohnehin mehrzeilig (Note) sein sollten, werden
        // sofort migriert und der Save EINMAL automatisch wiederholt. Alle
        // anderen Treffer werden in der Fehlermeldung beim Namen genannt.
        if (/invalid text value|text field contains invalid data/i.test(spMsg)) {
          const offenders = await this.findInvalidTextFields('DEX_Events', safeUpdates);
          if (offenders.length > 0) {
            console.warn('[DEX] updateEvent: Werte passen nicht in einzeilige Text-Spalten:',
              offenders.map((o) => `${o.internalName} (${o.length} Zeichen${o.intendedNote ? ', sollte Note sein' : ''})`).join(', '));
            const healable = offenders.filter((o) => o.intendedNote);
            if (!retried && healable.length > 0) {
              for (const o of healable) {
                await this._upgradeTextFieldToNote('DEX_Events', o.title);
              }
              return this.updateEvent(eventId, safeUpdates, true);
            }
            this.lastUpdateEventError += ` | ${offenders
              .map((o) => `Betroffenes Feld: „${o.title}" — ${o.length} Zeichen, die Spalte ist einzeiliger Text (max. 255 Zeichen)`)
              .join('; ')}`;
          }
        }
      }
      return response.ok;
    } catch (err) {
      this.lastUpdateEventError = `Netzwerkfehler — keine Verbindung zu SharePoint${err instanceof Error && err.message ? ` (${err.message.slice(0, 150)})` : ''}.`;
      return false;
    }
  }

  /**
   * v26.54: Diagnose-Helfer für „Invalid text value. A text field contains
   * invalid data." beim Event-Update. Findet alle String-Werte im Update-
   * Payload, die zu lang für eine einzeilige Text-Spalte sind (> 255 Zeichen
   * oder mit Zeilenumbrüchen), deren Ziel-Spalte auf der LIVE-Liste aber
   * tatsächlich als einzeiliger Text ('Text') liegt. `intendedNote` markiert
   * Spalten, die laut Schema-Definition (getEventsFieldDefinitions) eigentlich
   * mehrzeilig (Typ 3, Note) sein sollten — die dürfen automatisch per
   * _upgradeTextFieldToNote geheilt werden.
   */
  private async findInvalidTextFields(
    listName: string,
    updates: Record<string, unknown>
  ): Promise<Array<{ internalName: string; title: string; length: number; intendedNote: boolean }>> {
    const out: Array<{ internalName: string; title: string; length: number; intendedNote: boolean }> = [];
    try {
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false&$top=300`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return out;
      const data = await resp.json();
      const fields: Array<{ InternalName: string; Title: string; TypeAsString: string }> = data.value || [];
      const defs = this.getEventsFieldDefinitions();
      for (const key of Object.keys(updates)) {
        const v = updates[key];
        if (typeof v !== 'string') continue;
        if (v.length <= 255 && v.indexOf('\n') < 0) continue;
        const f = fields.filter((x) => x.InternalName === key)[0];
        if (!f || f.TypeAsString !== 'Text') continue;
        const def = defs.filter((d) => d.title === key)[0];
        out.push({ internalName: key, title: f.Title, length: v.length, intendedNote: !!def && def.type === 3 });
      }
    } catch { /* Diagnose darf den Fehlerpfad nie zusätzlich brechen */ }
    return out;
  }

  /**
   * Event vollständig löschen:
   * 1. Subsite löschen (inkl. Teilnehmerliste) - für neue Events
   * 2. Alte Registrierungsliste löschen (DEX_Reg_*) - für alte Events
   * 3. Event-Eintrag aus DEX_Events löschen
   */
  public async deleteEvent(eventId: number): Promise<boolean> {
    try {
      // Event-Daten laden um SubsiteUrl und RegistrationListName zu bekommen
      const event = await this.getEvent(eventId);
      if (!event) return false;

      // 0. Outlook-Kalendereintrag per Queue löschen (VOR allem anderen, damit
      //    CalendarLink noch vorhanden ist). Der DEX_Outlook_Einladungen-Flow
      //    greift den DeleteEvent-Eintrag auf und löscht den Kalender-Termin
      //    im Shared Mailbox über den Flow-Service-Account.
      //    Fehler hier ignorieren - Event-Delete soll trotzdem durchlaufen.
      if (event.CalendarLink) {
        try {
          await this.queueOutlookDeleteEvent(String(eventId), event.Title || '', event.CalendarLink);
        } catch { /* Queue-Fehler ignorieren */ }
      }
      // 1. Subsite RECYCEN (v9.0: nicht mehr per DELETE, sonst landet die
      //    Subsite permanent weg ohne Recycle-Bin-Eintrag. recycle() legt
      //    die Subsite mitsamt Teilnehmerliste 93 Tage in den Site
      //    Collection Recycle Bin → ein Tenant-Admin / Site Collection
      //    Admin kann sie dort wiederherstellen falls nötig.
      if (event.SubsiteUrl) {
        try {
          await this._post(`${event.SubsiteUrl}/_api/web/recycle`, {});
        } catch {
          console.warn('[DEX] Subsite konnte nicht in den Recycle Bin verschoben werden:', event.SubsiteUrl);
        }
      }

      // 2. Event-Bild ebenfalls RECYCEN statt löschen.
      if (event.EventImageUrl) {
        try {
          const url = new URL(event.EventImageUrl);
          const serverRelUrl = url.pathname;
          if (serverRelUrl.indexOf('DEX_EventImages') >= 0) {
            await this._post(
              `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelUrl}')/recycle`,
              {}
            );
          }
        } catch {
          console.warn('[DEX] Event-Bild konnte nicht in den Recycle Bin verschoben werden');
        }
      }

      // 3. Alte Registrierungsliste recyceln (legacy Events ohne Subsite).
      if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
        try {
          await this._post(
            `${this.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')/recycle`,
            {}
          );
        } catch {
          console.warn('[DEX] Alte Registrierungsliste konnte nicht recycelt werden:', event.RegistrationListName);
        }
      }

      // 3. DEX_Participants aufräumen: EventNumber aus allen Teilnehmern entfernen
      if (event.EventNumber) {
        try {
          const allParticipants = await this.getAllParticipants();
          // Parallelize participant cleanup for better performance
          const updatePromises = allParticipants
            .filter(p => {
              const en = String(event.EventNumber);
              const hasRegistered = p.EventRegistered?.split(',').map(s => s.trim()).includes(en);
              const hasWaitlist = p.EventOnWaitlist?.split(',').map(s => s.trim()).includes(en);
              return hasRegistered || hasWaitlist;
            })
            .map(p => this.removeParticipantEvent(p.Email, event.EventNumber));
          // Promise.all mit individueller Fehlerbehandlung (Promise.allSettled nicht verfügbar in ES2017)
          const safePromises = updatePromises.map(p => p.catch(() => null));
          await Promise.all(safePromises);
        } catch {
          console.warn('[DEX] DEX_Participants konnte nicht aufgeräumt werden');
        }
      }

      // 4. Event-Dokumente löschen (SiteAssets/DEX_EventDocs/Event_{number}_*)
      if (event.EventNumber) {
        try {
          const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
          const safeName = (event.Title || '').replace(/[#%&*:<>?/\\|"']/g, '').replace(/\s+/g, '_').substring(0, 50);
          const folderName = safeName ? `Event_${event.EventNumber}_${safeName}` : `Event_${event.EventNumber}`;
          await this._delete(`${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/${folderName}')`);
        } catch {
          // Fallback: alten Ordnernamen ohne Titel probieren
          try {
            const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
            await this._delete(`${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/Event_${event.EventNumber}')`);
          } catch { /* Ordner nicht gefunden */ }
        }
      }

      // 5. Event-Eintrag aus DEX_Events RECYCEN (v9.0: per recycle() statt
      //    delete(), damit ein Admin via SharePoint-Recycle-Bin das Item
      //    bei Bedarf 93 Tage lang wiederherstellen kann).
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/recycle`,
        {}
      );

      // 6. Audit-Eintrag in DEX_ChangeLog (v9.0). Best-effort, blockt
      //    den Lösch-Vorgang nicht falls Logging fehlschlägt.
      try {
        await this.writeChangeLog({
          action: 'EventDeletedTest', // wird vom Aufrufer überschrieben
          targetType: 'Event',
          targetId: String(eventId),
          targetName: event.Title || '',
          eventId: String(eventId),
          eventTitle: event.Title || '',
          details: {
            subsiteUrl: event.SubsiteUrl || '',
            eventNumber: event.EventNumber,
            recycledTo: 'SharePoint Recycle Bin (93 Tage)',
          },
        });
      } catch { /* */ }

      return response.ok;
    } catch {
      return false;
    }
  }

  // =====================================================================
  // v26.32: Löschkonzept — Teilnehmerliste 3 Monate nach Event-Ende löschen;
  // Event + KPIs bleiben im Statistik-Archiv (DEX_EventStats) erhalten.
  // =====================================================================

  private static readonly EVENTSTATS_LIST = 'DEX_EventStats';
  private static readonly EVENTSTATS_ITEM_TYPE = 'SP.Data.DEX_x005f_EventStatsListItem';

  /** Legt DEX_EventStats (Statistik-Archiv, KEINE PII) an, falls nicht vorhanden.
   *  Existiert die Liste bereits, werden nur fehlende (neue) Spalten nachgerüstet
   *  — z.B. `Organizer` (v26.33), das es in früheren Versionen noch nicht gab. */
  public async ensureEventStatsList(): Promise<void> {
    const listName = EventService.EVENTSTATS_LIST;
    const fields: Array<{ title: string; type: number }> = [
      { title: 'EventNumber', type: 9 }, { title: 'EventTitle', type: 2 },
      { title: 'EventType', type: 2 }, { title: 'EventLocation', type: 2 },
      { title: 'EventStart', type: 4 }, { title: 'EventEnd', type: 4 },
      { title: 'MaxParticipants', type: 9 }, { title: 'RegisteredCount', type: 9 },
      { title: 'QRSentCount', type: 9 }, { title: 'CheckedInCount', type: 9 },
      { title: 'NoShowCount', type: 9 }, { title: 'WaitlistCount', type: 9 },
      { title: 'DeregisteredCount', type: 9 }, { title: 'ArchivedByEmail', type: 2 },
      { title: 'ArchivedDate', type: 4 }, { title: 'Organizer', type: 3 },
    ];
    let exists = false;
    try { exists = await this.listExists(listName); } catch { exists = false; }
    if (!exists) {
      const createResp = await this._post(`${this.siteUrl}/_api/web/lists`, {
        '__metadata': { 'type': 'SP.List' },
        'Title': listName,
        'Description': 'Statistik-Archiv (v26): KPIs eines Events, nachdem die Teilnehmerliste nach 3 Monaten gelöscht wurde. Enthält KEINE personenbezogenen Daten.',
        'BaseTemplate': 100,
        'AllowContentTypes': false,
      });
      if (!createResp.ok) { console.warn('[DEX] DEX_EventStats konnte nicht angelegt werden — vermutlich fehlen Owner-Rechte.'); return; }
    }
    // Nur fehlende Felder anlegen (idempotent): bestehende Internal-Names holen.
    const have = new Set<string>();
    try {
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$top=200`,
        SPHttpClient.configurations.v1,
      );
      if (resp.ok) {
        const data = await resp.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const f of ((data.value || data.d?.results || []) as any[])) { if (f.InternalName) have.add(String(f.InternalName)); }
      }
    } catch { /* dann versuchen wir einfach alle anzulegen */ }
    for (const f of fields) {
      if (have.has(f.title)) continue;
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
        });
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
  }

  /** EventNumbers, für die bereits ein Statistik-Archiv-Eintrag existiert. */
  public async getArchivedStatsEventNumbers(): Promise<Set<number>> {
    const out = new Set<number>();
    try {
      if (!(await this.listExists(EventService.EVENTSTATS_LIST))) return out;
      let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('${EventService.EVENTSTATS_LIST}')/items?$select=EventNumber&$top=5000`;
      while (url) {
        const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!resp.ok) break;
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const it of (items as any[])) { const n = Number(it.EventNumber); if (!Number.isNaN(n)) out.add(n); }
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      }
    } catch { /* best-effort */ }
    return out;
  }

  /**
   * Liest alle Zeilen aus dem Statistik-Archiv (DEX_EventStats). Reine KPI-
   * Daten, KEINE PII. Für die Anzeige-Kachel „Statistik-Archiv" im Admin Center.
   */
  public async getEventStats(): Promise<Array<{
    id: number; eventNumber: number; eventTitle: string; eventType: string;
    location: string; startDate: string; endDate: string; maxParticipants: number | null;
    registeredCount: number; qrSentCount: number; checkedInCount: number;
    noShowCount: number; waitlistCount: number; deregisteredCount: number;
    organizer: string; archivedByEmail: string; archivedDate: string;
  }>> {
    const out: Array<{
      id: number; eventNumber: number; eventTitle: string; eventType: string;
      location: string; startDate: string; endDate: string; maxParticipants: number | null;
      registeredCount: number; qrSentCount: number; checkedInCount: number;
      noShowCount: number; waitlistCount: number; deregisteredCount: number;
      organizer: string; archivedByEmail: string; archivedDate: string;
    }> = [];
    try {
      if (!(await this.listExists(EventService.EVENTSTATS_LIST))) return out;
      const sel = 'Id,EventNumber,EventTitle,EventType,EventLocation,EventStart,EventEnd,MaxParticipants,RegisteredCount,QRSentCount,CheckedInCount,NoShowCount,WaitlistCount,DeregisteredCount,Organizer,ArchivedByEmail,ArchivedDate';
      let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('${EventService.EVENTSTATS_LIST}')/items?$select=${sel}&$orderby=ArchivedDate desc&$top=5000`;
      while (url) {
        const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!resp.ok) break;
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const it of (items as any[])) {
          const num = (v: unknown): number => { const n = Number(v); return Number.isNaN(n) ? 0 : n; };
          out.push({
            id: Number(it.Id),
            eventNumber: num(it.EventNumber),
            eventTitle: it.EventTitle || '',
            eventType: it.EventType || '',
            location: it.EventLocation || '',
            startDate: it.EventStart || '',
            endDate: it.EventEnd || '',
            maxParticipants: (it.MaxParticipants === null || it.MaxParticipants === undefined) ? null : num(it.MaxParticipants),
            registeredCount: num(it.RegisteredCount),
            qrSentCount: num(it.QRSentCount),
            checkedInCount: num(it.CheckedInCount),
            noShowCount: num(it.NoShowCount),
            waitlistCount: num(it.WaitlistCount),
            deregisteredCount: num(it.DeregisteredCount),
            organizer: EventService.stripNoteWrapper(it.Organizer) || '',
            archivedByEmail: it.ArchivedByEmail || '',
            archivedDate: it.ArchivedDate || '',
          });
        }
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      }
    } catch (err) { console.warn('[DEX] getEventStats failed:', err); }
    return out;
  }

  /**
   * Berechnet die KPIs eines Events aus der Teilnehmerliste und schreibt EINE
   * Zeile ins DEX_EventStats (keine PII). Muss VOR dem Löschen der Subsite
   * laufen. Idempotent: existiert schon eine Zeile für die EventNumber, wird
   * nichts geschrieben (true).
   */
  public async archiveEventStats(meta: {
    eventNumber: number; eventTitle: string; eventType?: string; location?: string;
    startDate?: string; endDate?: string; maxParticipants?: number; subsiteUrl?: string;
    organizer?: string;
  }): Promise<boolean> {
    try {
      await this.ensureEventStatsList();
      const existing = await this.getArchivedStatsEventNumbers();
      if (existing.has(meta.eventNumber)) return true;
      const regs = meta.subsiteUrl ? await this.getAllRegistrations(meta.subsiteUrl) : [];
      const countBy = (pred: (s: string) => boolean): number => regs.filter(r => pred(r.Status || '')).length;
      const me = (this.context.pageContext.user.email || '').toLowerCase();
      const resp = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${EventService.EVENTSTATS_LIST}')/items`,
        {
          '__metadata': { 'type': EventService.EVENTSTATS_ITEM_TYPE },
          'Title': `${meta.eventNumber}: ${(meta.eventTitle || '').slice(0, 200)}`,
          'EventNumber': meta.eventNumber, 'EventTitle': meta.eventTitle || '',
          'EventType': meta.eventType || '', 'EventLocation': meta.location || '',
          'EventStart': meta.startDate || null, 'EventEnd': meta.endDate || null,
          'MaxParticipants': typeof meta.maxParticipants === 'number' ? meta.maxParticipants : null,
          'RegisteredCount': countBy(s => s === 'Angemeldet' || s === 'QR versendet' || s === 'Eingecheckt'),
          'QRSentCount': countBy(s => s === 'QR versendet' || s === 'Eingecheckt'),
          'CheckedInCount': countBy(s => s === 'Eingecheckt'),
          'NoShowCount': countBy(s => s === 'No-Show'),
          'WaitlistCount': countBy(s => s === 'Warteliste'),
          'DeregisteredCount': countBy(s => s === 'Abgemeldet'),
          'ArchivedByEmail': me, 'ArchivedDate': new Date().toISOString(),
          'Organizer': meta.organizer || '',
        }
      );
      return resp.ok;
    } catch (err) { console.warn('[DEX] archiveEventStats failed:', err); return false; }
  }

  /** Rollback: entfernt die (soeben geschriebene) Statistik-Zeile(n) einer
   *  EventNumber — genutzt, wenn die anschließende Löschung fehlschlägt, damit
   *  das Event beim nächsten Lauf erneut verarbeitet wird. */
  public async deleteEventStatsRow(eventNumber: number): Promise<boolean> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('${EventService.EVENTSTATS_LIST}')/items?$select=Id&$filter=EventNumber eq ${eventNumber}&$top=50`;
      const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return false;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      let ok = true;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const it of (items as any[])) {
        try { await this._delete(`${this.siteUrl}/_api/web/lists/getbytitle('${EventService.EVENTSTATS_LIST}')/items(${it.Id})`); }
        catch { ok = false; }
      }
      return ok;
    } catch { return false; }
  }

  /**
   * Löscht die Teilnehmerliste (Subsite) eines Events, LÄSST das DEX_Events-Item
   * bestehen (Event bleibt sichtbar; KPIs stehen im DEX_EventStats). Bereinigt
   * zusätzlich DEX_Participants (EventNumber entfernen). Gegenstück zu
   * deleteEvent(), das auch das Event-Item entfernt.
   */
  public async deleteParticipantData(eventId: number): Promise<boolean> {
    try {
      const event = await this.getEvent(eventId);
      if (!event) return false;
      /**
       * v29.3: Reihenfolge umgedreht — erst das Register, dann die
       * unwiderrufliche Löschung.
       *
       * Bisher wurde zuerst die Subsite recycelt und danach das Register
       * aufgeräumt, und zwar so, dass ein Scheitern nicht auffiel:
       * `getAllParticipants()` liest NICHT strikt (eine unvollständige Seite
       * kam still als vollständige Liste zurück), alle Schreibvorgänge liefen
       * als EIN `Promise.all` gleichzeitig los (bei hunderten Teilnehmern die
       * Einladung zur SharePoint-Drosselung), und jeder Fehler wurde per
       * `.catch(() => null)` weggeworfen. Was danach im Register stand, wusste
       * niemand — die Teilnehmerliste war aber schon weg, also ließ sich der
       * Rest auch nicht mehr nachrechnen.
       *
       * Genau das ist der Rückstand, den die Register-Prüfung heute als
       * „Verweis ohne Zeile" meldet. Jetzt gilt: strikt lesen, sequentiell
       * schreiben, jeden Fehlschlag zählen — und bei Fehlern GAR NICHT
       * löschen. Der Aufrufer wertet `false` bereits als „später erneut
       * versuchen" und rollt die Statistik-Zeile zurück; nichts geht verloren,
       * weil die Subsite dann noch steht.
       */
      if (event.EventNumber) {
        let registryFailed = 0;
        try {
          const allParticipants = await this.fetchAllParticipantsOrThrow();
          const en = String(event.EventNumber);
          const affected = allParticipants.filter(p =>
            (p.EventRegistered?.split(',').map(s => s.trim()).includes(en))
            || (p.EventOnWaitlist?.split(',').map(s => s.trim()).includes(en)));
          for (const p of affected) {
            // eslint-disable-next-line no-await-in-loop
            const ok = await this.removeParticipantEvent(p.Email, event.EventNumber).catch(() => false);
            if (!ok) registryFailed += 1;
          }
        } catch (err) {
          console.warn('[DEX] deleteParticipantData: Register nicht vollständig lesbar — Löschung abgebrochen:', err);
          return false;
        }
        if (registryFailed > 0) {
          console.warn(`[DEX] deleteParticipantData: ${registryFailed} Register-Einträge nicht aktualisiert — Löschung abgebrochen (Event ${event.EventNumber}).`);
          return false;
        }
      }
      if (event.SubsiteUrl) {
        try { await this._post(`${event.SubsiteUrl}/_api/web/recycle`, {}); }
        catch { console.warn('[DEX] Teilnehmer-Subsite konnte nicht recycelt werden:', event.SubsiteUrl); }
      }
      if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
        try { await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')/recycle`, {}); }
        catch { /* */ }
      }
      try {
        await this.writeChangeLog({
          action: 'ParticipantListDeleted', targetType: 'Event', targetId: String(eventId),
          targetName: event.Title || '', eventId: String(eventId), eventTitle: event.Title || '',
          details: { subsiteUrl: event.SubsiteUrl || '', eventNumber: event.EventNumber, note: 'Teilnehmerliste gelöscht (3-Monats-Löschkonzept); KPIs im DEX_EventStats.' },
        });
      } catch { /* */ }
      return true;
    } catch { return false; }
  }

  /**
   * v11.69: Löscht NUR das DEX_Events-Listenitem — KEIN Cascade auf Subsite,
   * KEIN Outlook-DeleteEvent in die Queue, KEIN EventImage-Recycle, KEIN
   * DEX_Participants-Cleanup. Gegenstück zu `deleteEvent()`, das alles
   * mit-aufräumt.
   *
   * Nutzungs-Szenario: Outlook-Termin nachträglich auf einem bereits
   * angelegten Sub-Event aktivieren. Der `DEX_CreateOutlookEvent`-Flow
   * triggert ausschliesslich auf NEUE DEX_Events-Items (GetOnNewItems) —
   * ein MERGE-Update reicht nicht aus. Statt das ganze Sub-Event komplett
   * delete+recreate zu machen (was kaskadierend Subsite + Teilnehmer-
   * anmeldungen mitlöschen würde), wird hier nur die DEX_Events-Zeile
   * entfernt und gleich darauf eine neue mit `createEvent({ existingSubsiteUrl,
   * existingRegistrationListName })` angelegt. Die alte Subsite mit allen
   * Anmeldungen bleibt unangetastet und wird einfach an die neue Zeile
   * gekoppelt.
   *
   * **Garantie:** Diese Methode ruft KEIN `recycle()` auf der Subsite, KEIN
   * `recycle()` auf der Teilnehmerliste und KEIN `removeParticipantEvent()`.
   * Nur das DEX_Events-Item wird per REST-DELETE entfernt — alles andere
   * bleibt 1:1 erhalten.
   */
  public async deleteEventItemOnly(eventId: string | number): Promise<boolean> {
    try {
      const idNum = typeof eventId === 'string' ? parseInt(eventId, 10) : eventId;
      if (!idNum || Number.isNaN(idNum)) return false;
      const response = await this._delete(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${idNum})`
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== Subsites ====================

  /**
   * URL-Suffix aus Event-Titel generieren.
   * "B2Run Frankfurt 2026" → "b2run-frankfurt-2026-k8f3a"
   */
  private generateSubsiteUrl(title: string): string {
    const slug = title
      .toLowerCase()
      .replace(/[äÄ]/g, 'ae').replace(/[öÖ]/g, 'oe').replace(/[üÜ]/g, 'ue').replace(/ß/g, 'ss')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .substring(0, 40);
    const suffix = Date.now().toString(36).slice(-5);
    return `${slug}-${suffix}`;
  }

  /**
   * Subsite für ein Event erstellen.
   * Versucht mehrere Templates falls eines fehlschlägt.
   * Gibt die absolute URL der neuen Subsite zurück.
   */
  private async createEventSubsite(title: string, description: string): Promise<string | null> {
    const urlSuffix = this.generateSubsiteUrl(title);
    const desc = description || `Event-Subsite: ${title}`;

    // Templates in Reihenfolge versuchen:
    // STS#3 = Modern ohne Group, STS#0 = Classic Team Site, STS = Blank
    const templates = ['STS#3', 'STS#0', 'STS'];

    for (const template of templates) {
      try {
        const payload = {
          'parameters': {
            '__metadata': { 'type': 'SP.WebCreationInformation' },
            'Title': title,
            'Url': urlSuffix,
            'Description': desc,
            'Language': 1031,
            'WebTemplate': template,
            'UseSamePermissionsAsParentSite': false,
          }
        };

        const response = await this._post(`${this.siteUrl}/_api/web/webs/add`, payload);
        if (response.ok) {
          const result = await response.json();
          const subsiteAbsoluteUrl = result.d?.Url || result.Url;
          // Subsite erfolgreich erstellt
          return subsiteAbsoluteUrl || `${this.siteUrl}/${urlSuffix}`;
        }

        // Fehlerdetails loggen
        try {
          const err = await response.json();
          console.warn(`[DEX] Template ${template} fehlgeschlagen (${response.status}):`, err.error?.message?.value || err);
        } catch {
          console.warn(`[DEX] Template ${template} fehlgeschlagen: ${response.status}`);
        }
      } catch (e) {
        console.warn(`[DEX] Template ${template} Fehler:`, e);
      }
    }

    console.error('[DEX] Subsite konnte mit keinem Template erstellt werden');
    return null;
  }

  // ==================== Teilnehmerlisten (auf Subsites) ====================

  /**
   * Teilnehmerliste auf einer Subsite erstellen.
   * Liste heißt immer "Teilnehmer".
   */
  private async createRegistrationList(
    subsiteUrl: string,
    customFields: CustomField[],
    organizerEmail: string
  ): Promise<Record<string, string>> {
    // Liste erstellen
    await this._post(`${subsiteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': REG_LIST_NAME,
      'Description': 'Teilnehmerliste für dieses Event',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // Basis-Spalten
    const baseFields = [
      { title: 'TeilnehmerID', type: 9 }, // Number - fortlaufende ID
      { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'ParticipantName', type: 2 }, // Backward compat
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Department', type: 2 },
      { title: 'Location', type: 2 },
      { title: 'JobTitle', type: 2 },
      { title: 'Phone', type: 2 },
      // v24.29: Unternehmenszugehörigkeit / Rechtsträger (aus dem Profil).
      { title: 'Company', type: 2 },
      { title: 'Status', type: 6, choices: ['Angemeldet', 'QR versendet', 'Warteliste', 'Eingecheckt', 'No-Show', 'Abgemeldet'], metaType: 'SP.FieldChoice' },
      { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Typ-Auswahl
      { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Wunsch-Typ (wenn Fallback oder Warteliste)
      // v10.13: B2Run-Leistungsnachweis-Bestätigung. Virtuelles Feld der
      // RegistrationPage, das nur durchläuft wenn durchstarterRequiresProof
      // aktiv ist — die SP-Spalte muss aber existieren sonst kippt die
      // Anmeldung mit HTTP 400. Wird auf jeder neuen Teilnehmerliste angelegt
      // damit B2Run-Events nicht später nochmal manuell repariert werden müssen.
      { title: 'b2run_leistungsnachweis', type: 2 },
      { title: 'QuizScore', type: 9 }, // Number - Anzahl richtiger Antworten
      { title: 'QuizAnswers', type: 3 }, // Note - JSON der Antworten (für Statistik)
      { title: 'QuizCompletedAt', type: 4 }, // DateTime
      { title: 'RegistrationDate', type: 4 },
      { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgeführt hat
      { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgeführt hat
      { title: 'ProxyConsent', type: 3 },      // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung (Note)
      { title: 'LastModifiedDate', type: 4 },
      { title: 'ChangeLog', type: 3 }, // Note (multiline) - Änderungshistorie
      { title: 'CancellationDate', type: 4 },
      { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgelöst hat
      { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
      { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
      { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
      { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
      // v17.15: Nachrück-Audit (siehe SPRegistration-Interface):
      // - PromotedDate: gesetzt beim Promote auf die nachrückende Person.
      // - ReplacedParticipantEmail: E-Mail der Person, deren Cancel den
      //   Promote ausgelöst hat („Ersetzt wen") — auf der promoteten Person.
      // - ReplacedByParticipantEmail: E-Mail der nachrückenden Person
      //   („Ersetzt durch") — auf der cancelnden Person.
      { title: 'PromotedDate', type: 4 },
      { title: 'ReplacedParticipantEmail', type: 2 },
      { title: 'ReplacedByParticipantEmail', type: 2 },
      // v11.36: Überbuchungs-Review-Marker. '' = normal, 'Pending' = vom
      // „Überbuchung prüfen"-Lauf als über Kapazität erkannt; der Admin
      // entscheidet pro Person (auf Warteliste / Platz behalten).
      { title: 'OverbookReview', type: 2 },
      { title: 'ConsentReview', type: 2 }, // v26.47: Externe Anmeldung — 'Pending' = Datenschutz-Rückmeldung offen
      // v11.82: Team-Anmeldung — drei Spalten gruppieren Mitglieder eines
      // gemeinsam angemeldeten Teams. TeamId = UUID (gleicher Wert für alle
      // Mitglieder), TeamLead = true nur für die anmeldende Person, TeamName
      // = optionaler frei wählbarer Name (nur wenn das Event AskTeamName an
      // hat). Bei Nicht-Team-Anmeldungen bleiben alle drei Felder leer.
      { title: 'TeamId', type: 2 },
      { title: 'TeamLead', type: 8 },
      { title: 'TeamName', type: 2 },
      { title: 'CustomData', type: 3 },
    ];

    for (const f of baseFields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if ((f as { choices?: string[] }).choices) {
        payload['Choices'] = { 'results': (f as { choices: string[] }).choices };
      }
      await this._post(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload);
    }

    // Custom Fields als eigene Spalten anlegen + InternalName merken
    const customFieldViewNames: string[] = [];
    const fieldMap: Record<string, string> = {}; // cf.id -> SP InternalName
    for (const cf of customFields) {
      if (!cf.label) continue;
      // v19.0: Dokument-Felder bekommen KEINE Spalte — die Datei wird als
      // Attachment an die Teilnehmer-Zeile gehängt, nicht als Spaltenwert.
      if (cf.type === 'document') continue;
      let fieldPayload: Record<string, unknown>;

      if (cf.type === 'select' && cf.options && cf.options.length > 0) {
        fieldPayload = {
          '__metadata': { 'type': 'SP.FieldChoice' },
          'Title': cf.label,
          'FieldTypeKind': 6,
          'Required': false,
          'Choices': { 'results': cf.options },
        };
      } else if (cf.type === 'number') {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 9,
          'Required': false,
        };
      } else if (cf.type === 'checkbox') {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 8,
          'Required': false,
        };
      } else {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 2,
          'Required': false,
        };
      }

      try {
        const fieldResponse = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
          fieldPayload
        );
        if (fieldResponse.ok) {
          const fieldResult = await fieldResponse.json();
          const internalName = fieldResult.d?.InternalName || fieldResult.InternalName || cf.label;
          fieldMap[cf.id] = internalName;
          customFieldViewNames.push(internalName);
        }
      } catch {
        console.warn('[DEX] Custom Field konnte nicht angelegt werden:', cf.label);
      }
    }

    // FieldMap wird als Rückgabewert an den Caller zurückgegeben

    // Default View komplett neu aufbauen (Basis + Custom Fields). Mit rebuild:true
    // werden alle SP-Default-Spalten (Modified, Created, ID, Type, Compliance Asset,
    // App Created By, ...) aus der View rausgeworfen — nur funktionelle Felder.
    await this.configureDefaultView(REG_LIST_NAME, [
      'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail', 'Department', 'Location', 'JobTitle', 'Company', 'Phone', 'StarterType', 'PreferredStarterType', 'Status', 'RegistrationDate', 'RegisteredByName', 'RegisteredByEmail', 'ProxyConsent', 'CancellationDate', 'CancelledByName', 'CancelledByEmail',
      ...customFieldViewNames,
      // v11.82: Team-Spalten am Ende der View (nach allen Custom Fields, vor
      // System-Spalten). So bleibt die View bei Nicht-Team-Events unauffällig
      // und bei Team-Events sieht der Organizer auf einen Blick, wer mit wem
      // angemeldet ist.
      'TeamId', 'TeamLead', 'TeamName',
    ], subsiteUrl, { rebuild: true });

    // Item-Level Permissions
    await this.setItemLevelPermissions(subsiteUrl);

    // Berechtigungen
    await this.setRegistrationListPermissions(subsiteUrl, organizerEmail);

    // v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe anlegen
    // (Race-Condition-Schutz bei parallelen Anmeldungen).
    try {
      await this.ensureCounterList(subsiteUrl);
    } catch {
      // Nicht kritisch — falls das schiefgeht, fallback auf max+1 in upsertParticipant.
    }

    return fieldMap;
  }

  /**
   * v9.35: Berechtigungs-Sync für nachträglich hinzugefügte Organizer/Co-Organizer.
   *
   * Wird im Wizard im Edit-Modus nach updateEvent aufgerufen. Geht über die
   * komma-/semikolon-separierte Liste aller Organizer-Mails und stellt sicher,
   * dass jede Person Full Control auf der Subsite + auf der Teilnehmerliste hat.
   *
   * Idempotent: Personen, die bereits Full Control haben, werden von SharePoints
   * `addroleassignment` einfach durchgereicht (kein Fehler, kein Doppel-Eintrag).
   * Existierende Item-Level-Permissions auf der Liste bleiben unangetastet — wir
   * brechen die Inheritance hier NICHT erneut, sondern fügen nur fehlende Principals
   * obendrauf hinzu.
   */
  public async ensureOrganizerPermissions(subsiteUrl: string, organizerEmails: string): Promise<void> {
    await this.ensureOrganizerPermissionsMulti([subsiteUrl], organizerEmails);
  }

  /**
   * v30.37: Derselbe Sync über MEHRERE Subsites — Klammer **und** alle
   * Sub-Events.
   *
   * Warum das nötig war: `ensureOrganizerPermissions` lief im Wizard nur über
   * `editEvent.subsiteUrl`, also die Klammer. Jedes Sub-Event hat aber eine
   * EIGENE Subsite mit eigener Teilnehmerliste. Wer nachträglich als
   * Co-Organizer dazukam, bekam Full Control auf der Klammer und auf keinem
   * einzigen Termin — und weil `getAllRegistrations` bei 403 nicht wirft,
   * sondern `[]` liefert, sah diese Person überall „0 Teilnehmer" statt einer
   * Fehlermeldung. Ein Klammer-Event mit 19 Office-Tagen war damit für sie
   * vollständig leer (Befund 31.08.2026, Carolin R.).
   *
   * Die User-Id wird EINMAL je Person aufgelöst, nicht je Subsite — sonst sind
   * es bei 19 Terminen × 3 Organizern 57 Lookups gegen dieselben drei Konten.
   *
   * @returns Zähler fürs UI: wie viele Zuweisungen liefen, wie viele Personen
   *   gar nicht auflösbar waren. Ein „schon vorhanden" zählt als Erfolg —
   *   SharePoints `addroleassignment` ist idempotent.
   */
  public async ensureOrganizerPermissionsMulti(
    subsiteUrls: string[],
    organizerEmails: string
  ): Promise<{ sites: number; users: number; grants: number; unresolved: string[] }> {
    const sites = (subsiteUrls || []).map(s => (s || '').trim()).filter(Boolean);
    const emails = (organizerEmails || '').split(/[;,]/).map(s => s.trim()).filter(Boolean);
    const result = { sites: sites.length, users: 0, grants: 0, unresolved: [] as string[] };
    if (sites.length === 0 || emails.length === 0) return result;

    // 1) Personen auflösen — einmal, nicht je Subsite.
    const userIds: number[] = [];
    for (const em of emails) {
      try {
        const userResponse = await this._sp.get(
          `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (!userResponse.ok) { result.unresolved.push(em); continue; }
        const userData = await userResponse.json();
        const userId = userData.d?.Id || userData.Id;
        if (userId) userIds.push(userId); else result.unresolved.push(em);
      } catch { result.unresolved.push(em); }
    }
    result.users = userIds.length;

    // 2) Je Subsite: Web-Level + Teilnehmerliste. Beides einzeln gekapselt —
    //    eine recycelte Subsite darf die übrigen Termine nicht abbrechen.
    for (const site of sites) {
      for (const userId of userIds) {
        try {
          await this._post(
            `${site}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
            {}
          );
          result.grants++;
        } catch { /* idempotent — Person hatte schon Rechte */ }
        try {
          await this._post(
            `${site}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
            {}
          );
        } catch { /* idempotent */ }
      }
    }
    return result;
  }

  /**
   * Subsite-Berechtigungen: Owners Full Control, Members Read (damit User die Subsite betreten können).
   */
  private async setSubsitePermissions(subsiteUrl: string, organizerEmail: string): Promise<void> {
    try {
      // Owners der Hauptsite: Full Control auf der Subsite
      const ownersResponse = await this._sp.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        const ownersId = ownersData.d?.Id || ownersData.Id;
        await this._post(
          `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${ownersId}, roledefid=1073741829)`,
          {}
        );
      }

      // Visitors der Hauptsite: Read auf der Subsite (damit User die Subsite betreten können)
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
          {}
        );
      }

      // Organizer: Full Control auf der Subsite. v9.18: organizerEmail kann
      // ";"-separiert mehrere Emails enthalten — Hauptorganizer + Co-Organizer
      // bekommen alle Full Control auf der Subsite.
      if (organizerEmail) {
        const emails = organizerEmail.split(/[;,]/).map(s => s.trim()).filter(Boolean);
        for (const em of emails) {
          try {
            const userResponse = await this._sp.get(
              `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
              SPHttpClient.configurations.v1
            );
            if (userResponse.ok) {
              const userData = await userResponse.json();
              const userId = userData.d?.Id || userData.Id;
              await this._post(
                `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
                {}
              );
            }
          } catch { /* Organizer-Berechtigung optional */ }
        }
      }
    } catch {
      console.warn('[DEX] Subsite-Berechtigungen konnten nicht gesetzt werden');
    }
  }

  /**
   * Item-Level Permissions auf der Teilnehmerliste setzen.
   */
  private async setItemLevelPermissions(subsiteUrl: string): Promise<void> {
    // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400
    // unter SPFx odata-version 3.0; die Teilnehmerlisten waren dadurch bislang
    // ungeschützt auf 1/1 statt „nur eigene Elemente" 2/2).
    await this._setListSecurity(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')`, { ReadSecurity: 2, WriteSecurity: 2 });
  }

  /**
   * v20.5: Setzt nachträglich den Autor (Created By / SharePoint-Ersteller)
   * einer Teilnehmer-Zeile auf den TEILNEHMER selbst.
   *
   * Hintergrund: Die Teilnehmerlisten laufen mit Item-Level-Security
   * (ReadSecurity=2 / WriteSecurity=2) — ein User darf nur Items LESEN und
   * BEARBEITEN, die ER ERSTELLT hat (geprüft am Autor, NICHT am Feld
   * ParticipantEmail). Bei einer stellvertretenden Anmeldung (Organizer/Admin
   * meldet eine andere Person an) wäre sonst der Akteur der Autor — der
   * angemeldete Teilnehmer sähe seine eigene Anmeldung NICHT in "Meine
   * Events" und könnte sich nicht selbst abmelden. Indem der Teilnehmer zum
   * Autor wird, bekommt er Lese- + Abmelde-Zugriff auf SEINE Zeile.
   *
   * Best-effort: das Umsetzen von AuthorId erfordert "Listen verwalten" /
   * Full Control auf der Liste. Organizer (eigenes Event) und Admin haben das;
   * ein normaler Contribute-User (z.B. eine Assistenz) NICHT — dort schlägt
   * der MERGE mit 403 fehl und wird still ignoriert (die Zeile bleibt beim
   * Akteur als Autor). Der eigentliche Akteur ist ohnehin separat im Feld
   * RegisteredByEmail protokolliert, der Audit-Nachweis bleibt also erhalten.
   */
  /**
   * v24.40: Eine Teilnehmer-Zeile einer **Assistenz** zuordnen — der Admin
   * übergibt damit die Verwaltung der (Fremd-)Anmeldung an eine bestimmte
   * Assistenz. Setzt ZWEI Dinge:
   *  1. `RegisteredByEmail`/`RegisteredByName` auf die Assistenz (Audit + Filter
   *     der „Assistenz"-Kachel).
   *  2. Den **Zeilen-Autor** (`Created By` / `AuthorId`) auf die Assistenz —
   *     unter Item-Level-Security („nur eigene Elemente") ist das die
   *     Voraussetzung, damit eine NORMALE Assistenz die Zeile überhaupt
   *     lesen/bearbeiten darf (sonst sieht sie sie in ihrer Kachel nicht).
   * Best-effort: Schlägt der Autor-Wechsel mangels Rechten fehl, landet er in
   * der `DEX_AccessFix`-Queue (Flow setzt ihn nach). Gibt zurück, ob der
   * RegisteredBy-Schreibvorgang gelang.
   */
  public async assignRegistrationToAssistant(
    subsiteUrl: string,
    itemId: number,
    assistantEmail: string,
    assistantName: string
  ): Promise<boolean> {
    try {
      const merge = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'RegisteredByEmail': assistantEmail, 'RegisteredByName': assistantName }
      );
      // Zeilen-Autor auf die Assistenz (Voraussetzung für ILS-Lesezugriff).
      await this.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, itemId, assistantEmail);
      return merge.ok;
    } catch (err) {
      console.warn('[DEX] assignRegistrationToAssistant error:', err);
      return false;
    }
  }

  private async trySetItemAuthor(subsiteUrl: string, listName: string, itemId: number, participantEmail: string): Promise<void> {
    try {
      // 1. Teilnehmer als SP-User der Subsite sicherstellen + dessen Id holen.
      const ensureResp = await this._sp.post(
        `${subsiteUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
          },
          body: JSON.stringify({ logonName: participantEmail }),
        }
      );
      if (!ensureResp.ok) {
        // v20.7: z.B. Assistenz ohne ausreichende Rechte auf der Subsite →
        // Auftrag in die DEX_AccessFix-Queue, der Flow setzt den Autor.
        await this.queueAccessFix(subsiteUrl, itemId, participantEmail);
        return;
      }
      const u = await ensureResp.json();
      const userId: number = u?.Id || u?.d?.Id || 0;
      if (!userId) {
        await this.queueAccessFix(subsiteUrl, itemId, participantEmail);
        return;
      }
      // 2. AuthorId der Zeile auf den Teilnehmer setzen (nometadata-MERGE,
      //    daher KEIN __metadata im Body). 403 = fehlende Rechte → Queue.
      const m = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`,
        { 'AuthorId': userId }
      );
      if (!m.ok) {
        // v20.7: typischer Assistenz-Fall — Contribute reicht nicht, um den
        // Autor zu setzen (braucht "Listen verwalten"). Flow übernimmt.
        await this.queueAccessFix(subsiteUrl, itemId, participantEmail);
      }
    } catch {
      // Best-effort: auch hier den Flow-Auftrag versuchen — scheitert auch
      // der, bleibt die Zeile beim Akteur (Verhalten wie vor v20.5).
      try { await this.queueAccessFix(subsiteUrl, itemId, participantEmail); } catch { /* */ }
    }
  }

  /**
   * v20.6: Reparatur-Werkzeug (Admin) — prüft EINE Teilnehmerliste und
   * repariert den Zugriff bei Fremd-Anmeldungen. Zwei Schritte:
   *
   * 1. **Item-Level-Security verifizieren:** liest ReadSecurity/WriteSecurity
   *    der Liste. Steht sie NICHT auf 2/2 („nur eigene Elemente" — z.B. weil
   *    der Set beim Anlegen still fehlschlug, siehe Security-Audit v20.x),
   *    wird sie neu gesetzt und per Read-back verifiziert.
   * 2. **Fremd-Anmeldungen (Anmeldung durch Dritte):** lädt alle Items mit
   *    Autor und setzt bei jedem Item, dessen `RegisteredByEmail` von der
   *    `ParticipantEmail` abweicht UND dessen Autor noch nicht der Teilnehmer
   *    ist, den Autor auf den Teilnehmer (`ensureuser` → `AuthorId`-MERGE,
   *    pro E-Mail gecacht). Damit sieht die angemeldete Person ihre eigene
   *    Zeile in „Meine Events" und kann sich selbst abmelden (v20.5-Logik,
   *    rückwirkend für Bestands-Anmeldungen).
   *
   * Läuft sequentiell (SP-Throttling-Schonung). `onProgress` meldet den
   * Item-Fortschritt für die UI. Externe Teilnehmer (kein Tenant-Login)
   * scheitern am `ensureuser` und landen in `authorFailed` — erwartbar.
   */
  public async repairProxyRegistrationAccess(
    subsiteUrl: string,
    onProgress?: (done: number, total: number) => void
  ): Promise<{ ilsWasWrong: boolean; ilsFixed: boolean; itemsTotal: number; proxyFound: number; authorFixed: number; authorFailed: number }> {
    const result = { ilsWasWrong: false, ilsFixed: false, itemsTotal: 0, proxyFound: 0, authorFixed: 0, authorFailed: 0 };

    // ---- Schritt 1: Listen-Sicherheit „nur eigene Elemente" sicherstellen ----
    // v21 FIX: Der v20.6-Check meldete fälschlich ALLE Listen als unsicher
    // („27 falsch, 0 repariert"), obwohl die ILS nachweislich aktiv war
    // (Fremd-Zeilen unsichtbar). Ursache: Antwortformat/Typ der
    // ReadSecurity-Property nicht deterministisch behandelt. Jetzt: explizit
    // nometadata anfordern, Werte hart zu Zahlen koerzieren und bei
    // Unplausibilität die ROHE Antwort loggen statt „falsch" zu raten —
    // nur ein KLARER numerischer Wert ungleich 2 zählt als unsicher.
    const readSecurity = async (): Promise<{ rs: number; ws: number } | null> => {
      try {
        const resp = await this._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')?$select=ReadSecurity,WriteSecurity`,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (!resp.ok) return null;
        const d = await resp.json();
        const rawRs = d.ReadSecurity ?? d.d?.ReadSecurity;
        const rawWs = d.WriteSecurity ?? d.d?.WriteSecurity;
        const rs = Number(rawRs);
        const ws = Number(rawWs);
        if (!Number.isFinite(rs) || !Number.isFinite(ws)) {
          console.warn('[DEX][ILS-Check] ReadSecurity/WriteSecurity nicht lesbar — rohe Antwort:', subsiteUrl, JSON.stringify(d).slice(0, 400));
          return null;
        }
        return { rs, ws };
      } catch (e) {
        console.warn('[DEX][ILS-Check] Lesen fehlgeschlagen:', subsiteUrl, e);
        return null;
      }
    };
    const before = await readSecurity();
    if (before && (before.rs !== 2 || before.ws !== 2)) {
      console.warn(`[DEX][ILS-Check] Liste meldet ReadSecurity=${before.rs}/WriteSecurity=${before.ws} (erwartet 2/2):`, subsiteUrl);
      result.ilsWasWrong = true;
      await this.setItemLevelPermissions(subsiteUrl);
      const after = await readSecurity();
      result.ilsFixed = !!after && after.rs === 2 && after.ws === 2;
      if (!result.ilsFixed) {
        console.warn('[DEX][ILS-Check] Read-back nach Fix weiterhin abweichend:', subsiteUrl, after);
      }
    }

    // ---- Schritt 2: Items mit Autor laden (paged) ----
    type Row = { Id: number; ParticipantEmail?: string; RegisteredByEmail?: string; Author?: { EMail?: string } };
    const items: Row[] = [];
    let url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,RegisteredByEmail,Author/EMail&$expand=Author&$top=500`;
    while (url) {
      const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      const arr: Row[] = data.value || data.d?.results || [];
      items.push(...arr);
      url = data['odata.nextLink'] || data['@odata.nextLink'] || data.d?.__next || '';
    }
    result.itemsTotal = items.length;

    // ---- Schritt 3: Fremd-Anmeldungen → Autor auf Teilnehmer setzen ----
    const userIdCache: Record<string, number> = {};
    let done = 0;
    for (const it of items) {
      done++;
      const pe = (it.ParticipantEmail || '').toLowerCase().trim();
      const rb = (it.RegisteredByEmail || '').toLowerCase().trim();
      const au = (it.Author?.EMail || '').toLowerCase().trim();
      // Nur Fremd-Anmeldungen: RegisteredByEmail vorhanden UND != Teilnehmer.
      // (Alt-Bestand ohne RegisteredByEmail = vor v3.x — dort ist der Autor
      // ohnehin der Teilnehmer selbst, weil es nur Selbst-Anmeldung gab.)
      if (!pe || !rb || pe === rb) { if (onProgress) onProgress(done, items.length); continue; }
      result.proxyFound++;
      // Autor stimmt schon (z.B. v20.5-Anmeldung oder früherer Lauf) → ok.
      if (au === pe) { if (onProgress) onProgress(done, items.length); continue; }
      try {
        let uid = userIdCache[pe] || 0;
        if (!uid) {
          const er = await this._sp.post(
            `${subsiteUrl}/_api/web/ensureuser`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata',
              },
              body: JSON.stringify({ logonName: pe }),
            }
          );
          if (er.ok) {
            const u = await er.json();
            uid = u?.Id || u?.d?.Id || 0;
            if (uid) userIdCache[pe] = uid;
          }
        }
        if (!uid) {
          result.authorFailed++;
          if (onProgress) onProgress(done, items.length);
          continue;
        }
        const m = await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${it.Id})`,
          { 'AuthorId': uid }
        );
        if (m.ok) result.authorFixed++; else result.authorFailed++;
      } catch {
        result.authorFailed++;
      }
      if (onProgress) onProgress(done, items.length);
    }
    return result;
  }

  /**
   * Berechtigungen für Teilnehmerliste auf der Subsite setzen.
   */
  private async setRegistrationListPermissions(subsiteUrl: string, organizerEmail: string): Promise<void> {
    try {
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      // Site Owners der Hauptsite: Full Control
      const ownersResponse = await this._sp.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
          {}
        );
      }

      // Visitors: Contribute (damit User sich registrieren können)
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`,
          {}
        );
      }

      // Organizer: Full Control. v9.18: organizerEmail kann ";"-separiert
      // mehrere Emails enthalten (Hauptorganizer + Co-Organizer).
      if (organizerEmail) {
        const emails = organizerEmail.split(/[;,]/).map(s => s.trim()).filter(Boolean);
        for (const em of emails) {
          try {
            const userResponse = await this._sp.get(
              `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
              SPHttpClient.configurations.v1
            );
            if (userResponse.ok) {
              const userData = await userResponse.json();
              await this._post(
                `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userData.Id}, roledefid=1073741829)`,
                {}
              );
            }
          } catch { /* Organizer-Berechtigung optional */ }
        }
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden
    }
  }

  // ==================== Registrierungen ====================

  /**
   * Registrierung für ein Event erstellen.
   * Operiert auf der Subsite des Events.
   */
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
    try {
      // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
      // Serverseitige Prüfungen — nicht perfekt (SPFx läuft im Browser),
      // aber fangt naiven App-Bypass (F12, direkter Service-Aufruf) ab.
      const sessionEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const targetEmail = (participantEmail || '').toLowerCase();

      // Event-Metadaten laden (Deadline + OrganizerEmail) über SubsiteUrl.
      // Beide Checks nutzen die gleiche Abfrage — einmal laden, mehrfach prüfen.
      let eventDeadline = '';
      let eventOrganizerEmails: string[] = [];
      try {
        const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
        const evResp = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail`,
          SPHttpClient.configurations.v1
        );
        if (evResp.ok) {
          const evData = await evResp.json();
          const items = evData.value || evData.d?.results || [];
          if (items.length > 0) {
            eventDeadline = items[0].RegistrationDeadline || '';
            const orgStr: string = EventService.stripNoteWrapper(items[0].OrganizerEmail);
            eventOrganizerEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
          }
        }
      } catch { /* Bei Load-Fehler konservativ weitermachen — andere Checks greifen */ }

      // Check A: Darf der User für eine andere Person registrieren?
      // v19.9: Wenn der Client bereits bestätigt hat, dass der/die Anmeldende
      // Organizer/Co-Organizer dieses Events ist, vertrauen wir dem (gleiche
      // Datengrundlage wie die Button-Sichtbarkeit) und überspringen die
      // fragile serverseitige Ableitung. Sonst Fallback auf canRegisterForOthers
      // (deckt Admin-Rolle + Assistant-Ausnahme zuverlässig ab).
      if (targetEmail && targetEmail !== sessionEmail) {
        const allowed = actorIsEventOrganizer || clientAssistantAllowed || await this.canRegisterForOthers(subsiteUrl, participantEmail);
        if (!allowed) {
          console.warn(`[DEX] registerForEvent DENIED: ${sessionEmail} versuchte ${targetEmail} zu registrieren — weder Organizer noch Admin noch erlaubter Assistant-Fall.`);
          return { ok: false, reason: 'not-allowed' };
        }
      }

      // Check B: Deadline abgelaufen? Nur Event-Organizer + Admin dürfen nach
      // Deadline registrieren (auch für sich selbst). Assistant NICHT — das ist
      // wie ein normaler User.
      if (eventDeadline) {
        const deadlineDate = new Date(eventDeadline);
        if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
          const isEventOrganizer = eventOrganizerEmails.indexOf(sessionEmail) >= 0;
          let isAdmin = false;
          try {
            const esc = sessionEmail.replace(/'/g, "''");
            const roleResp = await this._sp.get(
              `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
              SPHttpClient.configurations.v1
            );
            if (roleResp.ok) {
              const rd = await roleResp.json();
              const rItems = rd.value || rd.d?.results || [];
              if (rItems.length > 0 && (rItems[0].Role === 'Admin' || rItems[0].Role === 'IT-Admin')) isAdmin = true;
            }
          } catch { /* ignore */ }

          if (!isEventOrganizer && !isAdmin) {
            console.warn(`[DEX] registerForEvent DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${eventDeadline} zu registrieren — weder Event-Organizer noch Admin.`);
            return { ok: false, reason: 'deadline' };
          }
        }
      }
      // ---- Ende Permission-Checks ----

      // v7.28 / v9.10: Nächste TeilnehmerID atomar über den Subsite-Counter
      // holen (ETag-CAS, verhindert Race-Conditions bei parallelen Anmeldungen).
      // Counter wird bei Bedarf on-demand angelegt + geseeded.
      //
      // v9.10: Der alte race-anfällige Fallback "max+1" wurde entfernt — bei
      // Massen-Anmeldungen (Go-Live große Events) hat er Duplikate produziert,
      // weil zwei Clients gleichzeitig den gleichen Max-Wert lesen und beide
      // mit Max+1 schreiben. Wenn der atomare Counter ausnahmsweise gar nicht
      // erreichbar ist, lassen wir TeilnehmerID undefined und der Admin
      // lädt anschliessend "IDs neu vergeben" — Lückenfreiheit ist nicht
      // hart kritisch, Eindeutigkeit ist es.
      const nextId = await this.getNextTeilnehmerId(subsiteUrl);

      // Profildaten laden - für den TATSAECHLICHEN Teilnehmer (nicht den eingeloggten User!)
      // Wenn jemand für eine andere Person registriert, muss deren Profil geladen werden,
      // sonst wird der eigene JobTitle/Department/Office in deren Teilnehmer-Eintrag geschrieben.
      const myEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const profile = participantEmail.toLowerCase() === myEmail
        ? await this.getCurrentUserProfile()
        : await this.getUserProfileByEmail(participantEmail);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        // v9.10: TeilnehmerID nur setzen wenn der atomare Counter sie geliefert hat.
        // Bei Counter-Outage bleibt das Feld leer — Admin kann nachträglich
        // "IDs neu vergeben" laufen lassen, was sequentielle IDs setzt.
        ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
        'Anrede': customData.salutation || '',
        'Vorname': firstName,
        'Nachname': surname,
        'ParticipantName': `${firstName} ${surname}`,
        'ParticipantEmail': participantEmail,
        'Department': profile.department,
        'Location': profile.location,
        'JobTitle': profile.jobTitle,
        'Phone': profile.phone,
        // v24.29: Unternehmenszugehörigkeit / Rechtsträger mitschreiben.
        'Company': profile.company,
        'Status': status,
        'RegistrationDate': new Date().toISOString(),
        'CustomData': JSON.stringify(customData),
      };

      // Audit: wer hat die Anmeldung ausgelöst?
      // Bei Self-Registration = der User selbst. Bei "Für andere Person registrieren"
      // = der Organizer/Admin der geklickt hat. Fallback wenn nichts übergeben: aus pageContext.
      const auditName = registeredByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (registeredByEmail || this.context.pageContext.user.email || '').toLowerCase();
      if (auditName) payload['RegisteredByName'] = auditName;
      if (auditEmail) payload['RegisteredByEmail'] = auditEmail;
      // v18.74: Zustimmungs-Nachweis bei stellvertretender Anmeldung.
      if (proxyConsent) payload['ProxyConsent'] = proxyConsent;

      // B2Run: Starter-Typ + Wunsch-Typ schreiben (bei normalen Events null)
      if (starterType) payload['StarterType'] = starterType;
      if (preferredStarterType) payload['PreferredStarterType'] = preferredStarterType;

      // Custom Field Werte in die echten SP-Spalten schreiben.
      // Wichtig: Wenn spInternalName fehlt (z.B. weil der Admin das Feld später
      // ergänzt hat ohne Spalte in der Teilnehmerliste), würde der Wert
      // SILENT VERLOREN GEHEN — deshalb ein console.warn damit der Admin im
      // Admin Center per "Custom Fields prüfen" das Mapping fixen kann.
      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            payload[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
          }
        }
      }

      let response = await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
        payload
      );
      // v19.9 BUG-FIX: Stellvertretende Anmeldung schlug auf älteren
      // Teilnehmerlisten fehl, weil die erst mit v18.74 eingeführte Spalte
      // `ProxyConsent` dort noch nicht existiert — der Insert mit
      // `ProxyConsent` im Body wird dann von SharePoint mit HTTP 400
      // ("property does not exist") abgewiesen. Da `ProxyConsent` NUR bei
      // Stellvertreter-Anmeldungen im Body steht, scheiterte ausschließlich der
      // Proxy-Pfad (Selbst-Anmeldung lief, weil dort kein ProxyConsent gesetzt
      // wird) — für den User sah es aus wie „Person bereits angemeldet", obwohl
      // gar nichts gespeichert wurde. Fix: Insert einmal OHNE das optionale
      // Audit-Feld wiederholen, damit die Anmeldung nicht an einer fehlenden
      // Spalte scheitert. Der Zustimmungs-Nachweis geht dann verloren (der
      // Admin kann die Spalte per „Spalten fixen" nachrüsten), die Anmeldung
      // selbst gelingt aber.
      // v24.32: Gleiches Schutz-Muster jetzt AUCH für die v24.29-Spalte
      // `Company` — auf Teilnehmerlisten, auf denen „Spalten fixen" noch nicht
      // lief, existiert sie nicht → der Insert mit `Company` im Body würde sonst
      // mit HTTP 400 scheitern und die GANZE Anmeldung kaputtmachen. Deshalb bei
      // einem fehlgeschlagenen Insert das optionale Feld strippen und erneut
      // versuchen. Folge ohne Spalte: Anmeldung gelingt, Unternehmens-Wert wird
      // nicht getrackt (Admin kann „Spalten fixen" nachziehen).
      if (!response.ok && (payload['ProxyConsent'] || payload['Company'])) {
        console.warn('[DEX] registerForEvent: Insert fehlgeschlagen — Retry OHNE ProxyConsent/Company (Spalte evtl. nicht vorhanden). Bitte im Admin Center "Spalten fixen" ausführen.');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const retryPayload: Record<string, any> = { ...payload };
        delete retryPayload['ProxyConsent'];
        delete retryPayload['Company'];
        response = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
          retryPayload
        );
      }
      if (!response.ok) {
        // v30.58: Den Grund AUSLESEN statt nur „insert-failed" zu melden.
        //
        // Bisher endete jeder abgelehnte Insert als anonymes „insert-failed" —
        // und genau das hat die Suche nach der fehlenden Klammer-Zeile so lange
        // aufgehalten: SharePoint sagt in der Antwort, WELCHE Spalte es nicht
        // gibt („The field or property 'X' does not exist"), nur hat es nie
        // jemand gelesen. Der Retry oben strippt bloß zwei fest verdrahtete
        // Spalten (`ProxyConsent`, `Company`); ein nachträglich angelegtes
        // Abfragefeld, dessen Spalte auf DIESER Liste fehlt, bringt den ganzen
        // Insert zu Fall — und zwar nur bei den Personen, die das Feld
        // ausfüllen. Das erklärt, warum es immer dieselben wenigen trifft und
        // nicht zufällig streut wie eine Drosselung.
        let detail = '';
        try {
          const txt = await response.text();
          const m = txt.match(/"value"\s*:\s*"([^"]{0,400})"/);
          detail = (m ? m[1] : txt).slice(0, 400);
        } catch { /* Body nicht lesbar */ }
        console.warn('[DEX] registerForEvent: Insert abgelehnt', {
          status: response.status, subsiteUrl, detail,
          felder: Object.keys(payload),
        });
        return { ok: false, reason: 'insert-failed', detail };
      }

      // Inserted-Item-Id EINMALIG aus der Response lesen (der Body lässt sich
      // nur einmal konsumieren) — wird sowohl für die Dedup-Prüfung als auch
      // für das Setzen des Autors (stellvertretende Anmeldung) gebraucht.
      let insertedId = 0;
      try {
        const respJson = await response.json();
        insertedId = respJson?.d?.Id || respJson?.Id || 0;
      } catch { /* Body nicht lesbar — Dedup/Autor-Set entfallen, Insert war ok */ }

      // v9.10: Post-Insert Safety Net — bei Massen-Anmeldungen (Go-Live)
      // gab es trotz ETag-Counter vereinzelt Duplikate. Ursache war der
      // alte max+1-Fallback (jetzt entfernt) und ggf. Edge-Cases im
      // Counter-Pfad. Als zusätzliche Versicherung: nach dem Insert
      // prüfen, ob jetzt zwei Einträge dieselbe TeilnehmerID haben.
      // Wenn ja: der mit der HOEHEREN SP-Item-Id verliert (= der spätere
      // Insert), holt sich frisch eine ID am Counter und patcht sich.
      // So bleiben die zuerst eingetroffenen Anmeldungen stabil.
      if (typeof nextId === 'number' && nextId > 0 && insertedId > 0) {
        try {
          const dupResp = await this._sp.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,TeilnehmerID&$filter=TeilnehmerID eq ${nextId}&$top=10`,
            SPHttpClient.configurations.v1
          );
          if (dupResp.ok) {
            const dupData = await dupResp.json();
            const dupItems: Array<{ Id: number; TeilnehmerID: number }> = dupData.value || dupData.d?.results || [];
            if (dupItems.length > 1) {
              const minId = Math.min(...dupItems.map(d => d.Id));
              if (insertedId !== minId) {
                // Wir haben verloren — fresh ID holen + patchen
                const fresh = await this.getNextTeilnehmerId(subsiteUrl);
                if (typeof fresh === 'number' && fresh > 0) {
                  await this._merge(
                    `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${insertedId})`,
                    { 'TeilnehmerID': fresh }
                  );
                  console.warn(`[DEX] Post-insert dedup: TeilnehmerID ${nextId} kollidierte, Item ${insertedId} hat jetzt #${fresh}.`);
                } else {
                  console.warn(`[DEX] Post-insert dedup: kollidierende TeilnehmerID ${nextId} entdeckt, aber Counter lieferte keine fresh ID. Admin sollte "IDs neu vergeben" laufen lassen.`);
                }
              }
            }
          }
        } catch (err) {
          // Safety-Net-Fehler nicht kritisch — Insert war erfolgreich
          console.warn('[DEX] Post-insert dedup check fehlgeschlagen:', err);
        }
      }

      // v20.5: Stellvertretende Anmeldung (Akteur != Teilnehmer) → den
      // Teilnehmer zum Autor der Zeile machen, damit er seine eigene Anmeldung
      // in "Meine Events" sieht und sich selbst abmelden kann. auditEmail ist
      // bereits lowercased; bei Selbst-Anmeldung sind beide gleich → kein Set.
      if (insertedId > 0 && auditEmail && participantEmail && participantEmail.toLowerCase().trim() !== auditEmail) {
        await this.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, insertedId, participantEmail);
      }

      return { ok: true };
    } catch {
      return { ok: false, reason: 'error' };
    }
  }

  /**
   * v11.82: Ein einzelnes Teilnehmer-Item im Team-Modus anlegen.
   *
   * Unterschied zu `registerForEvent`: kein eigener Permission-Check (der
   * Aufrufer hat schon im Team-Submit alle Mitglieder validiert), kein
   * Post-Insert Dedup-Loop (der ist im Team-Pfad überflüssig — wenn ein
   * Member mit Kollision verliert, fixt es der Folge-IDReorder). Nimmt
   * Profil-Daten und Anzeige-Namen direkt entgegen, weil der Lead-Submit
   * pro Member ohnehin schon das Graph-Profil geladen hat.
   */
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
    try {
      const nextId = await this.getNextTeilnehmerId(subsiteUrl);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': args.email,
        ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
        'Anrede': args.salutation || '',
        'Vorname': args.firstName,
        'Nachname': args.lastName,
        'ParticipantName': `${args.firstName} ${args.lastName}`.trim(),
        'ParticipantEmail': args.email,
        'Department': args.profile.department,
        'Location': args.profile.location,
        'JobTitle': args.profile.jobTitle,
        'Phone': args.profile.phone,
        'Company': args.profile.company || '',
        'Status': args.status,
        'RegistrationDate': new Date().toISOString(),
        'TeamId': args.teamId,
        'TeamLead': !!args.teamLead,
        'TeamName': args.teamName || '',
        'CustomData': JSON.stringify(args.customData || {}),
      };
      if (args.registeredByName) payload['RegisteredByName'] = args.registeredByName;
      if (args.registeredByEmail) payload['RegisteredByEmail'] = args.registeredByEmail;
      if (args.starterType) payload['StarterType'] = args.starterType;
      if (args.preferredStarterType) payload['PreferredStarterType'] = args.preferredStarterType;
      if (args.customFieldMap && args.customData) {
        for (const cfId of Object.keys(args.customData)) {
          if (cfId === 'salutation') continue;
          const v = args.customData[cfId];
          if (!v) continue;
          const spName = args.customFieldMap[cfId];
          if (spName) payload[spName] = v;
        }
      }
      let response = await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
        payload
      );
      // v24.32: Retry OHNE Company, falls die Spalte auf der Liste fehlt (s.
      // registerForEvent) — sonst bräche die Team-Anmeldung auf Alt-Listen.
      if (!response.ok && payload['Company']) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const retryPayload: Record<string, any> = { ...payload };
        delete retryPayload['Company'];
        response = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
          retryPayload
        );
      }
      if (!response.ok) return { ok: false };
      try {
        const respJson = await response.json();
        const itemId: number = respJson?.d?.Id || respJson?.Id || 0;
        return { ok: true, teilnehmerId: typeof nextId === 'number' ? nextId : undefined, itemId };
      } catch {
        return { ok: true, teilnehmerId: typeof nextId === 'number' ? nextId : undefined };
      }
    } catch {
      return { ok: false };
    }
  }

  /**
   * v11.82: Alle Mitglieder eines Teams (per TeamId) zu einer Registrierung
   * laden — wird in „Meine Events" zum Rendern des Team-Badges genutzt.
   */
  public async getTeamMembers(subsiteUrl: string, teamId: string): Promise<SPRegistration[]> {
    if (!teamId) return [];
    try {
      const tidEsc = teamId.replace(/'/g, "''");
      const response = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=TeamId eq '${tidEsc}'&$top=100&$orderby=TeamLead desc,Id asc`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return [];
      const data = await response.json();
      return data.value || data.d?.results || [];
    } catch {
      return [];
    }
  }

  /**
   * v11.83: Auf einer existierenden Teilnehmer-Zeile das Feld TeamLead
   * auf true setzen (Auto-Promote nach Lead-Cancel). MERGE auf der
   * Teilnehmerliste — die Subsite kennt das Item über `itemId`.
   */
  public async promoteToTeamLead(subsiteUrl: string, itemId: number): Promise<boolean> {
    try {
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
      const resp = await this._merge(url, { TeamLead: true });
      return !!resp.ok;
    } catch {
      return false;
    }
  }

  /**
   * v11.84: Lead-Rolle innerhalb eines Teams von einer Person auf eine andere
   * übergeben. Wird im Admin Center per Dropdown im Teams-Block ausgelöst.
   * Best-effort transaktional: erst die neue Lead-Zeile auf TeamLead=true
   * setzen, danach die alte auf TeamLead=false. Schlägt der zweite MERGE
   * fehl, gibt es kurzfristig zwei Leads — der Aufrufer kann dann erneut
   * versuchen oder die Liste manuell reparieren. Keine echte Transaktion,
   * SharePoint bietet sowas auf Listen-Ebene nicht.
   */
  public async transferTeamLead(
    subsiteUrl: string,
    fromLeadItemId: number,
    toNewLeadItemId: number
  ): Promise<boolean> {
    if (!subsiteUrl || !fromLeadItemId || !toNewLeadItemId || fromLeadItemId === toNewLeadItemId) {
      return false;
    }
    try {
      const newUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${toNewLeadItemId})`;
      const r1 = await this._merge(newUrl, { TeamLead: true });
      if (!r1.ok) return false;
      const oldUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${fromLeadItemId})`;
      const r2 = await this._merge(oldUrl, { TeamLead: false });
      return !!r2.ok;
    } catch {
      return false;
    }
  }

  /**
   * v11.83: Prüfen, ob eine bestimmte Email-Adresse bereits aktiv beim
   * Event angemeldet ist (Status in Angemeldet/QR versendet/Eingecheckt/
   * Warteliste). Wird vor jedem Team-Add (Initial, Add-Member, Beitritt)
   * benutzt, um Doppel-Anmeldungen sauber abzuweisen, bevor ein Sitzplatz
   * reserviert wird.
   *
   * Rückgabe: true = blockieren, false = frei (auch bei SP-Fehlern, weil
   * der Aufrufer dann auf die strikteren Stellen-internen Checks zurück-
   * fällt; ein lauter Throw würde den Pfad unnötig abbrechen).
   */
  public async isUserAlreadyOnEvent(subsiteUrl: string, email: string): Promise<boolean> {
    if (!subsiteUrl || !email) return false;
    try {
      const emEsc = email.trim().replace(/'/g, "''");
      const blockingStatuses = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
      const statusClause = blockingStatuses.map(s => `Status eq '${s}'`).join(' or ');
      const filter = `(ParticipantEmail eq '${emEsc}') and (${statusClause})`;
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${encodeURIComponent(filter)}&$top=1&$select=Id,Status,ParticipantEmail`;
      const response = await this._sp.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return false;
      const data = await response.json();
      const items = data.value || data.d?.results || [];
      return items.length > 0;
    } catch {
      return false;
    }
  }

  // ==================== DEX_TeamJoinRequests (v11.83) ====================

  /**
   * v11.83: Globale Liste für Team-Beitritts-Anfragen (Approve-Queue).
   * Liegt auf der Site-Collection-Ebene (nicht pro Subsite), damit alle
   * Events darauf zugreifen können und der Team-Lead alle ausstehenden
   * Anfragen in einer einzigen Query findet.
   *
   * Spalten:
   * - Title: Anzeige-Zusammenfassung "RequesterName -> Event-Title"
   * - EventId: ID des Events in DEX_Events
   * - TeamId: UUID der Team-Anmeldung
   * - RequesterEmail: Email des Anfragenden
   * - RequesterDisplayName: Anzeigename des Anfragenden
   * - Status: Pending / Approved / Rejected
   * - DecidedDate: Wann hat der Team-Lead entschieden
   * - DecidedByEmail: Email des entscheidenden Leads
   */
  public async ensureTeamJoinRequestsList(): Promise<void> {
    const listName = 'DEX_TeamJoinRequests';
    const exists = await this.listExists(listName);
    if (exists) {
      // v13.0: Backfill für ältere Installationen, die die Liste vor
      // v11.83 angelegt haben (DecidedDate/DecidedByEmail damals nicht
      // vorhanden). Ohne diesen Patch schlägt decideTeamJoinRequest
      // beim MERGE auf die fehlenden Felder mit HTTP 400 fehl.
      await this.ensureMissingTeamJoinRequestsFields(listName);
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Approve-Queue für Team-Beitritts-Anfragen (v11.83+).',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'EventId', type: 2 },
      { title: 'TeamId', type: 2 },
      { title: 'RequesterEmail', type: 2 },
      { title: 'RequesterDisplayName', type: 2 },
      { title: 'Status', type: 6, choices: ['Pending', 'Approved', 'Rejected'], metaType: 'SP.FieldChoice' },
      { title: 'DecidedDate', type: 4 },
      { title: 'DecidedByEmail', type: 2 },
      // v18.73: event-spezifische Antworten des Anfragenden als JSON — werden
      // beim Approve auf den neuen Member angewandt (Note = multi-line text).
      { title: 'CustomData', type: 3 },
    ];

    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) payload['Choices'] = { 'results': f.choices };
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    await this.configureDefaultView(listName, [
      'EventId', 'TeamId', 'RequesterEmail', 'RequesterDisplayName',
      'Status', 'Created', 'DecidedDate', 'DecidedByEmail',
    ]);

    // Schreibrechte für alle Authentifizierten (analog zu DEX_Emails-Queue):
    // jeder darf eine Anfrage erstellen, aber Item-Level-Security greift
    // sowieso über den Lead-Check beim Approve-Pfad.
    try {
      await this.setQueueListPermissions(listName);
    } catch { /* best-effort */ }
  }

  /**
   * v13.0: Backfill fehlender Felder in einer bestehenden DEX_TeamJoinRequests-
   * Liste. Greift bei Tenants die die Liste vor v11.83 angelegt haben.
   */
  private async ensureMissingTeamJoinRequestsFields(listName: string): Promise<void> {
    const wanted = [
      { title: 'DecidedDate', type: 4 },
      { title: 'DecidedByEmail', type: 2 },
      // v18.73: CustomData-Spalte auf Bestands-Listen nachziehen.
      { title: 'CustomData', type: 3 },
    ];
    for (const f of wanted) {
      try {
        const resp = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${f.title}')?$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (resp.ok) continue; // existiert
      } catch { /* anlegen */ }
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        });
      } catch (e) {
        console.warn(`[DEX] ensureMissingTeamJoinRequestsFields: failed to add '${f.title}':`, e);
      }
    }
  }

  /**
   * v18.48: Sperr-Liste für den Outlook-Einladungs-Flow (DEX_Outlook_Einladungen).
   *
   * Hintergrund: der Einladungs-Flow patcht pro Anmeldung/Abmeldung die
   * KOMPLETTE Teilnehmerliste eines Outlook-Termins (bei Grossevents bis zu
   * 1500 Personen) an Microsoft Graph. Lief der Flow seriell (Concurrency 1),
   * standen Anmeldungen für völlig UNTERSCHIEDLICHE Events stundenlang in
   * der Warteschlange. Lösung „Option B": die Trigger-Concurrency wird hoch-
   * gesetzt (z.B. 25 parallele Läufe), und ein Pro-Event-Lock verhindert,
   * dass zwei Läufe für dasSELBE Event gleichzeitig die Teilnehmerliste
   * lesen-und-schreiben (Race -> verlorene Einträge).
   *
   * Die Liste hat eine eindeutige (Enforce-Unique) Spalte `EventId`. Der Flow
   * „erwirbt" den Lock per Create-Item: gelingt das Create, hat er den Lock;
   * schlägt es wegen der Eindeutigkeits-Prüfung fehl, hält gerade ein
   * anderer Lauf desselben Events den Lock -> kurz warten und erneut
   * versuchen. Am Ende löscht der Flow das Lock-Item wieder (Release).
   *
   * Die UI-Schritt-für-Schritt-Anleitung steht in `docs/flow-jsons.md` unter
   * „UI-Anleitung 2026-06-02 (v18.48) — Option B: Pro-Event-Lock für
   * parallele Outlook-Läufe".
   */
  /**
   * v20.7: Queue-Liste `DEX_AccessFix` (Site-Collection-Root) für den
   * Assistenz-Fall der Fremd-Anmeldung. Wenn `trySetItemAuthor` mangels
   * „Listen verwalten"-Rechten scheitert (normaler Contribute-User, z.B.
   * Assistenz meldet einen Partner an), schreibt die App hier einen
   * Auftrag — der Power-Automate-Flow `DEX_AccessFix_Autor` (läuft mit
   * Service-Identität, hat Full Control) setzt dann den Zeilen-Autor auf
   * den Teilnehmer und markiert den Auftrag als Done/Failed.
   * Spalten: SubsiteUrl (Text), ItemId (Number), ParticipantEmail (Text),
   * Status (Text: Pending/Done/Failed). Schreibrechte für alle User via
   * setQueueListPermissions (analog DEX_Emails).
   */
  public async ensureAccessFixList(): Promise<void> {
    const listName = 'DEX_AccessFix';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Queue: Zeilen-Autor bei Fremd-Anmeldungen auf den Teilnehmer setzen (v20.7, Flow DEX_AccessFix_Autor).',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    const fields: Array<{ title: string; type: number }> = [
      { title: 'SubsiteUrl', type: 2 },
      { title: 'ItemId', type: 9 },
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Status', type: 2 },
    ];
    for (const f of fields) {
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        });
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
    try {
      await this.configureDefaultView(listName, ['SubsiteUrl', 'ItemId', 'ParticipantEmail', 'Status', 'Created']);
    } catch { /* View optional */ }
    try {
      await this.setQueueListPermissions(listName);
    } catch { /* best-effort */ }
    // Item-Level-Security: User sehen/ändern nur EIGENE Aufträge (sonst wäre
    // ablesbar, wer wen angemeldet hat). Der Flow läuft als Site-Owner und
    // sieht alle Items („Listen verwalten" hebelt die Item-Beschränkung aus).
    // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400).
    await this._setListSecurity(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 2, WriteSecurity: 2 });
  }

  /**
   * v20.7: Auftrag für den DEX_AccessFix_Autor-Flow einreihen (siehe
   * ensureAccessFixList). Best-effort — Fehler blocken die Anmeldung nie.
   */
  public async queueAccessFix(subsiteUrl: string, itemId: number, participantEmail: string): Promise<void> {
    try {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_AccessFix')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_AccessFixListItem' },
        'Title': `${participantEmail} -> Item ${itemId}`.slice(0, 250),
        'SubsiteUrl': subsiteUrl,
        'ItemId': itemId,
        'ParticipantEmail': participantEmail,
        'Status': 'Pending',
      });
    } catch (err) {
      console.warn('[DEX] queueAccessFix fehlgeschlagen (best-effort):', err);
    }
  }

  /**
   * v24.41: Liste `DEX_AssistantAccess` — Delegations-/Zugriffs-Queue für die
   * „Meine Assistenz beauftragen"-Funktion. Wenn ein Admin/Director sich für
   * ein Event anmeldet und eine Assistenz angibt, wird hier ein Eintrag
   * angelegt. Zwei Zwecke:
   *  1. **Flow-Auftrag:** Der Flow `DEX_AssistantAccess_Grant` setzt den
   *     Zeilen-Autor (`Created By`) der Teilnehmer-Anmeldung auf die Assistenz,
   *     damit diese die Anmeldung in ihrer „Assistenz"-Kachel sieht/bearbeitet
   *     (unter „nur eigene Elemente" geht das NUR über den Autor — siehe
   *     Recherche v24.41). Für Admins setzt die App den Autor direkt; für
   *     normale Directoren erledigt es der Flow.
   *  2. **Info-Zeile für den Director:** Da der Director nach der Delegation
   *     nicht mehr Autor der Teilnehmer-Zeile ist (und sie unter ILS nicht mehr
   *     sieht), liest „Meine Events" hier (der Director ist Autor SEINES
   *     Delegations-Eintrags → sieht ihn) eine schreibgeschützte Zeile
   *     „Angemeldet für X — verwaltet von Assistenz Y".
   * ReadSecurity/WriteSecurity=2: jeder sieht nur die EIGENEN Delegations-
   * Einträge; der Flow (Site-Owner) sieht alle.
   */
  public async ensureAssistantAccessList(): Promise<void> {
    const listName = 'DEX_AssistantAccess';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Queue + Info: Anmeldung an eine Assistenz delegieren — Zeilen-Autor auf die Assistenz setzen (v24.41, Flow DEX_AssistantAccess_Grant).',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    const fields: Array<{ title: string; type: number }> = [
      { title: 'SubsiteUrl', type: 2 },
      { title: 'ItemId', type: 9 },
      { title: 'EventId', type: 2 },
      { title: 'EventTitle', type: 2 },
      { title: 'ParticipantEmail', type: 2 },   // die angemeldete Person
      { title: 'ParticipantName', type: 2 },
      { title: 'AssistantEmail', type: 2 },     // die verknüpfte Assistenz
      { title: 'AssistantName', type: 2 },
      { title: 'OwnerEmail', type: 2 },         // wer die Anmeldung VERWALTET
      { title: 'LinkType', type: 2 },           // 'delegation' (Selbst-Anmeldung+Assistenz) | 'proxy' (Assistenz meldet an)
      { title: 'Status', type: 2 },             // 'Active' | 'Cancelled'
      { title: 'RequestType', type: 2 },        // '' | 'change' | 'cancel'
      { title: 'RequestNote', type: 3 },        // Note
      { title: 'RequestedByEmail', type: 2 },
      { title: 'RequestedByName', type: 2 },
      { title: 'RequestStatus', type: 2 },      // '' | 'Open' | 'Done' | 'Rejected'
    ];
    for (const f of fields) {
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        });
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
    try {
      await this.configureDefaultView(listName, ['SubsiteUrl', 'ItemId', 'EventTitle', 'ParticipantEmail', 'AssistantEmail', 'OwnerEmail', 'LinkType', 'Status', 'RequestType', 'RequestStatus', 'Created']);
    } catch { /* View optional */ }
    try {
      await this.setQueueListPermissions(listName);
    } catch { /* best-effort */ }
    // v24.41: ReadSecurity=1/WriteSecurity=1 — jeder darf lesen UND schreiben.
    // Die Liste enthält bewusst NUR Koordinations-Daten (Verknüpfung Person ↔
    // Assistenz + Anforderungs-Status), KEINE sensiblen Anmelde-Antworten (die
    // bleiben ILS-geschützt auf der Subsite). So sehen beide Seiten ihre
    // relevanten Einträge (App filtert) und können Anforderungen schreiben —
    // alles ohne Flow.
    // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400).
    await this._setListSecurity(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 1, WriteSecurity: 1 });
  }

  /**
   * v24.41: Verknüpfungs-/Koordinations-Eintrag in DEX_AssistantAccess anlegen.
   * `ownerEmail` = wer die Anmeldung verwaltet (Owner). `linkType`:
   * 'delegation' (Person meldet sich selbst an + benennt Assistenz) oder
   * 'proxy' (Assistenz meldet die Person an). Best-effort.
   */
  public async queueAssistantAccess(args: {
    subsiteUrl: string; itemId: number; eventId: string; eventTitle: string;
    participantEmail: string; participantName: string; assistantEmail: string; assistantName: string;
    ownerEmail: string; linkType: 'delegation' | 'proxy';
  }): Promise<void> {
    try {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_AssistantAccessListItem' },
        'Title': `${args.participantEmail} <-> ${args.assistantEmail} (${args.eventTitle || args.eventId})`.slice(0, 250),
        'SubsiteUrl': args.subsiteUrl,
        'ItemId': args.itemId,
        'EventId': args.eventId,
        'EventTitle': (args.eventTitle || '').slice(0, 250),
        'ParticipantEmail': args.participantEmail,
        'ParticipantName': (args.participantName || '').slice(0, 250),
        'AssistantEmail': args.assistantEmail,
        'AssistantName': (args.assistantName || '').slice(0, 250),
        'OwnerEmail': args.ownerEmail,
        'LinkType': args.linkType,
        'Status': 'Active',
        'RequestStatus': '',
      });
    } catch (err) {
      console.warn('[DEX] queueAssistantAccess fehlgeschlagen (best-effort):', err);
    }
  }

  /**
   * v24.41: Alle AKTIVEN Verknüpfungen lesen, die den eingeloggten User
   * betreffen (als angemeldete Person, als verknüpfte Assistenz ODER als Owner).
   * ReadSecurity=1 — der Server liefert alle; wir filtern serverseitig auf die
   * drei Email-Felder. Die App kategorisiert danach (Info-Ansichten + offene
   * Anforderungen an den Owner).
   */
  public async getAssistantLinksForUser(myEmail: string): Promise<AssistantLink[]> {
    const out: AssistantLink[] = [];
    const me = (myEmail || '').toLowerCase().trim();
    if (!me) return out;
    const esc = me.replace(/'/g, "''");
    const filter = `Status eq 'Active' and (ParticipantEmail eq '${esc}' or AssistantEmail eq '${esc}' or OwnerEmail eq '${esc}')`;
    try {
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items?$select=Id,SubsiteUrl,ItemId,EventId,EventTitle,ParticipantEmail,ParticipantName,AssistantEmail,AssistantName,OwnerEmail,LinkType,Status,RequestType,RequestNote,RequestedByEmail,RequestedByName,RequestStatus,Created&$filter=${encodeURIComponent(filter)}&$orderby=Created desc&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return out;
      const data = await resp.json();
      const items: Array<Record<string, string | number>> = data.value || data.d?.results || [];
      for (const it of items) {
        out.push({
          id: Number(it.Id) || 0,
          subsiteUrl: String(it.SubsiteUrl || ''),
          itemId: Number(it.ItemId) || 0,
          eventId: String(it.EventId || ''),
          eventTitle: String(it.EventTitle || ''),
          participantEmail: String(it.ParticipantEmail || ''),
          participantName: String(it.ParticipantName || ''),
          assistantEmail: String(it.AssistantEmail || ''),
          assistantName: String(it.AssistantName || ''),
          ownerEmail: String(it.OwnerEmail || ''),
          linkType: String(it.LinkType || ''),
          status: String(it.Status || ''),
          requestType: String(it.RequestType || ''),
          requestNote: String(it.RequestNote || ''),
          requestedByEmail: String(it.RequestedByEmail || ''),
          requestedByName: String(it.RequestedByName || ''),
          requestStatus: String(it.RequestStatus || ''),
          created: String(it.Created || ''),
        });
      }
    } catch { /* best-effort */ }
    return out;
  }

  /**
   * v24.42: Eine Änderungs-/Abmelde-Anforderung auf einem Link setzen
   * (RequestType/RequestNote/RequestStatus=Open + Anforderer). Schreibbar von
   * jedem (WriteSecurity=1) — der nicht-Owner stellt die Anforderung.
   */
  public async setAssistantLinkRequest(linkId: number, args: {
    requestType: 'change' | 'cancel'; note: string; requestedByEmail: string; requestedByName: string;
  }): Promise<boolean> {
    try {
      const resp = await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items(${linkId})`,
        {
          'RequestType': args.requestType,
          'RequestNote': (args.note || '').slice(0, 1000),
          'RequestedByEmail': args.requestedByEmail,
          'RequestedByName': (args.requestedByName || '').slice(0, 250),
          'RequestStatus': 'Open',
        }
      );
      return resp.ok;
    } catch { return false; }
  }

  /** v24.42: Anforderung als erledigt/abgelehnt markieren (Owner). */
  public async resolveAssistantLinkRequest(linkId: number, decision: 'Done' | 'Rejected'): Promise<boolean> {
    try {
      const resp = await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items(${linkId})`,
        { 'RequestStatus': decision }
      );
      return resp.ok;
    } catch { return false; }
  }

  /** v24.41: Link beim Abmelden auf 'Cancelled' setzen (Info verschwindet). */
  public async setAssistantLinkStatusForRegistration(itemId: number, subsiteUrl: string, status: 'Cancelled'): Promise<void> {
    try {
      const esc = (subsiteUrl || '').replace(/'/g, "''");
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items?$select=Id&$filter=ItemId eq ${itemId} and SubsiteUrl eq '${esc}' and Status eq 'Active'&$top=20`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return;
      const data = await resp.json();
      const items: Array<{ Id: number }> = data.value || data.d?.results || [];
      for (const it of items) {
        try { await this._merge(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items(${it.Id})`, { 'Status': status }); } catch { /* */ }
      }
    } catch { /* best-effort */ }
  }

  /**
   * v24.51: Liste `DEX_InactiveNotices` — Dedup-Marker für die „Organizer über
   * inaktives Konto informieren"-Mail. ReadSecurity=1 (alle Admins lesen
   * dieselben Marker → klickt ein zweiter Admin, wird NICHT erneut gesendet).
   * Spalten: EventId + ParticipantEmail (Schlüssel) + SentByEmail (informativ).
   */
  public async ensureInactiveNoticesList(): Promise<void> {
    const listName = 'DEX_InactiveNotices';
    const exists = await this.listExists(listName);
    if (exists) return;
    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Dedup: Organizer wurde über ein inaktives Konto informiert (v24.51). Eine Mail pro Event+Person, egal welcher Admin klickt.',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    const fields: Array<{ title: string; type: number }> = [
      { title: 'EventId', type: 2 },
      { title: 'ParticipantEmail', type: 2 },
      { title: 'SentByEmail', type: 2 },
    ];
    for (const f of fields) {
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
        });
      } catch { /* */ }
    }
    try { await this.configureDefaultView(listName, ['EventId', 'ParticipantEmail', 'SentByEmail', 'Created']); } catch { /* */ }
    try { await this.setQueueListPermissions(listName); } catch { /* */ }
    // ReadSecurity bleibt 1 (Default) — alle Admins müssen dieselben Marker
    // sehen können (sonst keine Cross-Admin-Dedup).
  }

  /** v24.51: Bereits benachrichtigte (Event+Email)-Paare für ein Event lesen. */
  public async getSentInactiveNotices(eventId: string): Promise<Set<string>> {
    const out = new Set<string>();
    try {
      const esc = (eventId || '').replace(/'/g, "''");
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_InactiveNotices')/items?$select=ParticipantEmail&$filter=EventId eq '${esc}'&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return out;
      const data = await resp.json();
      for (const it of (data.value || data.d?.results || [])) {
        const e = (it.ParticipantEmail || '').toLowerCase().trim();
        if (e) out.add(e);
      }
    } catch { /* */ }
    return out;
  }

  /** v24.51: Benachrichtigungs-Marker anlegen (claim). */
  public async recordInactiveNotice(eventId: string, participantEmail: string): Promise<void> {
    try {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_InactiveNotices')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_InactiveNoticesListItem' },
        'Title': `${eventId} | ${participantEmail}`.slice(0, 250),
        'EventId': eventId,
        'ParticipantEmail': participantEmail,
        'SentByEmail': (this.context.pageContext.user.email || '').toLowerCase(),
      });
    } catch (err) { console.warn('[DEX] recordInactiveNotice failed:', err); }
  }

  /**
   * v26.39: `DEX_PostEventMails` — PERSISTENTER Dedup-Marker für die Post-Event-
   * „Danke & Hinweis zur Aufbewahrung"-Mail an die Organizer. Vorher hing die
   * Entdopplung an der TRANSIENTEN `DEX_Emails`-Queue (nach Versand + Archivierung
   * wieder leer → `hasQueuedEmail` fand nichts → Mail wurde beim nächsten
   * App-Öffnen ERNEUT verschickt) + Browser-`localStorage` (pro Gerät/User). Ein
   * dauerhafter Marker pro EventId behebt das geräte- und zeitübergreifend.
   * Rückgabe: true = Liste wurde JETZT NEU angelegt (Erststart → Altbestand seeden).
   */
  public async ensurePostEventMailsList(): Promise<boolean> {
    const listName = 'DEX_PostEventMails';
    try { if (await this.listExists(listName)) return false; } catch { return false; }
    const cr = await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Dedup (v26.39): „Danke & Hinweis zur Aufbewahrung"-Mail wurde für dieses Event bereits an die Organizer verschickt. Ein Marker pro Event — verhindert Mehrfachversand.',
      'BaseTemplate': 100, 'AllowContentTypes': false,
    });
    if (!cr.ok) return false;
    for (const f of [{ title: 'EventId', type: 2 }, { title: 'SentByEmail', type: 2 }]) {
      try { await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, { '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false }); } catch { /* */ }
    }
    try { await this.configureDefaultView(listName, ['EventId', 'SentByEmail', 'Created']); } catch { /* */ }
    try { await this.setQueueListPermissions(listName); } catch { /* */ }
    return true;
  }

  /** v26.39: Alle EventIds, für die die Post-Event-Mail schon raus ist. */
  public async getPostEventMailSentEventIds(): Promise<Set<string>> {
    const out = new Set<string>();
    try {
      if (!(await this.listExists('DEX_PostEventMails'))) return out;
      let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_PostEventMails')/items?$select=EventId&$top=5000`;
      while (url) {
        const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!resp.ok) break;
        const data = await resp.json();
        for (const it of (data.value || data.d?.results || [])) { const e = String(it.EventId || '').trim(); if (e) out.add(e); }
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      }
    } catch { /* */ }
    return out;
  }

  /** v26.39: Marker setzen — Post-Event-Mail für dieses Event ist erledigt. */
  public async recordPostEventMail(eventId: string | number): Promise<void> {
    try {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_PostEventMails')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_PostEventMailsListItem' },
        'Title': `${eventId}`.slice(0, 250),
        'EventId': String(eventId),
        'SentByEmail': (this.context.pageContext.user.email || '').toLowerCase(),
      });
    } catch (err) { console.warn('[DEX] recordPostEventMail failed:', err); }
  }

  // =====================================================================
  // v26.41: DEX_EventComms — dauerhaftes Kommunikations-Log der EVENT-RUNDMAILS
  // (Einladung, Massenmail, Ankündigungen). Organizer sehen die Historie im
  // Organizer Center; Teilnehmer die Rundmails unter „Meine Events" (Nachrücker/
  // Spätanmelder können nachlesen). KEINE persönlichen Bestätigungsmails.
  // Lesbar für alle Site-Nutzer (Rundmails sind nicht vertraulich); geschrieben
  // nur bei Organizer-Aktionen (queueEmail-Begleitung).
  // =====================================================================
  public async ensureEventCommsList(): Promise<void> {
    const listName = 'DEX_EventComms';
    try { if (await this.listExists(listName)) return; } catch { return; }
    const cr = await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Kommunikations-Log (v26.41): Event-Rundmails (Einladung/Massenmail/Ankündigung) mit Zeitstempel — Organizer-Historie + Teilnehmer-Ansicht unter „Meine Events". Keine persönlichen Bestätigungsmails.',
      'BaseTemplate': 100, 'AllowContentTypes': false,
    });
    if (!cr.ok) { console.warn('[DEX] DEX_EventComms konnte nicht angelegt werden.'); return; }
    const fields: Array<{ title: string; type: number; note?: boolean }> = [
      { title: 'EventId', type: 2 }, { title: 'EventTitle', type: 2 },
      { title: 'Subject', type: 2 }, { title: 'BodyHtml', type: 3, note: true },
      { title: 'EmailType', type: 2 }, { title: 'SentByEmail', type: 2 }, { title: 'SentByName', type: 2 },
    ];
    for (const f of fields) {
      try {
        const payload: Record<string, unknown> = { '__metadata': { 'type': f.note ? 'SP.FieldMultiLineText' : 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false };
        if (f.note) { payload['RichText'] = false; payload['NumberOfLines'] = 12; }
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
    try { await this.configureDefaultView(listName, ['EventId', 'EventTitle', 'Subject', 'EmailType', 'SentByName', 'Created']); } catch { /* */ }
    // KEIN Inheritance-Break: alle Site-Nutzer dürfen die Rundmails lesen.
  }

  /** Eine gesendete Event-Rundmail ins Log schreiben. */
  public async logEventComm(meta: { eventId: string | number; eventTitle: string; subject: string; bodyHtml: string; emailType: string }): Promise<void> {
    try {
      await this.ensureEventCommsList();
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventCommsListItem' },
        'Title': `${meta.eventId}: ${(meta.subject || '').slice(0, 180)}`.slice(0, 250),
        'EventId': String(meta.eventId), 'EventTitle': meta.eventTitle || '',
        'Subject': meta.subject || '', 'BodyHtml': meta.bodyHtml || '',
        'EmailType': meta.emailType || '',
        'SentByEmail': (this.context.pageContext.user.email || '').toLowerCase(),
        'SentByName': this.context.pageContext.user.displayName || '',
      });
    } catch (err) { console.warn('[DEX] logEventComm failed:', err); }
  }

  /** Alle Rundmails eines Events (neueste zuerst). */
  public async getEventComms(eventId: string | number): Promise<EventCommRow[]> {
    const out: EventCommRow[] = [];
    try {
      if (!(await this.listExists('DEX_EventComms'))) return out;
      const esc = String(eventId).replace(/'/g, "''");
      const sel = 'Id,EventId,Subject,BodyHtml,EmailType,SentByName,SentByEmail,Created';
      let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items?$select=${sel}&$filter=EventId eq '${esc}'&$orderby=Created desc&$top=500`;
      while (url) {
        const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!resp.ok) break;
        const data = await resp.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const it of ((data.value || data.d?.results || []) as any[])) {
          out.push({ id: Number(it.Id), eventId: String(it.EventId || ''), subject: it.Subject || '', bodyHtml: it.BodyHtml || '', emailType: it.EmailType || '', sentByName: it.SentByName || '', sentByEmail: it.SentByEmail || '', created: it.Created || '' });
        }
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      }
    } catch (err) { console.warn('[DEX] getEventComms failed:', err); }
    return out;
  }

  /** v26.69: Eine Log-Zeile aus dem Kommunikations-Log löschen — z. B. wenn ein
   *  Eintrag versehentlich protokolliert wurde und den „Bereits versendete Infos"-
   *  Hinweis fälschlich auslöst. Gibt true bei Erfolg zurück. */
  public async deleteEventComm(id: number): Promise<boolean> {
    try {
      if (!(await this.listExists('DEX_EventComms'))) return false;
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items(${Number(id)})`;
      const resp = await this._delete(url);
      return resp.ok;
    } catch (err) { console.warn('[DEX] deleteEventComm failed:', err); return false; }
  }

  /** Gibt es überhaupt Rundmails zu diesem Event? (für den Anmeldemail-Hinweis) */
  /**
   * @param excludeTypes v29.11: Rundmail-Arten, die NICHT zählen sollen.
   *
   * Hintergrund: Der Hinweis „Bereits versendete Infos zu diesem Event“ in der
   * Anmeldebestätigung soll die Person auf Kommunikation aufmerksam machen, die
   * sie verpasst haben könnte. Die Einladung ist genau das nicht — über sie ist
   * die Person meist überhaupt erst gekommen, und wer sich ohne Einladung
   * anmeldet, kann sie in der App ohnehin nachlesen. Stand nur eine Einladung
   * im Log, verwies der Hinweis also auf die Mail, die man gerade in der Hand
   * hatte. Lohnend ist er erst, wenn es DARÜBER HINAUS etwas gab.
   *
   * Die Auswertung läuft bewusst im Code und nicht als OData-Filter: `ne` würde
   * Zeilen mit leerem EmailType je nach Auslegung verschlucken. Alt-Zeilen ohne
   * Art zählen hier als „weitere Mail“ — im Zweifel lieber hinweisen als eine
   * echte Ankündigung verschweigen.
   */
  public async hasEventComms(
    eventId: string | number,
    excludeTypes?: string[],
  ): Promise<boolean> {
    try {
      if (!(await this.listExists('DEX_EventComms'))) return false;
      const esc = String(eventId).replace(/'/g, "''");
      const resp = await this._sp.get(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items?$select=Id,EmailType&$filter=EventId eq '${esc}'&$top=200`, SPHttpClient.configurations.v1);
      if (!resp.ok) return false;
      const data = await resp.json();
      const items = (data.value || data.d?.results || []) as Array<{ EmailType?: string }>;
      if (!Array.isArray(items) || items.length === 0) return false;
      if (!excludeTypes || excludeTypes.length === 0) return true;
      const skip = excludeTypes.map(t => (t || '').trim().toLowerCase());
      return items.some(it => skip.indexOf((it.EmailType || '').trim().toLowerCase()) < 0);
    } catch { return false; }
  }

  public async ensureOutlookLocksList(): Promise<void> {
    const listName = 'DEX_OutlookLocks';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Pro-Event-Sperre für den Outlook-Einladungs-Flow (v18.48) — verhindert gleichzeitige Läufe für dasselbe Event.',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // EventId: der Lock-Schlüssel. Muss eindeutig + indiziert sein, damit das
    // Create-als-Lock-Erwerb-Muster funktioniert (zweiter gleichzeitiger
    // Create für dieselbe EventId schlägt fehl -> der Lauf wartet & retryt).
    await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': 'EventId',
      'FieldTypeKind': 2,
      'Required': false,
    });
    // LockedAt: rein informativ (Debugging hängengebliebener Locks).
    await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': 'LockedAt',
      'FieldTypeKind': 4,
      'Required': false,
    });

    // EventId indizieren und Eindeutigkeit erzwingen. Reihenfolge wichtig:
    // erst Indexed, dann EnforceUniqueValues (SP verlangt eine indizierte
    // Spalte für die Eindeutigkeits-Prüfung). Auf einer frischen, leeren
    // Liste ist das unkritisch (keine Duplikate vorhanden).
    const fieldUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('EventId')`;
    try {
      await this._merge(fieldUrl, { 'Indexed': true });
      await this._merge(fieldUrl, { 'EnforceUniqueValues': true });
    } catch (e) {
      console.warn('[DEX] ensureOutlookLocksList: EnforceUniqueValues konnte nicht gesetzt werden:', e);
    }

    await this.configureDefaultView(listName, ['EventId', 'LockedAt']);

    // Schreibrechte wie bei den anderen Queue-Listen (DEX_Emails etc.) —
    // der Flow-Connection-Account muss Lock-Items anlegen/löschen können.
    try {
      await this.setQueueListPermissions(listName);
    } catch { /* best-effort */ }
  }

  /**
   * v11.83: Neue Team-Beitritts-Anfrage anlegen.
   */
  public async createTeamJoinRequest(args: {
    eventId: string;
    eventTitle: string;
    teamId: string;
    requesterEmail: string;
    requesterDisplayName: string;
    // v18.73: event-spezifische Antworten als JSON (optional).
    customData?: string;
  }): Promise<{ ok: boolean; itemId?: number }> {
    try {
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_TeamJoinRequestsListItem' },
        'Title': `${args.requesterDisplayName} -> ${args.eventTitle}`.slice(0, 250),
        'EventId': args.eventId,
        'TeamId': args.teamId,
        'RequesterEmail': args.requesterEmail,
        'RequesterDisplayName': args.requesterDisplayName,
        'Status': 'Pending',
        ...(args.customData ? { 'CustomData': args.customData } : {}),
      };
      const resp = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items`,
        payload
      );
      if (!resp.ok) return { ok: false };
      try {
        const j = await resp.json();
        const id: number = j?.d?.Id || j?.Id || 0;
        return { ok: true, itemId: id };
      } catch {
        return { ok: true };
      }
    } catch {
      return { ok: false };
    }
  }

  /**
   * v11.83: Alle Pending-Beitritts-Anfragen — optional gefiltert nach
   * Event und/oder Team. Wird für die "Beitritts-Anfragen"-Box im
   * Team-Lead-UI in MyEventsPage aufgerufen.
   */
  public async listTeamJoinRequests(args: {
    eventId?: string;
    teamId?: string;
    status?: 'Pending' | 'Approved' | 'Rejected';
  }): Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string; DecidedDate?: string; DecidedByEmail?: string; CustomData?: string }>> {
    try {
      const clauses: string[] = [];
      if (args.eventId) clauses.push(`EventId eq '${args.eventId.replace(/'/g, "''")}'`);
      if (args.teamId) clauses.push(`TeamId eq '${args.teamId.replace(/'/g, "''")}'`);
      clauses.push(`Status eq '${args.status || 'Pending'}'`);
      const filter = clauses.join(' and ');
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items?$filter=${encodeURIComponent(filter)}&$top=200&$orderby=Created asc`;
      const resp = await this._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return [];
      const data = await resp.json();
      return data.value || data.d?.results || [];
    } catch {
      return [];
    }
  }

  /**
   * v11.83: Approve/Reject einer Beitritts-Anfrage — schreibt Status,
   * DecidedDate und DecidedByEmail. Die Folge-Logik (Member-Insert,
   * Mails) liegt im EventContext, weil dort die Subsite-/Event-Lookups
   * verfügbar sind.
   */
  public async decideTeamJoinRequest(
    requestId: number,
    decision: 'Approved' | 'Rejected',
    decidedByEmail: string
  ): Promise<boolean> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items(${requestId})`;
      const body = {
        'Status': decision,
        'DecidedDate': new Date().toISOString(),
        'DecidedByEmail': decidedByEmail || '',
      };
      const resp = await this._merge(url, body);
      return !!resp.ok;
    } catch {
      return false;
    }
  }

  /**
   * Bestehende abgemeldete Registrierung reaktivieren.
   * Setzt Status zurück auf Angemeldet/Warteliste, löscht CancellationDate,
   * aktualisiert RegistrationDate und CustomData.
   */
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
    proxyConsent?: string // v18.74: Zustimmungs-Nachweis bei stellvertretender Re-Anmeldung
  ): Promise<boolean> {
    try {
      // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
      // Lade die ParticipantEmail aus dem zu reaktivierenden Item und prüfe,
      // ob der aktuelle User dafür berechtigt ist. Plus Deadline-Check.
      // v7.30: Wir laden hier zusätzlich die existierende TeilnehmerID, damit
      // beim Reaktivieren die alte ID erhalten bleibt — Counter wird NUR
      // dann angefasst, wenn die alte ID null/0 ist (Legacy-Edge).
      let existingTeilnehmerId = 0;
      // v20.5: Teilnehmer-E-Mail in den aeusseren Scope heben, damit der
      // Autor-Set am Ende (stellvertretende Re-Anmeldung) sie nutzen kann.
      let reactivateParticipantEmail = '';
      try {
        const itemResp = await this._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${itemId})?$select=ParticipantEmail,TeilnehmerID`,
          SPHttpClient.configurations.v1
        );
        const sessionEmail = (this.context.pageContext.user.email || '').toLowerCase();
        let targetEmail = '';
        if (itemResp.ok) {
          const itemData = await itemResp.json();
          targetEmail = (itemData.ParticipantEmail || itemData.d?.ParticipantEmail || '').toLowerCase();
          reactivateParticipantEmail = targetEmail;
          const tnId = itemData.TeilnehmerID ?? itemData.d?.TeilnehmerID;
          if (typeof tnId === 'number' && tnId > 0) existingTeilnehmerId = tnId;
        }

        // Check A: für andere Person registrieren?
        if (targetEmail && targetEmail !== sessionEmail) {
          const allowed = await this.canRegisterForOthers(subsiteUrl, targetEmail);
          if (!allowed) {
            console.warn(`[DEX] reactivateRegistration DENIED: ${sessionEmail} versuchte ${targetEmail} zu reaktivieren — nicht berechtigt.`);
            return false;
          }
        }

        // Check B: Deadline-Check (Event über SubsiteUrl finden)
        const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
        const evResp = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail`,
          SPHttpClient.configurations.v1
        );
        if (evResp.ok) {
          const evData = await evResp.json();
          const items = evData.value || evData.d?.results || [];
          if (items.length > 0) {
            const deadline = items[0].RegistrationDeadline || '';
            const orgStr: string = EventService.stripNoteWrapper(items[0].OrganizerEmail);
            const orgEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
            if (deadline) {
              const deadlineDate = new Date(deadline);
              if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
                const isEventOrganizer = orgEmails.indexOf(sessionEmail) >= 0;
                let isAdmin = false;
                try {
                  const esc = sessionEmail.replace(/'/g, "''");
                  const roleResp = await this._sp.get(
                    `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
                    SPHttpClient.configurations.v1
                  );
                  if (roleResp.ok) {
                    const rd = await roleResp.json();
                    const rItems = rd.value || rd.d?.results || [];
                    if (rItems.length > 0 && (rItems[0].Role === 'Admin' || rItems[0].Role === 'IT-Admin')) isAdmin = true;
                  }
                } catch { /* ignore */ }
                if (!isEventOrganizer && !isAdmin) {
                  console.warn(`[DEX] reactivateRegistration DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${deadline} zu reaktivieren.`);
                  return false;
                }
              }
            }
          }
        }
      } catch { /* bei Load-Fehler konservativ: weitermachen */ }
      // ---- Ende Permission-Checks ----

      // Reaktivierung = funktional eine Neuanmeldung mit existierendem Listen-
      // Item. Deshalb wird hier — analog zu registerForEvent — atomar eine
      // neue TeilnehmerID am Counter gezogen. Wer mal #12 war und reaktiviert,
      // bekommt jetzt z.B. die #87, also die nächst-freie ID am Ende der
      // Liste — exakt wie ein Neuzugang. Ohne diesen Schritt blieb der
      // Eintrag mit TeilnehmerID=null hängen, weil im Reaktivierungs-Pfad
      // niemand den DEX_IDReorder-Flow triggert.
      // v9.10: race-anfälliger max+1-Fallback entfernt (siehe Kommentar in
      // registerForEvent). Bei Counter-Outage bleibt TeilnehmerID undefined.
      void existingTeilnehmerId;
      const nextId = await this.getNextTeilnehmerId(subsiteUrl);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'Vorname': firstName,
        'Nachname': surname,
        'ParticipantName': `${firstName} ${surname}`,
        'Status': status,
        // v9.10: TeilnehmerID nur setzen wenn Counter sie geliefert hat.
        ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
        'RegistrationDate': new Date().toISOString(),
        'CancellationDate': null,
        'CustomData': JSON.stringify(customData),
      };

      // Audit: wer hat die Re-Anmeldung ausgelöst? (überschreibt den Wert von
      // der ursprünglichen Anmeldung, weil das faktisch eine neue Anmeldung ist)
      const auditName = registeredByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (registeredByEmail || this.context.pageContext.user.email || '').toLowerCase();
      if (auditName) body['RegisteredByName'] = auditName;
      if (auditEmail) body['RegisteredByEmail'] = auditEmail;
      // v18.74: Zustimmungs-Nachweis bei stellvertretender Re-Anmeldung.
      if (proxyConsent) body['ProxyConsent'] = proxyConsent;

      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            body[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
          }
        }
      }

      let response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      // v19.9: Wie bei registerForEvent — falls die ProxyConsent-Spalte auf
      // diesem (älteren) Event noch fehlt, scheitert der MERGE mit HTTP 400.
      // Einmal ohne das optionale Audit-Feld wiederholen, damit die
      // Re-Aktivierung nicht an einer fehlenden Spalte scheitert.
      if (!response.ok && body['ProxyConsent']) {
        console.warn('[DEX] reactivateRegistration: MERGE fehlgeschlagen — Retry OHNE ProxyConsent (Spalte evtl. nicht vorhanden). Bitte im Admin Center "Spalten fixen" ausführen.');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const retryBody: Record<string, any> = { ...body };
        delete retryBody['ProxyConsent'];
        response = await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
          retryBody
        );
      }
      if (!response.ok) return false;

      // v9.10: Post-Update Safety Net (siehe registerForEvent). Bei
      // Counter-Edge-Cases könnte der nächste Wert kollidieren — der
      // aeltere Eintrag (kleinere SP-Id) gewinnt, der spätere bekommt
      // fresh ID.
      if (typeof nextId === 'number' && nextId > 0) {
        try {
          const dupResp = await this._sp.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,TeilnehmerID&$filter=TeilnehmerID eq ${nextId}&$top=10`,
            SPHttpClient.configurations.v1
          );
          if (dupResp.ok) {
            const dupData = await dupResp.json();
            const dupItems: Array<{ Id: number; TeilnehmerID: number }> = dupData.value || dupData.d?.results || [];
            if (dupItems.length > 1) {
              const minId = Math.min(...dupItems.map(d => d.Id));
              if (itemId !== minId) {
                const fresh = await this.getNextTeilnehmerId(subsiteUrl);
                if (typeof fresh === 'number' && fresh > 0) {
                  await this._merge(
                    `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
                    { 'TeilnehmerID': fresh }
                  );
                  console.warn(`[DEX] Post-update dedup (reactivate): TeilnehmerID ${nextId} kollidierte, Item ${itemId} hat jetzt #${fresh}.`);
                }
              }
            }
          }
        } catch (err) {
          console.warn('[DEX] Post-update dedup check (reactivate) fehlgeschlagen:', err);
        }
      }

      // v20.5: Stellvertretende Re-Anmeldung → Teilnehmer zum Autor der Zeile
      // machen (analog registerForEvent), damit er sie in "Meine Events" sieht
      // und sich selbst abmelden kann. Best-effort (nur mit "Listen verwalten").
      if (reactivateParticipantEmail && auditEmail && reactivateParticipantEmail !== auditEmail) {
        await this.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, itemId, reactivateParticipantEmail);
      }

      return true;
    } catch {
      return false;
    }
  }

  /**
   * Schlanker MERGE-Helper auf ein einzelnes Teilnehmerlisten-Item — baut
   * keine ChangeLog-Logik, keine FieldMap-Auflösung, keine Default-Felder ein.
   * Genutzt für One-Shot-Migrationen (z.B. T-Shirt-Größen-Import), die direkt
   * bestimmte Felder (inkl. CustomData-JSON + einzelne SP-Spalten) setzen wollen.
   */
  public async mergeRegistrationFields(
    subsiteUrl: string,
    itemId: number,
    body: Record<string, unknown>
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Custom Data einer Registrierung aktualisieren (Teilnehmer ändert eigene Angaben).
   */
  /**
   * v17.2: Bestehende Teilnehmer-Registrierung einem Team zuordnen
   * (PATCH der TeamId/TeamName/TeamLead-Felder auf einem schon existierenden
   * Item). Wird vom Admin-Center-Team-Management genutzt, wenn der
   * Organizer einen schon Angemeldeten ohne Team einem (neuen) Team
   * zuweist — vermeidet doppelte Anmeldung + Mail/Outlook-Spam.
   */
  public async assignRegistrationToTeam(
    subsiteUrl: string,
    itemId: number,
    teamId: string,
    teamName: string | undefined,
    isLead: boolean,
  ): Promise<boolean> {
    try {
      const body: Record<string, unknown> = {
        TeamId: teamId,
        TeamName: teamName || '',
        TeamLead: !!isLead,
      };
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] assignRegistrationToTeam failed:', err);
      return false;
    }
  }

  public async updateRegistrationData(
    subsiteUrl: string,
    itemId: number,
    customData: Record<string, string>,
    customFieldMap?: Record<string, string>,
    oldCustomData?: Record<string, string>,
    fieldLabelMap?: Record<string, string> // cf.id -> label
  ): Promise<boolean> {
    try {
      // Änderungen ermitteln
      const changes: string[] = [];
      const now = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      if (oldCustomData && fieldLabelMap) {
        for (const key of Object.keys(customData)) {
          if (key === 'salutation') continue;
          const label = fieldLabelMap[key] || key;
          const oldVal = oldCustomData[key] || '';
          const newVal = customData[key] || '';
          if (oldVal !== newVal) {
            changes.push(`${label}: "${oldVal}" → "${newVal}"`);
          }
        }
      }
      const changeEntry = changes.length > 0 ? `[${now}] ${changes.join(', ')}` : '';

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'CustomData': JSON.stringify(customData),
        'LastModifiedDate': new Date().toISOString(),
      };

      // ChangeLog anhängen (bestehenden Log behalten)
      if (changeEntry) {
        try {
          const existing = await this._sp.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
            SPHttpClient.configurations.v1
          );
          if (existing.ok) {
            const data = await existing.json();
            const oldLog = data.ChangeLog || '';
            body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
          }
        } catch {
          body['ChangeLog'] = changeEntry;
        }
      }

      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            body[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
          }
        }
      }

      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * v8.0: Admin/Organizer kann Teilnehmerdaten direkt aus dem Admin Center
   * editieren. Erlaubt das Ändern von Anrede/Vorname/Nachname/Email/Phone/
   * Department/Location/JobTitle/Status sowie aller Custom-Felder. Schreibt
   * automatisch ChangeLog-Eintrag mit Wer/Wann/Was-Diff und setzt
   * LastModifiedDate.
   *
   * patch: Nur die echten Spalten-Werte (keine __metadata nötig — _merge
   * sendet odata=nometadata).
   * actor: Audit-Info des aufrufenden Users.
   * oldValues: zum Diff-Bauen, nur Felder mit oldValues[key] !== patch[key]
   * landen im ChangeLog.
   * fieldLabelMap: optional, mappt internal column name -> display label.
   */
  public async adminUpdateRegistration(
    subsiteUrl: string,
    itemId: number,
    patch: Record<string, unknown>,
    actor: { name: string; email: string },
    oldValues?: Record<string, unknown>,
    fieldLabelMap?: Record<string, string>
  ): Promise<boolean> {
    try {
      const changes: string[] = [];
      const now = new Date().toLocaleString('de-DE', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit',
      });
      if (oldValues) {
        for (const key of Object.keys(patch)) {
          const oldV = oldValues[key];
          const newV = patch[key];
          // Vergleich als String, damit number vs string nicht stört
          const oldStr = oldV === null || oldV === undefined ? '' : String(oldV);
          const newStr = newV === null || newV === undefined ? '' : String(newV);
          if (oldStr !== newStr) {
            const label = (fieldLabelMap && fieldLabelMap[key]) || key;
            changes.push(`${label}: "${oldStr}" → "${newStr}"`);
          }
        }
      }
      const changeEntry = changes.length > 0
        ? `[${now}] ${actor.name || actor.email}: ${changes.join(', ')}`
        : '';

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = { ...patch, 'LastModifiedDate': new Date().toISOString() };

      // ChangeLog anhängen (bestehenden Log behalten, neuestes oben)
      if (changeEntry) {
        try {
          const existing = await this._sp.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
            SPHttpClient.configurations.v1
          );
          if (existing.ok) {
            const data = await existing.json();
            const oldLog = data.ChangeLog || data.d?.ChangeLog || '';
            body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
          } else {
            body['ChangeLog'] = changeEntry;
          }
        } catch {
          body['ChangeLog'] = changeEntry;
        }
      }

      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch (err) {
      console.warn('[DEX] adminUpdateRegistration error:', err);
      return false;
    }
  }

  /**
   * Eigene Registrierung für ein Event laden
   */
  public async getMyRegistration(
    subsiteUrl: string,
    email: string
  ): Promise<SPRegistration | null> {
    try {
      // v27.11: $orderby=Id desc — bei mehreren Zeilen derselben Person
      // (Alt-Duplikate) IMMER die neueste nehmen. Vorher konnte $top=1 ohne
      // Sortierung eine alte 'Abgemeldet'-Zeile erwischen und den
      // Reaktivierungs-Pfad statt des Duplikat-Blocks auslösen.
      const response = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.trim().replace(/'/g, "''")}'&$orderby=Id desc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      if (data.value && data.value.length > 0) return data.value[0];
      return null;
    } catch {
      return null;
    }
  }

  /**
   * v24.36: Alle Anmeldungen einer Teilnehmerliste, die der eingeloggte User
   * STELLVERTRETEND für eine andere Person durchgeführt hat (Assistenz-Funktion).
   * Filter: `RegisteredByEmail = Akteur` UND `ParticipantEmail ≠ Akteur`.
   *
   * Hinweis zur Item-Level-Security: Auf den Teilnehmerlisten greift
   * ReadSecurity=2 („nur eigene Elemente"). Eine Assistenz (Contribute-User)
   * sieht daher nur die Zeilen, deren SharePoint-Autor sie selbst ist. Nach dem
   * v20.5-Autor-Wechsel (AuthorId → Teilnehmer) verliert die Assistenz den
   * Lesezugriff auf genau diese Zeilen — solange der `DEX_AccessFix`-Autor-Flow
   * NICHT eingerichtet ist (häufigster Fall), bleibt die Assistenz Autor und
   * sieht ihre Fremd-Anmeldungen weiterhin. Admin/Organizer eigener Events
   * sehen ohnehin alle Zeilen. Best-effort: liefert nur, was der Server
   * tatsächlich zurückgibt.
   */
  public async getProxyRegistrationsByActor(
    subsiteUrl: string,
    actorEmail: string
  ): Promise<SPRegistration[]> {
    const me = (actorEmail || '').toLowerCase().trim();
    if (!me) return [];
    const out: SPRegistration[] = [];
    const esc = me.replace(/'/g, "''");
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=RegisteredByEmail eq '${esc}'&$orderby=Id asc&$top=5000`;
    while (url) {
      try {
        const response = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        const page: SPRegistration[] = data.value || data.d?.results || [];
        for (const r of page) {
          const pe = (r.ParticipantEmail || '').toLowerCase().trim();
          if (pe && pe !== me) out.push(r);
        }
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    return out;
  }

  /** v26.47: Externe stellvertretende Anmeldung als „Datenschutz-Rückmeldung
   *  offen" markieren (ConsentReview='Pending'). Lookup per E-Mail auf der
   *  jüngsten aktiven Zeile — läuft direkt nach der Registrierung. */
  public async markConsentPendingByEmail(subsiteUrl: string, participantEmail: string): Promise<boolean> {
    try {
      const emailLc = (participantEmail || '').trim().toLowerCase();
      if (!emailLc) return false;
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,Status&$orderby=Id desc&$top=200`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return false;
      const data = await resp.json();
      const items = (data.value || data.d?.results || []) as Array<{ Id: number; ParticipantEmail?: string; Status?: string }>;
      const hit = items.find(it => (it.ParticipantEmail || '').trim().toLowerCase() === emailLc && it.Status !== 'Abgemeldet');
      if (!hit) return false;
      const m = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${hit.Id})`,
        { 'ConsentReview': 'Pending' }
      );
      return m.ok;
    } catch (err) { console.warn('[DEX] markConsentPendingByEmail failed:', err); return false; }
  }

  /** v26.73: Fertigen .eml-Einladungs-Entwurf als Attachment an der Teilnehmer-
   *  Zeile ablegen (Zeile per E-Mail gefunden). Rückgabe = Item-Id (0 =
   *  fehlgeschlagen). Der Deeplink in der externen Instruktions-Mail holt genau
   *  diese Datei wieder — der Anhang darf per Deloitte-Mail-Regel nicht direkt
   *  mitgeschickt werden. Fester Dateiname `dxinvite--Einladung.eml`. */
  public async storeInviteEmlByEmail(subsiteUrl: string, participantEmail: string, emlContent: string): Promise<number> {
    try {
      const emailLc = (participantEmail || '').trim().toLowerCase();
      if (!emailLc || !emlContent) return 0;
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,Status&$orderby=Id desc&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return 0;
      const data = await resp.json();
      const items = (data.value || data.d?.results || []) as Array<{ Id: number; ParticipantEmail?: string; Status?: string }>;
      const hit = items.find(it => (it.ParticipantEmail || '').trim().toLowerCase() === emailLc && it.Status !== 'Abgemeldet');
      if (!hit) return 0;
      const fileName = 'dxinvite--Einladung.eml';
      // Vorherige Version best-effort löschen (sonst 409 bei erneuter Anmeldung).
      try {
        await this._sp.post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${hit.Id})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`,
          SPHttpClient.configurations.v1,
          { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'Accept': 'application/json;odata=nometadata' } }
        );
      } catch { /* gab es noch nicht */ }
      const buf = new TextEncoder().encode(emlContent);
      const add = await this._sp.post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${hit.Id})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' }, body: buf.buffer as ArrayBuffer }
      );
      return add.ok ? hit.Id : 0;
    } catch (err) { console.warn('[DEX] storeInviteEmlByEmail failed:', err); return 0; }
  }

  /** v26.73: Den an der Teilnehmer-Zeile abgelegten .eml-Entwurf (per Item-Id)
   *  wieder auslesen — für den Download-Deeplink. */
  public async getInviteEmlByItem(subsiteUrl: string, itemId: number): Promise<{ fileName: string; content: string } | null> {
    try {
      if (!subsiteUrl || !itemId) return null;
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${Number(itemId)})/AttachmentFiles/getByFileName('${encodeURIComponent('dxinvite--Einladung.eml')}')/$value`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (!resp.ok) return null;
      const content = await resp.text();
      if (!content) return null;
      return { fileName: 'Einladung.eml', content };
    } catch (err) { console.warn('[DEX] getInviteEmlByItem failed:', err); return null; }
  }

  /** v26.47: Datenschutz-Rückmeldung der externen Person bestätigen —
   *  ConsentReview zurücksetzen (Button in der Teilnehmerliste). */
  public async confirmConsentReview(subsiteUrl: string, itemId: number, meta?: { eventId?: string; eventTitle?: string; participantName?: string }): Promise<boolean> {
    try {
      const m = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'ConsentReview': '' }
      );
      if (m.ok && meta?.eventId) {
        try {
          await this.writeChangeLog({
            action: 'ExternalConsentConfirmed', targetType: 'Participant', targetId: String(itemId),
            targetName: meta.participantName || '', eventId: meta.eventId, eventTitle: meta.eventTitle || '',
            details: { note: 'Datenschutz-Rückmeldung der externen Person bestätigt (v26.47).' },
          });
        } catch { /* Audit best-effort */ }
      }
      return m.ok;
    } catch (err) { console.warn('[DEX] confirmConsentReview failed:', err); return false; }
  }

  /**
   * v30.48: Spalte `Startnummer` in der Teilnehmerliste sicherstellen.
   *
   * Eigene Spalte statt Custom-Field: Die Startnummer soll im Export, in der
   * Teilnehmerliste und in der Check-in-Suche auftauchen — ein Custom-Field
   * wäre pro Event konfigurierbar, aber nicht durchsuchbar. Text und nicht
   * Zahl, weil Veranstalter-Nummern führende Nullen und Präfixe haben können;
   * gerechnet wird damit ohnehin nie.
   *
   * Idempotent: Existiert die Spalte, passiert nichts. Wirft nicht — der
   * Import meldet den Fehlschlag selbst, statt hier abzubrechen.
   */
  public async ensureStartNumberColumn(subsiteUrl: string): Promise<boolean> {
    try {
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=InternalName eq 'Startnummer'&$select=InternalName`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const d = await resp.json();
        const arr = d.value || d.d?.results || [];
        if (arr.length > 0) return true;
      }
      await this._post(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': 'Startnummer',
        'FieldTypeKind': 2, // Text
        'Required': false,
      });
      return true;
    } catch (e) {
      console.warn('[DEX] ensureStartNumberColumn failed:', e);
      return false;
    }
  }

  /**
   * Alle Registrierungen für ein Event laden (nur für Organizer/Admin)
   */
  /**
   * @param onHttpError v29.3: Meldet, dass der Lesevorgang abgebrochen wurde,
   *   statt ihn — wie bisher — still zu verschlucken. Ohne diesen Rückruf ist
   *   „Liste existiert nicht" (404, z.B. nach dem 3-Monats-Löschkonzept) von
   *   „Liste ist leer" nicht zu unterscheiden: Beides kam als `[]` zurück.
   *   Genau daran hat sich `analyzeRegistryAgainstLists` verschluckt.
   *   `status` ist der HTTP-Status, `0` bei einem Netz-/Parse-Fehler.
   */
  public async getAllRegistrations(subsiteUrl: string, onHttpError?: (_status: number) => void): Promise<SPRegistration[]> {
    const allItems: SPRegistration[] = [];
    // $top=5000 ist das SP-REST-Maximum pro Page. Damit fallen Events bis zu
    // 5000 Teilnehmern in einen einzigen Response — keine Pagination-Edgecases
    // mit fehlendem nextLink. Bei größeren Listen folgen wir dem nextLink
    // weiter (Schleife unten). Vorher stand hier $top=500, was bei Events mit
    // ≥500 Teilnehmern zu fehlenden Einträgen führte: SharePoint liefert
    // bei $orderby+$top in Kombination mit Item-Level-Security nicht
    // zuverlässig nextLink, wenn die erste Page exakt voll ist.
    // v27.12: $select=*,Author/… + $expand=Author — der Zeilen-Autor dient als
    // Fallback für „Registriert von", wenn die Zeile nicht über die App
    // angelegt wurde (RegisteredBy* leer). '*' behält alle Skalar-Felder,
    // Verhalten ist sonst identisch zum bisherigen Query ohne $select.
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=*,Author/Title,Author/EMail&$expand=Author&$orderby=Id asc&$top=5000`;

    while (url) {
      try {
        const response = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) { if (onHttpError) onHttpError(response.status); break; }
        const data = await response.json();
        // Beide OData-Formate abdecken: nometadata (data.value) UND verbose
        // (data.d.results). Vorher nur data.value — bei verbose-Response
        // wären null Items dazugekommen.
        const page = data.value || data.d?.results || [];
        allItems.push(...page);
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        if (onHttpError) onHttpError(0);
        break;
      }
    }
    return allItems;
  }

  /**
   * TeilnehmerIDs sequentiell neu vergeben (1, 2, 3, ...).
   *
   * Seit v6.5: Zwei-Pass-Reorder.
   * 1. Erst alle **Angemeldeten** (Status ∈ Angemeldet / QR versendet / Eingecheckt)
   *    in Reihenfolge der Registrierung (SP-ItemId asc) → bekommen IDs 1..N.
   * 2. Danach alle **Warteliste**-Teilnehmer in Reihenfolge der Registrierung
   *    → bekommen IDs N+1..N+M (Warteliste hängt lückenlos hinten an).
   * 3. Abgemeldete bekommen TeilnehmerID = null.
   *
   * Damit stehen im Teilnehmerlisten-Grid die Angemeldeten sauber oben, die
   * Warteliste sauber unten — kein Durchmischen mehr.
   */
  /**
   * v28.70: Eine Person auf der Warteliste an eine bestimmte Position setzen
   * (z.B. „ganz nach vorn"), ohne die Reihenfolge der Angemeldeten anzufassen.
   *
   * Warum über die TeilnehmerID? Die Warteliste hat KEINE eigene
   * Positionsspalte — die Position ist der Rang innerhalb von
   * Status='Warteliste', sortiert nach `TeilnehmerID asc` (Gleichstand: Item-Id).
   * Genau so sortieren beide Stellen, die nachrücken: `promoteFirstWaitlistItem`
   * in der App UND der Flow `DEX_IDReorder_TeilnehmerIDs` (Order By
   * `TeilnehmerID asc`, s. docs/flow-jsons.md). Eine zusätzliche Prioritaets-
   * spalte würde der Flow ignorieren und weiter den Falschen nachrücken
   * lassen — deshalb wird hier die TeilnehmerID selbst umsortiert. Damit zieht
   * die neue Reihenfolge in beiden Pfaden.
   *
   * Die Angemeldeten behalten ihre IDs (ihre relative Reihenfolge ändert sich
   * nicht, sie werden auf dieselben 1..N abgebildet); es verschieben sich nur
   * die IDs innerhalb der Warteliste.
   *
   * @param targetPosition 1-basiert. Wird auf 1..(Anzahl Wartende) geklemmt.
   */
  public async setWaitlistPosition(
    subsiteUrl: string,
    itemId: number,
    targetPosition: number
  ): Promise<{ ok: boolean; from: number; to: number; changed: number; error?: string }> {
    const fail = (error: string): { ok: boolean; from: number; to: number; changed: number; error: string } =>
      ({ ok: false, from: 0, to: 0, changed: 0, error });
    try {
      const allItems: Array<{ Id: number; Status: string; TeilnehmerID: number | null }> = [];
      let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Status,TeilnehmerID&$orderby=Id asc&$top=5000`;
      while (url) {
        const response = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) return fail(`Teilnehmerliste konnte nicht gelesen werden (HTTP ${response.status}).`);
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      }
      const NO_TID = Number.MAX_SAFE_INTEGER;
      const byTidThenId = (a: { Id: number; TeilnehmerID: number | null }, b: { Id: number; TeilnehmerID: number | null }): number =>
        ((a.TeilnehmerID ?? NO_TID) - (b.TeilnehmerID ?? NO_TID)) || (a.Id - b.Id);
      const activeItems = allItems
        .filter(i => i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt')
        .sort(byTidThenId);
      const waitlist = allItems.filter(i => i.Status === 'Warteliste').sort(byTidThenId);

      const fromIdx = waitlist.findIndex(i => i.Id === itemId);
      if (fromIdx < 0) return fail('Diese Person steht nicht (mehr) auf der Warteliste — bitte die Liste neu laden.');
      if (waitlist.length < 2) return fail('Auf der Warteliste steht nur diese eine Person — es gibt nichts umzusortieren.');
      const toIdx = Math.max(0, Math.min(waitlist.length - 1, Math.round(targetPosition) - 1));
      if (toIdx === fromIdx) {
        return { ok: true, from: fromIdx + 1, to: toIdx + 1, changed: 0 };
      }
      const moved = waitlist.splice(fromIdx, 1)[0];
      waitlist.splice(toIdx, 0, moved);

      // Ziel-IDs: Angemeldete 1..N (unverändert), danach die neue
      // Warteliste-Reihenfolge, Abgemeldete null.
      const targetIds = new Map<number, number | null>();
      let nextId = 1;
      for (const item of activeItems) targetIds.set(item.Id, nextId++);
      for (const item of waitlist) targetIds.set(item.Id, nextId++);
      for (const item of allItems) if (!targetIds.has(item.Id)) targetIds.set(item.Id, null);

      let changed = 0;
      for (const item of allItems) {
        const newId = targetIds.get(item.Id) ?? null;
        if (newId === item.TeilnehmerID) continue;
        const resp = await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${item.Id})`,
          { 'TeilnehmerID': newId }
        );
        if (resp.ok || resp.status === 406) { changed++; }
        else { return fail(`Position konnte nicht vollständig gesetzt werden (HTTP ${resp.status}). Bitte „Teilnehmer-IDs neu vergeben" ausführen und erneut versuchen.`); }
      }
      try { await this.syncCounterToMax(subsiteUrl); } catch { /* best-effort */ }
      return { ok: true, from: fromIdx + 1, to: toIdx + 1, changed };
    } catch (err) {
      return fail(`Unerwarteter Fehler: ${err instanceof Error ? err.message.slice(0, 160) : 'unbekannt'}`);
    }
  }

  public async reorderParticipantIDs(
    subsiteUrl: string,
    onProgress?: (pct: number) => void
  ): Promise<{ success: number; errors: number }> {
    // Alle Items laden, sortiert nach SP Id (Erstellungsreihenfolge = Reihenfolge der Registrierung)
    const allItems: Array<{ Id: number; Status: string; TeilnehmerID: number | null }> = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Status,TeilnehmerID&$orderby=Id asc&$top=5000`;

    while (url) {
      try {
        const response = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch { break; }
    }

    // Ziel-IDs in einem ersten Durchlauf berechnen: erst Angemeldete, dann Warteliste.
    // v22.20: Innerhalb der Gruppen wird primär nach der VORHANDENEN TeilnehmerID
    // sortiert (aufsteigend, ohne ID ans Ende), Gleichstand nach Item-Id. Damit ist
    // die Client-Renummerierung deckungsgleich mit dem DEX_IDReorder-Flow (der
    // ebenfalls nach TeilnehmerID sortiert) — egal welcher Pfad zuletzt lief, das
    // Ergebnis ist identisch und die IDs „springen" nicht mehr zwischen
    // Admin-Button und Flow-Lauf hin und her.
    const targetIds = new Map<number, number | null>();
    let nextId = 1;
    const NO_TID = Number.MAX_SAFE_INTEGER;
    const byTidThenId = (a: { Id: number; TeilnehmerID: number | null }, b: { Id: number; TeilnehmerID: number | null }): number =>
      ((a.TeilnehmerID ?? NO_TID) - (b.TeilnehmerID ?? NO_TID)) || (a.Id - b.Id);
    // Pass 1: Angemeldete / QR versendet / Eingecheckt
    const activeItems = allItems
      .filter(item => item.Status === 'Angemeldet' || item.Status === 'QR versendet' || item.Status === 'Eingecheckt')
      .sort(byTidThenId);
    for (const item of activeItems) {
      targetIds.set(item.Id, nextId++);
    }
    // Pass 2: Warteliste
    const waitlistItems = allItems
      .filter(item => item.Status === 'Warteliste')
      .sort(byTidThenId);
    for (const item of waitlistItems) {
      targetIds.set(item.Id, nextId++);
    }
    // Pass 3: Abgemeldete (TeilnehmerID=null)
    for (const item of allItems) {
      if (!targetIds.has(item.Id)) {
        targetIds.set(item.Id, null);
      }
    }

    let success = 0;
    let errors = 0;
    const totalItems = allItems.length || 1;
    let processed = 0;
    if (onProgress) { try { onProgress(0); } catch { /* */ } }
    for (const item of allItems) {
      const newId = targetIds.get(item.Id) ?? null;
      if (newId === item.TeilnehmerID) {
        success++;
      } else {
        try {
          const resp = await this._merge(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${item.Id})`,
            { 'TeilnehmerID': newId }
          );
          if (resp.ok || resp.status === 406) { success++; } else { errors++; }
        } catch { errors++; }
      }
      processed++;
      if (onProgress) {
        // 0..95 % während der Merges; die letzten 5 % für syncCounterToMax.
        try { onProgress(Math.min(95, Math.round((processed / totalItems) * 95))); } catch { /* */ }
      }
    }

    // v7.31 / v9.14: Counter konsistent halten — syncCounterToMax patcht
    // Counter (monotonic up-only). ensureCounterList wurde hier ursprünglich
    // (v9.13) ebenfalls gerufen, hat aber Race-Conditions ausgelöst. Die
    // Counter-Liste sollte zum Zeitpunkt eines Reorders ohnehin existieren —
    // sonst hat die App ein anderes Problem das ein expliziter Klick auf
    // "Counter zurücksetzen" löst.
    try { await this.syncCounterToMax(subsiteUrl); } catch { /* best-effort */ }
    if (onProgress) { try { onProgress(100); } catch { /* */ } }

    return { success, errors };
  }

  // ==================== v11.36: Überbuchungs-Schutz + Bereinigung ====================

  private static readonly ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

  /**
   * Ungefilterte Gesamt-Item-Zahl einer Liste (List-Property `ItemCount`,
   * NICHT security-getrimmt). Dient als Vollständigkeits-Check: liefert eine
   * Item-Abfrage weniger Zeilen als `ItemCount`, beschneidet die
   * Element-Sicherheit („nur eigene Elemente") die Sicht des Aufrufers.
   * -1 bei Lesefehler.
   */
  private async getListItemCount(subsiteUrl: string, listName: string): Promise<number> {
    try {
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ItemCount`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return -1;
      const data = await resp.json();
      const raw = data?.ItemCount ?? data?.d?.ItemCount;
      const n = typeof raw === 'number' ? raw : parseInt(String(raw ?? ''), 10);
      return Number.isFinite(n) && n >= 0 ? n : -1;
    } catch { return -1; }
  }

  /**
   * Zählt die aktiven (= nicht Warteliste/Abgemeldet) Anmeldungen, gesamt und
   * pro Starter-Gruppe. Quelle ist die echte Teilnehmerliste — wird zum Seeden
   * und Reconcilen der Sitzplatz-Counter genutzt.
   *
   * v27.10 REGRESSIONS-FIX: Seit v26.87 greift die Element-Sicherheit („nur
   * eigene Elemente") auf den Teilnehmerlisten WIRKLICH — vorher schlug das
   * Setzen still mit HTTP 400 fehl. Für normale User liefert
   * getAllRegistrations seitdem nur noch die EIGENEN Zeilen. Eine darauf
   * basierende Zählung wäre katastrophal zu niedrig: real beobachtet hat ein
   * Self-Cancel über syncSeatsToActiveCount `SeatsTaken=0` geschrieben,
   * wonach Neu-Anmeldungen an der kompletten Warteliste vorbei direkt
   * „Angemeldet" wurden. Deshalb wird die gelesene Zeilenzahl gegen den
   * ungefilterten `ItemCount` der Liste geprüft — ist die Sicht unvollständig
   * (oder nicht verifizierbar), fliegt ein Fehler und die Aufrufer handeln
   * fail-safe (Sync schreibt nichts, reserveSeat nutzt den reinen Counter).
   */
  private async getActiveCounts(subsiteUrl: string): Promise<{ total: number; durch: number; fun: number; waitlist: number }> {
    const regs = await this.getAllRegistrations(subsiteUrl);
    const itemCount = await this.getListItemCount(subsiteUrl, REG_LIST_NAME);
    if (itemCount < 0 || regs.length < itemCount) {
      throw new Error(`[DEX] getActiveCounts: Sicht unvollständig (${regs.length} von ${itemCount} Items lesbar) — Zählung unbrauchbar (Item-Level-Security).`);
    }
    const active = regs.filter(r => EventService.ACTIVE_STATI.indexOf(r.Status) >= 0);
    return {
      total: active.length,
      durch: active.filter(r => r.StarterType === 'Durchstarter').length,
      fun: active.filter(r => r.StarterType === 'Funstarter').length,
      // v24.73: Warteliste-Zahl für den (privilegierten) Counter-Reconcile.
      waitlist: regs.filter(r => r.Status === 'Warteliste').length,
    };
  }

  private seatFieldFor(group: string): 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun' {
    if (group === 'Durchstarter') return 'SeatsTakenDurch';
    if (group === 'Funstarter') return 'SeatsTakenFun';
    return 'SeatsTaken';
  }

  /**
   * v11.36: Atomare Sitzplatz-Reservierung pro Gruppe via ETag-CAS auf der
   * DEX_TeilnehmerCounter-Liste — exakt dasselbe bewährte Muster wie
   * getNextTeilnehmerId (IF-MATCH, 412-Retry mit Backoff).
   *
   * Verhindert die Überbuchung bei zeitgleichen Anmeldungen: zwei parallele
   * Anmeldungen können nicht beide den letzten Platz greifen — die CAS
   * serialisiert das Increment, der Verlierer liest neu und sieht „voll".
   *
   * Rückgabe:
   * - 'reserved' → Platz wurde atomar belegt, Aufrufer darf 'Angemeldet' setzen
   * - 'full'     → Gruppe/Event ist voll → Aufrufer setzt 'Warteliste'
   * - 'error'    → Counter nicht nutzbar (Liste fehlt, Permission, Retries
   *                 erschöpft). Aufrufer MUSS fail-closed handeln (Warteliste),
   *                 NICHT optimistisch 'Angemeldet'.
   *
   * Self-Seed: ist das Seat-Feld noch nie gesetzt (null), wird es einmalig aus
   * der echten aktiven Anzahl der Gruppe initialisiert, bevor entschieden wird.
   */
  public async reserveSeat(
    subsiteUrl: string,
    group: '' | 'Durchstarter' | 'Funstarter',
    cap: number,
    count: number = 1
  ): Promise<'reserved' | 'full' | 'error'> {
    // cap <= 0 = unbegrenzt → kein Reservieren nötig.
    if (!cap || cap <= 0) return 'reserved';
    const inc = Math.max(1, Math.floor(count));
    const field = this.seatFieldFor(group);
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    // v18.8 (Überbuchungs-Fix): Der Counter allein ist NICHT verlässlich.
    // Der Power-Automate-Nachrück-Flow promotet Warteliste→Angemeldet, ohne
    // SeatsTaken zu erhöhen; läuft der app-seitige syncSeatsToActiveCount
    // zeitlich VOR dieser asynchronen Promotion, steht der Counter unter dem
    // echten Aktiv-Bestand. Folge (real beobachtet): trotz voller Warteliste
    // sah der nächste Registrant einen Phantom-Platz und überbuchte. Deshalb
    // lesen wir EINMAL pro Aufruf den echten Aktiv-Bestand der Gruppe und
    // floor-en den Counter-Wert dagegen (max). Das schließt die Drift-Lücke,
    // erhält die atomare CAS-Serialisierung paralleler Anmeldungen UND heilt
    // den Counter nach oben. Bei Lesefehler (Throttling): kein Floor (-1) →
    // Fallback auf reines Counter-Verhalten, nicht schlechter als vorher.
    let realActive = -1;
    try {
      const rc = await this.getActiveCounts(subsiteUrl);
      realActive = group === 'Durchstarter' ? rc.durch : group === 'Funstarter' ? rc.fun : rc.total;
    } catch { realActive = -1; }
    const MAX_RETRIES = 40;
    let triedLazyCreate = false;
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return 'error';
      }
      if (!getResp.ok) {
        if (getResp.status === 404 && !triedLazyCreate) {
          triedLazyCreate = true;
          try { await this.ensureCounterList(subsiteUrl); continue; } catch { return 'error'; }
        }
        return 'error';
      }
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return 'error';
      let data;
      try { data = await getResp.json(); } catch { return 'error'; }
      const rawVal = data?.[field] ?? data?.d?.[field];
      let current: number;
      if (rawVal === null || rawVal === undefined) {
        // Feld noch nie initialisiert → aus echtem Bestand seeden. v18.8:
        // den bereits oben gelesenen realActive wiederverwenden (kein zweiter
        // getActiveCounts-Roundtrip); nur falls der Read fehlschlug (-1),
        // konservativ auf 0.
        current = realActive >= 0 ? realActive : 0;
      } else {
        current = typeof rawVal === 'number' ? rawVal : (parseInt(String(rawVal), 10) || 0);
      }
      // v18.8: gegen echten Aktiv-Bestand floor-en (siehe Kommentar oben) —
      // fängt eine durch den Nachrück-Flow nach unten gedriftete Zählung ab.
      if (realActive >= 0 && realActive > current) current = realActive;
      // v11.82: Team-Anmeldungen reservieren N Plätze atomar. Wenn nicht alle
      // N in dieselbe Gruppe passen, schlägt die Reservierung als „full" fehl —
      // der Aufrufer setzt das gesamte Team auf Warteliste (kein Teil-Team
      // aktivieren). Bei count=1 (Solo) ist das Verhalten identisch zu vorher.
      if (current + inc > cap) return 'full';
      const patchResp = await this._mergeIfMatch(counterItemUrl, { [field]: current + inc }, etag);
      if (patchResp.ok) return 'reserved';
      if (patchResp.status !== 412) return 'error';
      const baseDelay = Math.min(500, 50 * Math.pow(1.4, attempt));
      await new Promise(res => setTimeout(res, Math.floor(baseDelay * (0.5 + Math.random()))));
    }
    // Retries erschöpft → fail-closed (Aufrufer setzt Warteliste).
    console.warn('[DEX] reserveSeat: 40 Retries erschöpft — fail-closed (Warteliste).');
    return 'error';
  }

  /**
   * v11.36: Sitzplatz-Counter mit dem echten aktiven Bestand abgleichen.
   * Nach Abmeldungen, Reorder und Überbuchungs-Bereinigung aufrufen. Die
   * Power-Automate-Nachrück-Promotion fasst den Counter nicht an — dieser
   * Reconcile (aktive Anzahl aus der Liste) hält ihn ehrlich. Best-effort,
   * ETag-CAS, blockiert nie den aufrufenden Flow.
   *
   * v27.10: Liefert zurück, ob tatsächlich synchronisiert wurde. `false`
   * heißt insbesondere: der Aufrufer sieht die Teilnehmerliste nur beschnitten
   * (Item-Level-Security, siehe getActiveCounts) — dann wird bewusst NICHTS
   * geschrieben und der Aufrufer muss ggf. additiv am Counter arbeiten
   * (releaseSeatAfterCancel).
   */
  public async syncSeatsToActiveCount(
    subsiteUrl: string,
    opts: { isSplit: boolean }
  ): Promise<boolean> {
    let counts: { total: number; durch: number; fun: number; waitlist: number };
    try { counts = await this.getActiveCounts(subsiteUrl); } catch (err) {
      console.warn('[DEX] syncSeatsToActiveCount übersprungen:', err);
      return false;
    }
    // v24.76: WaitlistTaken-Feld sicherstellen, sonst HTTP 400 beim MERGE auf
    // Bestands-Events (Feld noch nicht angelegt).
    await this.ensureCounterFieldsOnce(subsiteUrl);
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    // v24.73: WaitlistTaken (rein informativ, nicht überbuchungs-relevant) beim
    // Reconcile mitschreiben — so heilt eine durch Flow-Promotion gedriftete
    // Warteliste-Zahl. Läuft NUR in privilegierten Kontexten (getActiveCounts
    // braucht Vollzugriff), daher hier korrekt; in User-Self-Cancel-Pfaden wird
    // syncSeatsToActiveCount bewusst nicht aufgerufen.
    const desired = opts.isSplit
      ? { SeatsTakenDurch: counts.durch, SeatsTakenFun: counts.fun, SeatsTaken: counts.total, WaitlistTaken: counts.waitlist }
      : { SeatsTaken: counts.total, WaitlistTaken: counts.waitlist };
    for (let attempt = 0; attempt < 6; attempt++) {
      try {
        const getResp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
        if (!getResp.ok) return false;
        const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
        if (!etag) return false;
        const patchResp = await this._mergeIfMatch(counterItemUrl, desired, etag);
        if (patchResp.ok) return true;
        if (patchResp.status !== 412) return false;
        await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
      } catch { return false; }
    }
    return false;
  }

  /**
   * v27.10: Ein einzelnes Sitzplatz-Counter-Feld additiv anpassen (atomar per
   * ETag-CAS, floor bei 0) — Pendant zu adjustWaitlistCounter für SeatsTaken/
   * SeatsTakenDurch/SeatsTakenFun. Für Aufrufer OHNE Vollzugriff auf die
   * Teilnehmerliste, die keine Absolutwerte schreiben dürfen (siehe
   * getActiveCounts). Best-effort, blockiert nie.
   */
  private async adjustSeatCounterField(
    subsiteUrl: string,
    field: 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun',
    delta: number
  ): Promise<void> {
    if (!subsiteUrl || !delta) return;
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    for (let attempt = 0; attempt < 8; attempt++) {
      try {
        const getResp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
        if (!getResp.ok) return;
        const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
        if (!etag) return;
        const data = await getResp.json();
        const rawVal = data?.[field] ?? data?.d?.[field];
        // Feld nie initialisiert → additiv sinnlos; der nächste privilegierte
        // Reconcile seedet es korrekt.
        if (rawVal === null || rawVal === undefined) return;
        const current = typeof rawVal === 'number' ? rawVal : (parseInt(String(rawVal), 10) || 0);
        const next = Math.max(0, current + delta);
        const patchResp = await this._mergeIfMatch(counterItemUrl, { [field]: next }, etag);
        if (patchResp.ok) return;
        if (patchResp.status !== 412) return;
        await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
      } catch { return; }
    }
  }

  /**
   * v27.10 REGRESSIONS-FIX: Counter-Pflege nach einer Abmeldung — ersetzt den
   * direkten syncSeatsToActiveCount-Aufruf in den Abmelde-Pfaden (Selbst-
   * Abmeldung, Team-Lead-Abmeldung, Assistenz-Abmeldung), die auch von
   * NORMALEN Usern ausgelöst werden.
   *
   * Hintergrund: Seit die Element-Sicherheit („nur eigene Elemente", v26.87)
   * wirklich greift, sah ein normaler User beim Reconcile nur die eigene —
   * gerade abgemeldete — Zeile und schrieb SeatsTaken=0. Folge (real, 15.07.):
   * Neu-Anmeldungen wurden trotz voller Warteliste direkt „Angemeldet" und
   * überholten alle Wartenden.
   *
   * Ablauf jetzt:
   * 1. Voll-Reconcile versuchen (greift nur bei Vollzugriff, z.B. Organizer/
   *    Admin — exakteste Variante, heilt auch alte Drift).
   * 2. Sonst additiv (ILS-sicher):
   *    - Wartelisten-Zeile abgemeldet → WaitlistTaken −1.
   *    - Aktive Zeile abgemeldet → Platz nur freigeben (SeatsTaken −1), wenn
   *      die Warteliste laut Counter LEER ist. Steht jemand auf der Warteliste,
   *      besetzt der IDReorder-/Nachrück-Flow den Platz sofort FIFO-fair —
   *      SeatsTaken bleibt dann unverändert korrekt, und eine parallele
   *      Neu-Anmeldung kann die Wartenden nicht überholen. Bei unbekanntem
   *      Wartelisten-Stand: fail-closed nichts tun (privilegierter Reconcile
   *      heilt spätestens beim nächsten Admin-/Organizer-Boot).
   */
  public async releaseSeatAfterCancel(
    subsiteUrl: string,
    opts: { isSplit: boolean; previousStatus: string; starterType?: string; waitlistDisabled?: boolean }
  ): Promise<void> {
    // v30.63: Jeder Zweig sagt in der Konsole, WAS er getan hat und warum.
    // Der Zähler wird seit v30.62 auf der Anmeldeseite angezeigt; wenn eine
    // Zahl nicht stimmt, muss sich die Ursache ohne Rätselraten ablesen lassen.
    // eslint-disable-next-line no-console
    const log = (msg: string): void => console.log(`[DEX][seats] Abmeldung — ${msg}`);
    try {
      const synced = await this.syncSeatsToActiveCount(subsiteUrl, { isSplit: opts.isSplit });
      if (synced) { log('Voll-Abgleich gelaufen (Vollzugriff) — Zähler exakt neu gesetzt.'); return; }
      if (opts.previousStatus === 'Warteliste') {
        await this.adjustWaitlistCounter(subsiteUrl, -1);
        log('Wartelisten-Zeile abgemeldet → WaitlistTaken −1.');
        return;
      }
      if (EventService.ACTIVE_STATI.indexOf(opts.previousStatus) < 0) {
        log(`Vorheriger Status „${opts.previousStatus}" war nicht aktiv — kein Platz freizugeben.`);
        return;
      }
      // v27.11: Warteliste vom Organizer abgeschaltet → es rückt NIEMAND nach
      // (App-Gates + Flow-Bedingung). Der Platz muss dann direkt freigegeben
      // werden — sonst blieben frei gewordene Plätze dauerhaft als belegt
      // gezählt (Deadlock, bis ein privilegierter Reconcile läuft).
      if (opts.waitlistDisabled) {
        log('Warteliste ist abgeschaltet → SeatsTaken −1 (Platz sofort frei).');
        await this.adjustSeatCounterField(subsiteUrl, 'SeatsTaken', -1);
        if (opts.isSplit && opts.starterType === 'Durchstarter') {
          await this.adjustSeatCounterField(subsiteUrl, 'SeatsTakenDurch', -1);
        } else if (opts.isSplit && opts.starterType === 'Funstarter') {
          await this.adjustSeatCounterField(subsiteUrl, 'SeatsTakenFun', -1);
        }
        return;
      }
      const stats = await this.getCounterStats(subsiteUrl, opts.isSplit);
      // stats.waitlist: -1 = unbekannt (Feld nie gepflegt) → fail-closed.
      if (!stats) { log('Zähler nicht lesbar → nichts geändert (heilt beim nächsten Abgleich).'); return; }
      if (stats.waitlist < 0) {
        log('Wartelisten-Stand UNBEKANNT (Feld nie geschrieben) → nichts geändert. Seit v30.63 wird es bei neuen Events mit 0 angelegt; Bestands-Events heilt der Abgleich beim nächsten Organizer-/Admin-Start.');
        return;
      }
      if (stats.waitlist !== 0) {
        log(`${stats.waitlist} Person(en) auf der Warteliste → SeatsTaken bleibt (der Nachrück-Flow besetzt den Platz sofort).`);
        return;
      }
      log('Warteliste leer → SeatsTaken −1.');
      await this.adjustSeatCounterField(subsiteUrl, 'SeatsTaken', -1);
      if (opts.isSplit && opts.starterType === 'Durchstarter') {
        await this.adjustSeatCounterField(subsiteUrl, 'SeatsTakenDurch', -1);
      } else if (opts.isSplit && opts.starterType === 'Funstarter') {
        await this.adjustSeatCounterField(subsiteUrl, 'SeatsTakenFun', -1);
      }
    } catch (err) {
      console.warn('[DEX] releaseSeatAfterCancel failed (best-effort):', err);
    }
  }

  /**
   * v24.73: Warteliste-Zähler im Counter additiv anpassen (atomar per ETag-CAS).
   * REIN INFORMATIV — `WaitlistTaken` gatet keine Überbuchung; ein verlorener
   * Bump verfälscht nur kurz die angezeigte Warteliste-Zahl und wird vom
   * privilegierten `syncSeatsToActiveCount`-Reconcile wieder geheilt. Wird von
   * den Anmelde-/Abmelde-Pfaden mit delta +1/-1 aufgerufen (auch von normalen
   * Usern — der Counter ist für alle schreibbar). Best-effort, blockiert nie.
   */
  public async adjustWaitlistCounter(subsiteUrl: string, delta: number): Promise<void> {
    if (!subsiteUrl || !delta) return;
    // v24.76: Feld sicherstellen (greift für privilegierte Aufrufer; normale
    // User dürfen kein Feld anlegen → dann wird der Write unten übersprungen).
    await this.ensureCounterFieldsOnce(subsiteUrl);
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    for (let attempt = 0; attempt < 8; attempt++) {
      try {
        const getResp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
        if (!getResp.ok) return;
        const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
        if (!etag) return;
        const data = await getResp.json();
        const body = data?.d ?? data;
        // v24.76: Feld (noch) nicht vorhanden → MERGE würde 400 liefern → still
        // überspringen (kein Konsolen-Fehler). Sobald ein Admin reconcilet, ist
        // das Feld da und der Bump greift.
        if (body && typeof body === 'object' && !('WaitlistTaken' in body)) return;
        const rawVal = data?.WaitlistTaken ?? data?.d?.WaitlistTaken;
        const current = typeof rawVal === 'number' ? rawVal : (parseInt(String(rawVal), 10) || 0);
        const next = Math.max(0, current + delta);
        const patchResp = await this._mergeIfMatch(counterItemUrl, { WaitlistTaken: next }, etag);
        if (patchResp.ok) return;
        if (patchResp.status !== 412) return; // anderer Fehler → aufgeben (best-effort)
        await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
      } catch { return; }
    }
  }

  /**
   * v24.73: Live-Plätze aus dem Counter lesen — für ALLE lesbar (auch normale
   * Teilnehmer, im Gegensatz zur item-level-gesicherten Teilnehmerliste). Quelle
   * der Anzeige-Zahlen (aktiv = SeatsTaken, Warteliste = WaitlistTaken). Liefert
   * `null`, wenn der Counter (noch) nicht existiert/lesbar ist → Aufrufer fällt
   * dann auf den bisherigen (item-level-gefilterten) Zählweg zurück.
   */
  public async getCounterStats(subsiteUrl: string, isSplit: boolean): Promise<{ active: number; waitlist: number; seatsKnown: boolean } | null> {
    if (!subsiteUrl) return null;
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    try {
      const resp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      if (!resp.ok) return null;
      const data = await resp.json();
      const num = (v: unknown): number => typeof v === 'number' ? v : (parseInt(String(v ?? ''), 10) || 0);
      const total = num(data?.SeatsTaken ?? data?.d?.SeatsTaken);
      const durch = num(data?.SeatsTakenDurch ?? data?.d?.SeatsTakenDurch);
      const fun = num(data?.SeatsTakenFun ?? data?.d?.SeatsTakenFun);
      const wRaw = data?.WaitlistTaken ?? data?.d?.WaitlistTaken;
      // v30.62: Ist das Feld noch NIE gesetzt worden (null/undefined), liefert
      // `num()` eine 0 — und eine 0 liest sich wie „niemand angemeldet". Für die
      // Termin-Kacheln der Anmeldeseite ist das der Unterschied zwischen einer
      // Aussage und einer Erfindung, deshalb wird er hier durchgereicht.
      const sRaw = data?.SeatsTaken ?? data?.d?.SeatsTaken;
      const sDurchRaw = data?.SeatsTakenDurch ?? data?.d?.SeatsTakenDurch;
      const sFunRaw = data?.SeatsTakenFun ?? data?.d?.SeatsTakenFun;
      const known = (v: unknown): boolean => v !== null && v !== undefined && v !== '';
      const seatsKnown = known(sRaw) || (isSplit && (known(sDurchRaw) || known(sFunRaw)));
      // SeatsTaken ist der Gesamt-Aktiv-Wert; bei Split fällt er ggf. auf
      // Durch+Fun zurück, falls der Gesamtwert (noch) nicht gepflegt wurde.
      const active = total > 0 ? total : (isSplit ? durch + fun : total);
      const waitlist = (wRaw === null || wRaw === undefined) ? -1 : num(wRaw);
      return { active, waitlist, seatsKnown };
    } catch { return null; }
  }

  /**
   * v24.75: Echtzeit-Push auf eine Liste der Event-Subsite abonnieren.
   * kind='counter' → DEX_TeilnehmerCounter (für alle lesbar; Anmeldeformular),
   * kind='participants' → Teilnehmerliste (Organizer-Vollzugriff). Liefert eine
   * Cleanup-Funktion. Best-effort (siehe utils/spListRealtime).
   */
  public async subscribeListRealtime(
    subsiteUrl: string,
    kind: 'counter' | 'participants',
    onChange: () => void
  ): Promise<() => void> {
    if (!subsiteUrl) return () => { /* */ };
    const listTitle = kind === 'counter' ? COUNTER_LIST_NAME : REG_LIST_NAME;
    try { return await subscribeListChanges(this.context.spHttpClient, subsiteUrl, listTitle, onChange); }
    catch { return () => { /* */ }; }
  }

  /**
   * v11.36: Überbuchung erkennen + markieren (ändert KEINEN Status).
   * Pro Gruppe (bzw. gesamt bei Nicht-Split) werden die zuletzt angemeldeten
   * Einträge über Kapazität (höchste SP-Id = zuletzt registriert) mit
   * OverbookReview='Pending' markiert. First-come-first-served: wer zuerst
   * da war, behält den Platz.
   */
  public async detectOverbooking(
    subsiteUrl: string,
    opts: { isSplit: boolean; maxParticipants?: number; durchstarterCapacity?: number; funstarterCapacity?: number }
  ): Promise<{ groups: Array<{ group: string; cap: number; activeBefore: number; marked: number }>; total: number; errors: number }> {
    const regs = await this.getAllRegistrations(subsiteUrl); // Id asc
    const groups: Array<{ group: string; cap: number; activeBefore: number; marked: number }> = [];
    let total = 0;
    let errors = 0;
    const markExcess = async (items: SPRegistration[], cap: number, label: string): Promise<void> => {
      const before = items.length;
      let marked = 0;
      if (cap > 0 && before > cap) {
        const excess = items.slice(cap); // Id asc → ab Index cap = die neuesten
        for (const it of excess) {
          if (it.OverbookReview === 'Pending') { marked++; total++; continue; }
          try {
            const resp = await this._merge(
              `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${it.Id})`,
              { 'OverbookReview': 'Pending' }
            );
            if (resp.ok || resp.status === 406) { marked++; total++; } else { errors++; }
          } catch { errors++; }
        }
      }
      groups.push({ group: label, cap, activeBefore: before, marked });
    };
    const isActive = (r: SPRegistration): boolean => EventService.ACTIVE_STATI.indexOf(r.Status) >= 0;
    if (opts.isSplit) {
      await markExcess(regs.filter(r => isActive(r) && r.StarterType === 'Durchstarter'), opts.durchstarterCapacity || 0, 'Durchstarter');
      await markExcess(regs.filter(r => isActive(r) && r.StarterType === 'Funstarter'), opts.funstarterCapacity || 0, 'Funstarter');
    } else {
      await markExcess(regs.filter(isActive), opts.maxParticipants || 0, 'all');
    }
    return { groups, total, errors };
  }

  /** v11.36: Review-Marker einer Zeile entfernen (ohne Status-Änderung). */
  public async clearOverbookMark(subsiteUrl: string, itemId: number): Promise<boolean> {
    try {
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'OverbookReview': '' }
      );
      return resp.ok || resp.status === 406;
    } catch { return false; }
  }

  /**
   * v11.36: Markierte Person auf die Warteliste setzen (= „Bestätigen").
   * Gruppentreu: PreferredStarterType bleibt die Gruppe, StarterType wird
   * geleert (wie bei switchSplitGroup → Warteliste).
   */
  public async resolveOverbookToWaitlist(
    subsiteUrl: string,
    itemId: number,
    group: string
  ): Promise<boolean> {
    try {
      // Audit: festhalten, dass die Person kurz einen bestätigten Platz hatte
      // und wegen der technischen Überbuchung auf Warteliste korrigiert wurde
      // (inkl. Original-Registrierung) — dauerhaft nachvollziehbar, unabhängig
      // von der späteren TeilnehmerID-Neuvergabe.
      let changeLog = '';
      let origDate = '';
      try {
        const ex = await this._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog,RegistrationDate`,
          SPHttpClient.configurations.v1
        );
        if (ex.ok) {
          const d = await ex.json();
          changeLog = d.ChangeLog || d.d?.ChangeLog || '';
          origDate = d.RegistrationDate || d.d?.RegistrationDate || '';
        }
      } catch { /* ChangeLog optional */ }
      const stamp = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      const note = `[${stamp}] Überbuchung: war fälschlich angemeldet (technisches Problem bei zeitgleicher Anmeldung) → auf Warteliste korrigiert (Original-Registrierung: ${origDate || 'unbekannt'})`;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = group
        ? { 'Status': 'Warteliste', 'StarterType': '', 'PreferredStarterType': group, 'OverbookReview': '' }
        : { 'Status': 'Warteliste', 'OverbookReview': '' };
      body['ChangeLog'] = changeLog ? `${changeLog}\n${note}` : note;
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return resp.ok || resp.status === 406;
    } catch { return false; }
  }

  /**
   * v11.36: „Platz behalten" — Variante (a): bleibt angemeldet (Marker weg).
   * Die Gruppe bleibt damit ggf. +1 über Kapazität; der Power-Automate-Flow
   * (Check_<Typ>_Free bzw. Check_Nachrücken, strikt `<`) rückt beim nächsten
   * Frei-Werden in dieser Gruppe einmal NICHT nach — die Überzahl wird so
   * automatisch absorbiert. Identisch zu clearOverbookMark, eigener Name
   * nur fürs Audit/Lesbarkeit.
   */
  public async resolveOverbookKeepActive(subsiteUrl: string, itemId: number): Promise<boolean> {
    return this.clearOverbookMark(subsiteUrl, itemId);
  }

  /**
   * v11.36: „Platz behalten" — Variante (b): Person wird Erste(r) auf der
   * gruppeneigenen Warteliste. Der Nachrück-Flow sortiert Warteliste nach
   * RegistrationDate asc — daher setzen wir RegistrationDate knapp VOR den
   * frühesten aktuellen Wartelisten-Eintrag derselben Gruppe. Original-Datum
   * wird im ChangeLog vermerkt.
   */
  public async resolveOverbookKeepAsFirstWaitlist(
    subsiteUrl: string,
    itemId: number,
    group: string
  ): Promise<boolean> {
    try {
      const all = await this.getAllRegistrations(subsiteUrl);
      const sameGroupWaitlist = all.filter(r =>
        r.Status === 'Warteliste' && (!group || r.PreferredStarterType === group)
      );
      let newDateMs = Date.now();
      for (const w of sameGroupWaitlist) {
        const t = new Date(w.RegistrationDate).getTime();
        if (!isNaN(t) && t < newDateMs) newDateMs = t;
      }
      newDateMs -= 1000; // 1s vor den/die bisherige(n) Erste(n)
      const self = all.find(r => r.Id === itemId);
      const origDate = self?.RegistrationDate || '';
      const stamp = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      let changeLog = '';
      try {
        const ex = await this._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
          SPHttpClient.configurations.v1
        );
        if (ex.ok) { const d = await ex.json(); changeLog = d.ChangeLog || d.d?.ChangeLog || ''; }
      } catch { /* ChangeLog optional */ }
      const note = `[${stamp}] Überbuchung: Platz behalten → Erste(r) auf Warteliste (Original-Registrierung: ${origDate})`;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'Status': 'Warteliste',
        'StarterType': '',
        'PreferredStarterType': group || '',
        'RegistrationDate': new Date(newDateMs).toISOString(),
        'OverbookReview': '',
        'ChangeLog': changeLog ? `${changeLog}\n${note}` : note,
      };
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return resp.ok || resp.status === 406;
    } catch { return false; }
  }

  /**
   * v11.36: Vorgeschlagener Entschuldigungs-Mailtext (Deloitte-Wrap, DE/EN)
   * für „Bestätigen mit Mail". Der Admin kann den Text im Modal vor dem
   * Versand editieren.
   */
  // v13.0: Lädt das OverbookingApology-Template aus DEX_EmailTemplates;
  // wenn das Template existiert wird daraus die Mail gebaut (inkl.
  // Reseed-Funktionalität für Admins). Fallback ist der alte Inline-Text
  // damit ältere Tenants ohne Template-Update nicht ohne Mail dastehen.
  public async buildOverbookApologyEmail(
    name: string,
    eventTitle: string,
    lang: string,
    waitlistPos?: number
  ): Promise<{ subject: string; body: string }> {
    const isDe = (lang || 'EN').toUpperCase() === 'DE';
    const first = (name || '').split(' ')[0] || name;
    const hasPos = typeof waitlistPos === 'number' && waitlistPos > 0;
    const posBlock = hasPos
      ? (isDe
        ? `<p>Du stehst jetzt auf <strong>Warteliste-Platz ${waitlistPos}</strong>.</p>`
        : `<p>You are now <strong>waitlist position ${waitlistPos}</strong>.</p>`)
      : '';
    const tpl = await this.getEmailTemplate('OverbookingApology', lang).catch(() => null);
    const vars: Record<string, string> = {
      Name: first || name,
      EventTitle: eventTitle,
      WaitlistPositionBlock: posBlock,
      WaitlistPosition: hasPos ? String(waitlistPos) : '',
      AppUrl: `${this.siteUrl}/SitePages/DEX.aspx?env=WebView`,
    };
    if (tpl) {
      return buildEmailFromTemplate(tpl, vars);
    }
    // Fallback-Inline (alte Pfade)
    const heading = isDe ? 'Anmeldung korrigiert' : 'Registration corrected';
    if (isDe) {
      const inner = `<p>Hallo ${first},</p>`
        + `<p>leider müssen wir uns für ein technisches Problem entschuldigen: durch sehr viele zeitgleiche Anmeldungen wurde dir für <strong>${eventTitle}</strong> versehentlich ein Platz bestätigt, obwohl die Kapazität bereits erschöpft war.</p>`
        + `<p>Wir mussten deine Anmeldung daher auf die <strong>Warteliste</strong> korrigieren. Das tut uns aufrichtig leid — es lag nicht an dir, sondern an einem Ansturm auf die Anmeldung.</p>`
        + posBlock
        + `<p>Sobald ein Platz frei wird, rückst du automatisch nach und bekommst sofort eine Bestätigung. Du musst nichts weiter tun.</p>`
        + `<p style="margin-top:24px;"><strong>Vielen Dank für dein Verständnis</strong><br><br><strong>Dein Event-Team</strong></p>`;
      return {
        subject: `Wichtig: Korrektur deiner Anmeldung — ${eventTitle}`,
        body: wrapTemplateForStorage('#ed8b00', heading, eventTitle, inner),
      };
    }
    const inner = `<p>Hi ${first},</p>`
      + `<p>we sincerely apologize for a technical problem: due to a large number of simultaneous registrations, you were mistakenly confirmed a spot for <strong>${eventTitle}</strong> although capacity was already full.</p>`
      + `<p>We therefore had to move your registration to the <strong>waitlist</strong>. We're truly sorry — this was not your fault but caused by a registration rush.</p>`
      + posBlock
      + `<p>As soon as a spot opens up you will be promoted automatically and notified right away. Nothing else is needed from your side.</p>`
      + `<p style="margin-top:24px;"><strong>Thank you for your understanding</strong><br><br><strong>Your Event Team</strong></p>`;
    return {
      subject: `Important: correction of your registration — ${eventTitle}`,
      body: wrapTemplateForStorage('#ed8b00', heading, eventTitle, inner),
    };
  }

  /**
   * Spalten der Teilnehmerliste fixen: fehlende Spalten anlegen, View-Reihenfolge korrigieren.
   * Kann auf bestehenden Events ausgeführt werden um die Liste nachträglich zu aktualisieren.
   */
  /**
   * v30.58: Prüft EINE Teilnehmerliste gegen die Abfragefelder ihres Events —
   * ohne etwas zu ändern.
   *
   * Der Anlass: Fehlt auf einer Liste die Spalte zu einem Abfragefeld, lehnt
   * SharePoint den GANZEN Insert ab (HTTP 400, „The field or property … does
   * not exist"). Betroffen sind dann nur die Personen, die dieses Feld
   * ausfüllen — deshalb sieht es aus wie ein zufälliger Einzelfall und nicht
   * wie ein Struktur-Fehler. Genau so ist die fehlende Klammer-Zeile
   * entstanden: Die übergreifenden Hauptevent-Felder gehen NUR auf die
   * Klammer-Liste, also scheitert auch nur dort der Schreibvorgang.
   *
   * Bewusst lesend: „Was ist kaputt?" und „repariere es" sind zwei Fragen.
   * Wer erst sehen will, was los ist, soll dafür nichts verändern müssen.
   */
  public async diagnoseRegistrationList(
    subsiteUrl: string,
    customFields: Array<{ id: string; label: string; spInternalName?: string }>
  ): Promise<{ ok: boolean; listMissing: boolean; missingColumns: Array<{ id: string; label: string; column: string }>; error?: string }> {
    try {
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=500&$select=InternalName,Title`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) {
        // 404 heißt: Die Liste gibt es nicht (mehr). Das ist eine ANDERE
        // Aussage als „Spalte fehlt" und darf nicht damit vermischt werden —
        // dieselbe Unterscheidung wie bei getAllRegistrations (v29.3).
        return { ok: false, listMissing: resp.status === 404 || resp.status === 410, missingColumns: [], error: `HTTP ${resp.status}` };
      }
      const data = await resp.json();
      const rows: Array<{ InternalName?: string; Title?: string }> = data.value || data.d?.results || [];
      const have = new Set<string>();
      for (const r of rows) {
        if (r.InternalName) have.add(r.InternalName.toLowerCase());
        if (r.Title) have.add(r.Title.toLowerCase());
      }
      const missing: Array<{ id: string; label: string; column: string }> = [];
      for (const f of (customFields || [])) {
        const col = (f.spInternalName || '').trim();
        // Ohne spInternalName wurde das Feld noch nie auf eine Liste
        // geschrieben — dann ist nicht die Liste schuld, sondern die
        // Zuordnung fehlt. Beides meldet „Spalten fixen" als zu erledigen.
        if (!col) { missing.push({ id: f.id, label: f.label, column: '(noch nicht zugeordnet)' }); continue; }
        if (!have.has(col.toLowerCase())) missing.push({ id: f.id, label: f.label, column: col });
      }
      return { ok: true, listMissing: false, missingColumns: missing };
    } catch (e) {
      return { ok: false, listMissing: false, missingColumns: [], error: String((e as Error)?.message || e) };
    }
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
    const added: string[] = [];
    const removed: string[] = [];
    const duplicatesRemoved: string[] = [];
    const duplicatesWithData: string[] = [];

    // Bestehende Felder laden — InternalName + Title beide nehmen, damit wir per Title
    // dedupen können (siehe v11.56: alte Builds haben durch fehlgeschlagene Existenz-
    // checks beim wiederholten "Spalten fixen" pro Custom-Field 50+ Duplikate angelegt).
    const fieldsResp = await this._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=500&$select=InternalName,Title`,
      SPHttpClient.configurations.v1
    );
    const existingFieldsList: string[] = [];
    const existingByInternal: Set<string> = new Set();
    const existingByTitle: Map<string, Array<{ internalName: string }>> = new Map();
    if (fieldsResp.ok) {
      const fieldsData = await fieldsResp.json();
      const fields = fieldsData.value || fieldsData.d?.results || [];
      for (const f of fields) {
        const internalName: string = String(f.InternalName || '');
        const title: string = String(f.Title || '');
        if (!internalName) continue;
        existingFieldsList.push(internalName);
        existingByInternal.add(internalName);
        if (title) {
          const arr = existingByTitle.get(title) || [];
          arr.push({ internalName });
          existingByTitle.set(title, arr);
        }
      }
    }

    // ===== DEDUPE-PASS (v11.56) =====
    // Pro Title: wenn mehr als ein Feld diesen Titel hat → die uebrigen Felder löschen,
    // sofern sie leer sind (keine Items mit Wert in der Spalte). Erstes Feld bleibt
    // immer erhalten. Felder mit Daten werden gemeldet (duplicatesWithData) und nicht
    // automatisch gelöscht — der User soll sie manuell prüfen.
    const duplicateTitles: Array<{ title: string; entries: Array<{ internalName: string }> }> = [];
    existingByTitle.forEach((entries, title) => {
      if (entries.length > 1) {
        duplicateTitles.push({ title, entries });
      }
    });
    if (duplicateTitles.length > 0) {
      // Vor dem Löschen den Aufrufer fragen — Operation ist irreversibel.
      if (confirmDeleteDuplicates) {
        const count = duplicateTitles.reduce((sum, d) => sum + (d.entries.length - 1), 0);
        const titles = duplicateTitles.map(d => d.title);
        const ok = await Promise.resolve(confirmDeleteDuplicates(count, titles));
        if (!ok) {
          // Cleanup überspringen — nur den Hauptfix laufen lassen
          duplicateTitles.length = 0;
        }
      }
    }
    if (duplicateTitles.length > 0) {
      // Pro Duplikat-Set: den ersten Eintrag behalten, für alle weiteren prüfen ob leer.
      for (const dup of duplicateTitles) {
        // entries[0] bleibt erhalten
        for (let i = 1; i < dup.entries.length; i++) {
          const candidate = dup.entries[i];
          let isEmpty = false;
          // v11.67: SP truncated InternalNames auf 32 Zeichen. Wenn die Truncation
          // mitten in einer `_xXXXX_`-Encoding-Sequenz liegt (z.B.
          // `ADMIN_x0020__x002d__x0020_Who_x00` — die letzten Zeichen `_x00`
          // sind eine angeschnittene `_x0020_`-Sequenz), wirft SP HTTP 400 auf
          // jeden OData-`$filter`-Versuch. Solche Spalten werden hier nicht
          // geprüft → konservativ als „hat Daten" behandelt (kein Auto-
          // Löschen). Der Admin kann sie über die SP-Listen-UI manuell
          // entfernen, wenn sie wirklich leer sind.
          const looksTruncated = candidate.internalName.length === 32
            && /_x[0-9a-f]{1,3}$/i.test(candidate.internalName);
          if (looksTruncated) {
            if (duplicatesWithData.indexOf(dup.title) < 0) duplicatesWithData.push(dup.title);
            continue;
          }
          try {
            const checkUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${candidate.internalName} ne null&$top=1&$select=ID`;
            const checkResp = await this._sp.get(checkUrl, SPHttpClient.configurations.v1);
            if (checkResp.ok) {
              const data = await checkResp.json();
              const items = data.value || data.d?.results || [];
              isEmpty = items.length === 0;
            }
          } catch { isEmpty = false; }
          if (isEmpty) {
            try {
              const delResp = await this._post(
                `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields/getbyinternalnameortitle('${candidate.internalName}')/deleteObject`,
                {}
              );
              if (delResp.ok) {
                duplicatesRemoved.push(candidate.internalName);
                // Aus existingByInternal + existingFieldsList rausziehen
                existingByInternal.delete(candidate.internalName);
                const idx = existingFieldsList.indexOf(candidate.internalName);
                if (idx >= 0) existingFieldsList.splice(idx, 1);
              }
            } catch { /* löschen fehlgeschlagen — weiter */ }
          } else {
            // Daten vorhanden — Title für manuelle Prüfung melden (nur einmal pro Title)
            if (duplicatesWithData.indexOf(dup.title) < 0) duplicatesWithData.push(dup.title);
          }
        }
        // existingByTitle entsprechend bereinigen: nur die nicht-gelöschten Einträge behalten
        const remaining = dup.entries.filter(e => existingByInternal.has(e.internalName));
        existingByTitle.set(dup.title, remaining);
      }
    }

    // Fehlende Basis-Spalten anlegen. StarterType/Quiz-Felder sind feature-spezifisch:
    // - StarterType/PreferredStarterType: nur für B2Run-Events mit Split-Kapazität
    // - QuizScore/QuizAnswers/QuizCompletedAt: nur für Events mit Quizfragen
    // Wird das Event ohne eventContext gefixt (kein Aufrufer-seitiger Flag), lassen wir
    // feature-spezifische Spalten raus, damit sie nicht unbegründet auf jedem
    // Teilnehmerlisten-Schema auftauchen.
    const requiredFields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
      { title: 'Company', type: 2 },           // v24.29: Unternehmenszugehörigkeit / Rechtsträger (aus dem Profil)
      { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgeführt hat
      { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgeführt hat
      { title: 'ProxyConsent', type: 3 },      // v18.74: Zustimmungs-Nachweis bei stellvertretender Anmeldung (Note)
      { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgelöst hat
      { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
      { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
      { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
      { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
      // v19.3: Nachrück-Audit-Spalten auch beim „Spalten fixen" nachziehen, damit
      // der DEX_IDReorder-Flow (und der App-Button) sie auf Bestands-Events
      // beschreiben kann → „Nachgerückt am / Ersetzt / Ersetzt durch" in der App.
      { title: 'PromotedDate', type: 4 },              // DateTime — Zeitpunkt des Nachrückens
      { title: 'ReplacedParticipantEmail', type: 2 },  // E-Mail der Person, deren Cancel den Platz freigab
      { title: 'ReplacedByParticipantEmail', type: 2 },// E-Mail der nachrückenden Person (Spiegelbild)
      { title: 'OverbookReview', type: 2 },    // v11.36: Überbuchungs-Review-Marker
      { title: 'ConsentReview', type: 2 },     // v26.47: Externe Anmeldung — 'Pending' = Datenschutz-Rückmeldung offen
      { title: 'TeamId', type: 2 },            // v11.82: UUID einer Team-Anmeldung (leer = Solo)
      { title: 'TeamLead', type: 8 },          // v11.82: Boolean — true für die anmeldende Person
      { title: 'TeamName', type: 2 },          // v11.82: optionaler frei wählbarer Team-Name
    ];
    if (eventContext?.isB2Run) {
      requiredFields.push(
        { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' },
        { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' },
        // v10.13: b2run_leistungsnachweis ist ein virtuelles Feld das die
        // RegistrationPage hardcoded hinzufügt wenn durchstarterRequiresProof
        // aktiv ist — es ist NICHT Teil der regulären customFields, daher
        // muss die SP-Spalte hier explizit angelegt werden, sonst kippt die
        // Anmeldung mit HTTP 400 'Field not found'. Wird auf jedem B2Run-Event
        // angelegt (egal ob proof-flag aktuell aktiv ist) — die Spalte ist
        // klein und stört nicht wenn ungenutzt.
        { title: 'b2run_leistungsnachweis', type: 2 }
      );
    }
    if (eventContext?.hasQuiz) {
      requiredFields.push(
        { title: 'QuizScore', type: 9 },
        { title: 'QuizAnswers', type: 3 },
        { title: 'QuizCompletedAt', type: 4 }
      );
    }

    // Feature-spezifische Spalten, die auf diesem Event NICHT gebraucht werden,
    // aktiv löschen (z.B. StarterType auf einem Nicht-B2Run-Event). Das ist
    // irreversibel — eventuelle Daten in diesen Spalten gehen verloren. Ist aber
    // vom User explizit gewünscht, damit die Teilnehmerliste pro Event-Typ
    // sauber bleibt.
    const deletableFields: string[] = [];
    if (!eventContext?.isB2Run) {
      deletableFields.push('StarterType', 'PreferredStarterType');
    }
    if (!eventContext?.hasQuiz) {
      deletableFields.push('QuizScore', 'QuizAnswers', 'QuizCompletedAt');
    }
    for (const fieldName of deletableFields) {
      if (existingFieldsList.indexOf(fieldName) >= 0) {
        try {
          const delResp = await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields/getbytitle('${fieldName}')/deleteObject`,
            {}
          );
          if (delResp.ok) {
            removed.push(fieldName);
            // aus existingFieldsList rausziehen damit die View-Logik weiter unten
            // den Feldnamen nicht mehr als "noch vorhanden" betrachtet.
            const idx = existingFieldsList.indexOf(fieldName);
            if (idx >= 0) existingFieldsList.splice(idx, 1);
          }
        } catch { /* Feld konnte nicht gelöscht werden - weitermachen */ }
      }
    }

    for (const f of requiredFields) {
      if (existingFieldsList.indexOf(f.title) < 0) {
        const payload: Record<string, unknown> = {
          '__metadata': { 'type': f.metaType || 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        };
        if (f.choices) {
          payload['Choices'] = { 'results': f.choices };
        }
        const resp = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload
        );
        if (resp.ok) added.push(f.title);
      }
    }

    // Custom Fields pro Event: wenn spInternalName leer oder die Spalte fehlt,
    // jetzt anlegen. Der Aufrufer bekommt customFieldMap zurück und kann das
    // Event-Item persistieren (spInternalName für jede cf.id).
    const customFieldMap: Record<string, string> = {};
    if (eventContext?.customFields && eventContext.customFields.length > 0) {
      // Post-Fix Felder-Snapshot nach Basis-Anlage
      let currentFields = [...existingFieldsList, ...added];
      for (const cf of eventContext.customFields) {
        if (!cf.label || !cf.label.trim()) continue;
        // v19.0: Dokument-Felder bekommen keine Spalte (Datei = Attachment).
        if (cf.type === 'document') continue;
        // Wenn spInternalName schon gesetzt und Feld existiert: übernehmen.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const existingSp: string = String((cf as any).spInternalName || '');
        if (existingSp && currentFields.indexOf(existingSp) >= 0) {
          customFieldMap[cf.id] = existingSp;
          continue;
        }
        // v11.56: Wenn spInternalName fehlt oder nicht zur Liste passt, aber ein
        // Feld mit demselben Title bereits existiert: dieses InternalName übernehmen,
        // statt eine Duplikat-Spalte anzulegen. Das ist die Hauptursache der
        // 100x-Duplikate-Misere (P/D MEETING0, P/D MEETING1, ...).
        const titleMatches = existingByTitle.get(cf.label) || [];
        if (titleMatches.length > 0) {
          const firstInternal = titleMatches[0].internalName;
          customFieldMap[cf.id] = firstInternal;
          if (currentFields.indexOf(firstInternal) < 0) currentFields.push(firstInternal);
          continue;
        }
        // Feld-Payload je nach Typ
        let fieldPayload: Record<string, unknown>;
        if (cf.type === 'select' && cf.options && cf.options.length > 0) {
          fieldPayload = { '__metadata': { 'type': 'SP.FieldChoice' }, 'Title': cf.label, 'FieldTypeKind': 6, 'Required': false, 'Choices': { 'results': cf.options } };
        } else if (cf.type === 'number') {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 9, 'Required': false };
        } else if (cf.type === 'checkbox') {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 8, 'Required': false };
        } else if (cf.type === 'user') {
          // user-Picker wird als Text gespeichert ("Name <email>").
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 3, 'Required': false };
        } else {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 2, 'Required': false };
        }
        try {
          const resp = await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, fieldPayload
          );
          if (resp.ok) {
            const createdField = await resp.json().catch(() => null);
            const internalName = createdField?.InternalName || createdField?.d?.InternalName || '';
            if (internalName) {
              customFieldMap[cf.id] = internalName;
              added.push(internalName);
              currentFields = currentFields.concat([internalName]);
              // Title-Map aktualisieren, damit ein zweites cf mit gleichem Label
              // im selben Durchlauf (z.B. zwei Custom-Fields mit identischem Title)
              // das gerade angelegte Feld wiederverwendet, statt erneut zu erzeugen.
              const arr = existingByTitle.get(cf.label) || [];
              arr.push({ internalName });
              existingByTitle.set(cf.label, arr);
            }
          }
        } catch { /* nächstes Feld */ }
      }
    }

    // Default View komplett neu aufbauen (Reihenfolge: TeilnehmerID, Anrede, Vorname, Nachname, ...)
    let viewFixed = false;
    try {
      // Alle bestehenden Felder aus der View entfernen
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/removeallviewfields`,
        {}
      );

      // Felder in gewünschter Reihenfolge hinzufügen. StarterType/Quiz-Spalten
      // werden nur eingebaut, wenn sie tatsächlich auf der Liste existieren —
      // auf Nicht-B2Run- bzw. Nicht-Quiz-Events sollen sie nicht auftauchen.
      const viewFieldsCore = [
        'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail',
        'Department', 'Location', 'JobTitle', 'Company', 'Phone',
      ];
      const viewFields: string[] = [...viewFieldsCore];
      // Post-Fix-Feldliste (bestehende + gerade hinzugefügte) für die Existenz-Checks
      const postFixFields = existingFieldsList.concat(added);
      if (postFixFields.indexOf('StarterType') >= 0) viewFields.push('StarterType');
      if (postFixFields.indexOf('PreferredStarterType') >= 0) viewFields.push('PreferredStarterType');
      viewFields.push('Status', 'RegistrationDate');
      if (postFixFields.indexOf('RegisteredByName') >= 0) viewFields.push('RegisteredByName');
      if (postFixFields.indexOf('RegisteredByEmail') >= 0) viewFields.push('RegisteredByEmail');
      if (postFixFields.indexOf('ProxyConsent') >= 0) viewFields.push('ProxyConsent');
      viewFields.push('CancellationDate');

      // Wir blenden SP-System-Spalten komplett aus (Modified, Created, ID, Type,
      // Compliance Asset Id, Retention Label, etc.) — nur funktionelle Felder der
      // App + Custom Fields kommen in die View.
      const systemBlocklist = new Set([
        'ID', '_UIVersionString', 'Edit', 'LinkTitle', 'LinkTitleNoMenu',
        'LinkFilename', 'LinkFilenameNoMenu', 'DocIcon', 'FileLeafRef',
        'Modified', 'Created', 'Editor', 'Author', 'CreatedBy', 'ModifiedBy',
        'Title', 'ParticipantName',
        'ContentType', 'ContentTypeId', 'Attachments',
        'AppAuthor', 'AppEditor', 'App Created By', 'App Modified By',
        'Type', 'ItemChildCount', 'FolderChildCount',
        'ComplianceAssetId', '_ComplianceTag', '_ComplianceTagWrittenTime',
        '_ComplianceTagUserId', 'TaxCatchAll', 'TaxCatchAllLabel',
        'SMTotalFileStreamSize', 'SMTotalSize', 'SortBehavior',
        'OData__UIVersionString', 'OData__HasCopyDestinations',
        'LastModifiedDate', 'ChangeLog', 'CustomData',
        '_CopySource', 'owshiddenversion', 'WorkflowVersion', 'WorkflowInstanceID',
        'ItemIsRecord', '_HasEncryptedContent', '_IsRecord', '_IsRecordApplied',
        'InstanceID', 'Order', 'GUID', 'FileSizeDisplay', 'MetaInfo',
        'ParentUniqueId', 'AccessPolicy', 'HasUniqueRoleAssignments',
        'Restricted', 'Type0', 'ServerUrl', 'EncodedAbsUrl', 'BaseName',
        'FileType', 'HTML_x0020_File_x0020_Type', '_EditMenuTableStart',
        '_EditMenuTableStart2', '_EditMenuTableEnd', 'PermMask',
      ]);
      // Bereits zur View hinzugefügt — nicht doppelt anfassen
      const alreadyAdded = new Set(viewFields);
      // v11.82: Team-Spalten kommen ans Ende der View — nach allen
      // Custom-Fields, damit sie nicht zwischen den event-spezifischen
      // Antwortspalten landen. Hier merken und im Post-Loop überspringen.
      const teamTailFields = ['TeamId', 'TeamLead', 'TeamName'];
      const teamTailSet = new Set(teamTailFields);
      // Kompletter Feld-Stand NACH dem Fix (bestehende + neu angelegte),
      // damit neu angelegte Custom-Fields auch in die View kommen.
      for (const fn of postFixFields) {
        if (alreadyAdded.has(fn)) continue;
        if (systemBlocklist.has(fn)) continue;
        if (fn.charAt(0) === '_') continue;
        if (teamTailSet.has(fn)) continue; // ans Ende
        viewFields.push(fn);
        alreadyAdded.add(fn);
      }
      // Team-Spalten jetzt am Ende anhängen (nur die, die tatsächlich existieren).
      for (const fn of teamTailFields) {
        if (alreadyAdded.has(fn)) continue;
        if (postFixFields.indexOf(fn) < 0) continue;
        viewFields.push(fn);
        alreadyAdded.add(fn);
      }

      for (const fn of viewFields) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/addviewfield('${fn}')`,
          {}
        );
      }
      viewFixed = true;
    } catch {
      console.warn('[DEX] View-Reihenfolge konnte nicht gesetzt werden');
    }

    // v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe anlegen
    // (oder seeden mit dem aktuellen Max-Wert wenn schon vorhanden).
    try {
      const counterResult = await this.ensureCounterList(subsiteUrl);
      if (counterResult.created) {
        added.push(`Counter-Liste ${COUNTER_LIST_NAME} (atomare TeilnehmerID-Vergabe, seeded mit ${counterResult.seededValue})`);
      } else if (counterResult.seededValue !== undefined) {
        added.push(`Counter-Item nachgeseedet (NextValue=${counterResult.seededValue})`);
      }
    } catch {
      console.warn('[DEX] Counter-Liste konnte nicht angelegt werden');
    }

    return { added, removed, viewFixed, customFieldMap: Object.keys(customFieldMap).length > 0 ? customFieldMap : undefined };
  }


  /**
   * Quiz-Fortschritt in die Registrierung eines Teilnehmers schreiben.
   *
   * - answers: ausgewählte Antwort-Indices pro Frage (Array von Arrays, weil
   *   Fragen mehrere richtige Antworten haben können). Unbeantwortete Fragen
   *   bleiben als leeres Array `[]` stehen, damit der Index-Offset erhalten bleibt.
   * - score: aktuell erreichte Punkte (Anzahl korrekt beantworteter Fragen).
   * - isComplete: true wenn alle Fragen beantwortet sind — dann wird auch
   *   `QuizCompletedAt` gesetzt. Andernfalls bleibt QuizCompletedAt unverändert
   *   (null/leer), sodass der Teilnehmer als "teilweise beantwortet" gelistet wird.
   *
   * Ersetzt die früher nur-am-Ende aufgerufene `saveQuizResult()`. Wird jetzt
   * bei jedem "Weiter"-Klick im QuizPlayer aufgerufen (Auto-Save), damit der
   * Teilnehmer beim späteren Wiedereintritt an derselben Stelle weitermachen kann.
   */
  public async saveQuizProgress(
    subsiteUrl: string,
    itemId: number,
    score: number,
    answers: number[][],
    isComplete: boolean
  ): Promise<boolean> {
    try {
      // Vor dem Schreiben sicherstellen, dass die Quiz-Spalten auf der
      // Teilnehmer-Liste existieren. Bei Bestandsevents (vor Quiz-Feature
      // angelegt) fehlen sie oft; _merge mit odata=nometadata schluckt
      // unbekannte Felder stumm und das Save wirkt wie "gespeichert",
      // persistiert aber nichts.
      // Silent: wenn der aktuelle User keine Manage-Lists-Permission auf
      // der Subsite hat, schlägt das Anlegen fehl (Regular User). Dann
      // kann nur ein Admin/Organizer die Spalten anlegen — dafür gibt's
      // die "Spalten fixen"-Funktion im Admin Center.
      await this.ensureQuizColumnsOnRegList(subsiteUrl);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: { [key: string]: any } = {
        'QuizScore': score,
        'QuizAnswers': JSON.stringify(answers),
      };
      if (isComplete) {
        payload.QuizCompletedAt = new Date().toISOString();
      }
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        payload
      );
      return resp.ok || resp.status === 406;
    } catch {
      return false;
    }
  }

  /**
   * Quiz-Spalten auf der Teilnehmer-Liste einer Event-Subsite anlegen,
   * falls sie fehlen. Idempotent und silent: bei fehlender Permission
   * kein Crash, einfach kein-op.
   */
  private async ensureQuizColumnsOnRegList(subsiteUrl: string): Promise<void> {
    try {
      const fieldsResp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=200&$select=InternalName`,
        SPHttpClient.configurations.v1
      );
      if (!fieldsResp.ok) return;
      const fieldsData = await fieldsResp.json();
      const existing = new Set<string>(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (fieldsData.value || fieldsData.d?.results || []).map((f: any) => f.InternalName)
      );
      const required: Array<{ title: string; type: number }> = [
        { title: 'QuizScore', type: 9 },      // Number
        { title: 'QuizAnswers', type: 3 },    // Note (multiline)
        { title: 'QuizCompletedAt', type: 4 } // DateTime
      ];
      for (const f of required) {
        if (existing.has(f.title)) continue;
        try {
          await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
            {
              '__metadata': { 'type': 'SP.Field' },
              'Title': f.title,
              'FieldTypeKind': f.type,
              'Required': false,
            }
          );
          console.warn(`[DEX] ensureQuizColumnsOnRegList: ${f.title} nachgelegt auf ${subsiteUrl}`);
        } catch {
          // keine Permission -> silent. User braucht Admin der "Spalten fixen" macht.
        }
      }
    } catch { /* silent */ }
  }

  /**
   * Ersten Warteliste-Teilnehmer nachrücken: Status -> Angemeldet.
   * Wenn inheritStarterType übergeben wird (B2Run Split-Capacity), wird dieser Typ
   * dem Nachrücker zugewiesen (er erbt den Platz des Abgemeldeten).
   *
   * Wird **client-seitig** ausgeführt (von der App beim Abmelden), damit der
   * Power Automate DEX_IDReorder-Flow keinen doppelten Nachrück-Versuch macht.
   * Liefert den nachgerückten Teilnehmer (Email + Name) zurück für die
   * Nachrück-E-Mail.
   *
   * Schutz gegen Überbuchung: Wenn maxParticipants gesetzt ist und die Anzahl
   * der aktuell Angemeldeten (nach der Abmeldung) >= maxParticipants ist, wird
   * NICHT nachgerückt. Das verhindert, dass nach einer früheren Überbuchung
   * der Abbruch der Abmeldung nicht zu einer weiteren Überbuchung führt.
   */
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
    try {
      // Überbuchungs-Schutz: Nur nachrücken, wenn tatsächlich ein Platz frei ist.
      // Bei unlimited (maxParticipants === 0 oder undefined) immer nachrücken.
      //
      // WICHTIG: '>' statt '>='. Die Abmeldung (Status->Abgemeldet) ist kurz vor
      // diesem Call passiert; falls SharePoint den Statuswechsel noch nicht in
      // getRegistrationCount reflektiert (stale read), würden wir bei einem
      // vollen Event (z.B. 128/128) mit '>=' fälschlich skippen. Mit '>' ist
      // 'registered == max' noch erlaubt (= genau ein Platz wird nachgerückt),
      // und eine echte Überbuchung (401 > 128) wird weiterhin abgefangen.
      if (maxParticipants && maxParticipants > 0) {
        const counts = await this.getRegistrationCount(subsiteUrl);
        if (counts.registered > maxParticipants) {
          console.warn(`[DEX] promoteFirstWaitlistItem: skipping promotion - event is overbooked (${counts.registered}/${maxParticipants} registered).`);
          return { success: false, skippedOverbooked: true };
        }
      }

      // v12.10: Nachrück-Sortierung jetzt nach TeilnehmerID asc statt
      // RegistrationDate. Hintergrund: nach dem IDReorder-Flow sind die
      // TeilnehmerIDs durchlaufend (1..N aktiv, N+1.. Warteliste). Wenn
      // also Platz 100 frei wird, soll TID 101 (= erster auf der Liste)
      // nachrücken — unabhängig davon, ob TID 103 zeitlich gesehen vor
      // TID 101 registriert war (z.B. nach Re-Registration oder Wechsel
      // der Gruppe). RegistrationDate sortierte chronologisch, was bei
      // umverteilten IDs zur falschen Reihenfolge führte.
      // Bei B2Run-Split-Kapazitäten: nur die passende Warteliste durchsuchen
      // (PreferredStarterType == onlyWithPreferredType).
      let filter = `Status eq 'Warteliste'`;
      if (onlyWithPreferredType) {
        const esc = onlyWithPreferredType.replace(/'/g, "''");
        filter += ` and PreferredStarterType eq '${esc}'`;
      }
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${encodeURIComponent(filter)}&$orderby=TeilnehmerID asc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return { success: false };
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return { success: false };

      const firstWaiting = items[0];
      // v18.71: Kern-Update (Status -> Angemeldet, ggf. StarterType) STRIKT
      // getrennt von den optionalen Audit-Feldern. Hintergrund: bei Legacy-
      // Teilnehmerlisten, die noch nie per „Spalten fixen" aktualisiert wurden,
      // fehlt die Spalte PromotedDate (erst seit v17.15). Wenn PromotedDate im
      // selben MERGE-Body steht, lehnt SharePoint den GESAMTEN Request mit
      // HTTP 400 ab („The property 'PromotedDate' does not exist…") — der
      // Nachrück-Status wird dann gar nicht gesetzt und der Button „macht
      // nichts". Deshalb zuerst nur die Pflichtfelder schreiben.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mergeBody: Record<string, any> = { 'Status': 'Angemeldet' };
      // v19.14 AUDIT-FIX: StarterType IMMER mitsetzen — entweder den vom
      // Abgemeldeten geerbten Typ (inheritStarterType, Admin-Cancel-Pfad) ODER,
      // falls keiner mitgegeben wurde (z.B. der manuelle „Nachrücken"-Button, der
      // inheritStarterType=undefined übergibt), den EIGENEN Wunsch der
      // nachrückenden Person (PreferredStarterType). Vorher blieb StarterType bei
      // Promotes ohne inheritStarterType auf Split-Events LEER → es entstanden
      // angemeldete „typlose" Personen (Audit-Befund: Andreas Jehle), die aus den
      // Gruppen-Zahlen fielen und als „Wunsch: …" angezeigt wurden. Bei
      // Nicht-Split-Events ist PreferredStarterType leer → StarterType bleibt leer
      // (korrekt, da es dort keine Gruppen gibt).
      const effectiveStarter = inheritStarterType || firstWaiting.PreferredStarterType || '';
      if (effectiveStarter) {
        mergeBody['StarterType'] = effectiveStarter;
      }
      const mergeResp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${firstWaiting.Id})`,
        mergeBody
      );
      if (!(mergeResp.ok || mergeResp.status === 406)) return { success: false };

      // v17.15: Nachrück-Audit auf der promoteten Person — best-effort, in
      // einem SEPARATEN MERGE, damit eine fehlende Audit-Spalte (Legacy-Liste)
      // den eigentlichen Promote oben nicht kaputtmacht.
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const auditBody: Record<string, any> = { 'PromotedDate': new Date().toISOString() };
        if (replacedByCancel && replacedByCancel.participantEmail) {
          auditBody['ReplacedParticipantEmail'] = replacedByCancel.participantEmail;
        }
        await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${firstWaiting.Id})`,
          auditBody
        );
      } catch (err) {
        console.warn('[DEX] promoteFirstWaitlistItem: Nachrück-Audit (PromotedDate) konnte nicht geschrieben werden — Spalte fehlt evtl. auf einer Legacy-Liste:', err);
      }

      const vorname = firstWaiting.Vorname || '';
      const nachname = firstWaiting.Nachname || '';
      const name = (vorname && nachname) ? `${vorname} ${nachname}` : (firstWaiting.ParticipantName || '');
      const email = firstWaiting.ParticipantEmail || firstWaiting.Title || '';

      // v17.15: zweite PATCH auf die cancelnde Person — „Ersetzt durch".
      if (replacedByCancel && replacedByCancel.itemId && email) {
        try {
          await this._merge(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${replacedByCancel.itemId})`,
            { 'ReplacedByParticipantEmail': email }
          );
        } catch (err) {
          console.warn('[DEX] Nachrück-Audit auf cancelnder Person fehlgeschlagen:', err);
        }
      }

      console.warn(`[DEX] promoteFirstWaitlistItem: promoted ${name} <${email}> (item ${firstWaiting.Id}) to Angemeldet.`);
      return { success: true, email, name, itemId: firstWaiting.Id };
    } catch {
      return { success: false };
    }
  }

  /**
   * v11.0: Item-Attachments einer Teilnehmer-Registrierung listen.
   * Liefert ein Array mit FileName + ServerRelativeUrl, sodass die App
   * Download-Links rendern kann. Subsite-spezifisch (jede Teilnehmerliste
   * lebt in der Event-Subsite).
   */
  public async listRegistrationAttachments(
    subsiteUrl: string,
    itemId: number,
  ): Promise<Array<{ fileName: string; serverRelativeUrl: string }>> {
    try {
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles`;
      const resp = await this._sp.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
      });
      if (!resp.ok) return [];
      const data = await resp.json();
      const items = data.value || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return items.map((a: any) => ({
        fileName: a.FileName || '',
        serverRelativeUrl: a.ServerRelativeUrl || '',
      })).filter((x: { fileName: string }) => !!x.fileName);
    } catch (err) {
      console.warn('[DEX] listRegistrationAttachments failed:', err);
      return [];
    }
  }

  /**
   * v11.0: PDF / Datei als Item-Attachment an eine Teilnehmer-Zeile
   * hängen. SharePoint erlaubt mehrere Attachments pro Item; bei
   * gleichem Namen wirft die API einen 409, daher prefixen wir den
   * Dateinamen mit einem Timestamp wenn die App das aufruft.
   */
  public async addRegistrationAttachment(
    subsiteUrl: string,
    itemId: number,
    file: File,
    // v19.0: optionaler Präfix, um ein Attachment einem Dokument-Custom-Field
    // zuzuordnen (z.B. 'dxf-<fieldId>--'). Leer = generischer Attendee-Upload.
    fieldPrefix: string = '',
  ): Promise<boolean> {
    try {
      const buf = await file.arrayBuffer();
      // Dateiname säubern + Timestamp-prefix für Eindeutigkeit
      const safeName = (file.name || 'upload.pdf').replace(/[^a-zA-Z0-9._-]+/g, '_');
      const ts = new Date().toISOString().replace(/[:.]/g, '-').replace(/T/, '_').slice(0, 19);
      const finalName = `${fieldPrefix}${ts}_${safeName}`;
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(finalName)}')`;
      const resp = await this._sp.post(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
        body: buf,
      });
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] addRegistrationAttachment failed:', err);
      return false;
    }
  }

  /**
   * v11.0: Item-Attachment löschen. Wird sowohl vom User (eigener
   * Upload zurückziehen) als auch vom Admin (im Admin Center) genutzt.
   */
  public async deleteRegistrationAttachment(
    subsiteUrl: string,
    itemId: number,
    fileName: string,
  ): Promise<boolean> {
    try {
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`;
      const resp = await this._sp.post(url, SPHttpClient.configurations.v1, {
        headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'Accept': 'application/json;odata=nometadata' },
      });
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] deleteRegistrationAttachment failed:', err);
      return false;
    }
  }

  /**
   * v10.27: User wechselt seine Split-Capacity-Gruppe.
   *
   * Logik:
   * - Lädt aktuelle aktive Registrierungen der Subsite und zählt die Anzahl
   *   pro StarterType.
   * - Wenn die Ziel-Gruppe noch unter ihrer Kapazität liegt: User wird mit
   *   neuem StarterType direkt als 'Angemeldet' eingetragen.
   * - Wenn die Ziel-Gruppe bereits voll ist: User wandert auf die Warteliste
   *   mit PreferredStarterType=newType. StarterType bleibt leer (wie bei
   *   Erst-Anmeldung auf Warteliste). Nachgerückt wird er erst, wenn ein
   *   Platz in der Ziel-Gruppe frei wird (siehe Power-Automate-Flow).
   *
   * Liefert { ok, status, full } zurück — die App nutzt das, um die richtige
   * Mail (Anmeldung vs. Warteliste) zu queuen und dem User Feedback zu geben.
   */
  /**
   * v11.24: Tauscht StarterType (und PreferredStarterType) bei ALLEN
   * Registrierungen einer Subsite: jeder 'Durchstarter' wird zu
   * 'Funstarter' und umgekehrt. Wird vom Admin-Center aufgerufen, wenn
   * der Organizer im Wizard die Reihenfolge der Gruppen-Labels +
   * -Kapazitäten getauscht hat — die existierenden Anmeldungen sind
   * dann technisch noch in der „alten" Slot-Bedeutung. Dieser Flip
   * synchronisiert sie mit der neuen Reihenfolge.
   *
   * Liefert die Anzahl erfolgreich aktualisierter Items zurück.
   */
  public async flipAllStarterTypes(subsiteUrl: string): Promise<{ ok: boolean; updated: number; failed: number }> {
    try {
      const all = await this.getAllRegistrations(subsiteUrl);
      let updated = 0;
      let failed = 0;
      for (const r of all) {
        const flip = (t: string | undefined): string => {
          if (t === 'Durchstarter') return 'Funstarter';
          if (t === 'Funstarter') return 'Durchstarter';
          return t || '';
        };
        const newStarter = flip(r.StarterType);
        const newPref = flip(r.PreferredStarterType);
        if (newStarter === (r.StarterType || '') && newPref === (r.PreferredStarterType || '')) continue;
        try {
          const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${r.Id})`;
          const body: Record<string, unknown> = {};
          if (newStarter !== (r.StarterType || '')) body['StarterType'] = newStarter;
          if (newPref !== (r.PreferredStarterType || '')) body['PreferredStarterType'] = newPref;
          const resp = await this._merge(url, body);
          if (resp.ok) updated++;
          else failed++;
        } catch { failed++; }
      }
      return { ok: failed === 0, updated, failed };
    } catch (err) {
      console.warn('[DEX] flipAllStarterTypes error:', err);
      return { ok: false, updated: 0, failed: 0 };
    }
  }

  public async switchSplitGroup(
    subsiteUrl: string,
    itemId: number,
    newType: 'Durchstarter' | 'Funstarter',
    durchstarterCapacity: number,
    funstarterCapacity: number,
  ): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }> {
    try {
      // v27.10 REGRESSIONS-FIX: Die frühere Zählung über getAllRegistrations
      // ist für normale User unbrauchbar, seit die Element-Sicherheit („nur
      // eigene Elemente", v26.87) wirklich greift — sie sahen nur die eigene
      // Zeile, zählten die Ziel-Gruppe als leer und wechselten an einer
      // vollen Gruppe (inkl. deren Warteliste) vorbei direkt auf „Angemeldet".
      // Stattdessen jetzt die ATOMARE Sitzplatz-Reservierung über den für
      // alle lesbaren/schreibbaren Gruppen-Counter (gleicher Mechanismus wie
      // bei der Neu-Anmeldung, ETag-CAS): 'reserved' → Platz sicher belegt,
      // 'full'/'error' → fail-closed Warteliste. Der eigene Eintrag steckt
      // bei einem echten Wechsel nie im Ziel-Gruppen-Zähler (aktiv in der
      // ANDEREN Gruppe oder auf der Warteliste).
      // Vorherigen Zustand der eigenen Zeile lesen (unter Item-Level-Security
      // immer sichtbar) — für die additive Counter-Pflege unten.
      let prevStatus = '';
      try {
        const ownResp = await this._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=Status`,
          SPHttpClient.configurations.v1
        );
        if (ownResp.ok) {
          const ownData = await ownResp.json();
          const own = ownData?.d ?? ownData;
          prevStatus = own?.Status || '';
        }
      } catch { /* best-effort — Counter-Pflege unten fällt dann konservativ aus */ }
      const wasActive = EventService.ACTIVE_STATI.indexOf(prevStatus) >= 0;
      const wasWaitlist = prevStatus === 'Warteliste';
      const targetCap = newType === 'Durchstarter' ? durchstarterCapacity : funstarterCapacity;
      const seat = targetCap > 0
        ? await this.reserveSeat(subsiteUrl, newType, targetCap)
        : 'reserved'; // cap <= 0 = unbegrenzt
      const goWaitlist = seat !== 'reserved';
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = goWaitlist
        ? { 'Status': 'Warteliste', 'StarterType': '', 'PreferredStarterType': newType }
        : { 'Status': 'Angemeldet', 'StarterType': newType, 'PreferredStarterType': newType };
      // v22.20: FIFO-Fairness — wer auf die Warteliste der Zielgruppe wechselt,
      // verliert seine alte (niedrige) TeilnehmerID und reiht sich mit einer
      // frischen Counter-ID HINTEN ein. Sonst würde er beim typ-bewussten
      // Nachrücken (TeilnehmerID asc) alle überholen, die schon länger warten.
      // Best-effort: schlägt der Counter fehl, bleibt die alte ID stehen und der
      // anschließende Reorder-Lauf normalisiert wenigstens die Nummerierung.
      if (goWaitlist) {
        try {
          const freshId = await this.getNextTeilnehmerId(subsiteUrl);
          if (typeof freshId === 'number' && freshId > 0) body['TeilnehmerID'] = freshId;
        } catch { /* alte ID behalten */ }
      }
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
      const resp = await this._merge(url, body);
      if (!resp.ok) {
        console.warn('[DEX] switchSplitGroup MERGE failed:', resp.status);
        // Reservierten Ziel-Gruppen-Platz zurückgeben — die Zeile wurde ja
        // nicht umgestellt.
        if (!goWaitlist && targetCap > 0) {
          await this.adjustSeatCounterField(subsiteUrl, this.seatFieldFor(newType), -1);
        }
        return { ok: false, status: 'Failed', full: goWaitlist };
      }
      // v27.10: Additive Counter-Pflege (ILS-sicher, best-effort). Nur die
      // eindeutig sicheren Anpassungen — alles Unklare heilt der nächste
      // privilegierte Reconcile:
      // - Aktiv → Warteliste: WaitlistTaken +1 (der frei werdende Quell-Slot
      //   wird bewusst NICHT dekrementiert — fail-closed, Nachrücken/Reconcile
      //   übernimmt).
      // - Warteliste → Aktiv: WaitlistTaken −1 (Ziel-Gruppen-Zähler hat
      //   reserveSeat bereits atomar erhöht).
      if (goWaitlist && wasActive) {
        await this.adjustWaitlistCounter(subsiteUrl, +1);
      } else if (!goWaitlist && wasWaitlist) {
        await this.adjustWaitlistCounter(subsiteUrl, -1);
      }
      // Aktiv → Aktiv (Gruppenwechsel): Quell-Gruppen-Zähler bleibt bewusst
      // stehen (fail-closed) — der nächste privilegierte Reconcile setzt ihn
      // exakt.
      return { ok: true, status: goWaitlist ? 'Warteliste' : 'Angemeldet', full: goWaitlist };
    } catch (err) {
      console.warn('[DEX] switchSplitGroup error:', err);
      return { ok: false, status: 'Failed', full: false };
    }
  }

  /**
   * Registrierung stornieren
   */
  public async cancelRegistration(
    subsiteUrl: string,
    itemId: number,
    cancelledByName?: string,
    cancelledByEmail?: string
  ): Promise<boolean> {
    try {
      // Audit: wer hat die Abmeldung ausgelöst?
      // Bei Self-Cancel = der User selbst. Bei "Teilnehmer abmelden" durch den
      // Organizer/Admin im Admin Center = der eingeloggte Organizer/Admin.
      const auditName = cancelledByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (cancelledByEmail || this.context.pageContext.user.email || '').toLowerCase();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const corePayload: Record<string, any> = {
        'Status': 'Abgemeldet',
        'CancellationDate': new Date().toISOString(),
        'TeilnehmerID': null,
      };
      // Audit-Felder optional dazu - aeltere Subsites haben die Spalten evtl. noch
      // nicht (kommt erst mit Commit a10a608). Ein 400 von SP würde dann die
      // ganze Abmeldung blocken. Strategie: erst mit Audit-Feldern versuchen,
      // bei Misserfolg ohne sie nochmal probieren.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fullPayload: Record<string, any> = { ...corePayload };
      if (auditName) fullPayload['CancelledByName'] = auditName;
      if (auditEmail) fullPayload['CancelledByEmail'] = auditEmail;
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
      let response = await this._merge(url, fullPayload);
      if (!response.ok && (auditName || auditEmail)) {
        // Fallback ohne Audit-Felder (Subsite-Liste hat die Spalten noch nicht)
        console.warn('[DEX] cancelRegistration with audit failed (' + response.status + '), retrying without audit fields');
        response = await this._merge(url, corePayload);
      }
      // v7.31: Counter mit aktuellem Max syncen, damit er nicht "davonrast"
      // wenn der höchste ID-Inhaber sich abmeldet. Best-effort, blockiert
      // die Abmeldung nicht wenn's fehlschlägt. syncCounterToMax liest den
      // Max-Wert intern frisch (race-frei gegen parallele Anmeldungen).
      if (response.ok) {
        try { await this.syncCounterToMax(subsiteUrl); } catch { /* */ }
        // v24.41: Koordinations-Liste synchron halten — bei JEDER Abmeldung den
        // zugehörigen Assistenz-Link (falls vorhanden) auf 'Cancelled' setzen,
        // damit die Info bei der anderen Seite verschwindet. Deckt alle Cancel-
        // Pfade ab (Self / Proxy / Admin / Team), weil sie hier durchlaufen.
        try { await this.setAssistantLinkStatusForRegistration(itemId, subsiteUrl, 'Cancelled'); } catch { /* */ }
      }
      return response.ok;
    } catch (err) {
      console.warn('[DEX] cancelRegistration error:', err);
      return false;
    }
  }

  /**
   * v18.11: Proaktive Absage durch einen Teilnehmer, der sich NICHT angemeldet
   * hat („Ich nehme nicht teil"). Legt eine Teilnehmer-Zeile direkt mit
   * Status='Abgemeldet' an — KEINE Sitzplatz-Reservierung, KEINE TeilnehmerID.
   * Profil-Daten (Vorname/Nachname/Location/JobTitle/Department) werden geladen,
   * damit die Abmeldungs-Liste im Admin-Center dieselben Spalten füllen kann
   * wie Teilnehmer-/Warteliste. Marker `_declined` in CustomData unterscheidet
   * die proaktive Absage von einer regulären Abmeldung (die nach vorheriger
   * Anmeldung erfolgte).
   */
  public async declineRegistration(
    subsiteUrl: string,
    firstName: string,
    surname: string,
    participantEmail: string,
    actorName?: string,
    actorEmail?: string
  ): Promise<boolean> {
    try {
      const myEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const profile = (participantEmail || '').toLowerCase() === myEmail
        ? await this.getCurrentUserProfile()
        : await this.getUserProfileByEmail(participantEmail);
      const nowIso = new Date().toISOString();
      // v22.57: Claims-Token-Schutz. In manchen Kontexten liefert der Browser
      // als „displayName" das SharePoint-Claims-Login (z.B.
      // „i:0#.f|membership|user@deloitte.de" bzw. „0#.f|membership|…"). Das
      // landete bisher 1:1 als Vorname in der Absage-Zeile. Wir verwenden
      // deshalb bevorzugt den sauberen Namen aus dem Benutzerprofil und filtern
      // Claims-artige Werte raus.
      const looksLikeClaim = (s: string): boolean => /\|membership\||0#\.f\||^i:0#/i.test((s || '').trim());
      const cleanFirst = looksLikeClaim(firstName) ? '' : (firstName || '').trim();
      const cleanLast = looksLikeClaim(surname) ? '' : (surname || '').trim();
      const effFirst = (profile.firstName || '').trim() || cleanFirst;
      const effLast = (profile.lastName || '').trim() || cleanLast;
      // Anzeigename: bevorzugt Profil-PreferredName, sonst Vor-/Nachname,
      // sonst die E-Mail (nie das Claims-Token).
      const effName = (profile.displayName && !looksLikeClaim(profile.displayName) ? profile.displayName : '')
        || `${effFirst} ${effLast}`.trim()
        || participantEmail;
      const auditNameRaw = actorName || this.context.pageContext.user.displayName || '';
      const auditName = looksLikeClaim(auditNameRaw) ? effName : auditNameRaw;
      const auditEmail = (actorEmail || this.context.pageContext.user.email || '').toLowerCase();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        'Vorname': effFirst,
        'Nachname': effLast,
        'ParticipantName': effName,
        'ParticipantEmail': participantEmail,
        'Department': profile.department,
        'Location': profile.location,
        'JobTitle': profile.jobTitle,
        'Phone': profile.phone,
        'Company': profile.company,
        'Status': 'Abgemeldet',
        'RegistrationDate': nowIso,
        'CancellationDate': nowIso,
        // Marker: proaktive Absage (nie angemeldet gewesen).
        'CustomData': JSON.stringify({ _declined: 'true' }),
      };
      if (auditName) { payload['RegisteredByName'] = auditName; payload['CancelledByName'] = auditName; }
      if (auditEmail) { payload['RegisteredByEmail'] = auditEmail; payload['CancelledByEmail'] = auditEmail; }
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`;
      let resp = await this._post(url, payload);
      if (!resp.ok && (auditName || auditEmail || payload['Company'])) {
        // Fallback ohne Audit-Felder (alte Subsite-Liste ohne diese Spalten).
        // v24.32: zusätzlich Company strippen — fehlt die Spalte, würde der
        // Insert sonst auch im Fallback an Company scheitern.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const noAudit: Record<string, any> = { ...payload };
        delete noAudit['RegisteredByName']; delete noAudit['RegisteredByEmail'];
        delete noAudit['CancelledByName']; delete noAudit['CancelledByEmail'];
        delete noAudit['Company'];
        resp = await this._post(url, noAudit);
      }
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] declineRegistration error:', err);
      return false;
    }
  }

  /**
   * Teilnehmer einchecken (Status auf 'Eingecheckt' setzen).
   * v7.16: Erfasst zusätzlich, WANN und VON WEM der Check-In ausgelöst
   * wurde (CheckedInDate / CheckedInByName / CheckedInByEmail). Diese
   * Spalten werden bei neuen Events über createRegistrationList() automatisch
   * angelegt; für bestehende Events muss der Admin einmalig die Kachel
   * "Spalten fixen" im Admin-Center klicken, damit der Check-In nicht mit
   * HTTP 400 fehlschlägt.
   */
  public async checkInParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const me = this.context.pageContext.user;
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        {
          'Status': 'Eingecheckt',
          'CheckedInDate': new Date().toISOString(),
          'CheckedInByName': me.displayName || '',
          'CheckedInByEmail': me.email || me.loginName || '',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * v23.28/v23.29: Teilnehmer als „No-Show" markieren (war angemeldet, aber
   * nicht erschienen). Reuse der Check-in-Audit-Spalten (CheckedInBy*), damit
   * kein neues Schema nötig ist. **Nur für Events, deren Teilnehmerliste die
   * 'No-Show'-Choice kennt** (= ab v23.28 NEU angelegte Events). Bestehende
   * Events werden bewusst NICHT automatisch migriert — dort lehnt SharePoint
   * den Wert ab (HTTP 400) und die Methode liefert `false`.
   */
  public async markNoShowParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const me = this.context.pageContext.user;
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        {
          'Status': 'No-Show',
          'CheckedInDate': new Date().toISOString(),
          'CheckedInByName': me.displayName || '',
          'CheckedInByEmail': me.email || me.loginName || '',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Teilnehmer auschecken (Status zurück auf 'Angemeldet' setzen)
   */
  public async checkOutParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Status': 'Angemeldet' }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Status eines Teilnehmers auf 'QR versendet' setzen
   */
  public async setQRSentStatus(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Status': 'QR versendet' }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Registrierung per Email auf einer Subsite finden
   */
  public async getRegistrationByEmail(
    subsiteUrl: string,
    email: string
  ): Promise<SPRegistration | null> {
    try {
      const response = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,ParticipantName,ParticipantEmail,Status,RegistrationDate,RegisteredByName,RegisteredByEmail,CancellationDate,CancelledByName,CancelledByEmail,CustomData,Department,JobTitle,Location,Company&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      if (data.value && data.value.length > 0) return data.value[0];
      return null;
    } catch {
      return null;
    }
  }

  /**
   * Aktuelle Teilnehmeranzahl ermitteln
   */
  public async getRegistrationCount(subsiteUrl: string): Promise<{ registered: number; waitlist: number }> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allItems: any[] = [];
    // $top=5000 (SP-REST-Maximum) statt 500 — sonst werden bei Events mit
    // ≥500 Einträgen die Counts auf den Event-Karten falsch berechnet,
    // weil SharePoint bei $orderby+$top mit ILS nicht zuverlässig nextLink
    // liefert wenn die Page exakt voll ist.
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status&$top=5000`;

    while (url) {
      try {
        const response = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }

    const registered = allItems.filter((i: { Status: string }) => i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt').length;
    const waitlist = allItems.filter((i: { Status: string }) => i.Status === 'Warteliste').length;
    return { registered, waitlist };
  }

  /**
   * v22.74: Aktive + Warteliste-E-Mails einer Teilnehmerliste (lowercase) —
   * für die EINDEUTIGE Personenzählung einer Klammer über alle Sub-Events
   * (eine Person, die sich für mehrere Sub-Events anmeldet, zählt einmal).
   */
  public async getParticipantEmailsByStatus(subsiteUrl: string): Promise<{ active: string[]; waitlist: string[] }> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allItems: any[] = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,Status&$top=5000`;
    while (url) {
      try {
        const response = await this._sp.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    const active: string[] = [];
    const waitlist: string[] = [];
    for (const i of allItems) {
      // v23.3: Zeilen OHNE (gültige) ParticipantEmail wurden frueher komplett
      // übersprungen — dadurch zaehlte die entdoppelte Klammer-/Kachel-Zahl
      // weniger Koepfe als die Tabelle Zeilen hat (z.B. 188 statt 190). Eine
      // Anmeldung ohne E-Mail ist trotzdem ein realer Kopf (belegt einen Platz),
      // bekommt nur keine Mails. Deshalb als eigener Schlüssel (Zeilen-Id)
      // mitzaehlen, statt sie zu verschlucken.
      const email = (i.ParticipantEmail || '').toLowerCase().trim();
      const key = email || `__noemail#${i.Id}`;
      if (i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt') active.push(key);
      else if (i.Status === 'Warteliste') waitlist.push(key);
    }
    return { active, waitlist };
  }

  /**
   * Title-Feld (= Teilnehmer-ID) aktualisieren
   */
  public async updateRegistrationTitle(
    subsiteUrl: string,
    itemId: number,
    newTitle: string
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Title': newTitle }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== Bild-Upload ====================

  /**
   * SiteAssets-Unterordner sicherstellen:
   * - DEX_EventImages (Event-Bilder)
   * - DEX_Logos (Deloitte-Logo für E-Mail-Templates, manuell hochgeladen)
   */
  public async ensureAssetsFolders(): Promise<void> {
    const baseUrl = this.context.pageContext.web.serverRelativeUrl;
    const folders = ['DEX_EventImages', 'DEX_Logos'];

    for (const folder of folders) {
      const folderUrl = `${baseUrl}/SiteAssets/${folder}`;
      try {
        const check = await this._sp.get(
          `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')`,
          SPHttpClient.configurations.v1
        );
        if (check.ok) continue;
      } catch { /* */ }

      try {
        await this._post(`${this.siteUrl}/_api/web/folders`, {
          '__metadata': { 'type': 'SP.Folder' },
          'ServerRelativeUrl': folderUrl,
        });
        // Ordner erstellt
      } catch {
        console.warn(`[DEX] Konnte ${folder} Ordner nicht erstellen`);
      }
    }
  }

  // ==================== v26.50: Logo & Branding (Admin-Center) ====================
  // Zentrale Ablage des Default-Logos (PNG) + Logo-Videos. Das PNG lebt in der
  // _Config-Zeile von DEX_EmailTemplates (LogoBase64) — ALLE neu versendeten
  // Mails nutzen es automatisch — und wird zusätzlich als
  // SiteAssets/DEX_Logos/Deloitte_Logo.png gespiegelt (Fallback-Pfad von
  // loadLogosAsBase64). Das Video liegt als SiteAssets/DEX_Logos/dex-logo-video.*.

  private static readonly BRANDING_VIDEO_BASENAME = 'dex-logo-video';

  /** Aktuelles Branding: Deloitte-Logo + DEX-Orb (Data-URIs) + Video-URL
   *  (leer wenn keins da). v26.58: orbBase64 ergänzt — das eigentliche
   *  DEX-Logo (bunter Ring, _Config.DefaultImageBase64 bzw. dex-orb.png);
   *  LogoBase64 ist das Deloitte-Logo der E-Mail-Kopfzeile. */
  public async getBranding(): Promise<{ logoBase64: string; orbBase64: string; videoUrl: string; videoFileName: string }> {
    let logoBase64 = '';
    let orbBase64 = '';
    try {
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id,LogoBase64,DefaultImageBase64`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const d = await resp.json();
        const it = (d.value || d.d?.results || [])[0];
        if (it && it.LogoBase64) logoBase64 = String(it.LogoBase64);
        if (it && it.DefaultImageBase64) orbBase64 = String(it.DefaultImageBase64);
      }
    } catch { /* */ }
    if (!logoBase64) {
      try { logoBase64 = await this.loadFileAsBase64('DEX_Logos/Deloitte_Logo.png'); } catch { /* */ }
    }
    if (!orbBase64) {
      try { orbBase64 = await this.loadFileAsBase64('DEX_Logos/dex-orb.png'); } catch { /* */ }
    }
    // Video: feste Kandidaten-Namen prüfen (mp4 bevorzugt).
    let videoUrl = '';
    let videoFileName = '';
    const serverRel = this.context.pageContext.web.serverRelativeUrl;
    for (const ext of ['mp4', 'webm', 'mov']) {
      const name = `${EventService.BRANDING_VIDEO_BASENAME}.${ext}`;
      try {
        const check = await this._sp.get(
          `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos/${name}')?$select=Exists,ServerRelativeUrl`,
          SPHttpClient.configurations.v1
        );
        if (check.ok) {
          videoUrl = `${new URL(this.siteUrl).origin}${serverRel}/SiteAssets/DEX_Logos/${name}`;
          videoFileName = name;
          break;
        }
      } catch { /* nächster Kandidat */ }
    }
    return { logoBase64, orbBase64, videoUrl, videoFileName };
  }

  /** v26.58: Neues DEX-Logo (Orb, PNG als Data-URI) speichern:
   *  _Config.DefaultImageBase64 (Default-Mail-Bild / {{ORB_URL}}-Fallback)
   *  + Spiegelung nach SiteAssets/DEX_Logos/dex-orb.png. */
  public async saveBrandingOrb(orbDataUri: string): Promise<boolean> {
    if (!orbDataUri || orbDataUri.indexOf('data:image/') !== 0) return false;
    let ok = false;
    try {
      const listName = 'DEX_EmailTemplates';
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      let cfgId = 0;
      if (resp.ok) {
        const d = await resp.json();
        const it = (d.value || d.d?.results || [])[0];
        if (it) cfgId = Number(it.Id);
      }
      if (cfgId > 0) {
        const m = await this._merge(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${cfgId})`, { 'DefaultImageBase64': orbDataUri });
        ok = m.ok;
      } else {
        const c = await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailTemplatesListItem' },
          'Title': '_Config', 'TemplateType': '_Config', 'DefaultImageBase64': orbDataUri,
        });
        ok = c.ok;
      }
    } catch (err) { console.warn('[DEX] saveBrandingOrb (_Config) failed:', err); }
    // Spiegel nach SiteAssets (best-effort — Fallback-Pfad + Download-Quelle).
    try {
      await this.ensureAssetsFolders();
      const b64 = orbDataUri.split(',')[1] || '';
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const serverRel = this.context.pageContext.web.serverRelativeUrl;
      await this._sp.post(
        `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos')/Files/add(url='dex-orb.png',overwrite=true)`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' }, body: bytes.buffer as ArrayBuffer }
      );
    } catch (err) { console.warn('[DEX] saveBrandingOrb (SiteAssets mirror) failed:', err); }
    return ok;
  }

  /** Neues Default-Logo (PNG als Data-URI) speichern: _Config.LogoBase64 (Mails)
   *  + Spiegelung nach SiteAssets/DEX_Logos/Deloitte_Logo.png (Fallback). */
  public async saveBrandingLogo(logoDataUri: string): Promise<boolean> {
    if (!logoDataUri || logoDataUri.indexOf('data:image/') !== 0) return false;
    let ok = false;
    try {
      // _Config-Zeile finden bzw. anlegen, dann MERGE.
      const listName = 'DEX_EmailTemplates';
      const resp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      let cfgId = 0;
      if (resp.ok) {
        const d = await resp.json();
        const it = (d.value || d.d?.results || [])[0];
        if (it) cfgId = Number(it.Id);
      }
      if (cfgId > 0) {
        const m = await this._merge(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${cfgId})`, { 'LogoBase64': logoDataUri });
        ok = m.ok;
      } else {
        const c = await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailTemplatesListItem' },
          'Title': '_Config', 'TemplateType': '_Config', 'LogoBase64': logoDataUri,
        });
        ok = c.ok;
      }
    } catch (err) { console.warn('[DEX] saveBrandingLogo (_Config) failed:', err); }
    // Spiegel nach SiteAssets (best-effort — Fallback-Pfad + Download-Quelle).
    try {
      await this.ensureAssetsFolders();
      const b64 = logoDataUri.split(',')[1] || '';
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const serverRel = this.context.pageContext.web.serverRelativeUrl;
      await this._sp.post(
        `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos')/Files/add(url='Deloitte_Logo.png',overwrite=true)`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' }, body: bytes.buffer as ArrayBuffer }
      );
    } catch (err) { console.warn('[DEX] saveBrandingLogo (SiteAssets mirror) failed:', err); }
    return ok;
  }

  /** Neues Logo-Video nach SiteAssets/DEX_Logos hochladen (fester Name,
   *  overwrite). Liefert die absolute URL oder '' bei Fehler. */
  public async uploadBrandingVideo(file: File): Promise<string> {
    try {
      await this.ensureAssetsFolders();
      const ext = (file.name.split('.').pop() || 'mp4').toLowerCase();
      const safeExt = ['mp4', 'webm', 'mov'].indexOf(ext) >= 0 ? ext : 'mp4';
      const name = `${EventService.BRANDING_VIDEO_BASENAME}.${safeExt}`;
      const serverRel = this.context.pageContext.web.serverRelativeUrl;
      const buf = await file.arrayBuffer();
      const resp = await this._sp.post(
        `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRel}/SiteAssets/DEX_Logos')/Files/add(url='${name}',overwrite=true)`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' }, body: buf }
      );
      if (!resp.ok) return '';
      return `${new URL(this.siteUrl).origin}${serverRel}/SiteAssets/DEX_Logos/${name}`;
    } catch (err) { console.warn('[DEX] uploadBrandingVideo failed:', err); return ''; }
  }

  /**
   * Event-Bild als Attachment an ein DEX_Events-Item anhängen.
   * Löscht zuerst alle bestehenden Bild-Attachments (Präfix __eventimage__),
   * dann wird das neue Bild hochgeladen. Liefert die ServerRelativeUrl als absolute URL.
   * Vorteil: keine SiteAssets-Berechtigungen nötig - wer das Item editieren darf,
   * darf auch Attachments hinzufügen.
   */
  public async uploadEventImageAsAttachment(eventId: number, file: File): Promise<string> {
    const IMAGE_PREFIX = '__eventimage__';
    try {
      // 1. Bestehende Bild-Attachments löschen (nur __eventimage__-Präfixe)
      try {
        const listResp = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
          SPHttpClient.configurations.v1
        );
        if (listResp.ok) {
          const listData = await listResp.json();
          const files = listData.value || listData.d?.results || [];
          for (const f of files) {
            const fn: string = f.FileName || '';
            if (fn.indexOf(IMAGE_PREFIX) === 0) {
              try {
                await this._delete(
                  `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fn)}')`
                );
              } catch { /* ignore */ }
            }
          }
        }
      } catch { /* ignore */ }

      // 2. Neues Bild hochladen mit Präfix
      const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
      const safeName = `${IMAGE_PREFIX}${Date.now().toString(36)}.${ext}`;

      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
          body: file,
        } as ISPHttpClientOptions
      );

      if (response.ok) {
        const data = await response.json();
        const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
        if (relUrl) return `${window.location.origin}${relUrl}`;
        // Hochgeladen, aber ohne URL in der Antwort: Pfad ist bekannt und die
        // Datei liegt dort — hier darf geraten werden.
        const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
        return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${safeName}`;
      }
      // v29.34: Bei einer FEHLER-Antwort keine URL mehr raten. Die geratene
      // Adresse zeigte auf eine Datei, die nie ankam — gespeichert wurde sie
      // trotzdem, und das Bild fehlte danach dauerhaft. Leer heißt hier
      // „fehlgeschlagen"; die Aufrufer melden das dem Organizer.
      console.warn('[DEX] Image attachment upload status:', response.status);
    } catch (err) {
      console.warn('[DEX] uploadEventImageAsAttachment error:', err);
    }
    return '';
  }

  // v28.11: Präfix des UNBESCHNITTENEN Original-Bilds (bewusst KEIN
  // '__eventimage__'-Präfix-Match, sonst würde der normale Bild-Upload es
  // mitlöschen). Wird nur gespeichert, wenn ein Querformat-Original per
  // App-Zuschnitt rund/quadratisch wurde — die Anmeldeseite zeigt dann
  // lieber das Original im Querformat-Slot.
  private static readonly ORIG_IMAGE_PREFIX = '__eventimgorig__';

  /** v28.11: Alle Original-Bild-Attachments eines Events löschen (best-effort). */
  public async deleteEventOrigImageAttachment(eventId: number): Promise<void> {
    try {
      const listResp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
        SPHttpClient.configurations.v1
      );
      if (!listResp.ok) return;
      const listData = await listResp.json();
      const files = listData.value || listData.d?.results || [];
      for (const f of files) {
        const fn: string = f.FileName || '';
        if (fn.indexOf(EventService.ORIG_IMAGE_PREFIX) === 0) {
          try {
            await this._delete(
              `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fn)}')`
            );
          } catch { /* ignore */ }
        }
      }
    } catch { /* ignore */ }
  }

  /** v28.11: Unbeschnittenes Original-Bild als zweites Attachment speichern.
   *  Löscht vorher bestehende Originale; liefert die absolute URL. */
  public async uploadEventOrigImageAsAttachment(eventId: number, file: File): Promise<string> {
    try {
      await this.deleteEventOrigImageAttachment(eventId);
      const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
      const safeName = `${EventService.ORIG_IMAGE_PREFIX}${Date.now().toString(36)}.${ext}`;
      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
          body: file,
        } as ISPHttpClientOptions
      );
      if (response.ok) {
        const data = await response.json();
        const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
        if (relUrl) return `${window.location.origin}${relUrl}`;
        const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
        return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${safeName}`;
      }
      // v29.34: Bei Fehler-Antwort NICHT raten — die geratene URL landete in
      // `_imageOrigUrl`, und die Kachel bevorzugt diese Adresse vor dem
      // Event-Bild. Ergebnis war eine weiße Kachel bei einem Event, das im
      // Organizer Center sein Bild hatte. Leer lässt den bisherigen Wert
      // stehen (der Aufrufer patcht nur bei nicht-leerem Ergebnis).
      console.warn('[DEX] Orig image attachment upload status:', response.status);
    } catch (err) {
      console.warn('[DEX] uploadEventOrigImageAsAttachment error:', err);
    }
    return '';
  }

  /** v28.11: EINEN Schlüssel im EmailTemplateOverrides-JSON eines Events
   *  patchen (read-modify-write). Leerer Wert entfernt den Schlüssel.
   *  Nötig für Werte, die erst NACH dem Item-Save bekannt sind (z.B. die
   *  Attachment-URL des Original-Bilds). */
  public async patchEventOverridesKey(eventId: number, key: string, value: string): Promise<void> {
    try {
      const getResp = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=EmailTemplateOverrides`,
        SPHttpClient.configurations.v1
      );
      if (!getResp.ok) return;
      const data = await getResp.json();
      const raw = data.d?.EmailTemplateOverrides || data.EmailTemplateOverrides || '';
      let obj: Record<string, unknown> = {};
      try { obj = raw ? JSON.parse(raw) : {}; } catch { obj = {}; }
      if (value) obj[key] = value; else delete obj[key];
      await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        { 'EmailTemplateOverrides': JSON.stringify(obj) }
      );
    } catch (err) {
      console.warn('[DEX] patchEventOverridesKey fehlgeschlagen:', key, err);
    }
  }

  /**
   * v28.38: Wie `patchEventOverridesKey`, aber für beliebige JSON-Werte
   * (Arrays, Booleans, Objekte). `undefined`/`null`/leeres Array löschen den
   * Schlüssel, damit die Overrides nicht mit leeren Huellen zuwachsen.
   * Liefert true bei Erfolg.
   */
  public async patchEventOverridesValue(eventId: number, key: string, value: unknown): Promise<boolean> {
    const res = await this.patchEventOverridesValueEx(eventId, key, value);
    return res.ok;
  }

  /**
   * v28.60: Wie `patchEventOverridesValue`, liefert aber den Grund mit.
   *
   * Vorher gab es nur true/false — bei einem Fehlschlag stand im UI „konnte
   * nicht gespeichert werden" und sonst nichts, was die Ursachensuche
   * unmöglich machte. Jetzt kommt der HTTP-Status samt SharePoint-Meldung
   * zurück, und die drei typischen Stolpersteine sind abgefangen:
   *
   *  - **Transiente Fehler** (429 Throttling, 5xx): ein Wiederholungsversuch
   *    nach kurzer Pause statt sofort aufzugeben.
   *  - **Grössenlimit**: SharePoint lehnt Requests über 2 MB ab. Das Feld
   *    trägt bei Events mit eingebetteten Logos einiges — wir prüfen vorher
   *    und sagen es klar, statt in ein nacktes HTTP 400 zu laufen.
   *  - **Parallele Schreibvorgänge**: Der Aufruf ist ein Read-Modify-Write.
   *    Zwei gleichzeitige Aufrufe (z.B. schnell hintereinander geklickte
   *    Löschungen) würden sich gegenseitig überschreiben, deshalb laufen sie
   *    über `_ovQueue` streng nacheinander.
   */
  public async patchEventOverridesValueEx(
    eventId: number, key: string, value: unknown,
  ): Promise<{ ok: boolean; status: number; detail: string }> {
    const run = async (): Promise<{ ok: boolean; status: number; detail: string }> => {
      const attempt = async (): Promise<{ ok: boolean; status: number; detail: string }> => {
        const getResp = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=EmailTemplateOverrides`,
          SPHttpClient.configurations.v1
        );
        if (!getResp.ok) {
          return { ok: false, status: getResp.status, detail: `Lesen fehlgeschlagen (HTTP ${getResp.status})` };
        }
        const data = await getResp.json();
        const raw = data.d?.EmailTemplateOverrides || data.EmailTemplateOverrides || '';
        let obj: Record<string, unknown> = {};
        try { obj = raw ? JSON.parse(raw) : {}; } catch { obj = {}; }
        const empty = value === undefined || value === null || value === false
          || (Array.isArray(value) && value.length === 0);
        if (empty) { delete obj[key]; } else { obj[key] = value; }
        const payload = JSON.stringify(obj);
        if (payload.length > 1_900_000) {
          return {
            ok: false, status: 413,
            detail: `Der Event-Datensatz ist mit ${(payload.length / 1048576).toFixed(2)} MB zu gross für einen Schreibvorgang (Limit 2 MB).`,
          };
        }
        const resp = await this._merge(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
          { 'EmailTemplateOverrides': payload }
        );
        if (resp.ok || resp.status === 406) return { ok: true, status: resp.status, detail: '' };
        let body = '';
        try { body = await resp.text(); } catch { /* egal */ }
        // SharePoint verpackt die Meldung in {error:{message:{value}}}.
        let msg = '';
        try {
          const j = JSON.parse(body);
          msg = j?.error?.message?.value || j?.['odata.error']?.message?.value || '';
        } catch { msg = (body || '').substring(0, 200); }
        return { ok: false, status: resp.status, detail: msg || `HTTP ${resp.status}` };
      };

      for (let i = 0; i < 2; i++) {
        try {
          const r = await attempt();
          if (r.ok) return r;
          const transient = r.status === 429 || r.status >= 500;
          if (!transient || i === 1) {
            console.warn('[DEX] patchEventOverridesValue fehlgeschlagen:', key, r.status, r.detail);
            return r;
          }
        } catch (err) {
          if (i === 1) {
            console.warn('[DEX] patchEventOverridesValue Ausnahme:', key, err);
            return { ok: false, status: 0, detail: String((err as Error)?.message || err) };
          }
        }
        await new Promise<void>(res => setTimeout(res, 600));
      }
      return { ok: false, status: 0, detail: 'unbekannt' };
    };

    // Streng nacheinander — sonst gehen parallele Read-Modify-Writes verloren.
    const next = this._ovQueue.then(run, run);
    this._ovQueue = next.then(() => undefined, () => undefined);
    return next;
  }

  /** Serialisiert die Overrides-Schreibvorgänge (siehe oben). */
  private _ovQueue: Promise<void> = Promise.resolve();

  /**
   * EventImageUrl-Feld eines DEX_Events-Items setzen (kleines MERGE).
   */
  public async updateEventImageUrl(eventId: number, url: string): Promise<boolean> {
    try {
      const resp = await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        { 'EventImageUrl': url }
      );
      return resp.ok || resp.status === 406;
    } catch {
      return false;
    }
  }

  /**
   * Dokument als Attachment an ein DEX_Events-Item anfügen.
   * Nutzt native SharePoint List Item Attachments - keine Ordner nötig.
   */
  public async uploadEventDocument(eventId: number, file: File): Promise<string> {
    try {
      const fileName = file.name.replace(/[#%&*:<>?/\\|]/g, '_');

      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
          body: file,
        } as ISPHttpClientOptions
      );

      if (response.ok) {
        const data = await response.json();
        const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
        if (relUrl) return `${window.location.origin}${relUrl}`;
      } else {
        console.warn('[DEX] Attachment upload status:', response.status);
      }

      // Fallback: URL aus bekanntem Pfad
      const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
      return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${fileName}`;
    } catch (err) {
      console.warn('[DEX] uploadEventDocument error:', err);
    }
    return '';
  }

  /**
   * Dokument-Attachment von einem DEX_Events-Item löschen.
   * Wird beim Edit verwendet, wenn der User ein bestehendes Dokument entfernt.
   */
  public async deleteEventDocument(eventId: number, fileName: string): Promise<boolean> {
    try {
      const resp = await this._delete(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`
      );
      return resp.ok || resp.status === 200 || resp.status === 204;
    } catch (err) {
      console.warn('[DEX] deleteEventDocument error:', err);
      return false;
    }
  }

  /**
   * Attachments eines DEX_Events-Items laden.
   * Bilder mit Präfix __eventimage__ werden ausgefiltert (nur für EventImageUrl).
   */
  public async getEventAttachments(eventId: number): Promise<Array<{ name: string; url: string; size: number }>> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        const files = data.value || data.d?.results || [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return files
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .filter((f: any) => (f.FileName || '').indexOf('__eventimage__') !== 0)
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .map((f: any) => ({
            name: f.FileName || '',
            url: `${window.location.origin}${f.ServerRelativeUrl || ''}`,
            size: 0,
          }));
      }
    } catch { /* Attachments nicht verfügbar */ }
    return [];
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

  public async repairClaimNamesInRegistrations(
    subsiteUrl: string
  ): Promise<{ scanned: number; hits: number; fixed: number; failed: number }> {
    return profileData.repairClaimNamesInRegistrations(this, subsiteUrl);
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

  // v7.28: Atomare TeilnehmerID-Vergabe über die DEX_TeilnehmerCounter-Liste
  // pro Event-Subsite. Verhindert Race-Conditions wenn viele User gleichzeitig
  // anmelden — ohne das vorher passieren konnte, dass zwei User dieselbe ID
  // bekommen (siehe Bug-Report v7.27 → v7.28).
  //
  // Ablauf:
  //   1. Counter-Item GET'en, ETag aus Response-Header lesen.
  //   2. NextValue + 1 mit IF-MATCH: <etag> via MERGE schreiben.
  //   3. Bei 412 (ETag-Mismatch = jemand war schneller) → kurzes Jitter +
  //      Retry, max 8x.
  //   4. Bei Erfolg: NextValue+1 ist die neue TeilnehmerID des aufrufenden Users.
  //
  // Fallback: Wenn die Counter-Liste nicht existiert (z.B. legacy event ohne
  // "Spalten fixen"-Lauf), kommt undefined zurück — der Aufrufer fällt dann
  // auf das alte (race-anfällige) max+1-Verfahren zurück. Bestandsschutz.
  // v7.28 / v9.10: Nächste TeilnehmerID atomar holen.
  //   1. Counter-Item GET'en, ETag aus Response-Header lesen.
  //   2. Counter-Item PATCH'en mit IF-MATCH=<ETag>, NextValue=current+1.
  //   3. Bei 412 (ETag-Mismatch = jemand war schneller) → Exponential
  //      Backoff mit Full Jitter, dann Retry. Bis zu 40 Versuche.
  //   4. Bei Erfolg: NextValue+1 ist die neue TeilnehmerID des aufrufenden Users.
  //
  // v9.10: Counter-Liste wird ON-DEMAND angelegt+geseeded, falls sie fehlt
  // (z.B. weil das Event vor v7.28 erstellt wurde). Vorher gab undefined
  // zurück → Aufrufer fiel auf max+1 zurück → Race-Condition bei
  // Massen-Anmeldungen. Jetzt: einmalig ensureCounterList() rufen, dann
  // erneut versuchen. Damit ist der race-anfällige Fallback nur noch
  // erreicht, wenn auch das Anlegen scheitert (Permission-Issue).
  //
  // v9.10: Retries 8 → 40, Backoff von festem Jitter auf Exponential
  // Backoff mit Full Jitter (Cap 500ms). Bei 50+ parallelen Anmeldungen
  // wahren 8 Retries praktisch garantiert ausgeschöpft.
  private async getNextTeilnehmerId(subsiteUrl: string): Promise<number | undefined> {
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    const MAX_RETRIES = 40;
    let triedLazyCreate = false;
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return undefined;
      }
      if (!getResp.ok) {
        // 404 = Counter-Liste / Item existiert nicht.
        // v9.10: Statt direkt undefined zu liefern, einmalig versuchen die
        // Liste anzulegen + zu seeden (idempotent). Wenn das klappt, gleich
        // weiter — wenn nicht, geben wir auf.
        if (getResp.status === 404 && !triedLazyCreate) {
          triedLazyCreate = true;
          try {
            await this.ensureCounterList(subsiteUrl);
            // Kein delay — direkt nächste Iteration, die das frische Item liest.
            continue;
          } catch {
            return undefined;
          }
        }
        return undefined;
      }
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return undefined;
      let data;
      try { data = await getResp.json(); } catch { return undefined; }
      // v9.13: NextValue defensiv parsen — handhabt sowohl number als auch
      // (in seltenen SP-Konfigurationen) string. Sorgt dafür dass current
      // bei einer korrekt gespeicherten 165 nie auf 0 fallback'ed.
      const rawNextValue = data?.NextValue ?? data?.d?.NextValue;
      const current = typeof rawNextValue === 'number'
        ? rawNextValue
        : (typeof rawNextValue === 'string' ? (parseInt(rawNextValue, 10) || 0) : 0);
      // v9.13: Counter ist die Source of Truth. Wir vertrauen ihm und
      // inkrementieren atomar via ETag-CAS. KEIN zusätzlicher Lesezugriff
      // auf die Teilnehmerliste mehr — das Counter-Pattern existiert genau,
      // damit wir hier NICHT max(TID) aus der Teilnehmerliste rechnen
      // müssen. Wenn der Counter korrupt sein sollte (z.B. von altem
      // syncCounterToMax-Bug auf 0 gepatcht), gibt's den expliziten
      // "Counter zurücksetzen"-Button im Admin Center für den Fix.
      const next = current + 1;
      const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': next }, etag);
      if (patchResp.ok) return next;
      if (patchResp.status !== 412) {
        // Anderer Fehler (z.B. 500) → kein Sinn weiter zu retry'n
        return undefined;
      }
      // 412 = ETag-Mismatch = jemand war schneller → Exponential Backoff
      // mit Full Jitter (Cap 500ms). Cluster bei Massen-Anmeldungen
      // werden so zuverlässig entzerrt — ohne Backoff laufen alle
      // Clients sekundengleich in den nächsten Conflict.
      const baseDelay = Math.min(500, 50 * Math.pow(1.4, attempt));
      const delay = Math.floor(baseDelay * (0.5 + Math.random()));
      await new Promise(res => setTimeout(res, delay));
    }
    // Nach 40 Retries aufgeben — Aufrufer kann die Anmeldung sauber
    // mit TeilnehmerID=null durchziehen lassen und der Admin lädt
    // anschliessend "IDs neu vergeben".
    console.warn('[DEX] getNextTeilnehmerId: 40 retries erschöpft — TeilnehmerID bleibt unset, Admin sollte IDs neu vergeben.');
    return undefined;
  }

  // Hilfsroutine: prüft ob auf der Counter-Liste die NextValue-Spalte
  // existiert und legt sie an wenn sie fehlt. Idempotent.
  // v24.76: Counter-Felder (inkl. des neuen WaitlistTaken) EINMAL pro Subsite
  // pro Session sicherstellen, bevor darauf geschrieben wird — sonst liefert der
  // MERGE auf Bestands-Events ein HTTP 400 (Feld existiert noch nicht). Gecacht,
  // damit nicht bei jedem Reconcile/Bump erneut geprobt wird.
  private _counterFieldsEnsured: Set<string> = new Set<string>();
  private async ensureCounterFieldsOnce(subsiteUrl: string): Promise<void> {
    if (!subsiteUrl || this._counterFieldsEnsured.has(subsiteUrl)) return;
    try { await this.ensureCounterListField(subsiteUrl); } catch { /* best-effort */ }
    this._counterFieldsEnsured.add(subsiteUrl);
  }

  private async ensureCounterListField(subsiteUrl: string): Promise<void> {
    const fieldsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields`;
    // v7.28: NextValue (TeilnehmerID-Automat).
    // v11.36: SeatsTaken / SeatsTakenDurch / SeatsTakenFun — atomare
    // Sitzplatz-Reservierung pro Gruppe (gegen Überbuchung bei
    // zeitgleichen Anmeldungen). Alle Number-Felder, default 0/leer.
    const wanted = ['NextValue', 'SeatsTaken', 'SeatsTakenDurch', 'SeatsTakenFun', 'WaitlistTaken'];
    for (const name of wanted) {
      try {
        const probe = await this._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields/getbytitle('${name}')`,
          SPHttpClient.configurations.v1
        );
        if (probe.ok) continue;
        await this._post(
          fieldsUrl,
          { '__metadata': { 'type': 'SP.Field' }, 'Title': name, 'FieldTypeKind': 9, 'Required': false }
        );
      } catch { /* best-effort; reserveSeat fällt sonst sauber zurück */ }
    }
  }

  // v7.28: Counter-Liste für ein Event anlegen (1 Liste mit 1 Item) und
  // direkt mit dem aktuellen Max-Wert seeden — damit bestehende Events ohne
  // ID-Lückenproduktion umsteigen können.
  // Idempotent: tut nichts wenn die Liste schon existiert.
  // v7.29-Fix: Item-Inserts nutzen den korrekt _x005f_-encodeten Type-Namen
  // (genauso wie wir das für DEX_Events machen). Vorher wurde der Listen-
  // name 1:1 in den Type übernommen, was bei Unterstrich stillschweigend
  // zu HTTP 400 führt → leere Counter-Liste.
  private async ensureCounterList(subsiteUrl: string): Promise<{ created: boolean; seededValue?: number }> {
    const probe = await this._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')`,
      SPHttpClient.configurations.v1
    );
    const itemsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items`;
    const seedItem = async (): Promise<number> => {
      const maxId = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
      const resp = await this._post(itemsUrl, {
        '__metadata': { 'type': COUNTER_LIST_ITEM_TYPE },
        'Title': 'TeilnehmerID',
        'NextValue': maxId,
        // v30.63: Sitzplatz- und Wartelisten-Zähler AUSDRÜCKLICH auf 0 setzen.
        //
        // Bisher blieben beide Felder null, bis die erste Anmeldung sie anfasst
        // — und `WaitlistTaken` fasst niemand an, solange nie jemand auf der
        // Warteliste steht. `getCounterStats` meldet null aber als „unbekannt"
        // (-1), und `releaseSeatAfterCancel` handelt bei unbekanntem
        // Wartelisten-Stand fail-closed: Es zählt NICHT herunter. Ergebnis war
        // ein Zähler, der bei jeder Selbst-Abmeldung zu hoch stehen blieb, bis
        // ein Admin die App öffnete.
        //
        // Auf einer frischen Subsite ist die 0 keine Annahme, sondern die
        // Tatsache: Es gibt noch keine Anmeldung. Genau deshalb darf sie hier
        // geschrieben werden — und nur hier.
        'SeatsTaken': 0,
        'WaitlistTaken': 0,
      });
      if (!resp.ok) {
        const errBody = await resp.text().catch(() => '');
        console.warn('[DEX] ensureCounterList: Seed-Item fehlgeschlagen, status=', resp.status, errBody.substring(0, 300));
      }
      return maxId;
    };

    if (probe.ok) {
      // Liste existiert — sicherstellen dass das Schema komplett ist und ein Item drin liegt.
      try { await this.ensureCounterListField(subsiteUrl); } catch { /* */ }
      // v9.13/v9.14: setCounterListPermissions wurde hier ursprünglich
      // mitaufgerufen, hat aber bei laufender Event-Anlage Race-Conditions
      // ausgelöst (breakroleinheritance gegen frisch provisionierte Liste
      // in derselben Request-Welle). Permissions werden jetzt nur noch
      // explizit über den "Counter zurücksetzen"-Button gefixt — siehe
      // resetCounterToMax. Bestehende Events können damit per Admin-Klick
      // geheilt werden, neue Events bekommen ihre Permissions im
      // create-Branch unten gesetzt.
      const itemListResp = await this._sp.get(
        `${itemsUrl}?$top=1`,
        SPHttpClient.configurations.v1
      );
      if (itemListResp.ok) {
        const data = await itemListResp.json();
        const list = data.value || data.d?.results || [];
        if (list.length > 0) return { created: false }; // alles ok
      }
      // Liste ohne Item → nachseeden
      const seededValue = await seedItem();
      return { created: false, seededValue };
    }

    // Liste neu anlegen
    await this._post(
      `${subsiteUrl}/_api/web/lists`,
      {
        '__metadata': { 'type': 'SP.List' },
        'BaseTemplate': 100,
        'Title': COUNTER_LIST_NAME,
        'Description': 'Atomarer Counter für TeilnehmerID-Vergabe (ETag-basiert). Nicht manuell editieren.',
        'AllowContentTypes': false,
        'ContentTypesEnabled': false,
        'EnableVersioning': false,
        'EnableMinorVersions': false,
        'OnQuickLaunch': false,
      }
    );
    await this.ensureCounterListField(subsiteUrl);
    // v9.13: Counter-Liste muss explizit Contribute-Rechte für Visitors
    // bekommen, damit normale User die ETag-CAS-Inkrementierung
    // durchführen können. Ohne das schlägt PATCH NextValue mit 401/403
    // fehl → getNextTeilnehmerId gibt undefined zurück → TID landet null
    // (oder im allerersten Lazy-Create-Pfad bei 1).
    try { await this.setCounterListPermissions(subsiteUrl); } catch { /* */ }
    const seededValue = await seedItem();
    return { created: true, seededValue };
  }

  /**
   * Berechtigungen für DEX_TeilnehmerCounter setzen — analog zur
   * Teilnehmerliste:
   *   - Owners der Hauptsite: Full Control (1073741829)
   *   - Visitors (DEALL): Contribute (1073741827) → ETag-CAS-Inkrement
   *   - Organizer-Mail (falls bekannt): Full Control
   *
   * Idempotent: kann auf bestehenden Counter-Listen erneut aufgerufen
   * werden um v9.13-Permissions nachzupatchen. Die Funktion bricht
   * Rollen-Vererbung explizit (clearSubscopes=true), damit Read-Only-
   * Inheritance vom Subsite nicht versehentlich greift.
   */
  private async setCounterListPermissions(subsiteUrl: string, organizerEmail?: string): Promise<void> {
    try {
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );
      const ownersResp = await this._sp.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResp.ok) {
        const ownersData = await ownersResp.json();
        const ownersId = ownersData.Id ?? ownersData.d?.Id;
        if (ownersId) {
          await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${ownersId}, roledefid=1073741829)`,
            {}
          );
        }
      }
      // Visitors → Contribute. KRITISCH: damit normale User
      // den Counter atomar inkrementieren können.
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`,
          {}
        );
      }
      // Organizer optional → Full Control
      if (organizerEmail) {
        try {
          const userResp = await this._sp.get(
            `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(organizerEmail)}')?$select=Id`,
            SPHttpClient.configurations.v1
          );
          if (userResp.ok) {
            const userData = await userResp.json();
            const userId = userData.Id ?? userData.d?.Id;
            if (userId) {
              await this._post(
                `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
                {}
              );
            }
          }
        } catch { /* Organizer-Permission ist optional */ }
      }
    } catch (err) {
      console.warn('[DEX] setCounterListPermissions fehlgeschlagen:', err);
    }
  }

  // v7.28 / v9.13: Aktuellen Max-Wert von TeilnehmerID in der Teilnehmerliste
  // lesen. Wird beim **Seeden** des Counters und beim **Sync nach Reorder**
  // genutzt — die Counter-Liste selbst ist im Normalbetrieb die Source of
  // Truth (siehe getNextTeilnehmerId).
  //
  // **Bugfix v9.13:** Vorher hat $orderby=TeilnehmerID desc&$top=1 unter
  // bestimmten Bedingungen das null-Item zuerst geliefert (SP sortiert NULL-
  // Werte bei Number-Feldern oft als "größter Wert" in desc-Order).
  // Sobald irgendjemand abgemeldet war (TID=null) lief die Funktion ins
  // null-Branch und gab 0 zurück.
  //
  // Konsequenz im alten Code (vor v9.12): syncCounterToMax patcht den
  // Counter auf liveMax=0 RUNTER → nächste Anmeldung kriegt TID=1 →
  // Duplikat zu echten aktiven Teilnehmern. Genau der Fall den der User
  // beim Go-Live live gesehen hat (Theresa #1 obwohl 165 aktive Anmeldungen).
  //
  // Fix: $filter=TeilnehmerID gt 0 schliesst NULL und 0 explizit aus —
  // funktioniert unabhängig von SP-NULL-Sortier-Konventionen.
  private async getCurrentMaxTeilnehmerId(subsiteUrl: string): Promise<number> {
    try {
      const resp = await this._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=TeilnehmerID&$filter=TeilnehmerID gt 0&$orderby=TeilnehmerID desc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return 0;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length > 0 && items[0].TeilnehmerID != null) return items[0].TeilnehmerID;
    } catch { /* */ }
    return 0;
  }

  // v7.31 / v9.12: Counter mit aktuellem Max-Wert konsistent halten — wird
  // nach Cancel und nach reorderParticipantIDs aufgerufen. Die Logik ist
  // **monotonic up-only**: der Counter geht NIE runter (sonst werden
  // cancelled IDs reused — exakter Duplikat-Bug aus dem Go-Live).
  //
  // - current >= liveMax: nichts zu tun (Counter steht bereits hoch genug).
  // - current  < liveMax: Counter HOCH auf liveMax setzen (z.B. nach
  //   reorderParticipantIDs, der die TIDs im Bereich [1..N_active] vergibt
  //   wo N_active größer sein kann als der bisherige Counter-Stand).
  //
  // Vorher (Bug bis v9.11): exakt umgekehrt — Counter wurde auf liveMax
  // RUNTER gesetzt wenn current > liveMax. Das produzierte sowohl bei
  // Cancel als auch nach IDReorder Duplikate.
  //
  // ETag-CAS mit Retry, damit eine parallele Anmeldung den Counter nicht
  // zwischen Read und Write wegrasselt.
  // v9.13: Oeffentliche Recovery-Methode für den Admin-Button "Counter
  // zurücksetzen". Liest den aktuellen Max-TID aus der Teilnehmerliste und
  // setzt den Counter auf diesen Wert (per ETag-CAS, monotonic up-only via
  // syncCounterToMax). Gibt den neuen Counter-Wert zurück damit der Admin
  // direkt sehen kann auf was es gepatcht wurde.
  public async resetCounterToMax(subsiteUrl: string): Promise<{ counter: number; max: number }> {
    // v11.27: bidirektionaler Reset. Vorher rief diese Methode nur
    // syncCounterToMax auf — das ist monotonic up-only und liess einen
    // zu hohen Counter unverändert. Genau das hat der Maintainer beobachtet:
    // Counter=11, Max-TID=4, Klick auf "Counter zurücksetzen" → keine
    // Änderung, weiterhin 11. Jetzt setzen wir den Counter explizit
    // auf max(TID) — egal ob er drunter (gefährlich, Doppel-IDs möglich)
    // oder drüber stand (harmlos, nur Lücken-Springen).
    try { await this.ensureCounterList(subsiteUrl); } catch { /* */ }
    const max = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    let finalCounter = 0;
    // ETag-CAS-Loop, falls jemand parallel inserted und den Counter inkrementiert.
    for (let attempt = 0; attempt < 8; attempt++) {
      try {
        const getResp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
        if (!getResp.ok) break;
        const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
        if (!etag) break;
        const data = await getResp.json();
        const rawCurrent = data?.NextValue ?? data?.d?.NextValue;
        const current = typeof rawCurrent === 'number' ? rawCurrent : (typeof rawCurrent === 'string' ? (parseInt(rawCurrent, 10) || 0) : 0);
        if (current === max) {
          finalCounter = current;
          break;
        }
        const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': max }, etag);
        if (patchResp.ok) {
          finalCounter = max;
          console.warn(`[DEX] resetCounterToMax: counter von ${current} auf ${max} gesetzt (Subsite: ${subsiteUrl}).`);
          break;
        }
        if (patchResp.status !== 412) {
          finalCounter = current;
          break;
        }
        // 412 = jemand war schneller, nochmal lesen+patchen
        await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
      } catch (err) {
        console.warn('[DEX] resetCounterToMax error:', err);
        break;
      }
    }
    return { counter: finalCounter, max };
  }

  private async syncCounterToMax(subsiteUrl: string): Promise<void> {
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    for (let attempt = 0; attempt < 8; attempt++) {
      const liveMax = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return;
      }
      if (!getResp.ok) return;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return;
      let data;
      try { data = await getResp.json(); } catch { return; }
      const current = typeof data?.NextValue === 'number' ? data.NextValue : 0;
      if (current >= liveMax) return; // bereits konsistent — Counter ist nicht "zu klein"
      const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': liveMax }, etag);
      if (patchResp.ok) {
        console.warn(`[DEX] syncCounterToMax: counter von ${current} auf ${liveMax} hochgezogen.`);
        return;
      }
      if (patchResp.status !== 412) return;
      // 412 = jemand war schneller, nochmal lesen+patchen
      await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
    }
    // Nach 8 Retries aufgeben — best-effort, blockiert keine andere Aktion
  }

  /**
   * v19.28: Eine Teilnehmer-Registrierung endgültig aus der Subsite-Liste
   * löschen (hartes DELETE, kein Recycle-Bin). Use-Case: abgemeldete
   * Test-Anmeldungen aus der Abmeldungen-Liste entfernen, damit die Übersicht
   * sauber bleibt. Die Berechtigung (Admin/Organizer) wird in der UI geprüft.
   */
  public async deleteRegistration(subsiteUrl: string, itemId: number): Promise<boolean> {
    try {
      const resp = await this._delete(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`);
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] deleteRegistration failed:', err);
      return false;
    }
  }

  /**
   * v21: Item-Level-Security („nur eigene Elemente", 2/2) auf den globalen
   * Queue-Listen DEX_Outlook + DEX_IDReorder nachziehen. DEX_Emails und
   * DEX_AccessFix haben sie bereits; DEX_TeamJoinRequests bekommt sie
   * bewusst NICHT (der Team-Lead muss fremde Beitritts-Anfragen lesen).
   * Idempotent; wird vom Admin-Reparatur-Button mit ausgeführt.
   */
  public async hardenQueueListsIls(): Promise<{ fixed: string[]; failed: string[] }> {
    const targets = ['DEX_Outlook', 'DEX_IDReorder'];
    const fixed: string[] = [];
    const failed: string[] = [];
    for (const listName of targets) {
      try {
        // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400).
        const st = await this._setListSecurity(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 2, WriteSecurity: 2 });
        if (st >= 200 && st < 300) fixed.push(listName); else failed.push(listName);
      } catch { failed.push(listName); }
    }
    return { fixed, failed };
  }

  // ==================== v26.79: Berechtigungen aufräumen ====================
  // Scannt die GESAMTE Site-Collection (Haupt-Web + alle Listen/Bibliotheken +
  // alle Subsites + deren Listen) und findet EINZEL-Freigaben (direkte
  // Nutzer-Berechtigungen), die über das Rollen-Konzept hinaus SCHREIB-/Vollzugriff
  // geben. Soll-Konzept: Schreiben nur über die Gruppen (Owners, Members,
  // Visitors/DEALL) plus die im Rollen-Konzept vorgesehenen Einzelpersonen
  // (Admins global; Organizer auf DEX_Events/Web-Root + ihren eigenen
  // Event-Subsites). Alle anderen direkten Nutzer-Schreibrechte gelten als
  // „manuelle Über-Freigabe" und werden entfernt — LESERECHTE bleiben erhalten
  // (die betroffene Person liest weiterhin über die Visitors-/DEALL-Gruppe;
  // internationale Leser sind bewusst OK). Gruppen werden NIE angefasst.
  //
  // Zusätzlich wird die Element-Sicherheit (ReadSecurity/WriteSecurity=2 „nur
  // eigene Elemente") auf den sensiblen Listen (Teilnehmer, DEX_Emails,
  // DEX_Outlook, DEX_IDReorder) geprüft und korrigiert.
  //
  // apply=false → reiner Prüf-/Dry-Run-Bericht (ändert NICHTS).
  // apply=true  → entfernt die Über-Freigaben und korrigiert die Element-Sicherheit.
  public async auditOrCleanupPermissions(
    apply: boolean,
    ctx: { adminEmails: string[]; organizerEmails: string[]; subsiteOrganizers: Record<string, string>; selfEmail: string },
    onProgress?: (msg: string, done: number, total: number) => void
  ): Promise<PermCleanupReport> {
    const report: PermCleanupReport = {
      apply, websScanned: 0, listsScanned: 0, strayWriteFound: 0, strayWriteRemoved: 0,
      ilsIssues: 0, ilsFixed: 0, errors: 0, findings: [],
    };
    const MAX_FINDINGS = 800;
    const addFinding = (f: PermCleanupFinding): void => { if (report.findings.length < MAX_FINDINGS) report.findings.push(f); };
    const norm = (e: string): string => (e || '').trim().toLowerCase();
    const globalAllowed = new Set([...(ctx.adminEmails || []), ...(ctx.organizerEmails || []), ctx.selfEmail].map(norm).filter(Boolean));
    // 1073741825 = Limited Access (System-verwaltet), 1073741826 = Read → beide
    // gelten NICHT als „Schreibrecht". Alles andere (Contribute/Edit/Design/Full
    // Control + Custom-Level) zählt als elevated.
    const isElevated = (ids: number[]): boolean => ids.some(id => id !== 1073741825 && id !== 1073741826);
    const extractEmail = (login: string): string => {
      const l = login || '';
      const m = /\|membership\|([^|]+)$/i.exec(l) || /\|([^|]+@[^|]+)$/.exec(l);
      return m ? m[1] : '';
    };
    const isSystemPrincipal = (email: string, login: string): boolean => {
      const l = (login || '').toLowerCase();
      if (!email) return true; // ohne E-Mail nicht sicher klassifizierbar → nicht anfassen
      if (email.indexOf('@sharepoint') >= 0) return true;
      if (l.indexOf('app@sharepoint') >= 0 || l.indexOf('|spo-grid') >= 0 || l.indexOf('c:0(.s|true') >= 0) return true;
      return false;
    };

    const processSecurable = async (scopeBase: string, label: string, allowed: Set<string>): Promise<void> => {
      let assigns: Array<{ pid: number; type: number; title: string; login: string; email: string; roleIds: number[]; roleNames: string[] }>;
      try {
        assigns = await this._readRoleAssignments(scopeBase);
      } catch {
        report.errors++;
        addFinding({ scope: label, kind: 'error', detail: 'Berechtigungen konnten nicht gelesen werden.', fixed: false });
        return;
      }
      for (const a of assigns) {
        if (a.type !== 1) continue; // nur einzelne User (Gruppen NIE anfassen)
        const email = norm(a.email || extractEmail(a.login));
        if (isSystemPrincipal(email, a.login)) continue;
        if (allowed.has(email)) continue; // Admin/Organizer/Ich → legitim
        if (!isElevated(a.roleIds)) continue; // reine Leseberechtigung → OK (int. Leser)
        report.strayWriteFound++;
        let fixed = false;
        if (apply) {
          try {
            const r = await this._deletePrincipalAssignment(scopeBase, a.pid);
            fixed = r.ok || r.status === 200 || r.status === 204;
            if (fixed) report.strayWriteRemoved++; else report.errors++;
          } catch { report.errors++; }
        }
        addFinding({
          scope: label, kind: 'stray-write', principal: email || a.title,
          detail: `${(a.roleNames.filter(Boolean).join(', ') || 'Schreibzugriff')} — ${apply ? (fixed ? 'Schreibrecht entfernt (Lesen bleibt, sofern Gruppenmitglied)' : 'Entfernen fehlgeschlagen') : 'würde entfernt (Lesen bleibt über Gruppe)'}`,
          fixed,
        });
      }
    };

    // Liest ReadSecurity/WriteSecurity einer Liste (roh + geparst). Loggt bei
    // Bedarf die exakte Server-Antwort — Diagnose, warum eine Korrektur ggf.
    // nicht greift (Format/Stale/Rechte).
    const readIls = async (listBase: string, label: string): Promise<{ rs: number; ws: number; raw: string; status: number } | null> => {
      try {
        const resp = await this._sp.get(
          `${listBase}?$select=ReadSecurity,WriteSecurity`, SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (!resp.ok) { console.warn('[DEX PermFix] ILS-Read HTTP', resp.status, label); return null; }
        const d = await resp.json();
        const raw = JSON.stringify(d).slice(0, 200);
        return { rs: Number(d.ReadSecurity ?? d.d?.ReadSecurity), ws: Number(d.WriteSecurity ?? d.d?.WriteSecurity), raw, status: resp.status };
      } catch (e) { console.warn('[DEX PermFix] ILS-Read ERROR', label, e); return null; }
    };
    const checkIls = async (listBase: string, label: string): Promise<void> => {
      const before = await readIls(listBase, label);
      if (!before || !Number.isFinite(before.rs) || !Number.isFinite(before.ws)) {
        console.warn('[DEX PermFix] ILS unlesbar/format', label, 'raw=', before?.raw);
        return;
      }
      if (before.rs === 2 && before.ws === 2) return;
      report.ilsIssues++;
      console.warn(`[DEX PermFix] ILS FALSCH ${before.rs}/${before.ws} (soll 2/2)`, label, 'raw=', before.raw);
      let fixed = false;
      if (apply) {
        // v26.87: zuverlässiger nometadata-MERGE (Digest + kein __metadata).
        // Der bisherige verbose+__metadata-MERGE gab unter SPFx odata-version
        // 3.0 flächendeckend HTTP 400 zurück → nichts wurde je korrigiert.
        const mergeStatus = await this._setListSecurity(listBase, { ReadSecurity: 2, WriteSecurity: 2 });
        // ENTSCHEIDEND: Read-back — nur wenn der Server danach WIRKLICH 2/2
        // meldet, gilt es als korrigiert (nicht blind dem MERGE-Status trauen).
        const after = await readIls(listBase, label);
        fixed = !!after && after.rs === 2 && after.ws === 2;
        console.warn(`[DEX PermFix] ILS-FIX ${label} | MERGE-Status=${mergeStatus} | nachher=${after ? `${after.rs}/${after.ws}` : 'null'} | raw=${after?.raw} | => ${fixed ? 'OK' : 'WEITER FALSCH'}`);
        if (fixed) report.ilsFixed++; else report.errors++;
      }
      addFinding({ scope: label, kind: 'ils', detail: `Element-Sicherheit ${before.rs}/${before.ws} statt 2/2 („nur eigene Elemente") — ${apply ? (fixed ? 'korrigiert' : 'Korrektur fehlgeschlagen (siehe Konsole)') : 'würde korrigiert'}`, fixed });
    };
    const ILS_LISTS = new Set(['DEX_Emails', 'DEX_Outlook', 'DEX_IDReorder']);

    // ---- 1. Haupt-Web (Site-Root) ----
    onProgress?.('Hauptseite …', 0, 1);
    await processSecurable(`${this.siteUrl}/_api/web`, 'Hauptseite (Web-Root)', globalAllowed);
    report.websScanned++;

    // ---- 2. Listen/Bibliotheken des Haupt-Webs ----
    const rootLists = await this._listSecurables(this.siteUrl);
    // Alle Subsites der Collection einsammeln (BFS, max. 3 Ebenen, gedeckelt) —
    // Event-Subsites hängen direkt unter dem Root, tiefere Verschachtelung wird
    // vorsorglich mitgenommen, damit wirklich der GANZE SharePoint erfasst ist.
    const subwebs: Array<{ url: string; serverRel: string; title: string; unique: boolean }> = [];
    const seenWebs = new Set<string>([this.siteUrl.toLowerCase().replace(/\/+$/, '')]);
    let frontier = [this.siteUrl];
    for (let depth = 0; depth < 3 && frontier.length > 0 && subwebs.length < 500; depth++) {
      const next: string[] = [];
      for (const wurl of frontier) {
        const kids = await this._childWebs(wurl);
        for (const k of kids) {
          const kkey = k.url.toLowerCase().replace(/\/+$/, '');
          if (!k.url || seenWebs.has(kkey)) continue;
          seenWebs.add(kkey);
          subwebs.push(k);
          next.push(k.url);
          if (subwebs.length >= 500) break;
        }
        if (subwebs.length >= 500) break;
      }
      frontier = next;
    }
    const total = 1 + rootLists.length + subwebs.length;
    let done = 1;
    for (const l of rootLists) {
      onProgress?.(`Liste ${l.title} …`, done, total);
      if (l.unique) { await processSecurable(l.base, `Liste ${l.title}`, globalAllowed); report.listsScanned++; }
      if (ILS_LISTS.has(l.title)) await checkIls(l.base, `Liste ${l.title}`);
      done++;
    }

    // ---- 3. Subsites (Event-Subsites) + deren Listen ----
    for (const w of subwebs) {
      const wlabel = w.title || w.serverRel || w.url;
      onProgress?.(`Subsite ${wlabel} …`, done, total);
      const key = (s: string): string => norm(s).replace(/\/+$/, '');
      const orgStr = ctx.subsiteOrganizers[key(w.serverRel)] || ctx.subsiteOrganizers[key(w.url)] || '';
      const allowed = new Set(globalAllowed);
      orgStr.split(/[;,]/).map(norm).filter(Boolean).forEach(e => allowed.add(e));
      if (w.unique) { await processSecurable(`${w.url}/_api/web`, `Subsite ${wlabel} – Web`, allowed); }
      report.websScanned++;
      try {
        const subLists = await this._listSecurables(w.url);
        for (const sl of subLists) {
          if (sl.unique) { await processSecurable(sl.base, `Subsite ${wlabel} – Liste ${sl.title}`, allowed); report.listsScanned++; }
          if (sl.title === REG_LIST_NAME) await checkIls(sl.base, `Subsite ${wlabel} – Teilnehmerliste`);
        }
      } catch { /* Subsite-Listen nicht lesbar */ }
      done++;
    }
    onProgress?.('Fertig', total, total);
    return report;
  }

  /** Liest die Rollenzuweisungen eines Securables (Web oder Liste). scopeBase =
   *  voll-qualifizierte API-URL bis zum Securable (…/_api/web bzw.
   *  …/_api/web/lists/getbytitle('X')). */
  private async _readRoleAssignments(scopeBase: string): Promise<Array<{ pid: number; type: number; title: string; login: string; email: string; roleIds: number[]; roleNames: string[] }>> {
    // $top hoch setzen: ein über-freigegebenes Securable (genau der Fall, den
    // wir suchen) kann viele Einzel-Zuweisungen haben — ohne $top würde OData
    // bei 100 abschneiden und Über-Freigaben stillschweigend übersehen.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = [];
    let url: string | null = `${scopeBase}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/LoginName,Member/PrincipalType,Member/Email,RoleDefinitionBindings/Id,RoleDefinitionBindings/Name&$top=2000`;
    let guard = 0;
    while (url && guard < 20) {
      guard++;
      const resp: SPHttpClientResponse = await this._sp.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
      if (!resp.ok) { if (guard === 1) throw new Error(`roleassignments ${resp.status}`); break; }
      const d = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const page: any[] = d.value || d.d?.results || [];
      for (const p of page) rows.push(p);
      // Falls SharePoint doch paginiert: nextLink verfolgen (nometadata-Feldname).
      url = d['odata.nextLink'] || d['@odata.nextLink'] || null;
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return rows.map((ra: any) => {
      const m = ra.Member || {};
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const binds: any[] = ra.RoleDefinitionBindings || ra.RoleDefinitionBindings?.results || [];
      return {
        pid: Number(ra.PrincipalId),
        type: Number(m.PrincipalType) || 0,
        title: m.Title || '',
        login: m.LoginName || '',
        email: m.Email || '',
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        roleIds: binds.map((b: any) => Number(b.Id)),
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        roleNames: binds.map((b: any) => b.Name || ''),
      };
    });
  }

  /** Entfernt ALLE Rollenzuweisungen eines Principals auf einem Securable
   *  (Downgrade auf „kein direktes Recht" — Leserecht über Gruppen bleibt). */
  private async _deletePrincipalAssignment(scopeBase: string, principalId: number): Promise<SPHttpClientResponse> {
    // v26.81: Digest des Ziel-Webs mitschicken (Cross-Web-Schreibzugriff auf
    // Subsites würde sonst mit 403 abgelehnt).
    const digest = await this._webDigest(this._webOf(scopeBase));
    return this._sp.post(
      `${scopeBase}/roleassignments/getbyprincipalid(${principalId})`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'odata-version': '', 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', ...(digest ? { 'X-RequestDigest': digest } : {}) } }
    );
  }

  /** Alle Listen/Bibliotheken eines Webs mit Unique-Flag. base = API-URL zum
   *  Securable der Liste. Titel mit Sonderzeichen werden für getbytitle escaped. */
  private async _listSecurables(webUrl: string): Promise<Array<{ title: string; hidden: boolean; unique: boolean; base: string }>> {
    try {
      const resp = await this._sp.get(
        `${webUrl}/_api/web/lists?$select=Title,Hidden,HasUniqueRoleAssignments&$top=1000`,
        SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (!resp.ok) return [];
      const d = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rows: any[] = d.value || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return rows.map((l: any) => ({
        title: l.Title || '',
        hidden: !!l.Hidden,
        unique: !!l.HasUniqueRoleAssignments,
        base: `${webUrl}/_api/web/lists/getbytitle('${String(l.Title || '').replace(/'/g, "''")}')`,
      })).filter(l => l.title);
    } catch { return []; }
  }

  /** Direkte Kind-Webs eines Webs (eine Ebene — Event-Subsites hängen direkt
   *  unter dem Root). */
  private async _childWebs(webUrl: string): Promise<Array<{ url: string; serverRel: string; title: string; unique: boolean }>> {
    try {
      const resp = await this._sp.get(
        `${webUrl}/_api/web/webs?$select=Url,ServerRelativeUrl,Title,HasUniqueRoleAssignments&$top=1000`,
        SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (!resp.ok) return [];
      const d = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rows: any[] = d.value || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return rows.map((w: any) => ({ url: w.Url || '', serverRel: w.ServerRelativeUrl || '', title: w.Title || '', unique: !!w.HasUniqueRoleAssignments })).filter(w => w.url);
    } catch { return []; }
  }

  // ==================== v26.81: Verwaiste Subsites finden ====================
  // Vergleicht alle real existierenden Subsites (Webs) mit den von DEX_Events
  // referenzierten SubsiteUrls. Ein Web, das von KEINEM Event (Haupt- oder
  // Sub-Event) referenziert wird, ist ein „Rest" — z.B. eine Test-Subsite,
  // deren Event bereits gelöscht wurde. Reine Analyse, ändert NICHTS.
  public async findOrphanSubsites(
    onProgress?: (msg: string, done: number, total: number) => void
  ): Promise<OrphanScanResult> {
    const result: OrphanScanResult = { websScanned: 0, eventSubsites: 0, orphans: [] };
    const norm = (s: string): string => (s || '').trim().toLowerCase().replace(/\/+$/, '');

    // 1. Referenzierte Subsites aus DEX_Events (alle Events inkl. Sub-Events).
    onProgress?.('Events werden gelesen …', 0, 1);
    const events = await this.getAllEventsForKpi();
    const referenced = new Set<string>();
    for (const e of events) {
      if (e.subsiteUrl) {
        referenced.add(norm(e.subsiteUrl));
        try { referenced.add(norm(new URL(e.subsiteUrl).pathname)); } catch { /* */ }
      }
    }
    result.eventSubsites = referenced.size;

    // 2. Alle Subsites einsammeln (BFS, max. 3 Ebenen, cap 500).
    const allWebs: Array<{ url: string; serverRel: string; title: string }> = [];
    const seen = new Set<string>([norm(this.siteUrl)]);
    let frontier = [this.siteUrl];
    for (let depth = 0; depth < 3 && frontier.length > 0 && allWebs.length < 500; depth++) {
      const next: string[] = [];
      for (const wurl of frontier) {
        const kids = await this._childWebs(wurl);
        for (const k of kids) {
          const kk = norm(k.url);
          if (!k.url || seen.has(kk)) continue;
          seen.add(kk);
          allWebs.push({ url: k.url, serverRel: k.serverRel, title: k.title });
          next.push(k.url);
          if (allWebs.length >= 500) break;
        }
        if (allWebs.length >= 500) break;
      }
      frontier = next;
    }
    result.websScanned = allWebs.length;

    // 3. Nicht referenzierte Webs = Rest-Kandidaten; Metadaten nachladen.
    const candidates = allWebs.filter(w => !referenced.has(norm(w.url)) && !referenced.has(norm(w.serverRel)));
    let i = 0;
    for (const w of candidates) {
      i++;
      onProgress?.(`Prüfe „${w.title || w.serverRel}" …`, i, candidates.length);
      let created = '';
      try {
        const wr = await this._sp.get(`${w.url}/_api/web?$select=Created,Title`, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
        if (wr.ok) { const wd = await wr.json(); created = wd.Created || wd.d?.Created || ''; }
      } catch { /* */ }
      let hasParticipantList = false;
      let participantCount = 0;
      try {
        const lr = await this._sp.get(`${w.url}/_api/web/lists/getbytitle('${REG_LIST_NAME}')?$select=ItemCount`, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
        if (lr.ok) { const ld = await lr.json(); hasParticipantList = true; participantCount = Number(ld.ItemCount ?? ld.d?.ItemCount) || 0; }
      } catch { /* Liste fehlt → kein Event-Rest oder anders strukturiert */ }
      result.orphans.push({ url: w.url, serverRel: w.serverRel, title: w.title, created, hasParticipantList, participantCount });
    }
    onProgress?.('Fertig', candidates.length, candidates.length);
    return result;
  }

  /** Löscht eine (verwaiste) Subsite endgültig — inkl. aller Listen darin.
   *  Nur für Admins (Owner-Rechte nötig). SharePoint verlangt, dass das Web
   *  keine eigenen Unter-Webs mehr hat. */
  public async deleteSubsiteWeb(webUrl: string): Promise<boolean> {
    try {
      const digest = await this._webDigest(webUrl);
      const resp = await this._sp.post(
        `${webUrl}/_api/web`, SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'odata-version': '', 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', ...(digest ? { 'X-RequestDigest': digest } : {}) } }
      );
      return resp.ok || resp.status === 200 || resp.status === 204;
    } catch { return false; }
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
