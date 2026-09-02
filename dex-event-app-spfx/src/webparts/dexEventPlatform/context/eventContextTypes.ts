/**
 * Typen des Event-Contexts.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Reine Typdeklarationen — kein Laufzeit-Code, kein State-Zugriff.
 * `EventContext.tsx` re-exportiert die oeffentlichen Typen unveraendert,
 * damit keine Aufrufstelle angefasst werden musste.
 */

import { DeloitteEvent } from '../types';
import { CustomField, SPRegistration, SPParticipant, ReseedSummary, AssistantLink, EventCommRow } from '../services/EventService';
import { BundledItem } from '../utils/bundledComm';
import { FAConfig } from '../utils/faBilling';
import { CounterStats } from '../services/events/seats';

/** v18.33: Eingabe für den Self-Check-in-Deep-Link. Entweder `token` (statischer
 *  QR) ODER `eventNumber` + `code` + `windowIndex` (rotierender Live-QR). */
export interface SelfCheckInParams {
  token?: string;
  eventNumber?: number;
  code?: string;
  windowIndex?: number;
}

/** v18.33: Strukturiertes Ergebnis des Self-Check-ins für die Ergebnis-UI. */
export type SelfCheckInStatus =
  | 'success'        // erfolgreich eingecheckt
  | 'already'        // war bereits eingecheckt
  | 'not-registered' // nicht für dieses Event angemeldet
  | 'on-waitlist'    // auf der Warteliste — kein Check-in möglich
  | 'not-found'      // Event/Token nicht gefunden
  | 'disabled'       // Self-Check-in für dieses Event nicht aktiviert
  | 'closed'         // außerhalb des Check-in-Zeitfensters
  | 'expired'        // rotierender Code abgelaufen / ungültig
  | 'error';         // technischer Fehler

export interface SelfCheckInResult {
  status: SelfCheckInStatus;
  eventTitle?: string;
  eventStart?: string;
  opensAt?: string;   // ISO, bei status='closed'
  closesAt?: string;  // ISO, bei status='closed'
}

/** v26.33: Eine Zeile aus dem Statistik-Archiv (DEX_EventStats) — reine KPIs
 *  eines Events, dessen Teilnehmerliste nach dem 3-Monats-Löschkonzept gelöscht
 *  wurde. Enthält KEINE personenbezogenen Daten. */
export interface EventStatsRow {
  id: number;
  eventNumber: number;
  eventTitle: string;
  eventType: string;
  location: string;
  startDate: string;
  endDate: string;
  maxParticipants: number | null;
  registeredCount: number;
  qrSentCount: number;
  checkedInCount: number;
  noShowCount: number;
  waitlistCount: number;
  deregisteredCount: number;
  organizer: string;
  archivedByEmail: string;
  archivedDate: string;
}

/**
 * v30.58: Ein Befund je Event aus „Spalten fixen (alle Events)".
 *
 * `stillMissing` ist der wichtige Teil: Spalten, die auch nach dem Fix nicht
 * angelegt werden konnten. Genau die bringen jeden Insert zu Fall, in dem sie
 * vorkommen — bei einer Klammer-Liste also jede Anmeldung, bei der die Person
 * das betreffende Hauptevent-Feld ausgefüllt hat.
 */
export interface FixColumnsDetail {
  eventId: string;
  eventTitle: string;
  /** Klammer/Einzel-Event (true) oder Sub-Event (false). */
  isParent: boolean;
  /** Die Teilnehmerliste selbst ist weg (404/410) — keine Spalten-Frage. */
  listMissing: boolean;
  fixedColumns: string[];
  stillMissing: string[];
  error?: string;
}

export interface EventContextType {
  events: DeloitteEvent[];
  /** Top-Level-Events (ohne parentEventId) — was in EventListPage/MyEventsPage angezeigt wird. */
  topLevelEvents: DeloitteEvent[];
  /** Kind-Events eines Parents (Sub-Events / Trainingssessions), sortiert nach StartDate. */
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  isEventsLoading: boolean;
  /** v29.47: Dokumente eines Events bei Bedarf nachladen (Boot lädt sie nicht mehr). */
  ensureEventDocuments: (eventIds: string[]) => Promise<void>;
  /** v30.67 (Review): Anhänge eines Events verwerfen und neu laden — nach Änderungen am Context vorbei (Wizard). */
  refreshEventDocuments: (eventId: string) => Promise<void>;
  createEvent: (event: CreateEventInput) => Promise<number | null>;
  // v30.42: skipShadowParent — nur die Anmeldeseite setzt das; sie legt die
  // Klammer-Zeile selbst an, und zwar MIT den übergreifenden Antworten.
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string, opts?: { skipShadowParent?: boolean; suppressMail?: boolean; suppressOutlook?: boolean; extraCc?: string; proxyConsentConfirmed?: boolean; actorAllowedAsAssistant?: boolean; skipReload?: boolean; bundledItems?: BundledItem[] }) => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.82: Team-Anmeldung — Lead + N-1 Mitglieder gleichzeitig anmelden.
   *  Reserviert N Plätze atomar; bei Vollbelegung geht das ganze Team auf
   *  die Warteliste (keine Teil-Anmeldungen aus Kapazitätsmangel). */
  registerTeam: (
    eventId: string,
    leadData: { firstName: string; lastName: string; email: string; salutation?: string; customData: Record<string, string>; preferredStarterType?: string },
    members: Array<{ email: string; displayName: string; customData?: Record<string, string> }>,
    teamName: string | undefined
  ) => Promise<{ ok: boolean; teamId?: string; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.82: Andere Team-Mitglieder zu einer Registrierung laden — für das
   *  Team-Badge in „Meine Events". */
  getTeamMembers: (eventId: string, teamId: string) => Promise<SPRegistration[]>;
  /** v11.83: Ein Team-Lead kann nachträglich ein einzelnes Mitglied
   *  zum bereits angemeldeten Team hinzufügen (Plus-Button in MyEvents).
   *  Atomar einen Sitzplatz reservieren, neuen Member-Eintrag anlegen,
   *  Bestätigungs-Mail + Outlook-Termin queuen. */
  addTeamMember: (eventId: string, teamId: string, teamName: string | undefined, member: { email: string; displayName: string }, customData?: Record<string, string>, opts?: { suppressMemberMail?: boolean; suppressOthersMail?: boolean; ccEmail?: string }) => Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v22.49: „Neues Mitglied"-Info an bestehende Team-Mitglieder (scope: alle
   *  oder nur Lead) — vom Zuordnungs-Modal optional ausgelöst. */
  notifyExistingTeamMembers: (eventId: string, teamId: string, teamName: string | undefined, newMemberNames: string[], excludeEmails: string[], scope?: 'all' | 'lead') => Promise<void>;
  /** v17.2: Schon angemeldete Person (ohne TeamId) einem Team zuweisen.
   *  PATCHt nur die TeamId/TeamName/TeamLead-Felder, KEINE neue
   *  Registrierung, KEINE Bestätigungsmail, KEIN Outlook. */
  assignTeamlessToTeam: (eventId: string, teamId: string, teamName: string | undefined, existingRegId: number, isLead?: boolean, opts?: { sendMail?: boolean; recipientEmail?: string; recipientFirstName?: string; recipientLastName?: string; ccEmail?: string }) => Promise<boolean>;
  /** v11.83: Direkter Team-Beitritt aus der Anmeldeseite (wenn der
   *  Organizer "Beitritt erfordert Bestätigung" NICHT aktiviert hat).
   *  Verhalten wie `addTeamMember`, aber läuft mit dem eingeloggten User
   *  selbst als neuem Member. */
  joinTeam: (eventId: string, teamId: string, teamName: string | undefined, customData?: Record<string, string>) => Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.84: Team-Lead-Rolle innerhalb eines Teams übergeben — nur im
   *  Admin Center für Admin/Organizer eigener Events sichtbar. Setzt die
   *  alte Lead-Zeile auf TeamLead=false und die neue auf TeamLead=true,
   *  schickt anschliessend eine Info-Mail an alle aktiven Mitglieder. */
  transferTeamLead: (eventId: string, teamId: string, newLeadEmail: string) => Promise<{ ok: boolean; reason?: string }>;
  /** v11.83: Beitritts-Anfrage in DEX_TeamJoinRequests einreichen — für
   *  Events bei denen der Organizer Approval aktiviert hat. */
  createTeamJoinRequest: (eventId: string, teamId: string, customData?: Record<string, string>) => Promise<{ ok: boolean; itemId?: number; reason?: string }>;
  /** v11.83: Pending-Beitritts-Anfragen abrufen (nur für den
   *  eingeloggten User als Team-Lead — Filter auf TeamId, das er selber
   *  führt). */
  listTeamJoinRequestsForEvent: (eventId: string, teamId: string) => Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string }>>;
  /** v11.83: Approval-/Reject-Entscheidung eines Leads — bei „Approved"
   *  legt die App die Member-Anmeldung an und queued Mails; bei
   *  „Rejected" eine kurze Absage-Mail an den Anfragenden. */
  decideTeamJoinRequest: (requestId: number, decision: 'Approved' | 'Rejected') => Promise<boolean>;
  /** v11.83: Liste der Teams (gruppiert nach TeamId) eines Events für
   *  die „Offene Teams"-Anzeige auf der Registrierungs-Seite. Nur Teams
   *  mit aktivem Mitglied-Count < TeamSize werden aufgeführt. */
  listOpenTeamsForEvent: (eventId: string) => Promise<Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }>>;
  cancelRegistration: (eventId: string, opts?: { suppressNotifications?: boolean; skipReload?: boolean }) => Promise<boolean>;
  /** v18.11: Proaktive Absage durch einen (noch nicht angemeldeten) Teilnehmer
   *  — „Ich nehme nicht teil". Landet als Abgemeldet-Eintrag im Admin-Center. */
  declineEvent: (eventId: string) => Promise<boolean>;
  /** v11.86: Ein Team-Lead meldet über „Team verwalten" stellvertretend
   *  ein Team-Mitglied vom Event ab. Audit-Felder (CancelledBy*) werden
   *  mit dem eingeloggten Lead gefüllt, danach läuft derselbe
   *  Team-Post-Step wie beim Self-Cancel (Info-Mails an die uebrigen
   *  Mitglieder; Auto-Promote nicht relevant, weil der Lead sich nicht
   *  selbst löscht). */
  cancelTeamMember: (
    eventId: string,
    memberRegistration: SPRegistration
  ) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  /** v24.36: Assistenz — alle Anmeldungen (Haupt- + Sub-Events), die der
   *  eingeloggte User STELLVERTRETEND für eine andere Person durchgeführt hat
   *  (RegisteredByEmail = ich, ParticipantEmail ≠ ich). Liefert Event +
   *  Registrierung pro Treffer, dedupliziert über alle Subsites. */
  getMyProxyRegistrations: () => Promise<Array<{ event: DeloitteEvent; registration: SPRegistration }>>;
  /** v24.36: Stellvertretende Abmeldung einer Fremd-Anmeldung durch die
   *  Assistenz — voller Side-Effect-Pfad (Abmelde-Mail, Outlook-Ausladen,
   *  ID-Reorder, Sitzplatz-Sync), Audit als Akteur „assistant". */
  cancelProxyRegistration: (eventId: string, registration: SPRegistration) => Promise<boolean>;
  /** v24.36: Custom-Field-Antworten einer Fremd-Anmeldung aktualisieren
   *  (Assistenz). Schreibt CustomData + die gemappten SP-Spalten. */
  updateProxyRegistration: (eventId: string, registration: SPRegistration, customData: Record<string, string>) => Promise<boolean>;
  /** v24.43: Eine stellvertretend angelegte Anmeldung komplett an die Person
   *  selbst übergeben (Owner/Autor/RegisteredBy → Person). */
  handBackToParticipant: (eventId: string, registration: SPRegistration) => Promise<boolean>;
  /** v24.41: Nach einer Selbst-Anmeldung die Verwaltung an eine Assistenz
   *  delegieren — legt pro betroffener Zeile (Haupt-/Klammer-Event + alle
   *  Sub-Events, für die der User angemeldet ist) einen Delegations-Auftrag in
   *  DEX_AssistantAccess an. Der Flow setzt darauf den Zeilen-Autor +
   *  RegisteredBy auf die Assistenz (Zugriff in deren „Assistenz"-Kachel). */
  delegateRegistrationToAssistant: (eventId: string, assistant: { email: string; name: string }) => Promise<void>;
  /** v24.41 Szenario B: Nach stellvertretender Anmeldung einen Info-Link
   *  anlegen (Anmelder = Owner, angemeldete Person sieht Info). */
  recordProxyDelegation: (eventId: string, participant: { email: string; name: string }) => Promise<void>;
  /** v24.41: Alle aktiven Assistenz-Verknüpfungen des Users (als Person,
   *  Assistenz oder Owner) — Basis für Info-Ansichten + Anforderungen. */
  getMyAssistantLinks: () => Promise<AssistantLink[]>;
  /** v24.42: Änderungs-/Abmelde-Anforderung stellen (schreibt die Anforderung +
   *  schickt dem Owner eine Deeplink-Mail). */
  requestAssistantChange: (link: AssistantLink, requestType: 'change' | 'cancel', note: string) => Promise<boolean>;
  /** v24.42: Anforderung als erledigt/abgelehnt markieren (Owner). */
  resolveAssistantRequest: (linkId: number, decision: 'Done' | 'Rejected') => Promise<boolean>;
  /** v18.33: Self-Check-in über einen gescannten QR-Deep-Link. Löst das Event
   *  per Token (statischer QR) oder Event-Nummer + HMAC-Code (rotierender QR)
   *  auf, validiert Fenster/Frische und setzt die eigene Registrierung auf
   *  „Eingecheckt". Gibt ein strukturiertes Ergebnis für die Ergebnis-UI. */
  selfCheckIn: (params: SelfCheckInParams) => Promise<SelfCheckInResult>;
  /** v22.23: Organizer-Tutorial — solange aktiv, wird das synthetische
   *  Demo-Showcase-Event (read-only, nur client-seitig) in die Event-Liste
   *  injiziert, mit dem eingeloggten User als Organizer. So sieht der
   *  Organizer während der Tour ein Übungs-Event im Organizer Center. */
  setTutorialDemoActive: (on: boolean) => void;
  checkRegistrationByEmail: (eventId: string, email: string) => Promise<SPRegistration | null>;
  // v30.37: optionaler onHttpError — 403/404 ist NICHT „keine Teilnehmer".
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  deleteEvent: (eventId: string) => Promise<boolean>;
  /** v24.0: Anzahl der Anmeldungen über das Organizer-Team hinaus (echte
   *  Teilnehmer, Haupt- + Sub-Events, status-unabhängig). >0 ⇒ Lösch-Sperre
   *  (nur Admin, frühestens 1 Jahr nach Event-Ende). */
  countExternalRegistrations: (event: DeloitteEvent) => Promise<number>;
  /** v24.6: Organizer-Archiv (pro Person ausblenden, reiner Anzeige-Filter). */
  getOrganizerArchivedEventIds: () => Promise<Set<string>>;
  archiveEventForOrganizer: (eventId: string) => Promise<boolean>;
  unarchiveEventForOrganizer: (eventId: string) => Promise<boolean>;
  /** v11.69: Löscht NUR das DEX_Events-Listenitem, ohne Subsite-/Teilnehmer-
   *  Liste-Recycle und ohne Outlook-DeleteEvent-Queue. Wird gebraucht, um ein
   *  Sub-Event mit `existingSubsiteUrl` neu anzulegen, damit der
   *  `DEX_CreateOutlookEvent`-Flow triggert — die alte Subsite mit
   *  Anmeldungen bleibt erhalten. */
  deleteEventItemOnly: (eventId: string) => Promise<boolean>;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean }) => Promise<boolean>;
  /** v26.51: Klartext-Grund, warum das letzte updateEvent fehlschlug — für die
   *  Fehlermeldung im Wizard (vorher stand der Grund nur in der Konsole). */
  getLastEventUpdateError: () => string;
  /** v30.67: Klartext-Grund, warum das letzte deleteEvent fehlschlug — nennt
   *  die Termine, die nicht gelöscht werden konnten (die Klammer bleibt dann
   *  bewusst stehen, damit sie bedienbar bleibt). Leer nach Erfolg. */
  getLastEventDeleteError: (lang?: 'de' | 'en') => string;
  /** v22.42: Automatischer Hintergrund-Fix der Zeilen-Autoren (Sichtbarkeit
   *  von Fremd-Anmeldungen). Läuft beim Admin-Start gedrosselt (1×/24h) über
   *  alle aktiven Subsites — best-effort, blockiert nichts. */
  autoRepairProxyAccess: () => Promise<void>;
  /** v23.8: Wöchentlichen Admin-Bericht versenden, falls fällig (≥7 Tage seit
   *  dem letzten). Beim App-Start für Admins aufgerufen. Mit `force:true`
   *  (Settings-Testbutton) wird die 7-Tage-Sperre übersprungen. Best-effort;
   *  liefert ein kleines Ergebnis für UI-Feedback. */
  maybeSendWeeklyReport: (opts?: { force?: boolean }) => Promise<{ sent: boolean; admins: number; reason?: string }>;
  // v30.5: F&A-Abrechnung (Fachkonzept) — Verteiler, Versand, Abschluss, Auto-Job.
  getFAConfig: () => Promise<FAConfig>;
  saveFAConfig: (cfg: FAConfig) => Promise<boolean>;
  sendFAMail: (ev: DeloitteEvent, kind: 'info' | 'list', opts?: { auto?: boolean }) => Promise<{ ok: boolean; reason?: string }>;
  markEventSettled: (ev: DeloitteEvent) => Promise<boolean>;
  /** v30.61: Aktualisierte Sammelmail nach einer Änderung an den gebuchten
   *  Terminen (nur bei gebündelter Kommunikation). */
  sendBundledUpdateMail: (parentEvent: DeloitteEvent, recipientEmail: string, recipientName: string, items: BundledItem[]) => Promise<boolean>;
  /** v30.60: Von F&A nachgetragene Personalnummern/Kostenstellen in den
   *  gemeldeten Snapshot schreiben (Zuordnung über die E-Mail-Adresse). */
  saveFAPersonalNumbers: (ev: DeloitteEvent, values: Record<string, { personalNr?: string; costCenter?: string }>) => Promise<boolean>;
  /** v30.53: Rückfrage an F&A in der Kommunikationshistorie festhalten. */
  logFAContact: (ev: DeloitteEvent, to: string, subject: string) => Promise<boolean>;
  maybeSendBillingAutoMails: () => Promise<{ infoSent: number; listSent: number; reminders: number }>;
  /** v24.2: „Danke, wir hoffen es lief gut"-Mail an den Organizer beim
   *  App-Öffnen nach dem Event-Tag (1×/Event/Organizer, localStorage-Drossel). */
  maybeSendPostEventOrganizerMails: () => Promise<void>;
  /** v23.37: Antrag „Organizer werden" anlegen (+ Admin-Mail mit Deep-Link). */
  requestOrganizerRole: (email: string, name: string, location: string, message?: string) => Promise<{ ok: boolean; reason?: string }>;
  /** v26.24: Beim Event-Speichern für jeden benannten Co-Organizer, der noch
   *  KEIN Organizer/Admin ist, einen „Organizer werden"-Antrag (zur Admin-
   *  Freigabe) anlegen. orgNames/orgEmails sind die 1:1-gepaarten Strings aus
   *  sanitizeOrganizerPairs (Namen „; "-, Mails ";"-getrennt). Best-effort. */
  requestCoOrganizerApprovals: (orgNames: string, orgEmails: string, eventTitle: string) => Promise<void>;
  /** v26.34: Neu hinzugefügte (Co-)Organizer per Mail informieren (Zugriff auf
   *  die Teilnehmerliste) + Outlook-Kalendereinladung. Best-effort. */
  notifyNewCoOrganizers: (eventId: string, eventTitle: string, added: Array<{ name: string; email: string }>, isDe: boolean, disableOutlook?: boolean) => Promise<void>;
  /** v23.37: offene Organizer-Anträge (für den Admin-Hinweis beim App-Start). */
  getOpenOrganizerRequests: () => Promise<Array<{ id: number; email: string; name: string; location: string; message: string; created: string }>>;
  /** v23.37: Antrag entscheiden — Status setzen + Antragsteller informieren.
   *  Die eigentliche Rollenvergabe macht der Aufrufer über useRoles().addRole. */
  markOrganizerRequestDecided: (id: number, status: 'Approved' | 'Rejected', email: string, name: string, opts?: { suppressMail?: boolean }) => Promise<boolean>;
  /** v26.58: Einzelnen Antrag inkl. Entscheidungs-Metadaten (Status,
   *  DecidedByEmail, DecidedDate) — für den approveorg-Deep-Link. */
  getOrganizerRequestDetails: (id: number) => Promise<{ id: number; email: string; name: string; status: string; decidedByEmail: string; decidedDate: string } | null>;
  /** v22.45: Scannt die übergebenen Events auf Teilnehmer ohne aktives
   *  Deloitte-Konto (für die Landing-Page-Warnung der Organizer/Admins). */
  scanInactiveAccounts: (evs: Array<{ id: string; title: string; subsiteUrl?: string }>) => Promise<Array<{ eventId: string; title: string; people: Array<{ email: string; name: string }> }>>;
  /** v24.51: Organizer per Mail über (ein) inaktives Konto informieren — mit
   *  Dedup über alle Admins (nur einmal pro Event+Person). */
  notifyOrganizerOfInactive: (eventId: string, people: Array<{ email: string; name: string }>) => Promise<{ sent: number; skipped: number; noOrganizer?: boolean }>;
  /** v26.40: Erkannte Ex-Deloitte-Personen automatisch abmelden (Event-Setting
   *  inactiveHandling='autoderegister'). Gibt die abgemeldeten Personen zurück. */
  autoDeregisterInactive: (eventId: string, people: Array<{ email: string; name: string }>) => Promise<Array<{ email: string; name: string }>>;
  /** v26.41: Kommunikations-Log (Event-Rundmails) eines Events — für die
   *  Teilnehmer-Ansicht unter „Meine Events" und die Organizer-Historie. */
  getEventComms: (eventId: string) => Promise<EventCommRow[]>;
  /** v24.59: Bereits benachrichtigte Teilnehmer-E-Mails (lowercase) eines Events
   *  — damit die Landing-Page schon-benachrichtigte Konten ausblenden kann. */
  getSentInactiveNotices: (eventId: string) => Promise<Set<string>>;
  /** v21: Archivierung — zählt archivreife Zeilen abgelaufener Events. */
  getArchivableCount: () => Promise<{ total: number; perList: Record<string, number> }>;
  /** v21: Archivierung — verschiebt archivreife Zeilen ins DEX_Archive.
   *  v22.2: shouldCancel = Abbruch-Check aus dem Fortschrittsmodal. */
  runArchiveExpired: (onProgress?: (listIdx: number, listTotal: number, listName: string, done: number, total: number) => void, shouldCancel?: () => boolean) => Promise<{ archived: number; failed: number; cancelled: boolean; perList: Record<string, number> }>;
  /** v24.33: Globales „Spalten fixen" über ALLE Events inkl. Sub-Events + Company-Backfill bestehender Teilnehmer. */
  /** v30.58: `details` sagt PRO EVENT, was gefehlt hat und was danach noch fehlt. */
  fixAllEventColumns: (onProgress?: (done: number, total: number, label: string) => void) => Promise<{ lists: number; columnsAdded: number; backfilled: number; errors: number; anyChange: boolean; details: FixColumnsDetail[] }>;
  // v30.39: Organizer-Rechte über ALLE Event-Bäume nachziehen (Klammer + Sub-Events).
  repairAllOrganizerPermissions: (onProgress?: (done: number, total: number, label: string) => void) => Promise<{ trees: number; sites: number; grants: number; unresolved: string[]; errors: number }>;
  /** v26.13: Versehentlich gelöschte Custom-Field-Beschreibungen aus der
   *  SharePoint-Versionshistorie wiederherstellen. */
  restoreCustomFieldDescriptions: (onProgress?: (done: number, total: number, label: string) => void, dryRun?: boolean) => Promise<{ events: number; eventsChanged: number; fieldsRestored: number; errors: number; details: Array<{ eventId: string; eventTitle: string; fields: Array<{ label: string; props: string[] }> }> }>;
  /** v23.40: Löschkonzept — zählt DEX_Archive-Einträge älter als 1 Monat (v23.48). */
  getDeletableArchiveCount: () => Promise<number>;
  /** v23.40: Löschkonzept — löscht DEX_Archive-Einträge älter als 1 Monat (v23.48). */
  runDeleteOldArchive: (onProgress?: (done: number, total: number) => void, shouldCancel?: () => boolean) => Promise<{ deleted: number; failed: number; cancelled: boolean }>;
  /** v26.32: Löschkonzept — Teilnehmerlisten 3 Monate nach Event-Ende. Events im
   *  Vorwarn-Fenster (3 Mon. − 1 Woche) bzw. fällige (≥3 Mon.) + Ausführung +
   *  automatische Vorwarn-Mail an die Organizer. */
  getParticipantDeletionWarnings: () => Promise<DeloitteEvent[]>;
  getParticipantDeletionDue: () => Promise<DeloitteEvent[]>;
  runParticipantDeletion: (onProgress?: (done: number, total: number, label: string) => void) => Promise<{ deleted: number; failed: number }>;
  maybeSendParticipantDeletionWarnings: () => Promise<void>;
  /** v26.33: Liest das Statistik-Archiv (DEX_EventStats) — KPIs gelöschter
   *  Teilnehmerlisten für die Admin-Center-Kachel „Statistik-Archiv". */
  getEventStats: () => Promise<EventStatsRow[]>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  /** v10.27: Split-Capacity-Gruppen-Wechsel für die eigene Registrierung.
   *  Nimmt die App-internen Wert-IDs ('Durchstarter' | 'Funstarter') —
   *  liefert zurück, ob der Wechsel direkt in die Ziel-Gruppe gehen
   *  konnte oder ob der User auf die Warteliste der Ziel-Gruppe gesetzt
   *  wurde (full=true). */
  switchSplitGroup: (eventId: string, newType: 'Durchstarter' | 'Funstarter') => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }>;
  /** v11.0: Item-Attachments einer Teilnehmer-Zeile listen / hochladen /
   *  löschen — der itemId ist in beiden Fällen die SharePoint-ID des
   *  jeweiligen Teilnehmer-Items in der Subsite. Im User-Flow nutzt die
   *  App für 'eigene Anmeldung' getMyRegistration, im Admin-Flow gibt
   *  AdminPage die fremde Item-ID direkt mit. */
  listMyEventAttachments: (eventId: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string }>>;
  uploadMyEventAttachment: (eventId: string, file: File) => Promise<boolean>;
  deleteMyEventAttachment: (eventId: string, fileName: string) => Promise<boolean>;
  // v19.0: Dokument-Custom-Felder (pro-Feld-Attachments).
  uploadFieldDocument: (eventId: string, fieldId: string, file: File, participantEmail?: string) => Promise<boolean>;
  listFieldDocuments: (eventId: string, fieldId: string, participantEmail?: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>>;
  deleteFieldDocument: (eventId: string, fileName: string, participantEmail?: string) => Promise<boolean>;
  getMyEventNumbers: () => Promise<{ registered: number[]; waitlisted: number[] }>;
  /** v28.22: Event-Nummern einer beliebigen Person (DEX_Participants, ohne
   *  Item-Level-Security der Teilnehmerlisten) — für die Doppel-Anmelde-
   *  Vorwarnung, die auch fremd angelegte Zeilen erkennt. */
  getEventNumbersForEmail: (email: string) => Promise<{ registered: number[]; waitlisted: number[] }>;
  /** v22.50: globale Teilnehmer-Liste (DEX_Participants) für die Admin-Suche.
   *  Berechtigungs-Scoping (nur Events, die der Nutzer verwalten darf) macht
   *  der Aufrufer. */
  getAllParticipants: () => Promise<SPParticipant[]>;
  refreshEvents: () => Promise<void>;
  refreshParticipantCounts: (eventId?: string) => Promise<void>;
  /** v24.73: Live-Plätze aus dem (für alle lesbaren) Counter — korrekte
   *  Aktiv-/Warteliste-Zahl auch für normale Teilnehmer. null = Counter nicht
   *  verfügbar (Aufrufer fällt auf event.currentParticipants zurück). */
  /** v30.62: `seatsKnown` = false heißt, der Platzzähler wurde für diese
   *  Subsite noch nie geschrieben. Die 0 darin ist dann kein Messwert. */
  getLiveCounterStats: (eventId: string) => Promise<CounterStats | null>;
  /** v24.74: Sitzplatz-Counter aller aktiven Kapazitäts-Events aus dem echten
   *  Bestand frischziehen (SeatsTaken + WaitlistTaken). Braucht Vollzugriff →
   *  beim Admin-Start aufrufen, damit die für alle lesbaren Counter stimmen,
   *  bevor Teilnehmer das Anmeldeformular öffnen. Best-effort, sequentiell. */
  reconcileCounters: (opts?: { onlyMine?: boolean }) => Promise<void>;
  /** v24.75: Echtzeit-Push auf eine Liste der Event-Subsite. kind='counter'
   *  (Anmeldeformular, für alle lesbar) / 'participants' (Organizer-Liste).
   *  Liefert eine Cleanup-Funktion. Best-effort. */
  subscribeEventRealtime: (eventId: string, kind: 'counter' | 'participants', onChange: () => void) => Promise<() => void>;
  markExpiredEventsAsCompleted: () => Promise<number>;
  sendAdminInquiry: (requesterName: string, requesterEmail: string, eventName: string, message: string, requesterLocation?: string, requesterJobTitle?: string) => Promise<boolean>;
  /** v26.67: Erinnerungs-Mail an eine „verwaiste" (Geister-)Anmeldung — die
   *  Person (bzw. die anmeldende Person bei Fremd-Anmeldung) hat eine Klammer-
   *  Zeile, aber kein Sub-Event ausgewählt; die Mail bittet sie, die Anmeldung
   *  abzuschließen, und verlinkt direkt die Anmeldeseite des Events. */
  sendCompleteRegistrationReminder: (args: { eventId: string; eventTitle: string; participantEmail: string; participantName: string; registeredByEmail?: string; registeredByName?: string }) => Promise<boolean>;
  /** v26.57: Approve-Mail an die Admins, wenn Personen AUSSERHALB von
   *  @deloitte.de zur Zielgruppe eines Events hinzugefügt wurden — der
   *  SharePoint ist im Default nur für Deloitte DE ALL freigeschaltet,
   *  internationale Kolleg:innen brauchen also zusätzlich Site-Zugriff. */
  notifyAdminsExternalAudienceAccess: (eventTitle: string, persons: string[], requesterName: string) => Promise<void>;
  /** v12.12: Admin-Aktion zum Re-Seed der Default-Email-Templates in
   *  DEX_EmailTemplates. Überschreibt die aktuelle Subject/Heading/BodyHtml
   *  jedes Standard-Templates mit den Default-Werten aus dem Code. */
  reseedDefaultEmailTemplates: () => Promise<ReseedSummary>;
  /** v24.98: Globaler Mail-Vorlagen-Editor — alle Templates lesen + einzelne
   *  Felder (Subject/Heading/Subheading/Farbe/Body) speichern. */
  getAllEmailTemplates: () => Promise<Array<{ id: number; templateType: string; language: string; subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string }>>;
  updateEmailTemplate: (id: number, fields: { subject?: string; heading?: string; subheading?: string; headingColor?: string; bodyHtml?: string }) => Promise<boolean>;
  /** v11.52: Gecachte KPI-Werte (Events + Teilnehmer) aus _Config lesen —
   *  ein einziger schneller REST-Call, für Boot-Loader-Anzeige. */
  getKpiCache: () => Promise<{ participants: number; events: number } | null>;
  /** v11.52: Frische KPI-Werte in _Config schreiben. Wird nach vollem
   *  App-Load im Hintergrund aufgerufen, damit nächster Boot frisch ist. */
  updateKpiCache: (v: { participants: number; events: number }) => Promise<boolean>;
  /** v26.4: Korrekte KPI-Gesamtwerte über ALLE Events (paginiert, nicht nur die
   *  geladenen 100) — Admin-Recompute für den „bisher genutzt für"-Boot-Zähler. */
  getKpiTotals: () => Promise<{ participants: number; events: number } | null>;
  /** v26.63: NUR die Events-Zahl neu berechnen — allein aus DEX_Events, ohne den
   *  teuren Subsite-Teilnehmer-Scan. Liefert die neue Events-Zahl oder null. */
  recomputeEventKpiOnly: () => Promise<number | null>;
  /**
   * Onboarding-Mail an einen frisch ernannten Organizer/Admin verschicken.
   * Cc geht automatisch an die DEX-Verantwortlichen, der Body wird ins
   * Deloitte-Layout gewrappt (siehe organizerOnboardingEmail in EmailTemplates).
   */
  sendOrganizerOnboarding: (recipientEmail: string, recipientName: string, role: 'Organizer' | 'Admin') => Promise<boolean>;
  // v9.21: Globaler TestTeam-State entfernt — Test-Team ist ab jetzt
  // per-Event (auf event.testTeamEmails). Die globalen Methoden bleiben
  // im EventService dormant für Backward-Compat.
}

export interface CreateEventInput {
  title: string;
  type: string;
  status: string;
  description: string;
  location: string;
  locationAddress?: string; // JSON: { street, houseNo, zip, city }
  locationFilter: string;
  audience: string;
  /** v16.4: Pre-resolved DL members (';'-separated, lowercase). */
  audienceResolvedEmails?: string;
  filterMode: string;
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  lastDeregisterDate: string;
  /** v29.19: Auto-Aktivierung (UTC-ISO) — wurde beim Anlegen bisher nicht
   *  persistiert, obwohl der Wizard das Feld anbietet. */
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
  contactInfo?: string;
  contactOrganizerEmail?: string; // v26.18
  outlookEventId: string;
  outlookBody: string;
  agenda?: string; // JSON-Array mit Agenda-Einträgen
  transfers?: string; // JSON-Array mit Transferzeiten
  documents?: string; // JSON-Array mit Dokumenten
  funZone?: string; // JSON-Array mit Quiz-Fragen
  quizClusterSize?: number; // 1..4 Fragen pro Quiz-Ansicht
  /** Seit v6.4: wenn gesetzt, wird das Event als Sub-Event zum angegebenen Parent angelegt. */
  parentEventId?: string;
  emailLanguage?: string;
  /** v18.35: erzwungene Anmeldeseiten-Sprache ('de' | 'en'); leer = App-Sprache. */
  registrationLanguage?: 'de' | 'en';
  /** v18.40: manueller Outlook-Termin-Ort; leer = Auto aus Veranstaltungsort + Adresse. */
  outlookLocation?: string;
  /** v29.52: ganztägiger Termin — der Outlook-Flow macht daraus isAllDay. */
  allDay?: boolean;
  /** v29.54: Termin als „Frei" statt „Beschäftigt" anzeigen. */
  showAsFree?: boolean;
  /** v29.55: Organizer nicht in den Outlook-Termin eintragen. */
  skipOrganizerInvite?: boolean;
  /** v30.26: Flow soll für diesen Termin einen Teams-Link erzeugen (isOnlineMeeting). */
  outlookIsOnlineMeeting?: boolean;
  /** v18.42: Betreff des Outlook-Termins; leer = Event-Titel. */
  outlookSubject?: string;
  /** v18.44: abweichende Outlook-Start/-Ende (ISO); leer = Event-Datum. */
  outlookStart?: string;
  outlookEnd?: string;
  emailTemplateOverrides?: string;
  disableEmails?: boolean;
  disableRegistrationEmail?: boolean; // v19.21: keine Anmelde-Bestätigung
  disableCancellationEmail?: boolean; // v19.21: keine Abmelde-Bestätigung
  autoDeregisterOnDecline?: boolean; // v19.23: Outlook-Absage = Auto-Abmeldung
  inactiveHandling?: 'notify' | 'autoderegister'; // v26.40
  disableOutlook?: boolean;
  outlookDirty?: boolean; // v11.57: Outlook-Update ausstehend nach Bearbeitung
  notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
  notifyOrgRegisterFromDate?: string;
  notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';
  excludedUsers?: string[];
  isFictive?: boolean;
  durchstarterCapacity?: number;
  funstarterCapacity?: number;
  splitLabelA?: string;
  splitLabelB?: string;
  splitDescA?: string; // v26.72: Beschreibung Gruppe A (mehrzeilig)
  splitDescB?: string; // v26.72: Beschreibung Gruppe B (mehrzeilig)
  splitHelpText?: string; // v26.83: Hinweistext über der Gruppen-Auswahl
  splitSectionTitle?: string; // v26.83: frei wählbare Überschrift der Gruppen-Auswahl
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
  /** v11.80: Team-Namen abfragen (Default false). */
  askTeamName?: boolean;
  /** v11.81: Auch Teil-Teams erlauben (Default false = nur komplette Teams). */
  teamPartialAllowed?: boolean;
  /** v11.81: Offene Slots öffentlich für Beitritt sichtbar (Default false). */
  teamOpenSlotsVisible?: boolean;
  /** v11.81: Beitritt erfordert Bestätigung durch Team-Kapitän (Default false). */
  teamJoinRequiresApproval?: boolean;
  /** v17.20: Custom-Fields zweisprachig (DE + EN) anbieten. */
  bilingualFields?: boolean;
  customFields: CustomField[];
  /** v11.69: Optional — wenn gesetzt zusammen mit `existingRegistrationListName`,
   *  wird keine neue Subsite angelegt. Stattdessen wird die mitgegebene Subsite
   *  an die neue DEX_Events-Zeile gekoppelt. Genutzt für "Outlook nachträglich
   *  aktivieren ohne Teilnehmer-Verlust" (siehe `deleteEventItemOnly`). */
  existingSubsiteUrl?: string;
  /** v11.69: Optional — Listenname der bereits bestehenden Teilnehmerliste in
   *  der wiederverwendeten Subsite (i.d.R. "Teilnehmer"). Muss zusammen mit
   *  `existingSubsiteUrl` gesetzt sein, damit der Reuse-Pfad greift. */
  existingRegistrationListName?: string;
  /** v11.87: Optionaler Progress-Callback. Wird zu Beginn jeder Teil-
   *  Operation aufgerufen, sodass die UI den Fortschrittsbalken und die
   *  Unter-Caption sichtbar bewegen kann. Stages decken die langsamen
   *  SP-Operationen ab (Subsite-Create, Teilnehmer-Liste, Permissions,
   *  Item-Insert). */
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
}
