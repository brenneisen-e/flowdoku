/**
 * Event Context - zentraler State fuer alle Events
 *
 * Laedt Events aus der SharePoint-Liste DEX_Events.
 * Erstellt die Liste automatisch beim ersten Start.
 * Verwaltet Registrierungen ueber Event-Subsites mit Teilnehmerlisten.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DeloitteEvent } from '../types';
import { EventService, SPEvent, CustomField, SPRegistration } from '../services/EventService';
import { registrationEmail, waitlistEmail, cancellationEmail, buildEmailFromTemplate, loadLogosAsBase64, wrapTemplate, organizerOnboardingEmail, qrCodeEmail, teamInfoBlockHtml } from '../services/EmailTemplates';
import * as QRCode from 'qrcode';
import { APP_VERSION } from '../version';

/**
 * Organizer-Namen fuer Mail-Anreden sauber formatieren:
 *   Input:  ['Sathasivam, Philipp', 'Oesterle, Ines']
 *   Output: 'Philipp Sathasivam und Ines Oesterle'  (bei DE)
 *           'Philipp Sathasivam and Ines Oesterle'  (bei EN)
 *
 * - Namen koennen auch ';'-separiert als Einzel-String kommen, wird gesplittet.
 * - Nachname/Vorname-Pairs werden vorgetauscht (SP-Default ist "Nachname, Vorname").
 * - Bei 1 Name: nur der Name. Bei 2: "A und B" / "A and B". Bei 3+: "A, B und C" / "A, B and C".
 */
/**
 * Wendet Event-spezifische Template-Overrides auf die globale SP-Vorlage an.
 *
 * - Override-JSON-Format: { "Anmeldung": { subject, heading, bodyHtml }, ... }
 * - Pro Feld gilt: Override > globale SP-Vorlage. headingColor bleibt immer
 *   die globale (Overrides aendern keine Brand-Farben).
 * - Wenn weder Override noch SP-Template existieren, gibt die Funktion null
 *   zurueck und der Caller faellt auf das Code-Default zurueck.
 */
export function applyEventTemplateOverride(
  spTemplate: { subject: string; headingColor: string; heading: string; bodyHtml: string } | null,
  overridesJson: string | undefined,
  templateType: string
): { subject: string; headingColor: string; heading: string; bodyHtml: string } | null {
  if (!overridesJson) return spTemplate;
  try {
    const all = JSON.parse(overridesJson) as Record<string, { subject?: string; heading?: string; bodyHtml?: string }>;
    const o = all[templateType];
    if (!o || (!o.subject && !o.heading && !o.bodyHtml)) return spTemplate;
    return {
      subject: o.subject || spTemplate?.subject || '',
      heading: o.heading || spTemplate?.heading || '',
      bodyHtml: o.bodyHtml || spTemplate?.bodyHtml || '',
      headingColor: spTemplate?.headingColor || '#86bc25',
    };
  } catch {
    return spTemplate;
  }
}

/**
 * Strip SharePoint-Note-Field-Wrapper.
 *
 * Seit der Migration der Felder Organizer + OrganizerEmail von Single-Line-Text
 * auf Note (Multi-Line-Text, Plain) — nötig wegen 255-Char-Limit bei 10+ Co-
 * Organizern — wickelt SharePoint die Werte beim REST-Read in einen
 * `<div class="ExternalClassXXXX">…</div>`-Container. Das passiert obwohl
 * `RichText: false` gesetzt ist und ist eine bekannte SP-Quirk.
 *
 * Folge ohne Strip: `(e.Organizer || '').split(';')` zerhackt den Wrapper an
 * den Semikolons, das erste und letzte Stück enthalten dann die Tag-Reste
 * `<div class="…">…` bzw. `…</div>` und landen so in den Chip-Labels.
 *
 * Idempotent: Eingaben ohne Wrapper bleiben unverändert.
 */
export function stripSpNoteWrapper(value: string | null | undefined): string {
  if (!value) return '';
  let v = value.trim();
  v = v.replace(/^<div\b[^>]*>/i, '');
  v = v.replace(/<\/div>\s*$/i, '');
  return v.trim();
}

export function formatOrganizerList(organizers: string[], lang: string): string {
  const names: string[] = [];
  for (const entry of organizers || []) {
    // Akzeptiere ';' UND ',' als Top-Level-Trenner zwischen Personen.
    // Wenn die Anzahl der Komma-Tokens gerade und >=2 ist, behandeln wir sie als
    // Paare ('Lastname, Firstname, Lastname, Firstname, ...'). Sonst fallen wir
    // zurueck auf Semikolon-Split + 'Lastname, Firstname' pro Stueck.
    const raw = (entry || '').trim();
    if (!raw) continue;
    const semiPieces = raw.split(';').map(p => p.trim()).filter(Boolean);
    const pieces: string[] = [];
    for (const sp of semiPieces) {
      const commaTokens = sp.split(',').map(s => s.trim()).filter(Boolean);
      if (commaTokens.length >= 4 && commaTokens.length % 2 === 0) {
        // Paarweise interpretieren: ['Last','First','Last','First',...]
        for (let i = 0; i < commaTokens.length; i += 2) {
          pieces.push(`${commaTokens[i]}, ${commaTokens[i + 1]}`);
        }
      } else {
        pieces.push(sp);
      }
    }
    for (const piece of pieces) {
      const commaParts = piece.split(',').map(s => s.trim());
      if (commaParts.length === 2 && commaParts[0] && commaParts[1]) {
        names.push(`${commaParts[1]} ${commaParts[0]}`);
      } else {
        names.push(piece);
      }
    }
  }
  if (names.length === 0) return '';
  if (names.length === 1) return names[0];
  const conj = (lang || 'EN').toUpperCase() === 'DE' ? ' und ' : ' and ';
  if (names.length === 2) return `${names[0]}${conj}${names[1]}`;
  return `${names.slice(0, -1).join(', ')}${conj}${names[names.length - 1]}`;
}

interface EventContextType {
  events: DeloitteEvent[];
  /** Top-Level-Events (ohne parentEventId) — was in EventListPage/MyEventsPage angezeigt wird. */
  topLevelEvents: DeloitteEvent[];
  /** Kind-Events eines Parents (Sub-Events / Trainingssessions), sortiert nach StartDate. */
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  isEventsLoading: boolean;
  createEvent: (event: CreateEventInput) => Promise<number | null>;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string) => Promise<boolean>;
  /** v11.82: Team-Anmeldung — Lead + N-1 Mitglieder gleichzeitig anmelden.
   *  Reserviert N Plaetze atomar; bei Vollbelegung geht das ganze Team auf
   *  die Warteliste (keine Teil-Anmeldungen aus Kapazitaetsmangel). */
  registerTeam: (
    eventId: string,
    leadData: { firstName: string; lastName: string; email: string; salutation?: string; customData: Record<string, string>; preferredStarterType?: string },
    members: Array<{ email: string; displayName: string }>,
    teamName: string | undefined
  ) => Promise<{ ok: boolean; teamId?: string; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.82: Andere Team-Mitglieder zu einer Registrierung laden — fuer das
   *  Team-Badge in „Meine Events". */
  getTeamMembers: (eventId: string, teamId: string) => Promise<SPRegistration[]>;
  /** v11.83: Ein Team-Lead kann nachtraeglich ein einzelnes Mitglied
   *  zum bereits angemeldeten Team hinzufuegen (Plus-Button in MyEvents).
   *  Atomar einen Sitzplatz reservieren, neuen Member-Eintrag anlegen,
   *  Bestaetigungs-Mail + Outlook-Termin queuen. */
  addTeamMember: (eventId: string, teamId: string, teamName: string | undefined, member: { email: string; displayName: string }) => Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.83: Direkter Team-Beitritt aus der Anmeldeseite (wenn der
   *  Organizer "Beitritt erfordert Bestaetigung" NICHT aktiviert hat).
   *  Verhalten wie `addTeamMember`, aber laeuft mit dem eingeloggten User
   *  selbst als neuem Member. */
  joinTeam: (eventId: string, teamId: string, teamName: string | undefined) => Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.84: Team-Lead-Rolle innerhalb eines Teams uebergeben — nur im
   *  Admin Center fuer Admin/Organizer eigener Events sichtbar. Setzt die
   *  alte Lead-Zeile auf TeamLead=false und die neue auf TeamLead=true,
   *  schickt anschliessend eine Info-Mail an alle aktiven Mitglieder. */
  transferTeamLead: (eventId: string, teamId: string, newLeadEmail: string) => Promise<{ ok: boolean; reason?: string }>;
  /** v11.83: Beitritts-Anfrage in DEX_TeamJoinRequests einreichen — fuer
   *  Events bei denen der Organizer Approval aktiviert hat. */
  createTeamJoinRequest: (eventId: string, teamId: string) => Promise<{ ok: boolean; itemId?: number; reason?: string }>;
  /** v11.83: Pending-Beitritts-Anfragen abrufen (nur fuer den
   *  eingeloggten User als Team-Lead — Filter auf TeamId, das er selber
   *  fuehrt). */
  listTeamJoinRequestsForEvent: (eventId: string, teamId: string) => Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string }>>;
  /** v11.83: Approval-/Reject-Entscheidung eines Leads — bei „Approved"
   *  legt die App die Member-Anmeldung an und queued Mails; bei
   *  „Rejected" eine kurze Absage-Mail an den Anfragenden. */
  decideTeamJoinRequest: (requestId: number, decision: 'Approved' | 'Rejected') => Promise<boolean>;
  /** v11.83: Liste der Teams (gruppiert nach TeamId) eines Events fuer
   *  die „Offene Teams"-Anzeige auf der Registrierungs-Seite. Nur Teams
   *  mit aktivem Mitglied-Count < TeamSize werden aufgefuehrt. */
  listOpenTeamsForEvent: (eventId: string) => Promise<Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }>>;
  cancelRegistration: (eventId: string) => Promise<boolean>;
  /** v11.86: Ein Team-Lead meldet ueber „Team verwalten" stellvertretend
   *  ein Team-Mitglied vom Event ab. Audit-Felder (CancelledBy*) werden
   *  mit dem eingeloggten Lead gefuellt, danach laeuft derselbe
   *  Team-Post-Step wie beim Self-Cancel (Info-Mails an die uebrigen
   *  Mitglieder; Auto-Promote nicht relevant, weil der Lead sich nicht
   *  selbst loescht). */
  cancelTeamMember: (
    eventId: string,
    memberRegistration: SPRegistration
  ) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  checkRegistrationByEmail: (eventId: string, email: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
  deleteEvent: (eventId: string) => Promise<boolean>;
  /** v11.69: Loescht NUR das DEX_Events-Listenitem, ohne Subsite-/Teilnehmer-
   *  Liste-Recycle und ohne Outlook-DeleteEvent-Queue. Wird gebraucht, um ein
   *  Sub-Event mit `existingSubsiteUrl` neu anzulegen, damit der
   *  `DEX_CreateOutlookEvent`-Flow triggert — die alte Subsite mit
   *  Anmeldungen bleibt erhalten. */
  deleteEventItemOnly: (eventId: string) => Promise<boolean>;
  updateEvent: (eventId: string, updates: Record<string, unknown>) => Promise<boolean>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  /** v10.27: Split-Capacity-Gruppen-Wechsel für die eigene Registrierung.
   *  Nimmt die App-internen Wert-IDs ('Durchstarter' | 'Funstarter') —
   *  liefert zurück, ob der Wechsel direkt in die Ziel-Gruppe gehen
   *  konnte oder ob der User auf die Warteliste der Ziel-Gruppe gesetzt
   *  wurde (full=true). */
  switchSplitGroup: (eventId: string, newType: 'Durchstarter' | 'Funstarter') => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }>;
  /** v11.0: Item-Attachments einer Teilnehmer-Zeile listen / hochladen /
   *  loeschen — der itemId ist in beiden Fällen die SharePoint-ID des
   *  jeweiligen Teilnehmer-Items in der Subsite. Im User-Flow nutzt die
   *  App fuer 'eigene Anmeldung' getMyRegistration, im Admin-Flow gibt
   *  AdminPage die fremde Item-ID direkt mit. */
  listMyEventAttachments: (eventId: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string }>>;
  uploadMyEventAttachment: (eventId: string, file: File) => Promise<boolean>;
  deleteMyEventAttachment: (eventId: string, fileName: string) => Promise<boolean>;
  getMyEventNumbers: () => Promise<{ registered: number[]; waitlisted: number[] }>;
  refreshEvents: () => Promise<void>;
  refreshParticipantCounts: (eventId?: string) => Promise<void>;
  markExpiredEventsAsCompleted: () => Promise<number>;
  sendAdminInquiry: (requesterName: string, requesterEmail: string, eventName: string, message: string) => Promise<boolean>;
  /** v11.52: Gecachte KPI-Werte (Events + Teilnehmer) aus _Config lesen —
   *  ein einziger schneller REST-Call, fuer Boot-Loader-Anzeige. */
  getKpiCache: () => Promise<{ participants: number; events: number } | null>;
  /** v11.52: Frische KPI-Werte in _Config schreiben. Wird nach vollem
   *  App-Load im Hintergrund aufgerufen, damit naechster Boot frisch ist. */
  updateKpiCache: (v: { participants: number; events: number }) => Promise<boolean>;
  /**
   * Onboarding-Mail an einen frisch ernannten Organizer/Admin verschicken.
   * Cc geht automatisch an die DEX-Verantwortlichen, der Body wird ins
   * Deloitte-Layout gewrappt (siehe organizerOnboardingEmail in EmailTemplates).
   */
  sendOrganizerOnboarding: (recipientEmail: string, recipientName: string, role: 'Organizer' | 'Admin') => Promise<boolean>;
  // v9.21: Globaler TestTeam-State entfernt — Test-Team ist ab jetzt
  // per-Event (auf event.testTeamEmails). Die globalen Methoden bleiben
  // im EventService dormant fuer Backward-Compat.
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
  filterMode: string;
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  lastDeregisterDate: string;
  maxParticipants: number;
  waitlistEnabled: boolean;
  eventImageUrl: string;
  organizer: string;
  organizerEmail: string;
  /** v10.16: optionaler Ansprechpartner (Anzeige-Feld). */
  contactName?: string;
  contactEmail?: string;
  contactInfo?: string;
  outlookEventId: string;
  outlookBody: string;
  agenda?: string; // JSON-Array mit Agenda-Eintraegen
  transfers?: string; // JSON-Array mit Transferzeiten
  documents?: string; // JSON-Array mit Dokumenten
  funZone?: string; // JSON-Array mit Quiz-Fragen
  quizClusterSize?: number; // 1..4 Fragen pro Quiz-Ansicht
  /** Seit v6.4: wenn gesetzt, wird das Event als Sub-Event zum angegebenen Parent angelegt. */
  parentEventId?: string;
  emailLanguage?: string;
  emailTemplateOverrides?: string;
  disableEmails?: boolean;
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
  splitSharedWaitlist?: boolean;
  allowAttendeeUpload?: boolean;
  attendeeUploadHint?: string;
  attendeeUploadLabel?: string;
  /** v11.80: Anrede im Registrierungsformular abfragen (Default false). */
  askSalutation?: boolean;
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
  customFields: CustomField[];
  /** v11.69: Optional — wenn gesetzt zusammen mit `existingRegistrationListName`,
   *  wird keine neue Subsite angelegt. Stattdessen wird die mitgegebene Subsite
   *  an die neue DEX_Events-Zeile gekoppelt. Genutzt fuer "Outlook nachtraeglich
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

export const EventContext = React.createContext<EventContextType | undefined>(undefined);

export function EventProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [events, setEvents] = React.useState<DeloitteEvent[]>([]);
  const [isEventsLoading, setIsEventsLoading] = React.useState(true);
  // Map von EventId -> SubsiteUrl fuer schnellen Zugriff
  const subsiteMap = React.useRef<Record<string, string>>({});

  const eventService = React.useMemo(() => new EventService(props.context), []);

  // v9.16/v9.21: Test-Team war kurz global (TestTeamEmails in _Config),
  // ist jetzt per-Event (event.testTeamEmails). Globaler State raus.
  const currentUserEmail = props.context.pageContext.user.email;
  const currentUserName = props.context.pageContext.user.displayName;
  // Vorname fuer E-Mail-Anreden ({{Name}} im Template).
  // Deloitte-displayName ist "Nachname, Vorname" (mit Komma) -> Teil nach Komma.
  // Fallback: displayName ohne Komma -> erstes Wort (vereinzelt "Vorname Nachname").
  const getFirstName = (displayName: string): string => {
    if (!displayName) return '';
    const commaIdx = displayName.indexOf(',');
    if (commaIdx >= 0) return displayName.substring(commaIdx + 1).trim().split(/\s+/)[0];
    return displayName.trim().split(/\s+/)[0];
  };
  const currentUserFirstName = getFirstName(currentUserName);

  React.useEffect(() => {
    initEvents().catch(() => setIsEventsLoading(false));
  }, []);

  async function initEvents(): Promise<void> {
    // v11.74: Performance-Profiling — misst jede Boot-Phase und gibt am
    // Ende eine sortierte Tabelle aus. Nur in der Console sichtbar
    // (DevTools → Console), kein UI-Impact. Hilft beim Identifizieren
    // der echten Bottlenecks (ensure-Listen vs. getEvents vs. counts
    // vs. attachments).
    //
    // v11.76: Schema-Ensure-Gate. Die idempotenten `ensure*`/`upgrade*`-
    // Wartungs-Calls müssen NICHT bei jedem Page-Load laufen. Sie sind nur
    // beim ersten Boot nach einer App-Version mit Schema-Änderungen nötig.
    // Wir verwenden einen versions-gebundenen localStorage-Key — sobald
    // wir die Version 11.76 erfolgreich durchgeschmurgelt haben, sparen
    // wir uns alle ensure-Calls beim nächsten Boot.
    //
    // Außerdem: wenn die ensure-Calls DOCH laufen, parallelisieren wir sie
    // (Stage 1 = ensureEventsList alleine; Stage 2 = alles andere parallel
    // via Promise.allSettled), statt sie sequentiell hintereinander zu
    // ketten. Spart bei 11 Calls und ca. 6.7 s seriell ca. 4-5 s.
    const ENSURE_FLAG_KEY = 'dex.schema.ensured.v' + APP_VERSION;
    let skipEnsure = false;
    try {
      if (typeof window !== 'undefined' && window.localStorage.getItem(ENSURE_FLAG_KEY) === '1') {
        skipEnsure = true;
      }
    } catch { /* localStorage disabled */ }

    const perfMarks: Array<{ name: string; ms: number }> = [];
    const tBoot = performance.now();
    const stage = async (name: string, fn: () => Promise<unknown>): Promise<void> => {
      const t0 = performance.now();
      try { await fn(); } catch { /* swallow — matches existing try/catch pattern */ }
      perfMarks.push({ name, ms: Math.round(performance.now() - t0) });
    };
    const safeRun = async (name: string, fn: () => Promise<unknown>, acc: Array<{ name: string; ms: number }>): Promise<void> => {
      const t0 = performance.now();
      try { await fn(); } catch { /* swallow */ }
      acc.push({ name, ms: Math.round(performance.now() - t0) });
    };

    if (!skipEnsure) {
      // Stage 1: DEX_Events anlegen/sichern (Listen-Erstellung muss als erstes;
      // die upgrade*-Calls operieren auf DEX_Events).
      await stage('ensureEventsList', () => eventService.ensureEventsList());
      // Stage 2: alles andere parallel — keine inter-Abhaengigkeiten.
      const parallelMarks: Array<{ name: string; ms: number }> = [];
      const tPar = performance.now();
      // Hinweis: safeRun() swallowt Exceptions intern und resolved IMMER. Daher
      // ist Promise.all hier sicher (kein early-reject) und auch in ES2018-
      // Targets verfügbar — Promise.allSettled wäre erst ab ES2020.
      await Promise.all([
        safeRun('upgradeAudienceFieldToNote', () => eventService.upgradeAudienceFieldToNote(), parallelMarks),
        safeRun('upgradeOrganizerFieldsToNote', () => eventService.upgradeOrganizerFieldsToNote(), parallelMarks),
        safeRun('ensureEmailsList', () => eventService.ensureEmailsList(), parallelMarks),
        safeRun('ensureOutlookList', () => eventService.ensureOutlookList(), parallelMarks),
        safeRun('ensureParticipantsList', () => eventService.ensureParticipantsList(), parallelMarks),
        safeRun('ensureEmailTemplatesList', () => eventService.ensureEmailTemplatesList(), parallelMarks),
        safeRun('ensureIDReorderList', () => eventService.ensureIDReorderList(), parallelMarks),
        safeRun('ensureChangeLogList', () => eventService.ensureChangeLogList(), parallelMarks),
        safeRun('ensureTeamJoinRequestsList', () => eventService.ensureTeamJoinRequestsList(), parallelMarks),
        safeRun('ensureAssetsFolders', () => eventService.ensureAssetsFolders(), parallelMarks),
        safeRun('ensureLogosInConfig', () => eventService.ensureLogosInConfig(), parallelMarks),
      ]);
      const dPar = Math.round(performance.now() - tPar);
      // eslint-disable-next-line no-console
      console.log(`[DEX][perf][boot] ensure-parallel-stage = ${dPar} ms`);
      // Einzelne Sub-Zeiten in die Gesamt-Tabelle übernehmen.
      for (const m of parallelMarks) perfMarks.push(m);
      // Erfolg markieren — nächster Boot überspringt die ensure-Calls.
      try {
        if (typeof window !== 'undefined') window.localStorage.setItem(ENSURE_FLAG_KEY, '1');
      } catch { /* localStorage disabled */ }
    }

    // loadLogosAsBase64 ist KEIN ensure-Call — es füllt den In-Memory-Cache
    // mit den Logo-Daten, die für Mail-/Outlook-Templates gebraucht werden.
    // Muss bei jedem Boot laufen.
    await stage('loadLogosAsBase64', () => loadLogosAsBase64(props.context.spHttpClient, eventService.siteUrl));
    await stage('loadEvents (full chain)', () => loadEvents());
    setIsEventsLoading(false);
    const tTotal = Math.round(performance.now() - tBoot);
    const sorted = [...perfMarks].sort((a, b) => b.ms - a.ms);
    if (skipEnsure) {
      // eslint-disable-next-line no-console
      console.log(`[DEX][perf][boot] total = ${tTotal} ms (schema-ensure SKIPPED, version=v${APP_VERSION})`);
    } else {
      // eslint-disable-next-line no-console
      console.log(`[DEX][perf][boot] total = ${tTotal} ms (schema-ensure RAN, version=v${APP_VERSION})`);
    }
    // eslint-disable-next-line no-console
    console.table(sorted.map(m => ({ stage: m.name, ms: m.ms })));
  }

  async function loadEvents(): Promise<void> {
    // v11.74: Sub-Phase-Profiling — getEvents vs. Mapping vs. Counts vs. Attachments.
    const tGet = performance.now();
    const spEvents = await eventService.getEvents();
    const dGet = Math.round(performance.now() - tGet);
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][loadEvents] getEvents = ${dGet} ms (n=${spEvents.length})`);
    const tMap = performance.now();
    // v9.41: jedes Event-Mapping einzeln in try/catch wrappen — wenn EIN
    // Event-Mapping fehlschlägt (z.B. weil eine frisch erstellte Subsite noch
    // nicht API-konsistent ist), kippt nicht die ganze Eventliste in einen
    // Fehlerzustand. Stattdessen wird der einzelne kaputte Event ausgelassen
    // und beim nächsten Refresh erneut versucht. (Kein Promise.allSettled
    // benutzt, weil die SPFx-tsconfig auf ES2018 steht.)
    const safeMapped = await Promise.all(spEvents.map(async (e) => {
      try {
        return await mapSPEventToDeloitteEvent(e);
      } catch (err) {
        console.warn('[DEX] mapSPEventToDeloitteEvent fehlgeschlagen für Event', e?.Id, err);
        return null;
      }
    }));
    const mapped = safeMapped.filter((x): x is DeloitteEvent => x !== null);
    const dMap = Math.round(performance.now() - tMap);
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][loadEvents] mapSPEventToDeloitteEvent x ${spEvents.length} = ${dMap} ms`);
    // Teilnehmerzahlen fuer alle Events mit Subsite laden
    const tCnt = performance.now();
    const withCounts = await loadParticipantCountsForEvents(mapped);
    const dCnt = Math.round(performance.now() - tCnt);
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][loadEvents] participantCounts = ${dCnt} ms`);
    // Attachments (Dokumente) fuer alle Events laden
    const tAtt = performance.now();
    const withDocs = await Promise.all(withCounts.map(async (evt) => {
      try {
        const attachments = await eventService.getEventAttachments(Number(evt.id));
        return { ...evt, documents: attachments };
      } catch { return evt; }
    }));
    const dAtt = Math.round(performance.now() - tAtt);
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][loadEvents] attachments x ${mapped.length} = ${dAtt} ms`);
    setEvents(withDocs);
  }

  async function loadParticipantCountsForEvents(evts: DeloitteEvent[]): Promise<DeloitteEvent[]> {
    const results = await Promise.all(
      evts.map(async (evt) => {
        if (!evt.subsiteUrl) return evt;
        try {
          const counts = await eventService.getRegistrationCount(evt.subsiteUrl);
          return { ...evt, currentParticipants: counts.registered, waitlistCount: counts.waitlist };
        } catch {
          return evt;
        }
      })
    );
    return results;
  }

  async function refreshParticipantCounts(eventId?: string): Promise<void> {
    if (eventId) {
      const subsiteUrl = subsiteMap.current[eventId];
      if (!subsiteUrl) return;
      try {
        const counts = await eventService.getRegistrationCount(subsiteUrl);
        setEvents(current =>
          current.map(e =>
            e.id === eventId ? { ...e, currentParticipants: counts.registered, waitlistCount: counts.waitlist } : e
          )
        );
      } catch { /* default bleibt */ }
    } else {
      setEvents(current => {
        loadParticipantCountsForEvents(current).then(updated => setEvents(updated)).catch(() => { /* ignore */ });
        return current;
      });
    }
  }

  async function mapSPEventToDeloitteEvent(e: SPEvent): Promise<DeloitteEvent> {
    // SubsiteUrl merken
    if (e.SubsiteUrl) {
      subsiteMap.current[e.Id.toString()] = e.SubsiteUrl;
    }

    // Teilnehmeranzahl: default 0, wird lazy geladen wenn User ein Event oeffnet
    const currentParticipants = 0;
    const waitlistCount = 0;

    // Custom Fields parsen
    let customFields: CustomField[] = [];
    try {
      if (e.CustomFields) customFields = JSON.parse(e.CustomFields);
    } catch { /* ungueltig */ }
    // v11.18: Debug-Trace fuer den helpText-Roundtrip — den rohen SP-String
    // logge ich direkt aus, damit wir sehen koennen ob helpText/onlyForGroup
    // tatsaechlich in dem zurueckkommenden JSON drin sind. Wenn ja → das
    // Wizard-Loadmapping verschluckt sie. Wenn nicht → SP hat sie beim
    // Save gar nicht erst gespeichert.
    if (typeof e.CustomFields === 'string' && e.CustomFields.indexOf('helpText') >= 0) {
      // Nur ausfuehrlich loggen wenn das Event tatsaechlich helpText
      // beinhaltet — sonst lautes Logging fuer alle alten Events.
      // eslint-disable-next-line no-console
      console.log('[DEX][load] Raw CustomFields for event', e.Id, e.Title, ':\n', e.CustomFields);
    }

    return {
      id: e.Id.toString(),
      eventNumber: e.EventNumber || 0,
      title: e.Title || '',
      // v5.2: EventType-Spalte deprecated. Typ aus CustomFields ableiten
      // (Fallback auf alten SP-Wert wenn noch vorhanden).
      type: (e.EventType as DeloitteEvent['type'])
        || (customFields.some(f => f.id === 'b2run_startblock') ? 'B2Run' : 'Other'),
      status: (e.EventStatus as DeloitteEvent['status']) || 'Under Construction',
      organizers: (stripSpNoteWrapper(e.Organizer) || '').split(';').map((s: string) => s.trim()).filter((s: string) => s),
      organizerEmails: (stripSpNoteWrapper(e.OrganizerEmail) || '').split(';').map((s: string) => s.trim()).filter((s: string) => s),
      // v10.16: Optionaler Ansprechpartner. ContactInfo ist Note-Feld, daher
      // strippen — Name/Email sind Single-Line, kein Wrapper.
      contactName: e.ContactName || '',
      contactEmail: e.ContactEmail || '',
      contactInfo: stripSpNoteWrapper(e.ContactInfo),
      location: e.Location || '',
      locationAddress: (() => {
        try {
          if (!e.LocationAddress) return undefined;
          const o = JSON.parse(e.LocationAddress);
          return { street: o.street || '', houseNo: o.houseNo || '', zip: o.zip || '', city: o.city || '' };
        } catch { return undefined; }
      })(),
      locationAudience: e.LocationFilter ? e.LocationFilter.split(',').map(s => s.trim()) : [],
      audienceFilter: e.Audience ? e.Audience.split(',').map(s => s.trim()) : [],
      filterMode: (e.FilterMode as 'AND' | 'OR') || 'OR',
      startDate: e.StartDate || '',
      endDate: e.EndDate || '',
      registrationDeadline: e.RegistrationDeadline || '',
      lastDeregisterDate: e.LastDeregisterDate || '',
      description: e.Description || '',
      maxParticipants: e.MaxParticipants || 0,
      waitlistEnabled: e.WaitlistEnabled !== false, // default true wenn null/undefined
      autoSendQRCode: e.AutoSendQRCode === true, // v9.15 — explizites opt-in pro Event
      activeFrom: e.ActiveFrom || undefined, // v9.21 — Auto-Activate-Datum
      currentParticipants,
      waitlistCount,
      imageUrl: e.EventImageUrl || '',
      subsiteUrl: e.SubsiteUrl || '',
      outlookBody: e.OutlookBody || '',
      outlookEventId: e.OutlookEventId || '',
      // v11.61: CalendarLink (iCalUId) muss in den Event-Type, weil der
      // DEX_CreateOutlookEvent-Flow nur dieses Feld auf Erfolg setzt — die
      // v11.57-Modal-Erkennung hatte auf OutlookEventId geprueft (immer leer)
      // und das Outlook-Update-Confirm-Modal kam deshalb nie.
      calendarLink: e.CalendarLink || '',
      emailLanguage: e.EmailLanguage || 'EN',
      emailTemplateOverrides: e.EmailTemplateOverrides || '',
      disableEmails: !!e.DisableEmails,
      disableOutlook: !!e.DisableOutlook,
      // v11.57: bei alten Tenants kann die SP-Spalte fehlen — undefined wird
      // als false interpretiert (kein Hinweis anzeigen).
      outlookDirty: !!e.OutlookDirty,
      notifyOrgRegisterMode: ((): 'never' | 'always' | 'fromDate' => {
        const v = (e.NotifyOrgRegisterMode || '').toLowerCase();
        if (v === 'always') return 'always';
        if (v === 'fromdate') return 'fromDate';
        return 'never';
      })(),
      notifyOrgRegisterFromDate: e.NotifyOrgRegisterFromDate || '',
      notifyOrgCancelMode: ((): 'never' | 'always' | 'afterDeadline' => {
        const v = (e.NotifyOrgCancelMode || '').toLowerCase();
        if (v === 'always') return 'always';
        if (v === 'afterdeadline') return 'afterDeadline';
        return 'never';
      })(),
      excludedUsers: (e.ExcludedUsers || '').split(';').map((s: string) => s.trim().toLowerCase()).filter(Boolean),
      isFictive: !!e.IsFictive,
      durchstarterCapacity: typeof e.DurchstarterCapacity === 'number' ? e.DurchstarterCapacity : undefined,
      funstarterCapacity: typeof e.FunstarterCapacity === 'number' ? e.FunstarterCapacity : undefined,
      splitLabelA: e.SplitLabelA || undefined,
      splitLabelB: e.SplitLabelB || undefined,
      splitSharedWaitlist: !!e.SplitSharedWaitlist,
      allowAttendeeUpload: !!e.AllowAttendeeUpload,
      attendeeUploadHint: e.AttendeeUploadHint || undefined,
      attendeeUploadLabel: e.AttendeeUploadLabel || undefined,
      // v11.80: Anrede-Toggle + Team-Anmelde-Konfiguration durchreichen.
      // Alte Tenants ohne diese Spalten interpretieren undefined als false /
      // 0, das passt zum Default-Verhalten (Anrede aus, Team-Anmeldung aus).
      askSalutation: !!e.AskSalutation,
      teamRegistrationEnabled: !!e.TeamRegistrationEnabled,
      teamSize: typeof e.TeamSize === 'number' ? e.TeamSize : 0,
      askTeamName: !!e.AskTeamName,
      // v11.81: Erweiterte Team-Anmelde-Konfiguration (Beitritts-Modus).
      // Alte Tenants ohne diese Spalten interpretieren undefined als false
      // — das deckt sich mit dem konservativen Default „Nur komplette Teams,
      // keine offenen Slots, keine Approval-Queue".
      teamPartialAllowed: !!e.TeamPartialAllowed,
      teamOpenSlotsVisible: !!e.TeamOpenSlotsVisible,
      teamJoinRequiresApproval: !!e.TeamJoinRequiresApproval,
      // v6.15: Extra-B2Run-Config aus EmailTemplateOverrides._b2run (piggyback in
      // der bestehenden JSON-Struktur, keine neue SP-Spalte nötig).
      // v6.19: QR-Code-Scanner-Liste aus EmailTemplateOverrides._qrScanners (piggyback).
      // v9.18: Co-Organizer-Liste aus EmailTemplateOverrides._coOrganizers (piggyback, gleicher Pattern).
      // v9.21: Test-Team-Liste aus EmailTemplateOverrides._testTeam (per-Event statt global).
      ...(() => {
        try {
          const parsed = JSON.parse(e.EmailTemplateOverrides || '{}');
          if (!parsed || typeof parsed !== 'object') return { qrScannerNames: [], qrScannerEmails: [], coOrganizerNames: [], coOrganizerEmails: [], testTeamNames: [], testTeamEmails: [] };
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const b = (parsed as any)._b2run;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const qr = (parsed as any)._qrScanners;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const co = (parsed as any)._coOrganizers;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const tt = (parsed as any)._testTeam;
          // v11.25: pure Display-Reihenfolge-Umkehr fuer Split-Capacity-Karten.
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const splitDispRev = !!(parsed as any)._splitDisplayOrderReversed;
          const b2Part = b && typeof b === 'object' ? {
            durchstarterStartblock: typeof b.durchstarterStartblock === 'string' ? b.durchstarterStartblock : undefined,
            funstarterStartblock: typeof b.funstarterStartblock === 'string' ? b.funstarterStartblock : undefined,
            durchstarterRequiresProof: !!b.durchstarterRequiresProof,
          } : {};
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const qrNames: string[] = Array.isArray(qr) ? qr.map((x: any) => String(x?.name || '')) : [];
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const qrEmails: string[] = Array.isArray(qr) ? qr.map((x: any) => String(x?.email || '')) : [];
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const coNames: string[] = Array.isArray(co) ? co.map((x: any) => String(x?.name || '')) : [];
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const coEmails: string[] = Array.isArray(co) ? co.map((x: any) => String(x?.email || '')) : [];
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ttNames: string[] = Array.isArray(tt) ? tt.map((x: any) => String(x?.name || '')) : [];
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ttEmails: string[] = Array.isArray(tt) ? tt.map((x: any) => String(x?.email || '')) : [];
          return { ...b2Part, splitDisplayOrderReversed: splitDispRev, qrScannerNames: qrNames, qrScannerEmails: qrEmails, coOrganizerNames: coNames, coOrganizerEmails: coEmails, testTeamNames: ttNames, testTeamEmails: ttEmails };
        } catch { return { qrScannerNames: [], qrScannerEmails: [], coOrganizerNames: [], coOrganizerEmails: [], testTeamNames: [], testTeamEmails: [] }; }
      })(),
      agenda: (() => { try { return e.Agenda ? JSON.parse(e.Agenda) : []; } catch { return []; } })(),
      transferTimes: (() => { try { return e.Transfers ? JSON.parse(e.Transfers) : []; } catch { return []; } })(),
      quiz: (() => { try { return e.FunZone ? JSON.parse(e.FunZone) : []; } catch { return []; } })(),
      quizClusterSize: typeof e.QuizClusterSize === 'number' && e.QuizClusterSize >= 1 ? e.QuizClusterSize : undefined,
      parentEventId: e.ParentEventId || undefined,
      documents: [], // Wird per loadAttachments nachgeladen
      eventSpecificFields: customFields.map(cf => ({
        id: cf.id,
        label: cf.label,
        type: cf.type,
        required: cf.required,
        options: cf.options,
        // v7.20: helpText durchreichen, damit das Registrierungsformular ihn
        // im "i"-Tooltip neben dem Label anzeigen kann.
        helpText: cf.helpText || '',
        // v7.21: showIf-Bedingung durchreichen — RegistrationPage filtert
        // anhand davon, ob das Feld angezeigt wird.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        showIf: (cf as any).showIf,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        spInternalName: (cf as any).spInternalName || '',
        // v7.11: multi-Flag durchreichen, damit RegistrationPage Mehrfachauswahl rendern kann
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        multi: !!(cf as any).multi,
        // externalLinks ebenfalls durchreichen, damit AGB-Links fuer B2Run-Datenschutz
        // korrekt unter dem Feld angezeigt werden (war bisher nur ueber den Fallback in
        // RegistrationPage abgesichert).
        externalLinks: cf.externalLinks,
        // v11.16: onlyForGroup aus dem persistierten Feld durchreichen.
        // Wurde im Wizard sauber gespeichert (CustomFields-JSON enthaelt
        // den Schluessel), aber der Loader hat ihn nie zurueckgelesen —
        // Folge: die Gruppen-spezifische Sichtbarkeit (Funstarter only /
        // Durchstarter only) hat in der Registrierungs-UI nie gegriffen,
        // weil die Filter-Chain auf undefined gefallen ist.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        onlyForGroup: (cf as any).onlyForGroup,
      })),
    };
  }

  async function createEvent(input: CreateEventInput): Promise<number | null> {
    const eventId = await eventService.createEvent(input);
    if (eventId) {
      // v9.0: Audit-Log (fire-and-forget — Save-Flow nicht blocken)
      eventService.writeChangeLog({
        action: 'EventCreated',
        targetType: 'Event',
        targetId: String(eventId),
        targetName: input.title,
        eventId: String(eventId),
        eventTitle: input.title,
        details: { eventType: input.type, location: input.location, startDate: input.startDate, maxParticipants: input.maxParticipants },
      }).catch(() => { /* */ });
      // v11.53: KPI-Counter sofort hochzaehlen — nur fuer nicht-fictive Events
      // (Test-Events zaehlen nicht in der LandingPage-KPI).
      if (!input.isFictive) {
        eventService.bumpKpiEvents(1).catch(() => { /* best-effort */ });
      }
      // v9.41: KEIN Auto-Refresh mehr direkt nach Create. Grund: SharePoint braucht
      // einige Sekunden, bis die frisch angelegte Subsite + Teilnehmerliste +
      // DEX_TeilnehmerCounter-Liste API-seitig konsistent abrufbar sind. Wenn wir
      // hier sofort getEvents() + mapSPEventToDeloitteEvent() laufen lassen, fallen
      // die Subsite-Reads mit 400/404 ins Leere und das nachfolgende Event-List-
      // Rendering kann in eine Render-Loop laufen (React #300 → weißer Screen).
      //
      // Stattdessen wird der Refresh erst getriggert, wenn der User auf der Success-
      // Seite auf 'Events anzeigen' klickt — bis dahin hatte SP genug Zeit zum
      // Propagieren. Falls jemand auf der Erfolgs-Seite stehen bleibt und nichts
      // klickt, wird beim nächsten Page-Mount ohnehin gerefreshed.
    }
    return eventId;
  }

  async function registerForEvent(
    eventId: string,
    customData: Record<string, string>,
    participantFirstName?: string,
    participantLastName?: string,
    participantEmail?: string,
    preferredStarterType?: string // B2Run: 'Durchstarter' | 'Funstarter'
  ): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;

    // Vorname/Nachname aus displayName extrahieren falls nicht uebergeben.
    // Deloitte-Profile liefern den Namen typischerweise als "Nachname, Vorname"
    // (Komma-Format aus dem Active Directory). Frueher haben wir mit Space
    // gesplittet — das tauschte Vor- und Nachname und fuehrte u.a. dazu, dass
    // bei Sub-Event-Anmeldungen (die ohne explizite Vor-/Nachname-Args laufen)
    // die "Anrede" mit dem Nachnamen geschrieben wurde.
    const parseDisplayName = (raw: string): { firstName: string; lastName: string } => {
      const dn = (raw || '').trim();
      if (!dn) return { firstName: '', lastName: '' };
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        return { firstName: parts[1] || '', lastName: parts[0] || '' };
      }
      const parts = dn.split(/\s+/).filter(Boolean);
      if (parts.length === 0) return { firstName: '', lastName: '' };
      if (parts.length === 1) return { firstName: parts[0], lastName: '' };
      return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
    };
    const parsed = parseDisplayName(currentUserName);
    const firstNameToUse = participantFirstName || parsed.firstName;
    const lastNameToUse = participantLastName || parsed.lastName;
    const emailToUse = participantEmail || currentUserEmail;
    const nameToUse = `${firstNameToUse} ${lastNameToUse}`.trim();

    // Pruefen ob schon registriert
    const existing = await eventService.getMyRegistration(subsiteUrl, emailToUse);
    if (existing && existing.Status !== 'Abgemeldet') return false;

    // Pruefen ob Platz frei oder Waitlist
    const event = events.find(e => e.id === eventId);
    let status = 'Angemeldet';
    let effectiveStarterType: string | undefined = preferredStarterType;

    // B2Run Split-Capacity Logik (seit v6.5): getrennte Wartelisten pro StarterType.
    // Wenn der Wunsch-Typ noch freie Plätze hat → direkt angemeldet mit diesem Typ.
    // Wenn der Wunsch-Typ voll ist → landet auf der Warteliste MIT gesetztem
    // PreferredStarterType (kein stiller Fallback auf den anderen Typ mehr).
    // Die Entscheidung "möchte ich auf den anderen Typ umsteigen" trifft der User
    // explizit im UI (RegistrationPage Pre-Check-Dialog), bevor er hier reinkommt —
    // dann hat preferredStarterType bereits den neuen Wunsch.
    const isSplitGroup = !!event && typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    // v11.36: Atomare Sitzplatz-Reservierung statt des alten read-then-write
    // (TOCTOU-Race + fail-open-catch → Massen-Überbuchung bei Anmeldewellen).
    // reserveSeat inkrementiert den Gruppen-Counter per ETag-CAS:
    //   'reserved' → Platz sicher belegt → Angemeldet
    //   'full'     → Gruppe voll → Warteliste
    //   'error'    → Counter nicht nutzbar → FAIL-CLOSED → Warteliste
    //                (NIE optimistisch Angemeldet — genau das war der Bug).
    if (event && isSplitGroup && preferredStarterType) {
      const cap = preferredStarterType === 'Durchstarter'
        ? (event.durchstarterCapacity || 0)
        : (event.funstarterCapacity || 0);
      const seat = await eventService.reserveSeat(subsiteUrl, preferredStarterType as 'Durchstarter' | 'Funstarter', cap);
      if (seat === 'reserved') {
        effectiveStarterType = preferredStarterType;
      } else {
        // 'full' ODER 'error' → fail-closed auf Warteliste für genau diesen Typ.
        status = 'Warteliste';
        effectiveStarterType = undefined; // wird erst beim Nachrücken gesetzt
      }
    } else if (event && isSplitGroup && !preferredStarterType) {
      // Split-Event ohne Gruppenwahl: sicherste Variante ist Warteliste
      // (die UI erzwingt normalerweise eine Gruppenwahl; das ist der Schutz
      // gegen den Pfad, der früher ungebremst auf Angemeldet fiel).
      status = 'Warteliste';
      effectiveStarterType = undefined;
    } else if (event && event.maxParticipants > 0) {
      const seat = await eventService.reserveSeat(subsiteUrl, '', event.maxParticipants);
      if (seat !== 'reserved') {
        status = 'Warteliste';
      }
    }

    // FieldMap aus Custom Fields extrahieren (cf.id -> spInternalName)
    const fieldMap: Record<string, string> = {};
    if (event) {
      for (const f of event.eventSpecificFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const spName = (f as any).spInternalName;
        if (spName) fieldMap[f.id] = spName;
      }
    }

    // Audit-Trail: wer klickt gerade "Register"? = der eingeloggte User.
    // Bei Self-Registration ist das = der Teilnehmer selbst, bei "Fuer andere
    // Person registrieren" ist das der Organizer/Admin (Teilnehmer-Daten
    // wurden ueber participantFirstName/participantEmail uebergeben).
    const actorName = currentUserName;
    const actorEmail = currentUserEmail;

    let success: boolean;
    if (existing && existing.Status === 'Abgemeldet') {
      success = await eventService.reactivateRegistration(subsiteUrl, existing.Id, firstNameToUse, lastNameToUse, customData, status, fieldMap, actorName, actorEmail);
    } else {
      success = await eventService.registerForEvent(
        subsiteUrl, firstNameToUse, lastNameToUse, emailToUse, customData, status, fieldMap,
        effectiveStarterType, preferredStarterType, actorName, actorEmail
      );
    }

    if (success && event) {
      // v9.0: Audit-Log (fire-and-forget)
      eventService.writeChangeLog({
        action: existing ? 'ParticipantReactivated' : 'ParticipantRegistered',
        targetType: 'Participant',
        targetId: emailToUse,
        targetName: nameToUse,
        eventId: eventId,
        eventTitle: event.title,
        details: { status, asActor: emailToUse !== currentUserEmail ? 'on-behalf-of' : 'self' },
      }).catch(() => { /* */ });
      // Warteliste-Position ermitteln
      let waitlistPosition = 0;
      if (status === 'Warteliste') {
        try {
          const counts = await eventService.getRegistrationCount(subsiteUrl);
          waitlistPosition = counts.waitlist;
        } catch { /* Position nicht ermittelbar */ }
      }

      // Dual-Write: DEX_Participants aktualisieren
      if (event.eventNumber) {
        eventService.upsertParticipant(
          firstNameToUse, lastNameToUse, emailToUse, event.eventNumber, status
        ).catch(err => console.warn('[DEX] upsertParticipant failed:', err));
      }
      // E-Mail in Queue eintragen (SharePoint-Template, Fallback auf Code-Template)
      const templateType = status === 'Warteliste' ? 'Warteliste' : 'Anmeldung';
      const lang = event.emailLanguage || 'EN';
      const posText = waitlistPosition > 0 ? String(waitlistPosition) : '';
      // {{Name}} in E-Mail-Anreden: nur Vorname (firstNameToUse ist bei Self-Reg
      // aus dem displayName gesplittet, bei "Fuer andere registrieren" explizit gesetzt).
      const vars = { Name: firstNameToUse, EventTitle: event.title, Organizer: formatOrganizerList(event.organizers, lang), AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, WaitlistPosition: posText };
      let emailData: { subject: string; body: string };
      const spTemplateRaw = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
      const spTemplate = applyEventTemplateOverride(spTemplateRaw, event.emailTemplateOverrides, templateType);
      if (spTemplate) {
        emailData = buildEmailFromTemplate(spTemplate, vars);
      } else {
        emailData = status === 'Warteliste'
          ? waitlistEmail(firstNameToUse, event.title, waitlistPosition)
          : registrationEmail(firstNameToUse, event.title);
      }
      if (!event.disableEmails) {
        // v8.5: Organizer-BCC-Modus auswerten. Bei 'always' immer BCC,
        // bei 'fromDate' nur wenn das konfigurierte Datum bereits erreicht
        // ist, bei 'never'/undefined keinen BCC.
        let bcc: string | undefined;
        const mode = event.notifyOrgRegisterMode || 'never';
        if (mode === 'always' || (mode === 'fromDate' && event.notifyOrgRegisterFromDate && new Date() >= new Date(event.notifyOrgRegisterFromDate))) {
          const orgEmails = (event.organizerEmails || []).filter(Boolean);
          if (orgEmails.length > 0) bcc = orgEmails.join(';');
        }
        // v9.22: Externe Mail-Adresse erkennen — die Plattform ist nur fuer
        // Deloitte Deutschland (@deloitte.de) freigeschaltet. Auch @deloitte.com
        // (US/Global) zaehlt als extern. Bei externen Empfaengern wird die
        // Bestaetigungsmail an den Organizer umgeleitet, der sie ggf. unter
        // Beachtung der Datenschutzrichtlinien weiterleiten kann.
        const isExternalRecipient = !!emailToUse && !/@(.*\.)?deloitte\.de$/i.test(emailToUse);
        let finalRecipient = emailToUse;
        let finalSubject = emailData.subject;
        let finalBody = emailData.body;
        let finalRecipientName = nameToUse;
        if (isExternalRecipient) {
          const orgEmails = (event.organizerEmails || []).filter(Boolean);
          finalRecipient = orgEmails.length > 0 ? orgEmails.join(';') : currentUserEmail;
          finalRecipientName = 'Organizer';
          finalSubject = `[Externer Teilnehmer] ${emailData.subject} — bitte ggf. weiterleiten`;
          // Datenschutz-Hinweis-Box VOR dem Original-Body einbauen.
          const externalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
            + `<strong>Hinweis: externer Teilnehmer.</strong><br>`
            + `Diese Bestätigungsmail richtet sich eigentlich an <strong>${emailToUse}</strong>, einer externen Person die kein Deloitte-Postfach hat. Standardmäßig versendet die App keine Mails an externe Adressen, deshalb landet diese Mail bei dir als Organizer in der Inbox. `
            + `Du kannst die Mail bei Bedarf an den Empfänger weiterleiten — bitte beachte dabei die <a href="https://intranet.deloitte.com/datenschutz" style="color:#86bc25">Datenschutzrichtlinien Deloitte Deutschland</a> (insb. zur Verarbeitung personenbezogener Daten Dritter).`
            + `</div>`;
          // Body kommt schon als komplett-gewickeltes HTML (Deloitte-Template).
          // Wir injecten den Hinweis direkt nach dem opening-<body>-Tag.
          finalBody = finalBody.replace(/<body([^>]*)>/i, `<body$1>${externalHint}`);
        }
        eventService.queueEmail(
          finalSubject, finalRecipient, finalRecipientName, finalBody,
          templateType, event.title, eventId, undefined, bcc
        ).catch(err => console.warn('[DEX] queueEmail failed:', err));

        // v9.15: Auto-Send QR-Code wenn am Event aktiviert. Nur fuer
        // Status='Angemeldet' (Wartelistler bekommen keinen QR — sie sind
        // noch nicht confirmed). Setting wird im Admin-Center per QR-Versand-
        // Modal pro Event umgeschaltet (autoSendQRCode → SP-Feld AutoSendQRCode).
        if (event.autoSendQRCode && status === 'Angemeldet') {
          (async (): Promise<void> => {
            try {
              const qrData = `DEX|${event.eventNumber}|${emailToUse}`;
              let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
              try {
                const qrDataUrl = await QRCode.toDataURL(qrData, { width: 300, margin: 2 });
                qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
              } catch { /* fallback bleibt Text */ }
              const qrMail = qrCodeEmail(firstNameToUse, event.title, qrImageHtml, lang, nameToUse);
              // v9.22: Auto-Send-QR fuer externe Empfaenger ebenfalls an den
              // Organizer umleiten (mit klarem Subject-Praefix), nicht an den
              // externen Mail-Empfaenger.
              if (isExternalRecipient) {
                const orgEmails = (event.organizerEmails || []).filter(Boolean);
                const orgRecipient = orgEmails.length > 0 ? orgEmails.join(';') : currentUserEmail;
                const orgSubject = `[Externer Teilnehmer] QR-Code für ${nameToUse} — ${event.title}`;
                // Hinweis-Box vor dem QR-Code-Body — analog zur Bestaetigungsmail.
                const qrExternalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
                  + `<strong>QR-Code für externen Teilnehmer.</strong><br>`
                  + `Eigentlich für <strong>${emailToUse}</strong> (${nameToUse}). Da externe Adressen keinen Auto-Versand bekommen, landet der QR-Code bei dir — drucke ihn aus oder leite die Mail intern an den Empfänger weiter (Datenschutzrichtlinien Deloitte Deutschland beachten).`
                  + `</div>`;
                const qrBody = qrMail.body.replace(/<body([^>]*)>/i, `<body$1>${qrExternalHint}`);
                await eventService.queueEmail(
                  orgSubject, orgRecipient, 'Organizer', qrBody,
                  'QRCode', event.title, eventId
                );
              } else {
                await eventService.queueEmail(
                  qrMail.subject, emailToUse, nameToUse, qrMail.body,
                  'QRCode', event.title, eventId
                );
              }
              // Status auf 'QR versendet' setzen, damit der Admin-Center-View
              // sofort zeigt, dass der QR-Code raus ist (analog zum
              // manuellen Massen-QR-Versand).
              if (event.subsiteUrl) {
                const myReg = await eventService.getMyRegistration(event.subsiteUrl, emailToUse);
                if (myReg && myReg.Id) {
                  await eventService.setQRSentStatus(event.subsiteUrl, myReg.Id);
                }
              }
            } catch (err) { console.warn('[DEX] auto-send QR failed:', err); }
          })().catch(err => console.warn('[DEX] auto-send QR outer failed:', err));
        }
      }
      // Roommate-Benachrichtigung: nur Custom-Fields vom Typ 'roommate'
      // durchsuchen (seit v7.17 eigener Feldtyp; vorher waren es alle 'user'-
      // Felder, was bei "Assistent"-, "Mentor"- etc. Pickern zu ungewollten
      // Roommate-Mails fuehrte). Fuer jede ausgewaehlte E-Mail eine Roommate-
      // Anfrage-Mail im Deloitte-Template queuen.
      if (!event.disableEmails) {
        for (const f of event.eventSpecificFields) {
          if (f.type !== 'roommate') continue;
          const v = customData[f.id];
          if (!v) continue;
          const m = v.match(/<([^>]+@[^>]+)>/);
          const partnerEmail = m ? m[1].trim() : '';
          if (!partnerEmail || partnerEmail.toLowerCase() === emailToUse.toLowerCase()) continue;
          const partnerName = v.replace(/<[^>]*>/, '').trim() || partnerEmail;
          const partnerFirstName = partnerName.includes(',')
            ? (partnerName.split(',')[1] || '').trim()
            : (partnerName.split(/\s+/)[0] || '');
          const registrantFullName = `${firstNameToUse} ${lastNameToUse}`.trim() || nameToUse;
          const isDe = (lang || 'EN').toUpperCase() === 'DE';
          const subject = isDe
            ? `Zimmerpartner-Anfrage: ${event.title}`
            : `Roommate request: ${event.title}`;
          const inner = isDe
            ? `<p>Hallo ${partnerFirstName || partnerName},</p><p><strong>${registrantFullName}</strong> hat dich als <strong>Zimmerpartner</strong> für das Event <strong>${event.title}</strong> angegeben.</p><p>Wenn du das Match bestätigen möchtest, gib bei deiner Registrierung <strong>${registrantFullName}</strong> ebenfalls als Zimmerpartner an. Das Orga-Team sieht dann im Admin Center, dass ihr euch gegenseitig ausgewählt habt.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
            : `<p>Hello ${partnerFirstName || partnerName},</p><p><strong>${registrantFullName}</strong> has selected you as their <strong>roommate</strong> for the event <strong>${event.title}</strong>.</p><p>If you'd like to confirm the match, please pick <strong>${registrantFullName}</strong> as your roommate when registering. The organizers will then see a mutual match in the admin center.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`;
          const body = wrapTemplate('#86bc25', isDe ? 'Zimmerpartner-Anfrage' : 'Roommate request', event.title, inner);
          eventService.queueEmail(
            subject, partnerEmail, partnerName, body, 'Info', event.title, eventId
          ).catch(err => console.warn('[DEX] roommate queueEmail failed:', err));
        }
      }
      // Outlook-Termin-Einladung in Queue eintragen
      if (status !== 'Warteliste' && !event.disableOutlook) {
        eventService.queueOutlookEvent(
          emailToUse, eventId, event.title, 'Einladen'
        ).catch(err => console.warn('[DEX] queueOutlookEvent failed:', err));
      }
      // v11.53: KPI-Counter sofort hochzaehlen, damit der naechste Boot-
      // Loader die neue Zahl ohne Verzoegerung zeigt. Nur fuer 'Angemeldet'-
      // Status (Warteliste zaehlt nicht in 'Teilnehmer').
      if (status === 'Angemeldet') {
        eventService.bumpKpiParticipants(1).catch(() => { /* best-effort */ });
      }
      await loadEvents();
    }
    return success;
  }

  /**
   * v11.82: Team-Anmeldung — eine Person meldet sich + N-1 Mitglieder
   * gleichzeitig an. N Plaetze werden atomar reserviert (per `reserveSeat`
   * mit count=N). Sind nicht genug Plaetze frei, geht das gesamte Team
   * auf die Warteliste — kein Teil-Anmelden eines vollen Events.
   *
   * Jedes Mitglied bekommt einen eigenen Eintrag in der Subsite-Teilnehmer-
   * liste mit identischer `TeamId`. Genau ein Eintrag (der Lead, also der
   * Submitter) ist `TeamLead=true`. Jeder Mitglied bekommt eine eigene
   * Bestaetigungs-Mail (mit Hinweis dass er als Teil eines Teams angemeldet
   * wurde) und einen eigenen Outlook-Termin (sofern aktiviert).
   *
   * Die Member-Eintraege bekommen leere Custom-Field-Antworten — nur der
   * Lead beantwortet event-spezifische Fragen. Pflicht-Custom-Fields sollten
   * organisatorisch nicht mit Team-Anmeldung kombiniert werden; die App
   * setzt das nicht hart durch, der Wizard sollte den Organizer im Manual
   * darauf hinweisen.
   */
  async function registerTeam(
    eventId: string,
    leadData: {
      firstName: string;
      lastName: string;
      email: string;
      salutation?: string;
      customData: Record<string, string>;
      preferredStarterType?: string;
    },
    members: Array<{ email: string; displayName: string }>,
    teamName: string | undefined
  ): Promise<{ ok: boolean; teamId?: string; status?: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };

    // Doppel-Anmelde-Pruefung: weder der Lead noch ein Member darf bereits
    // (aktiv) angemeldet sein. v11.83: konsolidiert auf den zentralen Helper
    // `isUserAlreadyOnEvent`, der genau die blockierenden Status-Werte
    // beruecksichtigt (Angemeldet/QR versendet/Eingecheckt/Warteliste). Pfad
    // ist nicht performance-kritisch — sequentiell ist OK bei N ≤ 20.
    const allEmails = [leadData.email, ...members.map(m => m.email)];
    for (const em of allEmails) {
      const blocked = await eventService.isUserAlreadyOnEvent(subsiteUrl, em);
      if (blocked) {
        return { ok: false, reason: `already-registered:${em}` };
      }
    }

    // TeamId generieren — bevorzugt crypto.randomUUID, sonst Fallback.
    let teamId: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const c: any = (typeof crypto !== 'undefined') ? crypto : null;
    if (c && typeof c.randomUUID === 'function') {
      teamId = c.randomUUID();
    } else {
      teamId = `team-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    }

    const teamCount = 1 + members.length;
    let status: 'Angemeldet' | 'Warteliste' = 'Angemeldet';
    let effectiveStarterType: string | undefined = leadData.preferredStarterType;

    // Atomar N Plaetze reservieren — Split-Group oder klassisch.
    const isSplitGroup = typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    if (isSplitGroup && leadData.preferredStarterType) {
      const cap = leadData.preferredStarterType === 'Durchstarter'
        ? (event.durchstarterCapacity || 0)
        : (event.funstarterCapacity || 0);
      const seat = await eventService.reserveSeat(subsiteUrl, leadData.preferredStarterType as 'Durchstarter' | 'Funstarter', cap, teamCount);
      if (seat !== 'reserved') {
        status = 'Warteliste';
        effectiveStarterType = undefined;
      }
    } else if (event.maxParticipants > 0) {
      const seat = await eventService.reserveSeat(subsiteUrl, '', event.maxParticipants, teamCount);
      if (seat !== 'reserved') status = 'Warteliste';
    }

    // FieldMap aus Custom Fields extrahieren (cf.id -> spInternalName)
    const fieldMap: Record<string, string> = {};
    for (const f of event.eventSpecificFields) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const spName = (f as any).spInternalName;
      if (spName) fieldMap[f.id] = spName;
    }

    const actorName = currentUserName;
    const actorEmail = currentUserEmail;

    // Lead-Profil + Member-Profile parallel laden.
    const leadProfilePromise = leadData.email.toLowerCase() === currentUserEmail.toLowerCase()
      ? eventService.getCurrentUserProfile()
      : eventService.getUserProfileByEmail(leadData.email);
    const memberProfilePromises = members.map(m => eventService.getUserProfileByEmail(m.email));
    const [leadProfile, ...memberProfiles] = await Promise.all([leadProfilePromise, ...memberProfilePromises]);

    // Parse "Lastname, Firstname" → { firstName, lastName }. Deloitte-AD-Format.
    const parseDisplayName = (raw: string): { firstName: string; lastName: string } => {
      const dn = (raw || '').trim();
      if (!dn) return { firstName: '', lastName: '' };
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        return { firstName: parts[1] || '', lastName: parts[0] || '' };
      }
      const parts = dn.split(/\s+/).filter(Boolean);
      if (parts.length === 0) return { firstName: '', lastName: '' };
      if (parts.length === 1) return { firstName: parts[0], lastName: '' };
      return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
    };

    // Alle Team-Insert-Calls parallel — Counter-CAS in getNextTeilnehmerId
    // garantiert eindeutige TeilnehmerIDs auch bei N parallelen Inserts.
    const insertPromises: Array<Promise<{ ok: boolean; email: string; firstName: string; lastName: string }>> = [];
    // Lead
    insertPromises.push((async (): Promise<{ ok: boolean; email: string; firstName: string; lastName: string }> => {
      const r = await eventService.registerTeamMember(subsiteUrl, {
        firstName: leadData.firstName,
        lastName: leadData.lastName,
        email: leadData.email,
        profile: leadProfile,
        status,
        teamId,
        teamLead: true,
        teamName,
        customData: leadData.customData,
        customFieldMap: fieldMap,
        starterType: effectiveStarterType,
        preferredStarterType: leadData.preferredStarterType,
        registeredByName: actorName,
        registeredByEmail: actorEmail,
        salutation: leadData.salutation,
      });
      return { ok: r.ok, email: leadData.email, firstName: leadData.firstName, lastName: leadData.lastName };
    })());
    // Members
    members.forEach((m, idx) => {
      const profile = memberProfiles[idx] || { department: '', location: '', jobTitle: '', phone: '' };
      const parsed = parseDisplayName(m.displayName);
      insertPromises.push((async (): Promise<{ ok: boolean; email: string; firstName: string; lastName: string }> => {
        const r = await eventService.registerTeamMember(subsiteUrl, {
          firstName: parsed.firstName,
          lastName: parsed.lastName,
          email: m.email,
          profile,
          status,
          teamId,
          teamLead: false,
          teamName,
          customData: {},
          customFieldMap: fieldMap,
          starterType: effectiveStarterType,
          preferredStarterType: leadData.preferredStarterType,
          registeredByName: actorName,
          registeredByEmail: actorEmail,
          // Anrede der Mitglieder bleibt leer — kein Picker fuer Member-Anreden.
          salutation: '',
        });
        return { ok: r.ok, email: m.email, firstName: parsed.firstName, lastName: parsed.lastName };
      })());
    });

    const results = await Promise.all(insertPromises);
    const anyOk = results.some(r => r.ok);
    if (!anyOk) return { ok: false, reason: 'insert-failed' };

    // Pro erfolgreiche Anmeldung: Bestaetigungs-Mail + Outlook-Termin queuen.
    const lang = event.emailLanguage || 'EN';
    const isDe = lang.toUpperCase() === 'DE';
    // v11.87: Team-Info-Block — Mitglieder-Liste, Belegung, Cancel-Hinweis.
    // Baue die Mitglieder-Liste aus den erfolgreichen Inserts auf — Reihenfolge
    // entspricht dem Insert-Pfad (Lead zuerst, dann Members in der Eingabe-
    // Reihenfolge). TeamSize aus dem Event-Config, Fallback auf die Anzahl
    // der tatsächlich angemeldeten Personen.
    const successResults = results.filter(r => r.ok);
    const teamMembersForBlock = successResults.map((r, i) => ({
      firstName: r.firstName,
      lastName: r.lastName,
      isLead: i === 0,
    }));
    const teamSizeForBlock = (typeof event.teamSize === 'number' && event.teamSize > 0)
      ? event.teamSize
      : successResults.length;

    for (const r of results) {
      if (!r.ok) continue;
      const templateType: string = status === 'Warteliste' ? 'Warteliste' : 'Anmeldung';
      const vars = {
        Name: r.firstName,
        EventTitle: event.title,
        Organizer: formatOrganizerList(event.organizers, lang),
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
        WaitlistPosition: '',
      };
      let emailData: { subject: string; body: string };
      const spTemplateRaw = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
      const spTemplate = applyEventTemplateOverride(spTemplateRaw, event.emailTemplateOverrides, templateType);
      if (spTemplate) {
        emailData = buildEmailFromTemplate(spTemplate, vars);
      } else {
        emailData = status === 'Warteliste'
          ? waitlistEmail(r.firstName, event.title, 0)
          : registrationEmail(r.firstName, event.title);
      }
      // v11.87: Team-Info-Block + Consent-Hinweis nach <body> injecten.
      const teamInfoHtml = teamInfoBlockHtml({
        teamName,
        members: teamMembersForBlock,
        teamSize: teamSizeForBlock,
        isDe,
        registeredByName: currentUserName,
        consentRequired: true,
      });
      const bodyWithHint = emailData.body.replace(/<body([^>]*)>/i, `<body$1>${teamInfoHtml}`);
      if (!event.disableEmails) {
        const fullName = `${r.firstName} ${r.lastName}`.trim();
        eventService.queueEmail(
          emailData.subject, r.email, fullName, bodyWithHint,
          templateType, event.title, eventId
        ).catch(err => console.warn('[DEX] team queueEmail failed:', err));
      }
      if (status !== 'Warteliste' && !event.disableOutlook) {
        eventService.queueOutlookEvent(
          r.email, eventId, event.title, 'Einladen'
        ).catch(err => console.warn('[DEX] team queueOutlookEvent failed:', err));
      }
      if (event.eventNumber) {
        eventService.upsertParticipant(
          r.firstName, r.lastName, r.email, event.eventNumber, status
        ).catch(err => console.warn('[DEX] team upsertParticipant failed:', err));
      }
    }

    // Audit-Log (fire-and-forget).
    eventService.writeChangeLog({
      action: 'TeamRegistered',
      targetType: 'Participant',
      targetId: leadData.email,
      targetName: `${leadData.firstName} ${leadData.lastName}`.trim(),
      eventId,
      eventTitle: event.title,
      details: { teamId, teamSize: teamCount, status, members: members.map(m => m.email) },
    }).catch(() => { /* */ });

    if (status === 'Angemeldet') {
      eventService.bumpKpiParticipants(teamCount).catch(() => { /* best-effort */ });
    }
    await loadEvents();
    return { ok: true, teamId, status };
  }

  /**
   * v11.83: Einzelnes Mitglied zu einem bestehenden Team hinzufuegen.
   * Wird vom „+ Mitglied"-Button im MyEvents-Team-Badge benutzt — nur fuer
   * Leads sichtbar, daher hier keine separate Lead-Authorisierung; die
   * UI versteckt den Button.
   *
   * Schritte:
   *   1) Doppel-Anmelde-Check via `isUserAlreadyOnEvent`. Wenn die Person
   *      schon angemeldet ist, brechen wir mit klarem Reason ab.
   *   2) Atomar 1 Sitzplatz reservieren — split-aware. Bei Vollbelegung
   *      landet das neue Mitglied auf der Warteliste (kein Hard-Fail).
   *   3) `registerTeamMember` mit identischer TeamId, `teamLead=false`.
   *   4) Bestaetigungs-Mail + Outlook-Termin queuen.
   *   5) Optional: Info-Mail an die anderen Mitglieder „X ist eurem Team
   *      beigetreten" (best-effort).
   */
  async function addTeamMember(
    eventId: string,
    teamId: string,
    teamName: string | undefined,
    member: { email: string; displayName: string }
  ): Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };
    if (!teamId) return { ok: false, reason: 'invalid-team-id' };

    const blocked = await eventService.isUserAlreadyOnEvent(subsiteUrl, member.email);
    if (blocked) return { ok: false, reason: `already-registered:${member.email}` };

    // Vorhandene Mitglieder laden — um die richtige Gruppe (Split) und
    // die existierenden Custom-Field-Antworten als Vorlage zu erben.
    const existingMembers = await eventService.getTeamMembers(subsiteUrl, teamId);
    const activeMembers = existingMembers.filter(m => m.Status !== 'Abgemeldet');
    const teamSizeCfg = event.teamSize || (activeMembers.length + 1);
    if (activeMembers.length >= teamSizeCfg) {
      return { ok: false, reason: 'team-full' };
    }
    const inheritedStarterType = activeMembers.find(m => !!m.PreferredStarterType)?.PreferredStarterType || '';

    let status: 'Angemeldet' | 'Warteliste' = 'Angemeldet';
    let effectiveStarterType: string | undefined = inheritedStarterType || undefined;
    const isSplitGroup = typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    if (isSplitGroup && inheritedStarterType) {
      const cap = inheritedStarterType === 'Durchstarter'
        ? (event.durchstarterCapacity || 0)
        : (event.funstarterCapacity || 0);
      const seat = await eventService.reserveSeat(subsiteUrl, inheritedStarterType as 'Durchstarter' | 'Funstarter', cap, 1);
      if (seat !== 'reserved') {
        status = 'Warteliste';
        effectiveStarterType = undefined;
      }
    } else if (event.maxParticipants > 0) {
      const seat = await eventService.reserveSeat(subsiteUrl, '', event.maxParticipants, 1);
      if (seat !== 'reserved') status = 'Warteliste';
    }

    // Profil laden + DisplayName parsen.
    const profile = await eventService.getUserProfileByEmail(member.email);
    const parseDisplayName = (raw: string): { firstName: string; lastName: string } => {
      const dn = (raw || '').trim();
      if (!dn) return { firstName: '', lastName: '' };
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        return { firstName: parts[1] || '', lastName: parts[0] || '' };
      }
      const parts = dn.split(/\s+/).filter(Boolean);
      if (parts.length === 0) return { firstName: '', lastName: '' };
      if (parts.length === 1) return { firstName: parts[0], lastName: '' };
      return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
    };
    const parsed = parseDisplayName(member.displayName);

    const r = await eventService.registerTeamMember(subsiteUrl, {
      firstName: parsed.firstName,
      lastName: parsed.lastName,
      email: member.email,
      profile,
      status,
      teamId,
      teamLead: false,
      teamName,
      customData: {},
      starterType: effectiveStarterType,
      preferredStarterType: inheritedStarterType || undefined,
      registeredByName: currentUserName,
      registeredByEmail: currentUserEmail,
      salutation: '',
    });
    if (!r.ok) return { ok: false, reason: 'insert-failed' };

    // Bestaetigungs-Mail + Outlook + DEX_Participants — same pattern as
    // registerTeam aber fuer EINEN Member.
    const lang = event.emailLanguage || 'EN';
    const isDe = (lang || 'EN').toUpperCase() === 'DE';
    const templateType = status === 'Warteliste' ? 'Warteliste' : 'Anmeldung';
    const vars = {
      Name: parsed.firstName,
      EventTitle: event.title,
      Organizer: formatOrganizerList(event.organizers, lang),
      AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      WaitlistPosition: '',
    };
    let emailData: { subject: string; body: string };
    const spTplRaw = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
    const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, templateType);
    if (spTpl) {
      emailData = buildEmailFromTemplate(spTpl, vars);
    } else {
      emailData = status === 'Warteliste'
        ? waitlistEmail(parsed.firstName, event.title, 0)
        : registrationEmail(parsed.firstName, event.title);
    }
    // v11.87: Team-Info-Block mit allen aktiven Mitgliedern inkl. dem neuen.
    // activeMembers wurde vor dem Insert geladen — wir hängen den frisch
    // angemeldeten User vorne dran als Nicht-Lead an. Den Lead identifizieren
    // wir per TeamLead-Flag.
    const allActiveForBlock: Array<{ firstName: string; lastName: string; isLead: boolean }> = [
      ...activeMembers.map(m => ({
        firstName: m.Vorname || '',
        lastName: m.Nachname || '',
        isLead: !!m.TeamLead,
      })),
      { firstName: parsed.firstName, lastName: parsed.lastName, isLead: false },
    ];
    // Lead zuerst sortieren, danach in Insert-Reihenfolge belassen.
    allActiveForBlock.sort((a, b) => (a.isLead === b.isLead) ? 0 : (a.isLead ? -1 : 1));
    const teamSizeForBlock = (typeof event.teamSize === 'number' && event.teamSize > 0)
      ? event.teamSize
      : allActiveForBlock.length;
    const teamInfoHtml = teamInfoBlockHtml({
      teamName,
      members: allActiveForBlock,
      teamSize: teamSizeForBlock,
      isDe,
      registeredByName: currentUserName,
      consentRequired: true,
    });
    const bodyWithHint = emailData.body.replace(/<body([^>]*)>/i, `<body$1>${teamInfoHtml}`);
    if (!event.disableEmails) {
      const fullName = `${parsed.firstName} ${parsed.lastName}`.trim() || member.email;
      eventService.queueEmail(
        emailData.subject, member.email, fullName, bodyWithHint,
        templateType, event.title, eventId
      ).catch(err => console.warn('[DEX] addTeamMember queueEmail failed:', err));
    }
    if (status !== 'Warteliste' && !event.disableOutlook) {
      eventService.queueOutlookEvent(
        member.email, eventId, event.title, 'Einladen'
      ).catch(err => console.warn('[DEX] addTeamMember queueOutlookEvent failed:', err));
    }
    if (event.eventNumber) {
      eventService.upsertParticipant(
        parsed.firstName, parsed.lastName, member.email, event.eventNumber, status
      ).catch(err => console.warn('[DEX] addTeamMember upsertParticipant failed:', err));
    }

    // Info-Mail an die anderen aktiven Mitglieder — knapp, im Layout
    // gewrappt. Best-effort.
    if (!event.disableEmails) {
      for (const other of activeMembers) {
        const otherFirst = other.Vorname || '';
        const subject = isDe
          ? `Neues Team-Mitglied bei Event ${event.title}`
          : `New team member for event ${event.title}`;
        const inner = isDe
          ? `<p>Hallo ${otherFirst},</p>`
            + `<p><strong>${parsed.firstName} ${parsed.lastName}</strong> ist eurem Team${teamName ? ` „${teamName}"` : ''} beigetreten.</p>`
            + `<p>Beste Gruesse,<br>Dein Event-Team</p>`
          : `<p>Hello ${otherFirst},</p>`
            + `<p><strong>${parsed.firstName} ${parsed.lastName}</strong> joined your team${teamName ? ` „${teamName}"` : ''}.</p>`
            + `<p>Best regards,<br>Your event team</p>`;
        const html = wrapTemplate('#86bc25', isDe ? 'Team-Update' : 'Team update', `Event ${event.title}`, inner);
        eventService.queueEmail(
          subject, other.ParticipantEmail,
          `${other.Vorname || ''} ${other.Nachname || ''}`.trim() || other.ParticipantEmail,
          html, 'Info', event.title, eventId
        ).catch(() => { /* best-effort */ });
      }
    }

    if (status === 'Angemeldet') {
      eventService.bumpKpiParticipants(1).catch(() => { /* best-effort */ });
    }
    eventService.writeChangeLog({
      action: 'TeamMemberAdded',
      targetType: 'Participant',
      targetId: member.email,
      targetName: `${parsed.firstName} ${parsed.lastName}`.trim(),
      eventId,
      eventTitle: event.title,
      details: { teamId, addedBy: currentUserEmail, status },
    }).catch(() => { /* */ });

    await loadEvents();
    return { ok: true, status };
  }

  /**
   * v11.83: Direkter Team-Beitritt aus der Anmeldeseite (ohne Approval).
   * Funktional identisch zu `addTeamMember`, aber laeuft mit dem
   * eingeloggten User als Member. Der Submit-Pfad in RegistrationPage
   * unterscheidet zwischen `joinTeam` (Approval OFF) und
   * `createTeamJoinRequest` (Approval ON).
   */
  async function joinTeam(
    eventId: string,
    teamId: string,
    teamName: string | undefined
  ): Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    return addTeamMember(eventId, teamId, teamName, {
      email: currentUserEmail,
      displayName: currentUserName,
    });
  }

  /**
   * v11.84: Lead-Rolle innerhalb eines Teams uebergeben. Nur fuer Admin
   * Center gedacht — die UI versteckt den Button fuer alle anderen Rollen.
   * Wirft kein Mail zur "alten" Person, sondern eine Info-Mail an alle
   * aktiven Team-Mitglieder mit dem Hinweis "Die Team-Lead-Rolle wurde
   * an <Name> uebergeben". Audit-Eintrag im ChangeLog.
   */
  async function transferTeamLead(
    eventId: string,
    teamId: string,
    newLeadEmail: string
  ): Promise<{ ok: boolean; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };
    if (!teamId || !newLeadEmail) return { ok: false, reason: 'invalid-input' };

    const members = await eventService.getTeamMembers(subsiteUrl, teamId);
    const active = members.filter(m => m.Status !== 'Abgemeldet');
    const oldLead = active.find(m => !!m.TeamLead);
    const newLead = active.find(m => (m.ParticipantEmail || '').toLowerCase() === newLeadEmail.toLowerCase());
    if (!newLead) return { ok: false, reason: 'new-lead-not-found' };
    if (!oldLead) {
      // Kein aktiver Lead — einfach den neuen promoten, kein Demote noetig.
      const okPromote = await eventService.promoteToTeamLead(subsiteUrl, newLead.Id);
      if (!okPromote) return { ok: false, reason: 'promote-failed' };
    } else {
      if (oldLead.Id === newLead.Id) return { ok: false, reason: 'already-lead' };
      const ok = await eventService.transferTeamLead(subsiteUrl, oldLead.Id, newLead.Id);
      if (!ok) return { ok: false, reason: 'transfer-failed' };
    }

    // Info-Mails an alle aktiven Mitglieder — best-effort.
    if (!event.disableEmails) {
      const lang = (event.emailLanguage || 'EN').toUpperCase();
      const isDe = lang === 'DE';
      const newLeadName = `${newLead.Vorname || ''} ${newLead.Nachname || ''}`.trim() || newLead.ParticipantEmail;
      const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
      for (const other of active) {
        const otherFirst = other.Vorname || '';
        const subject = isDe
          ? `Team-Lead-Rolle uebergeben — Event ${event.title}`
          : `Team lead role transferred — event ${event.title}`;
        const isNewLeadMember = other.Id === newLead.Id;
        const inner = isDe
          ? `<p>Hallo ${otherFirst},</p>`
            + `<p>Die Team-Lead-Rolle in eurem Team${teamName ? ` „${teamName}"` : ''} wurde an <strong>${newLeadName}</strong> uebergeben.</p>`
            + (isNewLeadMember ? `<p>Du bist ab jetzt Team-Lead — du kannst neue Mitglieder ueber „Meine Events" hinzufuegen und ggf. Beitritts-Anfragen entscheiden.</p>` : '')
            + `<p>Beste Gruesse,<br>Dein Event-Team</p>`
          : `<p>Hello ${otherFirst},</p>`
            + `<p>The team lead role in your team${teamName ? ` „${teamName}"` : ''} has been transferred to <strong>${newLeadName}</strong>.</p>`
            + (isNewLeadMember ? `<p>You are now the team lead — you can add new members via „My Events" and decide on join requests if any.</p>` : '')
            + `<p>Best regards,<br>Your event team</p>`;
        const html = wrapTemplate('#86bc25', isDe ? 'Team-Lead-Wechsel' : 'Team lead change', `Event ${event.title}`, inner);
        eventService.queueEmail(
          subject, other.ParticipantEmail,
          `${other.Vorname || ''} ${other.Nachname || ''}`.trim() || other.ParticipantEmail,
          html, 'Info', event.title, eventId
        ).catch(() => { /* best-effort */ });
      }
    }

    eventService.writeChangeLog({
      action: 'TeamLeadTransferred',
      targetType: 'Participant',
      targetId: newLeadEmail,
      targetName: `${newLead.Vorname || ''} ${newLead.Nachname || ''}`.trim(),
      eventId,
      eventTitle: event.title,
      details: { teamId, fromLeadEmail: oldLead?.ParticipantEmail || '', toLeadEmail: newLeadEmail, actor: currentUserEmail },
    }).catch(() => { /* */ });

    return { ok: true };
  }

  /**
   * v11.83: Eine Beitritts-Anfrage in DEX_TeamJoinRequests anlegen +
   * Lead-Notification queuen. Die App liest die TeamId vom UI, holt sich
   * den Lead aus der Subsite-Teilnehmerliste (TeamId-Match,
   * TeamLead=true) und schreibt dann eine Mail an die Lead-Email.
   */
  async function createTeamJoinRequest(
    eventId: string,
    teamId: string
  ): Promise<{ ok: boolean; itemId?: number; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };

    const blocked = await eventService.isUserAlreadyOnEvent(subsiteUrl, currentUserEmail);
    if (blocked) return { ok: false, reason: 'already-registered' };

    // Lead finden — die Mail-Notification soll an ihn gehen.
    const members = await eventService.getTeamMembers(subsiteUrl, teamId);
    const lead = members.find(m => !!m.TeamLead && m.Status !== 'Abgemeldet');
    if (!lead) return { ok: false, reason: 'team-has-no-lead' };

    const teamNameFromMembers = members.find(m => !!m.TeamName)?.TeamName || '';
    const result = await eventService.createTeamJoinRequest({
      eventId,
      eventTitle: event.title,
      teamId,
      requesterEmail: currentUserEmail,
      requesterDisplayName: currentUserName,
    });
    if (!result.ok) return { ok: false, reason: 'queue-failed' };

    // Lead-Notification-Mail.
    if (!event.disableEmails) {
      const lang = (event.emailLanguage || 'EN').toUpperCase();
      const isDe = lang === 'DE';
      const leadFirst = lead.Vorname || '';
      const subject = isDe
        ? `Team-Beitritts-Anfrage fuer Event ${event.title}`
        : `Team join request for event ${event.title}`;
      const appUrl = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView&action=teamjoin&request=${result.itemId || 0}`;
      const inner = isDe
        ? `<p>Hallo ${leadFirst},</p>`
          + `<p><strong>${currentUserName}</strong> moechte deinem Team${teamNameFromMembers ? ` „${teamNameFromMembers}"` : ''} beim Event „${event.title}" beitreten. Bitte entscheide:</p>`
          + `<p style="text-align:center;margin:18px 0;"><a href="${appUrl}&decision=approve" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Bestaetigen</a> <a href="${appUrl}&decision=reject" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Ablehnen</a></p>`
          + `<p style="font-size:0.85rem;color:#666;">Hinweis: die Buttons fuehren dich auf die App; dort findest du den Beitritts-Anfragen-Block in „Meine Events".</p>`
          + `<p>Beste Gruesse,<br>Dein Event-Team</p>`
        : `<p>Hello ${leadFirst},</p>`
          + `<p><strong>${currentUserName}</strong> would like to join your team${teamNameFromMembers ? ` „${teamNameFromMembers}"` : ''} for event „${event.title}". Please decide:</p>`
          + `<p style="text-align:center;margin:18px 0;"><a href="${appUrl}&decision=approve" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Approve</a> <a href="${appUrl}&decision=reject" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Reject</a></p>`
          + `<p style="font-size:0.85rem;color:#666;">Note: the buttons lead you to the app; the request block lives in „My Events".</p>`
          + `<p>Best regards,<br>Your event team</p>`;
      const html = wrapTemplate('#86bc25', isDe ? 'Team-Beitritts-Anfrage' : 'Team join request', `Event ${event.title}`, inner);
      eventService.queueEmail(
        subject, lead.ParticipantEmail,
        `${lead.Vorname || ''} ${lead.Nachname || ''}`.trim() || lead.ParticipantEmail,
        html, 'Info', event.title, eventId
      ).catch(() => { /* best-effort */ });
    }

    return { ok: true, itemId: result.itemId };
  }

  /**
   * v11.83: Pending-Anfragen fuer ein bestimmtes Team eines Events
   * abrufen — wird in der „Beitritts-Anfragen"-Box im MyEvents-Team-
   * Badge angezeigt (nur Leads sehen sie).
   */
  async function listTeamJoinRequestsForEvent(
    eventId: string,
    teamId: string
  ): Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string }>> {
    if (!eventId || !teamId) return [];
    const items = await eventService.listTeamJoinRequests({ eventId, teamId, status: 'Pending' });
    return items;
  }

  /**
   * v11.83: Approve/Reject einer Beitritts-Anfrage durch den Team-Lead.
   * Bei Approve: Member-Anmeldung via `addTeamMember`-Logik + Mail an
   * Anfragenden „du wurdest aufgenommen". Bei Reject: kurze Absage-Mail.
   * Beide Pfade setzen anschliessend den Status der DEX_TeamJoinRequests-
   * Zeile auf Approved/Rejected.
   */
  async function decideTeamJoinRequest(
    requestId: number,
    decision: 'Approved' | 'Rejected'
  ): Promise<boolean> {
    // Erst die Request-Zeile holen, damit wir Event-/Team-Kontext kennen.
    const all = await eventService.listTeamJoinRequests({ status: 'Pending' });
    const req = all.find(r => r.Id === requestId);
    if (!req) return false;
    const event = events.find(e => e.id === req.EventId);
    if (!event) return false;
    const subsiteUrl = subsiteMap.current[req.EventId];
    if (!subsiteUrl) return false;

    if (decision === 'Approved') {
      // Bestehenden Team-Namen ableiten.
      const members = await eventService.getTeamMembers(subsiteUrl, req.TeamId);
      const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
      const addRes = await addTeamMember(req.EventId, req.TeamId, teamName || undefined, {
        email: req.RequesterEmail,
        displayName: req.RequesterDisplayName,
      });
      if (!addRes.ok) {
        // Wir markieren die Anfrage trotzdem als Approved, wenn der Add
        // fehlschlug — der Lead bekommt ein UI-Feedback und kann manuell
        // nachsetzen. Status bleibt Pending nur bei System-Fehlern auf der
        // List-Selbst.
        return false;
      }
      await eventService.decideTeamJoinRequest(requestId, 'Approved', currentUserEmail);
      // „Du wurdest aufgenommen"-Mail wurde bereits durch addTeamMember
      // gequeued (Bestaetigungs-Mail), daher hier keine doppelte Mail.
      return true;
    }

    // Reject
    const ok = await eventService.decideTeamJoinRequest(requestId, 'Rejected', currentUserEmail);
    if (!event.disableEmails) {
      const lang = (event.emailLanguage || 'EN').toUpperCase();
      const isDe = lang === 'DE';
      const subject = isDe
        ? `Team-Beitritts-Anfrage abgelehnt — Event ${event.title}`
        : `Team join request declined — event ${event.title}`;
      const inner = isDe
        ? `<p>Hallo ${req.RequesterDisplayName.split(',').pop()?.trim() || req.RequesterDisplayName},</p>`
          + `<p>deine Beitritts-Anfrage zum Team beim Event „${event.title}" wurde vom Team-Lead abgelehnt.</p>`
          + `<p>Du kannst dich gerne einzeln anmelden, falls die Kapazitaet noch reicht — oder einem anderen offenen Team beitreten.</p>`
          + `<p>Beste Gruesse,<br>Dein Event-Team</p>`
        : `<p>Hello ${req.RequesterDisplayName.split(',').pop()?.trim() || req.RequesterDisplayName},</p>`
          + `<p>your join request for the team at event „${event.title}" was declined by the team lead.</p>`
          + `<p>You can register individually if capacity allows — or join another open team.</p>`
          + `<p>Best regards,<br>Your event team</p>`;
      const html = wrapTemplate('#86bc25', isDe ? 'Team-Beitritts-Anfrage' : 'Team join request', `Event ${event.title}`, inner);
      eventService.queueEmail(
        subject, req.RequesterEmail, req.RequesterDisplayName,
        html, 'Info', event.title, req.EventId
      ).catch(() => { /* best-effort */ });
    }
    return ok;
  }

  /**
   * v11.83: Aktive Teams eines Events fuer die „Offene Teams"-Box.
   * Filter: nur Teams mit aktivem Mitglied-Count < event.teamSize.
   * Mitgliedernamen werden bewusst NICHT zurueckgegeben (Privatsphaere) —
   * nur Belegungs-Anzahl, TeamName und LeadEmail (LeadEmail wird ohnehin
   * gebraucht, weil der Beitritts-Pfad eine Lead-Notification queued).
   */
  async function listOpenTeamsForEvent(eventId: string): Promise<Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }>> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return [];
    const event = events.find(e => e.id === eventId);
    if (!event || !event.teamRegistrationEnabled) return [];
    const teamSizeCfg = event.teamSize || 0;
    if (teamSizeCfg < 2) return [];

    const all = await eventService.getAllRegistrations(subsiteUrl);
    // Gruppieren nach TeamId.
    const byTeam: Record<string, SPRegistration[]> = {};
    for (const r of all) {
      if (!r.TeamId) continue;
      if (r.Status === 'Abgemeldet') continue;
      if (!byTeam[r.TeamId]) byTeam[r.TeamId] = [];
      byTeam[r.TeamId].push(r);
    }
    const open: Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }> = [];
    for (const tid of Object.keys(byTeam)) {
      const list = byTeam[tid];
      if (list.length >= teamSizeCfg) continue;
      const lead = list.find(m => !!m.TeamLead) || list[0];
      open.push({
        teamId: tid,
        teamName: list.find(m => !!m.TeamName)?.TeamName || '',
        activeCount: list.length,
        teamSize: teamSizeCfg,
        leadEmail: lead?.ParticipantEmail || '',
        leadDisplayName: `${lead?.Vorname || ''} ${lead?.Nachname || ''}`.trim() || lead?.ParticipantEmail || '',
      });
    }
    return open;
  }

  async function cancelRegistration(eventId: string): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;

    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;

    // Audit: wer klickt gerade 'Abmelden'? = der eingeloggte User. Bei Self-Cancel
    // ist das = der Teilnehmer selbst. Aus der App heraus gibt's aktuell keinen
    // "Abmeldung fuer andere"-Pfad (das macht der Organizer/Admin ueber AdminPage,
    // dort wird eventService.cancelRegistration direkt aufgerufen).
    // v11.53: vorherigen Status merken, damit wir den KPI-Counter nur dann
    // dekrementieren, wenn der User tatsaechlich 'Angemeldet' war (Wartelist-
    // Cancel beruehrt den Teilnehmer-KPI nicht).
    const wasActive = myReg.Status === 'Angemeldet';
    // v11.83: Team-Anmeldungs-Kontext snapshotten, BEVOR der eigene Status
    // auf 'Abgemeldet' kippt — danach liefert getTeamMembers den eigenen
    // Eintrag schon mit dem alten Lead-Flag aus und der Promote-Pfad
    // verlaesst sich nicht mehr darauf. Wir speichern hier den eigenen
    // TeamId/TeamLead/TeamName-Stand und filtern nach dem Cancel die
    // verbleibenden Mitglieder.
    const wasTeamCancel = !!myReg.TeamId;
    const wasTeamLead = wasTeamCancel && !!myReg.TeamLead;
    const teamId = myReg.TeamId || '';
    const teamName = myReg.TeamName || '';
    const success = await eventService.cancelRegistration(subsiteUrl, myReg.Id, currentUserName, currentUserEmail);
    if (success) {
      const event = events.find(e => e.id === eventId);
      if (wasActive) {
        eventService.bumpKpiParticipants(-1).catch(() => { /* best-effort */ });
      }
      // v9.0: Audit-Log (fire-and-forget)
      eventService.writeChangeLog({
        action: 'ParticipantCancelled',
        targetType: 'Participant',
        targetId: currentUserEmail,
        targetName: currentUserName,
        eventId: eventId,
        eventTitle: event?.title || '',
        details: { participantId: myReg.Id, asActor: 'self' },
      }).catch(() => { /* */ });
      if (event) {
        // Dual-Write: DEX_Participants aktualisieren
        if (event.eventNumber) {
          try {
            await eventService.removeParticipantEvent(currentUserEmail, event.eventNumber);
          } catch (err) { console.warn('[DEX] removeParticipantEvent failed:', err); }
        }
        // Abmelde-E-Mail in Queue eintragen (SharePoint-Template, Fallback auf Code-Template)
        if (!event.disableEmails) {
          try {
            const lang = event.emailLanguage || 'EN';
            // {{Name}} in Anreden: nur Vorname (displayName ist im Deloitte-Tenant
            // "Nachname, Vorname" -> getFirstName extrahiert den Vornamen).
            const cancelVars = { Name: currentUserFirstName, EventTitle: event.title, AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView` };
            let emailData: { subject: string; body: string };
            const spTplRaw = await eventService.getEmailTemplate('Abmeldung', lang).catch(() => null);
            const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, 'Abmeldung');
            if (spTpl) {
              emailData = buildEmailFromTemplate(spTpl, cancelVars);
            } else {
              emailData = cancellationEmail(currentUserFirstName, event.title);
            }
            // v8.5: Organizer-BCC bei Abmeldung auswerten. 'always' = immer,
            // 'afterDeadline' = nur wenn lastDeregisterDate ueberschritten ist.
            let bcc: string | undefined;
            const mode = event.notifyOrgCancelMode || 'never';
            if (mode === 'always' || (mode === 'afterDeadline' && event.lastDeregisterDate && new Date() > new Date(event.lastDeregisterDate))) {
              const orgEmails = (event.organizerEmails || []).filter(Boolean);
              if (orgEmails.length > 0) bcc = orgEmails.join(';');
            }
            const emailOk = await eventService.queueEmail(
              emailData.subject, currentUserEmail, currentUserName, emailData.body,
              'Abmeldung', event.title, eventId, undefined, bcc
            );
            if (!emailOk) console.warn('[DEX] queueEmail for cancellation returned false');
          } catch (err) { console.warn('[DEX] queueEmail for cancellation failed:', err); }
        }
        // Outlook-Termin-Einladung zurückziehen
        if (!event.disableOutlook) {
          try {
            await eventService.queueOutlookEvent(
              currentUserEmail, eventId, event.title, 'Ausladen'
            );
          } catch (err) { console.warn('[DEX] queueOutlookEvent failed:', err); }
        }
        // Nachrücken wird komplett vom Power-Automate-Flow DEX_IDReorder_TeilnehmerIDs
        // übernommen (seit v6.7). Der Flow ist typen-bewusst für B2Run-Split-
        // Wartelisten: er promotet den ersten Warteliste-Teilnehmer mit passendem
        // PreferredStarterType und verschickt Nachrück-Mail + Outlook-Einladung.
        // Die App macht nur noch Abmeldung + IDReorder-Queue-Trigger — keine
        // parallele Client-Promote-Logik mehr (die vorher zu Race-Conditions mit
        // dem Flow geführt hat).
        // ID-Reorder in Queue eintragen (triggert den DEX_IDReorder-Flow, der
        // danach ID-Neuvergabe + Nachrücken abwickelt).
        if (subsiteUrl) {
          try {
            const reorderOk = await eventService.queueIDReorder(
              eventId, event.eventNumber || 0, subsiteUrl, event.title
            );
            if (!reorderOk) console.warn('[DEX] queueIDReorder returned false');
          } catch (err) { console.warn('[DEX] queueIDReorder failed:', err); }
          // v11.36: Sitzplatz-Counter nach der Abmeldung mit dem echten
          // Bestand abgleichen, damit der frei gewordene Platz für die
          // nächste Anmeldung wieder reservierbar ist (best-effort).
          try {
            const isSplit = typeof event.durchstarterCapacity === 'number'
              && typeof event.funstarterCapacity === 'number'
              && ((event.durchstarterCapacity || 0) > 0 || (event.funstarterCapacity || 0) > 0);
            await eventService.syncSeatsToActiveCount(subsiteUrl, { isSplit });
          } catch { /* best-effort */ }
        }
      } else {
        console.warn('[DEX] cancelRegistration: event not found in state for id', eventId);
      }
      // v11.83: Team-Cancel-Nachlauf — Auto-Promote des frueheren Members
      // zum neuen Lead (falls Self-Cancel der Lead war), Info-Mails an die
      // verbleibenden Mitglieder, Hinweis welche Optionen ihnen offenstehen.
      // Der Sitzplatz-Counter wird im normalen Reconcile oben schon
      // dekrementiert — der frei werdende Platz darf von anderen Teilnehmern
      // belegt werden (oder vom Team-Lead nachbesetzt werden, siehe
      // addTeamMember). Die App entscheidet hier bewusst NICHT, ob der
      // Slot fuer das Team reserviert bleibt — das passt zur Beschreibung
      // im Spec, weil der frei werdende Sitz neutral ist: der Team-Lead
      // kann ihn ueber "Mitglied hinzufuegen" wieder fuellen, ansonsten
      // landet er in der normalen Sitzplatz-Verwaltung.
      if (wasTeamCancel && event) {
        await handleTeamCancelPostStep(event, eventId, subsiteUrl, teamId, teamName, wasTeamLead, myReg).catch(err => {
          console.warn('[DEX] team-cancel post-step failed:', err);
        });
      }
      await loadEvents();
    }
    return success;
  }

  /**
   * v11.86: Team-Lead meldet stellvertretend ein Team-Mitglied vom Event
   * ab — ausgeloest aus dem „Team verwalten"-Modal in MyEvents. Audit
   * wird auf den eingeloggten Lead geschrieben (CancelledByName/Email),
   * danach laeuft derselbe Team-Post-Step wie beim Self-Cancel:
   * Sitzplatz-Reconcile, IDReorder-Queue, Outlook-Ausladung,
   * Abmelde-Bestaetigung an die abgemeldete Person und Info-Mails an die
   * uebrigen Team-Mitglieder. Der Lead darf sich ueber diesen Pfad
   * NICHT selbst loeschen — das uebernimmt der normale Self-Cancel ueber
   * `cancelRegistration` (inkl. Auto-Promote des fruehesten Members).
   */
  async function cancelTeamMember(
    eventId: string,
    memberRegistration: SPRegistration
  ): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl || !memberRegistration?.Id) return false;
    if (!memberRegistration.TeamId) return false;
    // Self-Schutz: der Lead loescht sich nicht ueber diesen Pfad — sein
    // eigener Cancel laeuft via cancelRegistration mit Auto-Promote.
    if ((memberRegistration.ParticipantEmail || '').toLowerCase() === (currentUserEmail || '').toLowerCase()) {
      console.warn('[DEX] cancelTeamMember: Lead cannot cancel itself via this path');
      return false;
    }
    const wasActive = memberRegistration.Status === 'Angemeldet';
    const teamId = memberRegistration.TeamId;
    const teamName = memberRegistration.TeamName || '';
    // Audit = der eingeloggte Lead (stellvertretender Cancel).
    const ok = await eventService.cancelRegistration(
      subsiteUrl, memberRegistration.Id, currentUserName, currentUserEmail
    );
    if (!ok) return false;
    const event = events.find(e => e.id === eventId);
    if (wasActive) {
      eventService.bumpKpiParticipants(-1).catch(() => { /* */ });
    }
    eventService.writeChangeLog({
      action: 'ParticipantCancelled',
      targetType: 'Participant',
      targetId: memberRegistration.ParticipantEmail,
      targetName: `${memberRegistration.Vorname || ''} ${memberRegistration.Nachname || ''}`.trim() || memberRegistration.ParticipantEmail,
      eventId: eventId,
      eventTitle: event?.title || '',
      details: { participantId: memberRegistration.Id, asActor: 'teamLead', actorEmail: currentUserEmail },
    }).catch(() => { /* */ });
    if (event) {
      // Dual-Write: DEX_Participants aktualisieren
      if (event.eventNumber) {
        try {
          await eventService.removeParticipantEvent(memberRegistration.ParticipantEmail, event.eventNumber);
        } catch (err) { console.warn('[DEX] removeParticipantEvent failed:', err); }
      }
      // Abmelde-Mail an die abgemeldete Person.
      if (!event.disableEmails) {
        try {
          const lang = event.emailLanguage || 'EN';
          const cancelledFirst = memberRegistration.Vorname
            || (memberRegistration.ParticipantName || '').split(/[ ,]+/)[0]
            || '';
          const cancelVars = {
            Name: cancelledFirst,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          };
          let emailData: { subject: string; body: string };
          const spTplRaw = await eventService.getEmailTemplate('Abmeldung', lang).catch(() => null);
          const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, 'Abmeldung');
          if (spTpl) {
            emailData = buildEmailFromTemplate(spTpl, cancelVars);
          } else {
            emailData = cancellationEmail(cancelledFirst, event.title);
          }
          let bcc: string | undefined;
          const mode = event.notifyOrgCancelMode || 'never';
          if (mode === 'always' || (mode === 'afterDeadline' && event.lastDeregisterDate && new Date() > new Date(event.lastDeregisterDate))) {
            const orgEmails = (event.organizerEmails || []).filter(Boolean);
            if (orgEmails.length > 0) bcc = orgEmails.join(';');
          }
          await eventService.queueEmail(
            emailData.subject,
            memberRegistration.ParticipantEmail,
            `${memberRegistration.Vorname || ''} ${memberRegistration.Nachname || ''}`.trim() || memberRegistration.ParticipantEmail,
            emailData.body,
            'Abmeldung', event.title, eventId, undefined, bcc
          );
        } catch (err) { console.warn('[DEX] queueEmail for team-lead cancel failed:', err); }
      }
      // Outlook-Ausladung.
      if (!event.disableOutlook) {
        try {
          await eventService.queueOutlookEvent(
            memberRegistration.ParticipantEmail, eventId, event.title, 'Ausladen'
          );
        } catch (err) { console.warn('[DEX] queueOutlookEvent (team-lead cancel) failed:', err); }
      }
      // ID-Reorder + Sitzplatz-Sync.
      if (subsiteUrl) {
        try {
          await eventService.queueIDReorder(
            eventId, event.eventNumber || 0, subsiteUrl, event.title
          );
        } catch (err) { console.warn('[DEX] queueIDReorder (team-lead cancel) failed:', err); }
        try {
          const isSplit = typeof event.durchstarterCapacity === 'number'
            && typeof event.funstarterCapacity === 'number'
            && ((event.durchstarterCapacity || 0) > 0 || (event.funstarterCapacity || 0) > 0);
          await eventService.syncSeatsToActiveCount(subsiteUrl, { isSplit });
        } catch { /* best-effort */ }
      }
      // Info-Mails an die uebrigen Team-Mitglieder. Wir loeschen NICHT
      // den Lead, daher `wasTeamLead = false` → kein Auto-Promote.
      await handleTeamCancelPostStep(event, eventId, subsiteUrl, teamId, teamName, false, memberRegistration)
        .catch(err => { console.warn('[DEX] team-cancel post-step (lead-initiated) failed:', err); });
    }
    await loadEvents();
    return true;
  }

  /**
   * v11.83: Nach einem Team-Mitglied-Cancel (Self-Cancel) erledigt diese
   * Routine:
   *   1) Verbleibende aktive Team-Mitglieder laden (ohne den gerade
   *      Abgemeldeten, der jetzt 'Abgemeldet' ist).
   *   2) Falls die abgemeldete Person Lead war UND mindestens ein Member
   *      uebrig ist, das frueheste aktive Mitglied per MERGE-Patch zum
   *      neuen Lead promoten.
   *   3) Pro verbleibendem Mitglied eine Info-Mail in DEX_Emails queuen,
   *      die den Cancel ankuendigt und die naechsten Schritte erklaert.
   *
   * Fail-safe: alle Sub-Operationen sind best-effort und schlucken Fehler
   * still — das Cancel selbst hat oben schon erfolgreich auf dem Item
   * geschrieben, ein Mail-/Promote-Fehler darf den User-Flow nicht
   * blockieren.
   */
  async function handleTeamCancelPostStep(
    event: DeloitteEvent,
    eventId: string,
    subsiteUrl: string,
    teamId: string,
    teamName: string,
    wasTeamLead: boolean,
    cancelledReg: SPRegistration
  ): Promise<void> {
    const members = await eventService.getTeamMembers(subsiteUrl, teamId);
    // Verbleibende = aktive (NICHT 'Abgemeldet') und NICHT der gerade
    // abgemeldete Eintrag (Id-Vergleich, weil ein parallel-Member denselben
    // Vor-/Nachnamen haben koennte).
    const remaining = members.filter(m => m.Status !== 'Abgemeldet' && m.Id !== cancelledReg.Id);
    if (remaining.length === 0) {
      // Team aufgeloest — kein Promote, keine Info-Mails noetig.
      return;
    }

    // Auto-Promote: wenn der Cancel ein Lead war, das frueheste aktive
    // Member zum neuen Lead machen. Sortier-Kriterium: kleinste
    // TeilnehmerID, sonst frueheste RegistrationDate, sonst kleinste Id.
    let newLeadId: number | null = null;
    if (wasTeamLead) {
      const sorted = [...remaining].sort((a, b) => {
        const aTid = typeof a.TeilnehmerID === 'number' ? a.TeilnehmerID : Number.MAX_SAFE_INTEGER;
        const bTid = typeof b.TeilnehmerID === 'number' ? b.TeilnehmerID : Number.MAX_SAFE_INTEGER;
        if (aTid !== bTid) return aTid - bTid;
        const aRd = new Date(a.RegistrationDate || 0).getTime();
        const bRd = new Date(b.RegistrationDate || 0).getTime();
        if (aRd !== bRd) return aRd - bRd;
        return a.Id - b.Id;
      });
      const promoteTarget = sorted[0];
      if (promoteTarget) {
        try {
          await eventService.promoteToTeamLead(subsiteUrl, promoteTarget.Id);
          newLeadId = promoteTarget.Id;
        } catch (err) {
          console.warn('[DEX] promoteToTeamLead failed:', err);
        }
      }
    }

    // Info-Mails an die verbleibenden Mitglieder.
    if (event.disableEmails) return;
    const lang = (event.emailLanguage || 'EN').toUpperCase();
    const isDe = lang === 'DE';
    const teamSizeCfg = event.teamSize || (remaining.length + 1);
    const cancelledFullName = `${cancelledReg.Vorname || ''} ${cancelledReg.Nachname || ''}`.trim() || cancelledReg.ParticipantEmail;
    const subject = isDe
      ? `Team-Update fuer Event ${event.title}`
      : `Team update for event ${event.title}`;
    const heading = isDe ? 'Team-Update' : 'Team update';
    const subheading = `Event ${event.title}`;

    for (const m of remaining) {
      const mFirst = m.Vorname || (m.ParticipantName || '').split(/[ ,]+/)[0] || '';
      const isNewLead = newLeadId !== null && m.Id === newLeadId;
      const greeting = isDe ? `Hallo ${mFirst},` : `Hello ${mFirst},`;
      const intro = isDe
        ? `<p>ein Mitglied deines Teams${teamName ? ` „${teamName}"` : ''} hat sich vom Event abgemeldet:</p>`
          + `<p style="padding:8px 12px;background:#f7f7f7;border-left:3px solid #86bc25;font-weight:600;">${cancelledFullName}</p>`
        : `<p>a member of your team${teamName ? ` „${teamName}"` : ''} has cancelled their registration:</p>`
          + `<p style="padding:8px 12px;background:#f7f7f7;border-left:3px solid #86bc25;font-weight:600;">${cancelledFullName}</p>`;
      const occupancy = isDe
        ? `<p>Aktuelle Team-Belegung: <strong>${remaining.length}/${teamSizeCfg}</strong></p>`
        : `<p>Current team occupancy: <strong>${remaining.length}/${teamSizeCfg}</strong></p>`;
      const leadPromote = isNewLead
        ? (isDe
          ? `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>Du bist jetzt der neue Team-Lead.</strong> Du kannst ueber „Meine Events" eine neue Person hinzufuegen, falls der frei gewordene Platz wieder gefuellt werden soll.</p>`
          : `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>You are the new team lead now.</strong> You can add a replacement member via „My Events" if you want to fill the freed slot.</p>`)
        : '';
      const options = isDe
        ? `<p>Was du jetzt machen kannst:</p>`
          + `<ul>`
          + `<li>Nichts tun — euer Platz bleibt erstmal fuer das Team reserviert.</li>`
          + `<li>Als Team-Lead: ueber „Meine Events" eine andere Person nachtraeglich hinzufuegen.</li>`
          + `<li>Andere Teilnehmer koennen ggf. den freien Slot ueber die Event-Anmeldeseite belegen (sofern der Organizer „Unvollstaendige Teams oeffentlich sichtbar" aktiviert hat).</li>`
          + `</ul>`
        : `<p>What you can do now:</p>`
          + `<ul>`
          + `<li>Do nothing — your seat stays reserved for the team for now.</li>`
          + `<li>As team lead: add a replacement person via „My Events".</li>`
          + `<li>Other participants can join the open slot via the registration page (if the organizer enabled „Public open slots").</li>`
          + `</ul>`;
      const closing = isDe ? `<p>Beste Gruesse,<br>Dein Event-Team</p>` : `<p>Best regards,<br>Your event team</p>`;
      const innerHtml = `<p>${greeting}</p>${intro}${occupancy}${leadPromote}${options}${closing}`;
      const html = wrapTemplate('#86bc25', heading, subheading, innerHtml);
      try {
        await eventService.queueEmail(
          subject,
          m.ParticipantEmail,
          `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail,
          html,
          'Info',
          event.title,
          eventId
        );
      } catch (err) {
        console.warn('[DEX] team-cancel info mail failed:', err);
      }
    }
  }

  async function getMyRegistration(eventId: string): Promise<SPRegistration | null> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return null;
    return eventService.getMyRegistration(subsiteUrl, currentUserEmail);
  }

  async function checkRegistrationByEmail(eventId: string, email: string): Promise<SPRegistration | null> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl || !email) return null;
    return eventService.getMyRegistration(subsiteUrl, email);
  }

  async function getAllRegistrations(eventId: string): Promise<SPRegistration[]> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return [];
    return eventService.getAllRegistrations(subsiteUrl);
  }

  async function updateEvent(eventId: string, updates: Record<string, unknown>): Promise<boolean> {
    const success = await eventService.updateEvent(Number(eventId), updates);
    if (success) {
      // v9.0: Audit-Log (fire-and-forget — UI-Save soll nicht haengen
      // falls SP-ChangeLog-Liste fehlt oder Permissions fehlen).
      const ev = events.find(e => e.id === eventId);
      eventService.writeChangeLog({
        action: 'EventUpdated',
        targetType: 'Event',
        targetId: eventId,
        targetName: ev?.title || '',
        eventId: eventId,
        eventTitle: ev?.title || '',
        details: { changedFields: Object.keys(updates) },
      }).catch(() => { /* */ });
      // v9.41: loadEvents im try/catch — wenn ein einzelner Event-Mapping (z.B.
      // ein frisch erstellter Sibling) fehlschlägt, soll das den updateEvent-
      // Erfolg nicht zu einem white-screen-blow-up führen. allSettled in loadEvents
      // selbst sollte das auch schon abfangen, hier nur belt-and-suspenders.
      try { await loadEvents(); } catch (err) { console.warn('[DEX] post-update loadEvents fehlgeschlagen:', err); }
    }
    return success;
  }

  async function deleteEvent(eventId: string): Promise<boolean> {
    // Seit v6.4: Sub-Events sind eigene DEX_Events-Items. Vor dem Löschen des
    // Parent-Events müssen alle Child-Events gelöscht werden, damit auch deren
    // Outlook-Kalendertermine, Subsites und Teilnehmerlisten aufgeräumt werden.
    const children = events.filter(e => e.parentEventId === eventId);
    for (const child of children) {
      try {
        await eventService.deleteEvent(Number(child.id));
        delete subsiteMap.current[child.id];
      } catch (err) {
        console.warn('[DEX] Child-Event-Delete fehlgeschlagen:', child.id, err);
      }
    }
    // v11.53: vor dem Loeschen merken, wie viele aktive Anmeldungen wir
    // vom KPI-Counter abziehen muessen — Parent + alle Children, nur
    // nicht-fictive Events. Wird im Hintergrund einmalig abgezogen.
    const ev = events.find(e => e.id === eventId);
    const childActive = children
      .filter(c => !c.isFictive)
      .reduce((s, c) => s + (c.currentParticipants || 0), 0);
    const parentActive = (ev && !ev.isFictive) ? (ev.currentParticipants || 0) : 0;
    const childEventsToDecrement = children.filter(c => !c.isFictive).length
      + ((ev && !ev.isFictive) ? 1 : 0);
    const success = await eventService.deleteEvent(Number(eventId));
    delete subsiteMap.current[eventId];
    if (success) {
      if (childEventsToDecrement > 0) {
        eventService.bumpKpiEvents(-childEventsToDecrement).catch(() => { /* */ });
      }
      const totalActive = childActive + parentActive;
      if (totalActive > 0) {
        eventService.bumpKpiParticipants(-totalActive).catch(() => { /* */ });
      }
    }
    // Events immer neu laden, auch wenn Subsite-Loeschung fehlschlug
    await loadEvents();
    return success;
  }

  /**
   * v11.69: Loescht ausschliesslich das DEX_Events-Listenitem — KEIN Cascade
   * auf Subsite, Teilnehmerliste oder Outlook-DeleteEvent-Queue.
   * Wird genutzt, um ein Sub-Event mit `existingSubsiteUrl` an einer neuen
   * DEX_Events-Zeile wieder anzulegen, damit der `DEX_CreateOutlookEvent`-
   * Flow (GetOnNewItems-Trigger) triggert — die alte Subsite mit allen
   * Teilnehmer-Anmeldungen bleibt unangetastet erhalten.
   */
  async function deleteEventItemOnly(eventId: string): Promise<boolean> {
    const success = await eventService.deleteEventItemOnly(eventId);
    if (success) {
      delete subsiteMap.current[eventId];
      eventService.writeChangeLog({
        action: 'EventItemOnlyDeleted',
        targetType: 'Event',
        targetId: eventId,
        targetName: (events.find(e => e.id === eventId)?.title) || '',
        eventId: eventId,
        eventTitle: (events.find(e => e.id === eventId)?.title) || '',
        details: { reason: 'Outlook-Recreate ohne Subsite-Verlust (v11.69)' },
      }).catch(() => { /* */ });
    }
    return success;
  }

  async function updateMyRegistration(eventId: string, customData: Record<string, string>): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;

    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;

    // FieldMap aus Event extrahieren
    const event = events.find(e => e.id === eventId);
    const fieldMap: Record<string, string> = {};
    if (event) {
      for (const f of event.eventSpecificFields) {
        if (f.spInternalName) fieldMap[f.id] = f.spInternalName;
      }
    }

    // Alte Daten und Labels fuer ChangeLog
    let oldCustomData: Record<string, string> = {};
    try {
      if (myReg.CustomData) oldCustomData = JSON.parse(myReg.CustomData);
    } catch { /* */ }

    const fieldLabelMap: Record<string, string> = {};
    if (event) {
      for (const f of event.eventSpecificFields) {
        fieldLabelMap[f.id] = f.label;
      }
    }

    const success = await eventService.updateRegistrationData(subsiteUrl, myReg.Id, customData, fieldMap, oldCustomData, fieldLabelMap);
    if (success) await loadEvents();
    return success;
  }

  async function getMyEventNumbers(): Promise<{ registered: number[]; waitlisted: number[] }> {
    try {
      const record = await eventService.getParticipantByEmail(currentUserEmail);
      if (!record) return { registered: [], waitlisted: [] };
      const registered = record.EventRegistered
        ? record.EventRegistered.split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n))
        : [];
      const waitlisted = record.EventOnWaitlist
        ? record.EventOnWaitlist.split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n))
        : [];
      return { registered, waitlisted };
    } catch {
      return { registered: [], waitlisted: [] };
    }
  }

  async function refreshEvents(): Promise<void> {
    await loadEvents();
  }

  // v11.0: Item-Attachments — Wrapper für die eigene Registrierung.
  async function listMyEventAttachments(eventId: string): Promise<Array<{ fileName: string; serverRelativeUrl: string }>> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return [];
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return [];
    return eventService.listRegistrationAttachments(subsiteUrl, myReg.Id);
  }
  async function uploadMyEventAttachment(eventId: string, file: File): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;
    return eventService.addRegistrationAttachment(subsiteUrl, myReg.Id, file);
  }
  async function deleteMyEventAttachment(eventId: string, fileName: string): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;
    return eventService.deleteRegistrationAttachment(subsiteUrl, myReg.Id, fileName);
  }

  // v10.27: Split-Capacity-Gruppen-Wechsel — wrappt EventService.switchSplitGroup,
  // ergänzt um Mail/Outlook-Sideeffects und Reload.
  async function switchSplitGroup(eventId: string, newType: 'Durchstarter' | 'Funstarter'): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, status: 'Failed', full: false };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, status: 'Failed', full: false };
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return { ok: false, status: 'Failed', full: false };
    const result = await eventService.switchSplitGroup(
      subsiteUrl,
      myReg.Id,
      newType,
      event.durchstarterCapacity || 0,
      event.funstarterCapacity || 0,
    );
    if (result.ok) {
      // Audit-Log + Mail/Outlook anstoßen — analog zu cancelRegistration.
      eventService.writeChangeLog({
        action: 'ParticipantSwitchedGroup',
        targetType: 'Participant',
        targetId: currentUserEmail,
        targetName: currentUserName,
        eventId, eventTitle: event.title,
        details: { from: myReg.StarterType || myReg.PreferredStarterType || '', to: newType, finalStatus: result.status },
      }).catch(() => { /* */ });
      if (!event.disableEmails) {
        try {
          const lang = event.emailLanguage || 'EN';
          const isDeMail = lang.toUpperCase() === 'DE';
          const labelA = (event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
          const labelB = (event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
          const newLabel = newType === 'Durchstarter' ? labelA : labelB;
          const subj = isDeMail
            ? (result.status === 'Warteliste'
              ? `Gruppen-Wechsel — auf Warteliste: ${event.title}`
              : `Gruppen-Wechsel bestätigt: ${event.title}`)
            : (result.status === 'Warteliste'
              ? `Group switch — added to waitlist: ${event.title}`
              : `Group switch confirmed: ${event.title}`);
          const innerBody = isDeMail
            ? (result.status === 'Warteliste'
              ? `<p>Du hast den Wechsel in die Gruppe <strong>${newLabel}</strong> für <strong>${event.title}</strong> angefragt. Diese Gruppe ist aktuell voll, daher steht deine Anmeldung auf der <strong>Warteliste der Gruppe ${newLabel}</strong>. Sobald jemand absagt, rückst du automatisch nach.</p>`
              : `<p>Dein Gruppen-Wechsel zu <strong>${newLabel}</strong> für <strong>${event.title}</strong> ist bestätigt. Du bist jetzt regulär in dieser Gruppe angemeldet.</p>`)
            : (result.status === 'Warteliste'
              ? `<p>You requested to switch to the <strong>${newLabel}</strong> group for <strong>${event.title}</strong>. The group is currently full, so your registration is on the <strong>${newLabel} waitlist</strong>. You will be promoted automatically as soon as a spot frees up.</p>`
              : `<p>Your group switch to <strong>${newLabel}</strong> for <strong>${event.title}</strong> is confirmed. You are now regularly registered in this group.</p>`);
          const heading = isDeMail ? 'Gruppen-Wechsel' : 'Group switch';
          const body = wrapTemplate('#86bc25', heading, event.title, innerBody);
          await eventService.queueEmail(subj, currentUserEmail, currentUserName, body, 'Info', event.title, eventId)
            .catch(err => console.warn('[DEX] switchSplitGroup mail failed:', err));
        } catch { /* */ }
      }
      await loadEvents();
    }
    return result;
  }

  /**
   * Admin-Cleanup: Events mit Status='Active' + EndDate < jetzt auf 'Completed' setzen.
   * Anschliessend wird die Event-Liste neu geladen, damit die UI die neuen Status sieht.
   */
  async function markExpiredEventsAsCompleted(): Promise<number> {
    const n = await eventService.markExpiredEventsAsCompleted();
    if (n > 0) await loadEvents();
    return n;
  }

  /**
   * Anfrage von der Landing Page an die DEX-Admins. Verwendet DEX_SEND_MAIL via
   * der DEX_Emails-Queue, mit dem Anfrager im Cc-Feld. Body wird ins Deloitte-
   * Template (gruener Header, Footer) gewrappt.
   */
  async function sendAdminInquiry(
    requesterName: string,
    requesterEmail: string,
    eventName: string,
    message: string
  ): Promise<boolean> {
    const adminTo = 'ebrenneisen@deloitte.de;nifelten@deloitte.de;aenk@deloitte.de';
    const subject = `DEX-Anfrage: ${eventName || 'Event ohne Titel'} (von ${requesterName || 'unbekannt'})`;
    const escape = (s: string): string => s
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    const messageHtml = escape(message).replace(/\r?\n/g, '<br>');
    const bodyInner = `
      <p>Hallo DEX-Team,</p>
      <p>es gibt eine neue Anfrage zur DEX Event Experience Platform:</p>
      <table cellpadding="6" cellspacing="0" style="border-collapse:collapse;margin:8px 0;">
        <tr><td style="color:#555;font-weight:600;vertical-align:top;">Name:</td><td>${escape(requesterName)}</td></tr>
        <tr><td style="color:#555;font-weight:600;vertical-align:top;">E-Mail:</td><td><a href="mailto:${escape(requesterEmail)}">${escape(requesterEmail)}</a></td></tr>
        <tr><td style="color:#555;font-weight:600;vertical-align:top;">Event:</td><td>${escape(eventName)}</td></tr>
      </table>
      <p style="color:#555;font-weight:600;margin-bottom:4px;">Worum geht es:</p>
      <p>${messageHtml}</p>
      <p style="margin-top:24px;color:#888;font-size:0.85rem;">${escape(requesterName)} ist im Cc und kann direkt geantwortet werden.</p>
    `;
    const body = wrapTemplate('#86bc25', 'Neue DEX-Anfrage', `Event: ${eventName || '-'}`, bodyInner);
    // EventId muss '0' sein (nicht ''), damit der DEX_SEND_MAIL Flow Get_Event
    // mit "ID eq 0" als gueltigem OData-Filter aufrufen kann. Bei leerem
    // EventId baut der Flow "ID eq " was kein gueltiger OData-Ausdruck ist
    // und der Flow direkt in Get_Event failed (clientRequestId-Fehler).
    // Get_Event liefert dann 0 Items, Compose_Image faellt automatisch auf
    // das Default-Bild aus _Config zurueck - die Mail geht trotzdem raus.
    return eventService.queueEmail(
      subject, adminTo, 'DEX Admin Team', body, 'Info', eventName || 'DEX-Anfrage', '0', requesterEmail
    );
  }

  /**
   * Onboarding-Mail an einen neu ernannten Organizer (oder Admin) verschicken.
   * Subject + Body kommen aus EmailTemplates.organizerOnboardingEmail (Deloitte-
   * Layout inkl. Header/Footer). Die DEX-Verantwortlichen werden im Cc
   * informiert. EventId='0' damit der DEX_SEND_MAIL Flow den Get_Event-Step
   * mit gueltigem OData-Filter ausfuehren kann (analog sendAdminInquiry).
   */
  async function sendOrganizerOnboarding(
    recipientEmail: string,
    recipientName: string,
    role: 'Organizer' | 'Admin'
  ): Promise<boolean> {
    if (!recipientEmail || !recipientName) return false;
    const cc = 'ebrenneisen@deloitte.de;nifelten@deloitte.de';
    const { subject, body } = organizerOnboardingEmail(recipientName, role);
    return eventService.queueEmail(
      subject, recipientEmail, recipientName, body, 'Info', 'DEX-Onboarding', '0', cc
    );
  }


  // ==================== Sub-Event-Helper (v6.4+) ====================
  // Seit v6.4 sind Sub-Events keine separaten JSON-Arrays mehr, sondern
  // eigene DEX_Events-Items mit gesetztem parentEventId. Damit funktionieren
  // alle bestehenden Flows (DEX_CreateOutlookEvent, DEX_Outlook_Einladungen,
  // Teilnehmerliste, Organizer-Kalendereinladungen, Declines, QR-Codes,
  // Warteliste, ...) unverändert — ein Sub-Event ist einfach ein Event.
  const topLevelEvents = React.useMemo(
    () => events.filter(e => !e.parentEventId),
    [events]
  );
  const childEventsOf = React.useCallback(
    (parentEventId: string): DeloitteEvent[] => {
      if (!parentEventId) return [];
      return events
        .filter(e => e.parentEventId === parentEventId)
        .slice()
        .sort((a, b) => (a.startDate || '').localeCompare(b.startDate || ''));
    },
    [events]
  );

  return React.createElement(
    EventContext.Provider,
    {
      value: {
        events, topLevelEvents, childEventsOf, isEventsLoading,
        createEvent, registerForEvent, registerTeam,
        getTeamMembers: async (eventId: string, teamId: string): Promise<SPRegistration[]> => {
          const subsiteUrl = subsiteMap.current[eventId];
          if (!subsiteUrl) return [];
          return eventService.getTeamMembers(subsiteUrl, teamId);
        },
        addTeamMember,
        joinTeam,
        transferTeamLead,
        createTeamJoinRequest,
        listTeamJoinRequestsForEvent,
        decideTeamJoinRequest,
        listOpenTeamsForEvent,
        cancelRegistration,
        cancelTeamMember,
        getMyRegistration, checkRegistrationByEmail, getAllRegistrations, deleteEvent, deleteEventItemOnly, updateEvent, updateMyRegistration, switchSplitGroup, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, getMyEventNumbers, refreshEvents, refreshParticipantCounts, markExpiredEventsAsCompleted,
        sendAdminInquiry,
        sendOrganizerOnboarding,
        getKpiCache: () => eventService.getKpiCache(),
        updateKpiCache: (v) => eventService.updateKpiCache(v),
      },
    },
    props.children
  );
}

export function useEvents(): EventContextType {
  const ctx = React.useContext(EventContext);
  if (!ctx) throw new Error('useEvents must be used within EventProvider');
  return ctx;
}
