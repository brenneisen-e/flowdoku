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
import { EventService, SPEvent, CustomField, SPRegistration, ReseedSummary } from '../services/EventService';
import { verifyRotatingCode, isWithinCheckInWindow } from '../utils/selfCheckIn';
import { registrationEmail, waitlistEmail, cancellationEmail, buildEmailFromTemplate, loadLogosAsBase64, wrapTemplate, organizerOnboardingEmail, qrCodeEmail, teamInfoBlockHtml, injectIntoEmailContent } from '../services/EmailTemplates';
import * as QRCode from 'qrcode';
import { APP_VERSION } from '../version';
import { buildDemoShowcaseEvents, isDemoShowcaseId, buildDemoRegistrations } from '../services/demoShowcaseEvent';

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
  spTemplate: { subject: string; headingColor: string; heading: string; subheading?: string; bodyHtml: string } | null,
  overridesJson: string | undefined,
  templateType: string
): { subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string; headingFontSize?: string; headingBold?: boolean; headingItalic?: boolean; subheadingColor?: string; subheadingFontSize?: string; subheadingBold?: boolean; subheadingItalic?: boolean; imageWidth?: number; imagePaddingV?: number; imagePaddingH?: number } | null {
  // v15.19: Subheading-Override pro Event mitziehen. Color/Size bleiben
  // weiterhin aus dem Standard-Template (wrapTemplate-Layout fest), nur
  // die Text-Werte (Subject, Heading, Subheading, Body) sind editierbar.
  if (!overridesJson) {
    if (!spTemplate) return null;
    return {
      subject: spTemplate.subject,
      headingColor: spTemplate.headingColor || '#86bc25',
      heading: spTemplate.heading,
      subheading: spTemplate.subheading || '',
      bodyHtml: spTemplate.bodyHtml,
    };
  }
  try {
    const all = JSON.parse(overridesJson) as Record<string, { subject?: string; heading?: string; subheading?: string; bodyHtml?: string; headingColor?: string; headingFontSize?: string; headingBold?: boolean; headingItalic?: boolean; subheadingColor?: string; subheadingFontSize?: string; subheadingBold?: boolean; subheadingItalic?: boolean }>;
    // v18.73: globales Header-Bild-Layout (Breite + Innenabstand). Liegt unter
    // dem reservierten Piggyback-Key `_headerImageLayout` und gilt für ALLE
    // Template-Typen des Events — daher hier einmal gelesen und in jeden
    // Rückgabe-Zweig gespreadet (auch wenn der konkrete Typ keinen Text-
    // Override hat).
    const il = (all as unknown as { _headerImageLayout?: { width?: number; paddingV?: number; paddingH?: number } })._headerImageLayout || {};
    const imgSpread = {
      ...(typeof il.width === 'number' && il.width > 0 ? { imageWidth: il.width } : {}),
      ...(typeof il.paddingV === 'number' && il.paddingV >= 0 ? { imagePaddingV: il.paddingV } : {}),
      ...(typeof il.paddingH === 'number' && il.paddingH >= 0 ? { imagePaddingH: il.paddingH } : {}),
    };
    const o = all[templateType];
    if (!o || (!o.subject && !o.heading && o.subheading === undefined && !o.bodyHtml && !o.headingColor && !o.headingFontSize && o.headingBold === undefined && o.headingItalic === undefined && !o.subheadingColor && !o.subheadingFontSize && o.subheadingBold === undefined && o.subheadingItalic === undefined)) {
      if (!spTemplate) return null;
      return {
        subject: spTemplate.subject,
        headingColor: spTemplate.headingColor || '#86bc25',
        heading: spTemplate.heading,
        subheading: spTemplate.subheading || '',
        bodyHtml: spTemplate.bodyHtml,
        ...imgSpread,
      };
    }
    return {
      subject: o.subject || spTemplate?.subject || '',
      heading: o.heading || spTemplate?.heading || '',
      // Override.subheading ist „intentional set" — auch leerer String
      // soll respektiert werden, damit man die zweite Zeile abschalten kann.
      subheading: o.subheading !== undefined ? o.subheading : (spTemplate?.subheading || ''),
      bodyHtml: o.bodyHtml || spTemplate?.bodyHtml || '',
      // v18.19: Überschrift-Farbe + -Größe pro Event überschreibbar.
      headingColor: o.headingColor || spTemplate?.headingColor || '#86bc25',
      ...(o.headingFontSize ? { headingFontSize: o.headingFontSize } : {}),
      // v18.22: Fett/Kursiv (Überschrift) + Unter-Überschrift-Formatierung.
      ...(o.headingBold !== undefined ? { headingBold: o.headingBold } : {}),
      ...(o.headingItalic !== undefined ? { headingItalic: o.headingItalic } : {}),
      ...(o.subheadingColor ? { subheadingColor: o.subheadingColor } : {}),
      ...(o.subheadingFontSize ? { subheadingFontSize: o.subheadingFontSize } : {}),
      ...(o.subheadingBold !== undefined ? { subheadingBold: o.subheadingBold } : {}),
      ...(o.subheadingItalic !== undefined ? { subheadingItalic: o.subheadingItalic } : {}),
      ...imgSpread,
    };
  } catch {
    if (!spTemplate) return null;
    return {
      subject: spTemplate.subject,
      headingColor: spTemplate.headingColor || '#86bc25',
      heading: spTemplate.heading,
      subheading: spTemplate.subheading || '',
      bodyHtml: spTemplate.bodyHtml,
    };
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

/** v18.41: Sammelt die E-Mail-Adressen aus People-Picker-Feldern (user/roommate),
 *  die der Organizer als „CC bei An-/Abmelde-Mail" markiert hat. Format des
 *  Feldwerts ist „Anzeigename <email>". Liefert einen ';'-getrennten CC-String
 *  (ohne den Teilnehmer selbst, dedupliziert). NUR für Mails — nicht Outlook. */
export function collectCcEmailsFromFields(
  fields: Array<{ id: string; type: string; ccOnEmails?: boolean }> | undefined,
  customData: Record<string, string>,
  excludeEmail?: string
): string {
  const seen = new Set<string>();
  const out: string[] = [];
  const exclude = (excludeEmail || '').toLowerCase();
  for (const f of (fields || [])) {
    if (!f.ccOnEmails) continue;
    if (f.type !== 'user' && f.type !== 'roommate') continue;
    const v = customData[f.id];
    if (!v) continue;
    const m = v.match(/<([^>]+@[^>]+)>/);
    const em = m ? m[1].trim() : '';
    const lc = em.toLowerCase();
    if (em && lc !== exclude && !seen.has(lc)) { seen.add(lc); out.push(em); }
  }
  return out.join(';');
}

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

interface EventContextType {
  events: DeloitteEvent[];
  /** Top-Level-Events (ohne parentEventId) — was in EventListPage/MyEventsPage angezeigt wird. */
  topLevelEvents: DeloitteEvent[];
  /** Kind-Events eines Parents (Sub-Events / Trainingssessions), sortiert nach StartDate. */
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  isEventsLoading: boolean;
  createEvent: (event: CreateEventInput) => Promise<number | null>;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string, opts?: { suppressMail?: boolean; suppressOutlook?: boolean; extraCc?: string; proxyConsentConfirmed?: boolean }) => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' }>;
  /** v11.82: Team-Anmeldung — Lead + N-1 Mitglieder gleichzeitig anmelden.
   *  Reserviert N Plaetze atomar; bei Vollbelegung geht das ganze Team auf
   *  die Warteliste (keine Teil-Anmeldungen aus Kapazitaetsmangel). */
  registerTeam: (
    eventId: string,
    leadData: { firstName: string; lastName: string; email: string; salutation?: string; customData: Record<string, string>; preferredStarterType?: string },
    members: Array<{ email: string; displayName: string; customData?: Record<string, string> }>,
    teamName: string | undefined
  ) => Promise<{ ok: boolean; teamId?: string; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.82: Andere Team-Mitglieder zu einer Registrierung laden — fuer das
   *  Team-Badge in „Meine Events". */
  getTeamMembers: (eventId: string, teamId: string) => Promise<SPRegistration[]>;
  /** v11.83: Ein Team-Lead kann nachtraeglich ein einzelnes Mitglied
   *  zum bereits angemeldeten Team hinzufuegen (Plus-Button in MyEvents).
   *  Atomar einen Sitzplatz reservieren, neuen Member-Eintrag anlegen,
   *  Bestaetigungs-Mail + Outlook-Termin queuen. */
  addTeamMember: (eventId: string, teamId: string, teamName: string | undefined, member: { email: string; displayName: string }, customData?: Record<string, string>) => Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v17.2: Schon angemeldete Person (ohne TeamId) einem Team zuweisen.
   *  PATCHt nur die TeamId/TeamName/TeamLead-Felder, KEINE neue
   *  Registrierung, KEINE Bestaetigungsmail, KEIN Outlook. */
  assignTeamlessToTeam: (eventId: string, teamId: string, teamName: string | undefined, existingRegId: number, isLead?: boolean) => Promise<boolean>;
  /** v11.83: Direkter Team-Beitritt aus der Anmeldeseite (wenn der
   *  Organizer "Beitritt erfordert Bestaetigung" NICHT aktiviert hat).
   *  Verhalten wie `addTeamMember`, aber laeuft mit dem eingeloggten User
   *  selbst als neuem Member. */
  joinTeam: (eventId: string, teamId: string, teamName: string | undefined, customData?: Record<string, string>) => Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  /** v11.84: Team-Lead-Rolle innerhalb eines Teams uebergeben — nur im
   *  Admin Center fuer Admin/Organizer eigener Events sichtbar. Setzt die
   *  alte Lead-Zeile auf TeamLead=false und die neue auf TeamLead=true,
   *  schickt anschliessend eine Info-Mail an alle aktiven Mitglieder. */
  transferTeamLead: (eventId: string, teamId: string, newLeadEmail: string) => Promise<{ ok: boolean; reason?: string }>;
  /** v11.83: Beitritts-Anfrage in DEX_TeamJoinRequests einreichen — fuer
   *  Events bei denen der Organizer Approval aktiviert hat. */
  createTeamJoinRequest: (eventId: string, teamId: string, customData?: Record<string, string>) => Promise<{ ok: boolean; itemId?: number; reason?: string }>;
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
  cancelRegistration: (eventId: string, opts?: { suppressNotifications?: boolean }) => Promise<boolean>;
  /** v18.11: Proaktive Absage durch einen (noch nicht angemeldeten) Teilnehmer
   *  — „Ich nehme nicht teil". Landet als Abgemeldet-Eintrag im Admin-Center. */
  declineEvent: (eventId: string) => Promise<boolean>;
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
  /** v18.33: Self-Check-in über einen gescannten QR-Deep-Link. Löst das Event
   *  per Token (statischer QR) oder Event-Nummer + HMAC-Code (rotierender QR)
   *  auf, validiert Fenster/Frische und setzt die eigene Registrierung auf
   *  „Eingecheckt". Gibt ein strukturiertes Ergebnis für die Ergebnis-UI. */
  selfCheckIn: (params: SelfCheckInParams) => Promise<SelfCheckInResult>;
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
  // v19.0: Dokument-Custom-Felder (pro-Feld-Attachments).
  uploadFieldDocument: (eventId: string, fieldId: string, file: File, participantEmail?: string) => Promise<boolean>;
  listFieldDocuments: (eventId: string, fieldId: string, participantEmail?: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>>;
  deleteFieldDocument: (eventId: string, fileName: string, participantEmail?: string) => Promise<boolean>;
  getMyEventNumbers: () => Promise<{ registered: number[]; waitlisted: number[] }>;
  refreshEvents: () => Promise<void>;
  refreshParticipantCounts: (eventId?: string) => Promise<void>;
  markExpiredEventsAsCompleted: () => Promise<number>;
  sendAdminInquiry: (requesterName: string, requesterEmail: string, eventName: string, message: string) => Promise<boolean>;
  /** v12.12: Admin-Aktion zum Re-Seed der Default-Email-Templates in
   *  DEX_EmailTemplates. Überschreibt die aktuelle Subject/Heading/BodyHtml
   *  jedes Standard-Templates mit den Default-Werten aus dem Code. */
  reseedDefaultEmailTemplates: () => Promise<ReseedSummary>;
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
  /** v16.4: Pre-resolved DL members (';'-separated, lowercase). */
  audienceResolvedEmails?: string;
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
  /** v18.35: erzwungene Anmeldeseiten-Sprache ('de' | 'en'); leer = App-Sprache. */
  registrationLanguage?: 'de' | 'en';
  /** v18.40: manueller Outlook-Termin-Ort; leer = Auto aus Veranstaltungsort + Adresse. */
  outlookLocation?: string;
  /** v18.42: Betreff des Outlook-Termins; leer = Event-Titel. */
  outlookSubject?: string;
  /** v18.44: abweichende Outlook-Start/-Ende (ISO); leer = Event-Datum. */
  outlookStart?: string;
  outlookEnd?: string;
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
        safeRun('ensureOutlookLocksList', () => eventService.ensureOutlookLocksList(), parallelMarks),
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
      // v11.89: 'Under Construction' aus Legacy-Daten transparent auf 'Active'
      // mappen — der Entwurfs-Zustand lebt jetzt auf IsFictive.
      status: (e.EventStatus === 'Under Construction' ? 'Active' : (e.EventStatus as DeloitteEvent['status'])) || 'Active',
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
      // v16.4: vor-aufgeloeste Member-E-Mails der Audience-DLs.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      audienceResolvedEmails: ((e as any).AudienceResolvedEmails || '')
        .split(';').map((s: string) => s.trim().toLowerCase()).filter(Boolean),
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
      outlookSubject: e.OutlookSubject || undefined,
      outlookStart: e.OutlookStart || undefined,
      outlookEnd: e.OutlookEnd || undefined,
      outlookLocation: e.OutlookLocation || undefined,
      outlookEventId: e.OutlookEventId || '',
      // v11.61: CalendarLink (iCalUId) muss in den Event-Type, weil der
      // DEX_CreateOutlookEvent-Flow nur dieses Feld auf Erfolg setzt — die
      // v11.57-Modal-Erkennung hatte auf OutlookEventId geprueft (immer leer)
      // und das Outlook-Update-Confirm-Modal kam deshalb nie.
      calendarLink: e.CalendarLink || '',
      emailLanguage: e.EmailLanguage || 'EN',
      // v18.35: erzwungene Anmeldeseiten-Sprache (nur 'de'/'en' gültig, sonst undefined).
      registrationLanguage: (e.RegistrationLanguage === 'de' || e.RegistrationLanguage === 'en') ? e.RegistrationLanguage : undefined,
      emailTemplateOverrides: e.EmailTemplateOverrides || '',
      disableEmails: !!e.DisableEmails,
      disableOutlook: !!e.DisableOutlook,
      // v14.5: requireSubEventSelection als Piggyback im EmailTemplateOverrides-
      // JSON (kein neues SP-Feld nötig).
      // v14.8: subEventsOnlyMode + childEventTerm zusätzlich aus dem
      // Piggyback-Blob auslesen. subEventsOnlyMode impliziert
      // requireSubEventSelection (auch wenn der Flag nicht explizit gesetzt
      // ist).
      requireSubEventSelection: ((): boolean => {
        try {
          const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
          return !!(ov && (ov._requireSubEventSelection || ov._subEventsOnlyMode));
        } catch { return false; }
      })(),
      subEventsOnlyMode: ((): boolean => {
        try {
          const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
          return !!(ov && ov._subEventsOnlyMode);
        } catch { return false; }
      })(),
      // v18.9: Organizer-Anzeige ausblenden (Piggyback _hideOrganizer).
      hideOrganizer: ((): boolean => {
        try {
          const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
          return !!(ov && ov._hideOrganizer);
        } catch { return false; }
      })(),
      childEventTermSingular: ((): string | undefined => {
        try {
          const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
          const v = ov && ov._childEventTerm && typeof ov._childEventTerm.singular === 'string' ? ov._childEventTerm.singular : '';
          return v || undefined;
        } catch { return undefined; }
      })(),
      childEventTermPlural: ((): string | undefined => {
        try {
          const ov = JSON.parse(e.EmailTemplateOverrides || '{}');
          const v = ov && ov._childEventTerm && typeof ov._childEventTerm.plural === 'string' ? ov._childEventTerm.plural : '';
          return v || undefined;
        } catch { return undefined; }
      })(),
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
      // v11.89: Legacy-Events mit EventStatus='Under Construction' werden
      // auch ohne explizites IsFictive-Flag als Entwurf erkannt — bis die
      // Migration im Hintergrund das SP-Item neu geschrieben hat.
      isFictive: !!e.IsFictive || e.EventStatus === 'Under Construction',
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
      confirmDialogEnabled: !!e.ConfirmDialogEnabled,
      confirmDialogMode: e.ConfirmDialogMode || '',
      confirmDialogText: e.ConfirmDialogText || '',
      // v18.33: Self-Check-in. Alte Tenants ohne diese Spalten lesen undefined
      // als false / leer — Self-Check-in bleibt dann schlicht aus.
      selfCheckInEnabled: !!e.SelfCheckInEnabled,
      selfCheckInToken: e.SelfCheckInToken || undefined,
      selfCheckInFrom: e.SelfCheckInFrom || undefined,
      selfCheckInTo: e.SelfCheckInTo || undefined,
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
      // v17.20: Bilingual-Toggle fuer Custom-Fields (DE + EN).
      bilingualFields: !!e.BilingualFields,
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
        // v18.18: Darstellungs-Stil der Beschreibung (tooltip|inline).
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        helpTextStyle: (cf as any).helpTextStyle === 'inline' ? 'inline' : 'tooltip',
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
        // v18.41: CC-bei-Mail-Flag durchreichen — collectCcEmailsFromFields liest es.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ccOnEmails: !!(cf as any).ccOnEmails,
        // v11.16: onlyForGroup aus dem persistierten Feld durchreichen.
        // Wurde im Wizard sauber gespeichert (CustomFields-JSON enthaelt
        // den Schluessel), aber der Loader hat ihn nie zurueckgelesen —
        // Folge: die Gruppen-spezifische Sichtbarkeit (Funstarter only /
        // Durchstarter only) hat in der Registrierungs-UI nie gegriffen,
        // weil die Filter-Chain auf undefined gefallen ist.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        onlyForGroup: (cf as any).onlyForGroup,
        // v17.19: confirmLabel (Text neben Checkbox) im Mapping nachgezogen
        // — vorher hier vergessen, Folge: Wizard speicherte den Text korrekt,
        // RegistrationPage fiel aber immer auf den Default „Ja, bestätigen"
        // zurueck, weil das Field-Mapping confirmLabel droppte.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        confirmLabel: (cf as any).confirmLabel,
        // v17.20: Englische Varianten durchreichen — nur dann wirksam, wenn
        // auf Event-Ebene `bilingualFields=true` ist; die RegistrationPage
        // entscheidet zur Laufzeit, ob sie die EN-Spalte zieht.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        labelEn: (cf as any).labelEn,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        helpTextEn: (cf as any).helpTextEn,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        confirmLabelEn: (cf as any).confirmLabelEn,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        optionsEn: (cf as any).optionsEn,
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
    preferredStarterType?: string, // B2Run: 'Durchstarter' | 'Funstarter'
    // v18.13: Massenimport — pro Anmeldung Bestätigungsmail bzw. Outlook-Termin
    // unterdrücken („stille Anmeldung").
    // v18.53: extraCc — zusätzliche CC-Adressen, die NICHT aus den
    // event-eigenen Feldern stammen. Genutzt im subEventsOnlyMode: das
    // „Hauptevent" ist dort nicht anmeldbar, die CC-Felder (z.B. Assistenz)
    // sind übergreifend und gelten für die Sub-Events — also müssen die
    // Sub-Event-Bestätigungsmails ebenfalls an die Assistenz auf CC gehen.
    // v18.74: proxyConsentConfirmed — bei stellvertretender Anmeldung wurde die
    // Zustimmung der Person bestätigt (Pflicht-Checkbox auf der Anmeldeseite).
    // Wird als Nachweis in die SP-Spalte ProxyConsent geschrieben.
    opts?: { suppressMail?: boolean; suppressOutlook?: boolean; extraCc?: string; proxyConsentConfirmed?: boolean }
  ): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' }> {
    // v17.25: Demo-Showcase-Event → No-Op, kein SP-Roundtrip. Die Register-
    // Seite blockt den Submit ohnehin mit einem Demo-Hinweis; dieser Guard
    // ist die zweite Verteidigungslinie.
    if (isDemoShowcaseId(eventId)) return { ok: true, status: 'Angemeldet' };
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, status: 'Warteliste' };

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
    if (existing && existing.Status !== 'Abgemeldet') return { ok: false, status: 'Warteliste' };

    // Pruefen ob Platz frei oder Waitlist
    const event = events.find(e => e.id === eventId);
    let status: 'Angemeldet' | 'Warteliste' = 'Angemeldet';
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

    // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung. Eine
    // Anmeldung gilt als „stellvertretend", wenn die Teilnehmer-E-Mail von der
    // des eingeloggten Users abweicht. Bei externen Adressen (kein @deloitte.de)
    // ist die Zustimmung schriftlich einzuholen — das wird im Nachweis vermerkt.
    const isProxyRegistration = (emailToUse || '').toLowerCase() !== (currentUserEmail || '').toLowerCase();
    const isExternalParticipant = !!emailToUse && !/@(.*\.)?deloitte\.de$/i.test(emailToUse);
    const proxyConsentStr = (isProxyRegistration && opts?.proxyConsentConfirmed)
      ? `${isExternalParticipant ? 'Schriftliche ' : ''}Zustimmung der Person zur stellvertretenden Anmeldung bestätigt durch ${actorName} (${actorEmail}) am ${new Date().toLocaleString('de-DE')}`
      : '';

    let success: boolean;
    if (existing && existing.Status === 'Abgemeldet') {
      success = await eventService.reactivateRegistration(subsiteUrl, existing.Id, firstNameToUse, lastNameToUse, customData, status, fieldMap, actorName, actorEmail, proxyConsentStr);
    } else {
      success = await eventService.registerForEvent(
        subsiteUrl, firstNameToUse, lastNameToUse, emailToUse, customData, status, fieldMap,
        effectiveStarterType, preferredStarterType, actorName, actorEmail, proxyConsentStr
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
      // v15.25: Im subEventsOnlyMode wird das Hauptevent nur als
      // „Schatten-Registrierung" angelegt (Daten-Zeile fuer Parent-CFs).
      // Der User nimmt nicht am Hauptevent teil und soll dafuer KEINE
      // Bestaetigungs-Mail und KEINEN Outlook-Termin bekommen — die
      // tatsaechlichen Teilnahme-Mails kommen pro Sub-Event.
      const suppressParentNotifications = !!event.subEventsOnlyMode;
      if (!event.disableEmails && !suppressParentNotifications && !opts?.suppressMail) {
        // v8.5: Organizer-BCC-Modus auswerten. Bei 'always' immer BCC,
        // bei 'fromDate' nur wenn das konfigurierte Datum bereits erreicht
        // ist, bei 'never'/undefined keinen BCC.
        let bcc: string | undefined;
        const mode = event.notifyOrgRegisterMode || 'never';
        if (mode === 'always' || (mode === 'fromDate' && event.notifyOrgRegisterFromDate && new Date() >= new Date(event.notifyOrgRegisterFromDate))) {
          const orgEmails = (event.organizerEmails || []).filter(Boolean);
          if (orgEmails.length > 0) bcc = orgEmails.join(';');
        }
        // v9.22: Externe Mail-Adresse erkennen — kein Deloitte-Postfach
        // (@deloitte.de; auch @deloitte.com/Global zaehlt nicht als intern).
        // v18.74: Bei externen Empfaengern wird die Bestaetigungsmail jetzt
        // DIREKT an die externe Person versendet (vorher an den Organizer
        // umgeleitet) — mit dem Organizer auf CC (Nachweis/Kopie). Ein
        // Outlook-Kalendereintrag wird fuer externe Adressen weiterhin NICHT
        // versendet (Microsoft blockt das ohne Federation, s.u.
        // skipOutlookForExternal).
        const isExternalRecipient = !!emailToUse && !/@(.*\.)?deloitte\.de$/i.test(emailToUse);
        const finalRecipient = emailToUse;
        const finalSubject = emailData.subject;
        let finalBody = emailData.body;
        const finalRecipientName = nameToUse;
        // CC-Adressen, die zusaetzlich zu den Feld-CCs gelten (Organizer bei
        // externer Anmeldung).
        let externalCcExtra = '';
        if (isExternalRecipient) {
          const orgEmails = (event.organizerEmails || []).filter(Boolean);
          externalCcExtra = orgEmails.join(';');
          // Hinweis-Box VOR dem Original-Body — adressiert an die externe Person.
          const externalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
            + `<strong>Externe Anmeldung — kein automatischer Kalendereintrag.</strong><br>`
            + `Diese Anmeldebestätigung gehört zu <strong>${nameToUse}</strong> (<strong>${emailToUse}</strong>). Da es sich um eine externe Adresse (kein Deloitte-Postfach) handelt, wird <strong>kein Outlook-Kalendereintrag</strong> versendet — bitte trage dir den Termin manuell in deinen Kalender ein. `
            + `Der Organizer ist zur Bestätigung auf Kopie (CC).`
            + `</div>`;
          // Body kommt schon als komplett-gewickeltes HTML (Deloitte-Template).
          // Wir injecten den Hinweis direkt nach dem opening-<body>-Tag.
          finalBody = injectIntoEmailContent(finalBody, externalHint);
        }
        // v18.41: People-Picker-Felder mit „CC bei Mail" → ausgewählte
        // Person(en) auf CC der An-/Warteliste-Mail (NICHT im Outlook-Termin).
        // v18.53: opts.extraCc (übergreifende CC aus dem Hauptformular im
        // subEventsOnlyMode) mit den event-eigenen CC-Feldern mergen + deduppen.
        const ccOwn = collectCcEmailsFromFields(event.eventSpecificFields, customData, emailToUse);
        const ccMerged = (() => {
          const seen = new Set<string>();
          const out: string[] = [];
          // v18.74: externalCcExtra (Organizer bei externer Anmeldung) mitmergen.
          for (const part of [ccOwn, opts?.extraCc || '', externalCcExtra]) {
            for (const e of part.split(';').map(s => s.trim()).filter(Boolean)) {
              const lc = e.toLowerCase();
              if (lc !== (emailToUse || '').toLowerCase() && !seen.has(lc)) { seen.add(lc); out.push(e); }
            }
          }
          return out.join(';');
        })();
        const ccFromFields = ccMerged || undefined;
        eventService.queueEmail(
          finalSubject, finalRecipient, finalRecipientName, finalBody,
          templateType, event.title, eventId, ccFromFields, bcc
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
                const qrBody = injectIntoEmailContent(qrMail.body, qrExternalHint);
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
          // v13.0: Roommate-Mail aus TemplateType=RoommateRequest.
          const tpl = await eventService.getEmailTemplate('RoommateRequest', lang).catch(() => null);
          const vars: Record<string, string> = {
            Name: partnerFirstName || partnerName,
            RegistrantName: registrantFullName,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          };
          let mail: { subject: string; body: string };
          if (tpl) {
            mail = buildEmailFromTemplate(tpl, vars);
          } else {
            const inner = isDe
              ? `<p>Hallo ${partnerFirstName || partnerName},</p><p><strong>${registrantFullName}</strong> hat dich als <strong>Zimmerpartner</strong> für das Event <strong>${event.title}</strong> angegeben.</p>`
              : `<p>Hello ${partnerFirstName || partnerName},</p><p><strong>${registrantFullName}</strong> has selected you as their <strong>roommate</strong> for the event <strong>${event.title}</strong>.</p>`;
            mail = {
              subject: isDe ? `Zimmerpartner-Anfrage: ${event.title}` : `Roommate request: ${event.title}`,
              body: wrapTemplate('#86bc25', isDe ? 'Zimmerpartner-Anfrage' : 'Roommate request', event.title, inner),
            };
          }
          eventService.queueEmail(
            mail.subject, partnerEmail, partnerName, mail.body, 'RoommateRequest', event.title, eventId
          ).catch(err => console.warn('[DEX] roommate queueEmail failed:', err));
        }
      }
      // Outlook-Termin-Einladung in Queue eintragen
      // v15.16: Für externe Empfänger (kein @deloitte.de) keinen
      // Outlook-Termin queuen — Microsoft blockt das Versenden an externe
      // Adressen ohne Federation, deshalb ist der Eintrag immer ein
      // Bounce. Der Organizer bekommt stattdessen die Bestätigungsmail
      // mit Betreff „Weiterleitung notwendig" (s.o.) und kann darüber
      // den externen Teilnehmer informieren.
      const skipOutlookForExternal = !!emailToUse && !/@(.*\.)?deloitte\.de$/i.test(emailToUse);
      // v15.25: Schatten-Parent-Registrierung im subEventsOnlyMode bekommt
      // keinen Outlook-Termin (s.o. — der User „nimmt teil" an Sub-Events,
      // nicht am Parent).
      if (status !== 'Warteliste' && !event.disableOutlook && !skipOutlookForExternal && !suppressParentNotifications && !opts?.suppressOutlook) {
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
    // v18.67: echten Status zurueckgeben (Angemeldet/Warteliste), damit die
    // RegistrationPage das Ergebnis-Modal nicht mehr aus der gecachten
    // currentParticipants-Schaetzung (isFull) ableiten muss — die war nach
    // Cancel/Re-Register veraltet und zeigte faelschlich "Warteliste".
    return { ok: success, status };
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
    members: Array<{ email: string; displayName: string; customData?: Record<string, string> }>,
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
          // v18.12: Custom-Field-Antworten des Mitglieds (z.B. Essenspraeferenz).
          customData: m.customData || {},
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
      // v11.87: Team-Info-Block + Consent-Hinweis injecten.
      // v16.3: Vorher wurde der Block direkt nach <body> eingefuegt — damit
      // landete er VOR dem Deloitte-Template-Header (Logo, gruener Balken,
      // Headline). Stattdessen jetzt INNERHALB des Content-<td> einsetzen,
      // also direkt vor dem eigentlichen Mail-Body. Wir matchen das Content-
      // Padding (`padding:0 30px 30px 30px`) als Marker — dieser Style ist
      // im wrapTemplate eindeutig fuer das Body-Td.
      const teamInfoHtml = teamInfoBlockHtml({
        teamName,
        members: teamMembersForBlock,
        teamSize: teamSizeForBlock,
        isDe,
        registeredByName: currentUserName,
        consentRequired: true,
      });
      const bodyWithHint = injectIntoEmailContent(emailData.body, teamInfoHtml);
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
  async function assignTeamlessToTeam(
    eventId: string,
    teamId: string,
    teamName: string | undefined,
    existingRegId: number,
    isLead: boolean = false,
  ): Promise<boolean> {
    const event = subsiteMap.current[eventId] ? events.find(e => e.id === eventId) : events.find(e => e.id === eventId);
    if (!event || !event.subsiteUrl) return false;
    return eventService.assignRegistrationToTeam(event.subsiteUrl, existingRegId, teamId, teamName, isLead);
  }

  async function addTeamMember(
    eventId: string,
    teamId: string,
    teamName: string | undefined,
    member: { email: string; displayName: string },
    // v18.73: optionale event-spezifische Antworten des Beitretenden — werden
    // beim Insert mitgeschrieben (vorher immer leer). Genutzt vom Self-Join
    // über die Anmeldeseite, damit der Beitritt die Pflicht-Custom-Felder nicht
    // mehr überspringt.
    customData?: Record<string, string>
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
      customData: customData || {},
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
    const bodyWithHint = injectIntoEmailContent(emailData.body, teamInfoHtml);
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

    // v12.14: Info-Mail an verbleibende Mitglieder kommt jetzt aus
    // DEX_EmailTemplates (TemplateType=TeamMemberJoined). Pre-Wrap +
    // Variable-Substitution durch buildEmailFromTemplate.
    if (!event.disableEmails) {
      const tpl = await eventService.getEmailTemplate('TeamMemberJoined', lang).catch(() => null);
      const newMemberFullName = `${parsed.firstName} ${parsed.lastName}`.trim();
      const teamNameStr = teamName ? `„${teamName}"` : '';
      for (const other of activeMembers) {
        const otherFirst = other.Vorname || '';
        const otherFull = `${other.Vorname || ''} ${other.Nachname || ''}`.trim() || other.ParticipantEmail;
        const vars: Record<string, string> = {
          Name: otherFirst || otherFull,
          NewMemberName: newMemberFullName,
          TeamName: teamNameStr,
          EventTitle: event.title,
          AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
        };
        let mail: { subject: string; body: string };
        if (tpl) {
          mail = buildEmailFromTemplate(tpl, vars);
        } else {
          // Fallback: alter Inline-Text falls Template nicht geseeded ist.
          const inner = isDe
            ? `<p>Hallo ${otherFirst},</p><p><strong>${newMemberFullName}</strong> ist eurem Team ${teamNameStr} beigetreten.</p>`
            : `<p>Hello ${otherFirst},</p><p><strong>${newMemberFullName}</strong> joined your team ${teamNameStr}.</p>`;
          mail = {
            subject: isDe ? `Neues Team-Mitglied — ${event.title}` : `New team member — ${event.title}`,
            body: wrapTemplate('#86bc25', isDe ? 'Team-Update' : 'Team update', `Event ${event.title}`, inner),
          };
        }
        eventService.queueEmail(
          mail.subject, other.ParticipantEmail, otherFull,
          mail.body, 'TeamMemberJoined', event.title, eventId
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
    teamName: string | undefined,
    // v18.73: event-spezifische Antworten des Beitretenden durchreichen.
    customData?: Record<string, string>
  ): Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    return addTeamMember(eventId, teamId, teamName, {
      email: currentUserEmail,
      displayName: currentUserName,
    }, customData);
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

    // v12.14: Info-Mails an alle aktiven Mitglieder kommen jetzt aus
    // TemplateType=TeamLeadTransferred. {{NewLeadBlock}}-Platzhalter wird
    // pro Empfänger gefüllt — für den neuen Lead ein zusätzlicher Hinweis,
    // sonst leer.
    if (!event.disableEmails) {
      const lang = event.emailLanguage || 'EN';
      const isDe = lang.toUpperCase() === 'DE';
      const tpl = await eventService.getEmailTemplate('TeamLeadTransferred', lang).catch(() => null);
      const newLeadName = `${newLead.Vorname || ''} ${newLead.Nachname || ''}`.trim() || newLead.ParticipantEmail;
      const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
      const teamNameStr = teamName ? `„${teamName}"` : '';
      const newLeadBlockHtml = isDe
        ? `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>Du bist ab jetzt Team-Lead.</strong> Du kannst neue Mitglieder über „Meine Events" hinzufügen und ggf. Beitritts-Anfragen entscheiden.</p>`
        : `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>You are now the team lead.</strong> You can add new members via „My Events" and decide on join requests if any.</p>`;
      for (const other of active) {
        const otherFirst = other.Vorname || '';
        const otherFull = `${other.Vorname || ''} ${other.Nachname || ''}`.trim() || other.ParticipantEmail;
        const isNewLeadMember = other.Id === newLead.Id;
        const vars: Record<string, string> = {
          Name: otherFirst || otherFull,
          NewLeadName: newLeadName,
          TeamName: teamNameStr,
          EventTitle: event.title,
          NewLeadBlock: isNewLeadMember ? newLeadBlockHtml : '',
          AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
        };
        let mail: { subject: string; body: string };
        if (tpl) {
          mail = buildEmailFromTemplate(tpl, vars);
        } else {
          const inner = isDe
            ? `<p>Hallo ${otherFirst},</p><p>Die Team-Lead-Rolle in deinem Team ${teamNameStr} wurde an <strong>${newLeadName}</strong> übergeben.</p>${isNewLeadMember ? newLeadBlockHtml : ''}`
            : `<p>Hello ${otherFirst},</p><p>The team lead role in your team ${teamNameStr} has been transferred to <strong>${newLeadName}</strong>.</p>${isNewLeadMember ? newLeadBlockHtml : ''}`;
          mail = {
            subject: isDe ? `Team-Lead-Wechsel — ${event.title}` : `Team lead change — ${event.title}`,
            body: wrapTemplate('#86bc25', isDe ? 'Team-Lead-Wechsel' : 'Team lead change', `Event ${event.title}`, inner),
          };
        }
        eventService.queueEmail(
          mail.subject, other.ParticipantEmail, otherFull,
          mail.body, 'TeamLeadTransferred', event.title, eventId
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
    teamId: string,
    // v18.73: event-spezifische Antworten des Anfragenden — werden in der
    // Request-Zeile gespeichert und beim Approve auf den neuen Member angewandt.
    customData?: Record<string, string>
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
      // v18.73: Antworten als JSON mitschreiben (leer = '{}').
      customData: JSON.stringify(customData || {}),
    });
    if (!result.ok) return { ok: false, reason: 'queue-failed' };

    // v12.14: Lead-Notification kommt jetzt aus TemplateType=TeamJoinRequest.
    if (!event.disableEmails) {
      const lang = event.emailLanguage || 'EN';
      const isDe = lang.toUpperCase() === 'DE';
      const tpl = await eventService.getEmailTemplate('TeamJoinRequest', lang).catch(() => null);
      const leadFirst = lead.Vorname || '';
      const leadFull = `${lead.Vorname || ''} ${lead.Nachname || ''}`.trim() || lead.ParticipantEmail;
      const appUrl = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView&action=teamjoin&request=${result.itemId || 0}`;
      const teamNameStr = teamNameFromMembers ? `„${teamNameFromMembers}"` : '';
      const vars: Record<string, string> = {
        Name: leadFirst || leadFull,
        RequesterName: currentUserName,
        TeamName: teamNameStr,
        EventTitle: event.title,
        ApproveUrl: `${appUrl}&decision=approve`,
        RejectUrl: `${appUrl}&decision=reject`,
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      };
      let mail: { subject: string; body: string };
      if (tpl) {
        mail = buildEmailFromTemplate(tpl, vars);
      } else {
        const inner = isDe
          ? `<p>Hallo ${leadFirst},</p><p><strong>${currentUserName}</strong> möchte deinem Team ${teamNameStr} beitreten.</p><p style="text-align:center;margin:18px 0;"><a href="${appUrl}&decision=approve" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Bestätigen</a> <a href="${appUrl}&decision=reject" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Ablehnen</a></p>`
          : `<p>Hello ${leadFirst},</p><p><strong>${currentUserName}</strong> would like to join your team ${teamNameStr}.</p><p style="text-align:center;margin:18px 0;"><a href="${appUrl}&decision=approve" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Approve</a> <a href="${appUrl}&decision=reject" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Reject</a></p>`;
        mail = {
          subject: isDe ? `Team-Beitritts-Anfrage — ${event.title}` : `Team join request — ${event.title}`,
          body: wrapTemplate('#86bc25', isDe ? 'Team-Beitritts-Anfrage' : 'Team join request', `Event ${event.title}`, inner),
        };
      }
      eventService.queueEmail(
        mail.subject, lead.ParticipantEmail, leadFull,
        mail.body, 'TeamJoinRequest', event.title, eventId
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
      // v18.73: bei der Anfrage gespeicherte event-spezifische Antworten
      // wiederherstellen und auf den neuen Member anwenden.
      let reqCustomData: Record<string, string> | undefined;
      try { reqCustomData = req.CustomData ? JSON.parse(req.CustomData) : undefined; } catch { reqCustomData = undefined; }
      const addRes = await addTeamMember(req.EventId, req.TeamId, teamName || undefined, {
        email: req.RequesterEmail,
        displayName: req.RequesterDisplayName,
      }, reqCustomData);
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
    // v12.14: Absage-Mail aus TemplateType=TeamJoinRejected.
    if (!event.disableEmails) {
      const lang = event.emailLanguage || 'EN';
      const isDe = lang.toUpperCase() === 'DE';
      const tpl = await eventService.getEmailTemplate('TeamJoinRejected', lang).catch(() => null);
      const requesterFirst = req.RequesterDisplayName.split(',').pop()?.trim() || req.RequesterDisplayName;
      const vars: Record<string, string> = {
        Name: requesterFirst,
        EventTitle: event.title,
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      };
      let mail: { subject: string; body: string };
      if (tpl) {
        mail = buildEmailFromTemplate(tpl, vars);
      } else {
        const inner = isDe
          ? `<p>Hallo ${requesterFirst},</p><p>deine Beitritts-Anfrage zum Team beim Event „${event.title}" wurde vom Team-Lead abgelehnt.</p>`
          : `<p>Hello ${requesterFirst},</p><p>your join request for the team at event „${event.title}" was declined by the team lead.</p>`;
        mail = {
          subject: isDe ? `Team-Beitritts-Anfrage abgelehnt — ${event.title}` : `Team join request declined — ${event.title}`,
          body: wrapTemplate('#ed8b00', isDe ? 'Team-Beitritts-Anfrage abgelehnt' : 'Team join request declined', `Event ${event.title}`, inner),
        };
      }
      eventService.queueEmail(
        mail.subject, req.RequesterEmail, req.RequesterDisplayName,
        mail.body, 'TeamJoinRejected', event.title, req.EventId
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

  async function cancelRegistration(eventId: string, opts?: { suppressNotifications?: boolean }): Promise<boolean> {
    // v17.25: Demo-Showcase-Event → No-Op.
    if (isDemoShowcaseId(eventId)) return true;
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
        // v17.19/v17.22: Notifications werden NUR unterdrueckt, wenn der Aufrufer
        // das explizit anfordert (`opts.suppressNotifications`). Das passiert
        // ausschliesslich beim automatischen Schatten-Parent-Cancel im
        // subEventsOnlyMode (MyEventsPage: letzte Sub-Event-Abmeldung raeumt
        // den Schatten-Parent ab — die Sub-Event-Abmeldung hat ihre eigene
        // Mail schon verschickt). v17.22-Fix: vorher wurde pauschal auf
        // `event.subEventsOnlyMode` geprueft, wodurch Alt-Anmeldungen (User
        // hat sich noch im Normal-Modus direkt beim Parent angemeldet, bevor
        // der Organizer das Event auf subEventsOnlyMode umstellte) beim
        // direkten Abmelden weder Bestaetigungs-Mail noch Outlook-Ausladen
        // bekamen.
        const suppressParentNotificationsCancel = !!opts?.suppressNotifications;
        if (!event.disableEmails && !suppressParentNotificationsCancel) {
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
            // v18.41: People-Picker-Felder mit „CC bei Mail" → ausgewählte
            // Person(en) auch bei der Abmelde-Mail auf CC (nicht im Outlook-Termin).
            // v18.58: Im subEventsOnlyMode liegen die übergreifenden CC-Felder
            // (z.B. Assistenz) in der Schatten-Parent-Registrierung, NICHT in der
            // Sub-Event-Registrierung. Beim Abmelden einer Section daher
            // zusätzlich die CC aus dem Parent-Event nachladen und mergen, damit
            // die Assistenz auch die Abmelde-Mail in Kopie bekommt (Anmeldung
            // war bereits via extraCc abgedeckt).
            let cancelCc: string | undefined;
            try {
              const cd = myReg.CustomData ? JSON.parse(myReg.CustomData) as Record<string, string> : {};
              const ownCc = collectCcEmailsFromFields(event.eventSpecificFields, cd, currentUserEmail);
              let parentCc = '';
              if (event.parentEventId) {
                const parentEvent = events.find(ev => ev.id === event.parentEventId);
                if (parentEvent && parentEvent.subEventsOnlyMode) {
                  try {
                    const parentReg = await getMyRegistration(event.parentEventId);
                    const pcd = parentReg?.CustomData ? JSON.parse(parentReg.CustomData) as Record<string, string> : {};
                    parentCc = collectCcEmailsFromFields(parentEvent.eventSpecificFields, pcd, currentUserEmail);
                  } catch { /* parent-CC best-effort */ }
                }
              }
              const seen = new Set<string>();
              const merged: string[] = [];
              for (const part of [ownCc, parentCc]) {
                for (const em of part.split(';').map(s => s.trim()).filter(Boolean)) {
                  const lc = em.toLowerCase();
                  if (lc !== (currentUserEmail || '').toLowerCase() && !seen.has(lc)) { seen.add(lc); merged.push(em); }
                }
              }
              cancelCc = merged.length ? merged.join(';') : undefined;
            } catch { cancelCc = undefined; }
            const emailOk = await eventService.queueEmail(
              emailData.subject, currentUserEmail, currentUserName, emailData.body,
              'Abmeldung', event.title, eventId, cancelCc, bcc
            );
            if (!emailOk) console.warn('[DEX] queueEmail for cancellation returned false');
          } catch (err) { console.warn('[DEX] queueEmail for cancellation failed:', err); }
        }
        // Outlook-Termin-Einladung zurückziehen
        if (!event.disableOutlook && !suppressParentNotificationsCancel) {
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
              eventId, event.eventNumber || 0, subsiteUrl, event.title, currentUserName
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
          // v18.41: CC-Felder der abgemeldeten Person berücksichtigen.
          let memberCc: string | undefined;
          try {
            const cd = memberRegistration.CustomData ? JSON.parse(memberRegistration.CustomData) as Record<string, string> : {};
            memberCc = collectCcEmailsFromFields(event.eventSpecificFields, cd, memberRegistration.ParticipantEmail) || undefined;
          } catch { memberCc = undefined; }
          await eventService.queueEmail(
            emailData.subject,
            memberRegistration.ParticipantEmail,
            `${memberRegistration.Vorname || ''} ${memberRegistration.Nachname || ''}`.trim() || memberRegistration.ParticipantEmail,
            emailData.body,
            'Abmeldung', event.title, eventId, memberCc, bcc
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
            eventId, event.eventNumber || 0, subsiteUrl, event.title,
            `${memberRegistration.Vorname || ''} ${memberRegistration.Nachname || ''}`.trim() || memberRegistration.ParticipantName || undefined
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

    // v12.14: Info-Mails kommen aus TemplateType=TeamMemberCancelled.
    // {{NewLeadBlock}}-Platzhalter wird für den Auto-Promote-Empfänger
    // gefüllt, für alle anderen leer.
    if (event.disableEmails) return;
    const lang = event.emailLanguage || 'EN';
    const isDe = lang.toUpperCase() === 'DE';
    const tpl = await eventService.getEmailTemplate('TeamMemberCancelled', lang).catch(() => null);
    const teamSizeCfg = event.teamSize || (remaining.length + 1);
    const cancelledFullName = `${cancelledReg.Vorname || ''} ${cancelledReg.Nachname || ''}`.trim() || cancelledReg.ParticipantEmail;
    const teamNameStr = teamName ? `„${teamName}"` : '';
    const newLeadBlockHtml = isDe
      ? `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>Du bist jetzt der neue Team-Lead.</strong> Du kannst über „Meine Events" eine neue Person hinzufügen, falls der frei gewordene Platz wieder gefüllt werden soll.</p>`
      : `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>You are the new team lead now.</strong> You can add a replacement member via „My Events" if you want to fill the freed slot.</p>`;

    for (const m of remaining) {
      const mFirst = m.Vorname || (m.ParticipantName || '').split(/[ ,]+/)[0] || '';
      const mFull = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail;
      const isNewLead = newLeadId !== null && m.Id === newLeadId;
      const vars: Record<string, string> = {
        Name: mFirst || mFull,
        CancelledName: cancelledFullName,
        TeamName: teamNameStr,
        EventTitle: event.title,
        ActiveCount: String(remaining.length),
        TeamSize: String(teamSizeCfg),
        NewLeadBlock: isNewLead ? newLeadBlockHtml : '',
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      };
      let mail: { subject: string; body: string };
      if (tpl) {
        mail = buildEmailFromTemplate(tpl, vars);
      } else {
        const inner = isDe
          ? `<p>Hallo ${mFirst},</p><p>${cancelledFullName} hat sich vom Team ${teamNameStr} abgemeldet (${remaining.length}/${teamSizeCfg}).</p>${isNewLead ? newLeadBlockHtml : ''}`
          : `<p>Hello ${mFirst},</p><p>${cancelledFullName} cancelled their registration from team ${teamNameStr} (${remaining.length}/${teamSizeCfg}).</p>${isNewLead ? newLeadBlockHtml : ''}`;
        mail = {
          subject: isDe ? `Team-Update — ${event.title}` : `Team update — ${event.title}`,
          body: wrapTemplate('#ed8b00', isDe ? 'Team-Update' : 'Team update', `Event ${event.title}`, inner),
        };
      }
      try {
        await eventService.queueEmail(
          mail.subject,
          m.ParticipantEmail,
          mFull,
          mail.body,
          'TeamMemberCancelled',
          event.title,
          eventId
        );
      } catch (err) {
        console.warn('[DEX] team-cancel info mail failed:', err);
      }
    }
  }

  // v18.11: Proaktive Absage durch den eingeloggten User („Ich nehme nicht
  // teil"). Wenn schon eine aktive/Warteliste-Anmeldung existiert, wird sie
  // regulär abgemeldet (Seat-Sync, Mail, IDReorder laufen mit). Sonst wird
  // eine reine Absage-Zeile (Status=Abgemeldet, Marker _declined) angelegt.
  async function declineEvent(eventId: string): Promise<boolean> {
    if (isDemoShowcaseId(eventId)) return true;
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    // Name aus displayName ableiten („Nachname, Vorname" oder „Vorname Nachname").
    const dn = (currentUserName || '').trim();
    let firstName = ''; let lastName = '';
    if (dn.indexOf(',') >= 0) {
      const p = dn.split(',').map(s => s.trim());
      lastName = p[0] || ''; firstName = p[1] || '';
    } else {
      const p = dn.split(/\s+/).filter(Boolean);
      firstName = p[0] || ''; lastName = p.slice(1).join(' ');
    }
    const existing = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (existing) {
      // Bereits abgemeldet/abgesagt → nichts zu tun. Aktiv/Warteliste →
      // regulärer Cancel-Pfad (gibt Sitzplatz frei, Mail, IDReorder).
      if (existing.Status === 'Abgemeldet') return true;
      return await cancelRegistration(eventId);
    }
    const ok = await eventService.declineRegistration(
      subsiteUrl, firstName, lastName, currentUserEmail, currentUserName, currentUserEmail
    );
    if (ok) {
      const event = events.find(e => e.id === eventId);
      eventService.writeChangeLog({
        action: 'ParticipantDeclined',
        targetType: 'Participant',
        targetId: currentUserEmail,
        targetName: `${firstName} ${lastName}`.trim() || currentUserEmail,
        eventId,
        eventTitle: event?.title || '',
        details: { asActor: 'self', proactiveDecline: true },
      }).catch(() => { /* */ });
    }
    return ok;
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

  // v18.33: Self-Check-in über gescannten QR-Deep-Link.
  async function selfCheckIn(params: SelfCheckInParams): Promise<SelfCheckInResult> {
    try {
      // 1) Event auflösen — per Token (statischer QR) ODER Event-Nummer (rotierend).
      let sp: SPEvent | null = null;
      if (params.token) {
        sp = await eventService.getEventBySelfCheckInToken(params.token);
      } else if (typeof params.eventNumber === 'number' && !isNaN(params.eventNumber)) {
        sp = await eventService.getEventByEventNumber(params.eventNumber);
      }
      if (!sp) return { status: 'not-found' };

      const eventTitle = sp.Title;
      const eventStart = sp.StartDate;

      // 2) Self-Check-in muss aktiviert sein.
      if (!sp.SelfCheckInEnabled) {
        return { status: 'disabled', eventTitle, eventStart };
      }

      // 3) Rotierender Code: Frische + HMAC prüfen.
      if (!params.token) {
        const secret = sp.SelfCheckInToken || '';
        const ok = await verifyRotatingCode(
          secret,
          params.code || '',
          typeof params.windowIndex === 'number' ? params.windowIndex : NaN
        );
        if (!ok) return { status: 'expired', eventTitle, eventStart };
      }

      // 4) Zeitfenster prüfen (from/to ODER Default „Event-Tag").
      const win = isWithinCheckInWindow(
        sp.StartDate,
        sp.EndDate,
        sp.SelfCheckInFrom,
        sp.SelfCheckInTo
      );
      if (!win.ok) {
        return {
          status: 'closed',
          eventTitle,
          eventStart,
          opensAt: win.opensAt ? win.opensAt.toISOString() : undefined,
          closesAt: win.closesAt ? win.closesAt.toISOString() : undefined,
        };
      }

      // 5) Eigene Registrierung auf der Subsite finden.
      const subsiteUrl = sp.SubsiteUrl;
      if (!subsiteUrl) return { status: 'error', eventTitle, eventStart };
      const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
      if (!myReg) return { status: 'not-registered', eventTitle, eventStart };
      if (myReg.Status === 'Eingecheckt') return { status: 'already', eventTitle, eventStart };
      if (myReg.Status === 'Warteliste') return { status: 'on-waitlist', eventTitle, eventStart };
      if (myReg.Status === 'Abgemeldet') return { status: 'not-registered', eventTitle, eventStart };

      // 6) Einchecken (eigener Eintrag — Item-Level-Security erlaubt das Schreiben).
      const ok = await eventService.checkInParticipant(subsiteUrl, myReg.Id);
      return { status: ok ? 'success' : 'error', eventTitle, eventStart };
    } catch (e) {
      console.warn('[DEX] selfCheckIn error:', e);
      return { status: 'error' };
    }
  }

  async function getAllRegistrations(eventId: string): Promise<SPRegistration[]> {
    // v18: Demo-Event → synthetische Teilnehmerliste (~25 Demo-User inkl.
    // Team, Warteliste, Abmeldungen), damit der Admin die Teilnehmer-
    // Verwaltung im Demo-Modus durchspielen kann.
    if (isDemoShowcaseId(eventId)) return buildDemoRegistrations();
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
    // v18.3: Demo-Showcase-Event → No-Op (kein SP-Backend). Defense in depth;
    // die UI blendet den Löschen-Button fuer das Demo-Event ohnehin aus.
    if (isDemoShowcaseId(eventId)) return false;
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

  // v19.0: Dokument-Custom-Felder. Ein Attachment wird über einen Dateinamen-
  // Präfix (`dxf-<fieldId>--`) genau EINEM Dokument-Feld zugeordnet, sodass ein
  // Event mehrere Dokument-Felder haben kann. participantEmail erlaubt den
  // Upload für eine andere Person (stellvertretende Anmeldung); Default = der
  // eingeloggte User (Self-Anmeldung + „Meine Events").
  const docFieldPrefix = (fieldId: string): string => `dxf-${(fieldId || '').replace(/[^a-zA-Z0-9]/g, '')}--`;
  const stripDocPrefix = (fileName: string): string =>
    fileName
      .replace(/^dxf-[a-zA-Z0-9]+--\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '')
      .replace(/^dxf-[a-zA-Z0-9]+--/, '');
  async function uploadFieldDocument(eventId: string, fieldId: string, file: File, participantEmail?: string): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const email = (participantEmail || currentUserEmail || '').trim();
    const reg = await eventService.getMyRegistration(subsiteUrl, email);
    if (!reg || !reg.Id) return false;
    return eventService.addRegistrationAttachment(subsiteUrl, reg.Id, file, docFieldPrefix(fieldId));
  }
  async function listFieldDocuments(eventId: string, fieldId: string, participantEmail?: string): Promise<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return [];
    const email = (participantEmail || currentUserEmail || '').trim();
    const reg = await eventService.getMyRegistration(subsiteUrl, email);
    if (!reg || !reg.Id) return [];
    const all = await eventService.listRegistrationAttachments(subsiteUrl, reg.Id);
    const prefix = docFieldPrefix(fieldId);
    return all.filter(f => f.fileName.startsWith(prefix)).map(f => ({ ...f, displayName: stripDocPrefix(f.fileName) }));
  }
  async function deleteFieldDocument(eventId: string, fileName: string, participantEmail?: string): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    const email = (participantEmail || currentUserEmail || '').trim();
    const reg = await eventService.getMyRegistration(subsiteUrl, email);
    if (!reg || !reg.Id) return false;
    return eventService.deleteRegistrationAttachment(subsiteUrl, reg.Id, fileName);
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
          // v13.0: Mail aus DB-Template — Confirmed bzw. Waitlist-Variante.
          const isWaitlist = result.status === 'Warteliste';
          const templateType = isWaitlist ? 'GroupSwitchWaitlist' : 'GroupSwitchConfirmed';
          const tpl = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
          const firstName = currentUserName.split(',').pop()?.trim().split(' ')[0] || currentUserName.split(' ')[0];
          const vars: Record<string, string> = {
            Name: firstName,
            GroupLabel: newLabel,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          };
          let mail: { subject: string; body: string };
          if (tpl) {
            mail = buildEmailFromTemplate(tpl, vars);
          } else {
            const innerBody = isDeMail
              ? (isWaitlist
                ? `<p>Du hast den Wechsel in die Gruppe <strong>${newLabel}</strong> für <strong>${event.title}</strong> angefragt. Diese Gruppe ist aktuell voll, daher steht deine Anmeldung auf der <strong>Warteliste der Gruppe ${newLabel}</strong>.</p>`
                : `<p>Dein Gruppen-Wechsel zu <strong>${newLabel}</strong> für <strong>${event.title}</strong> ist bestätigt.</p>`)
              : (isWaitlist
                ? `<p>You requested to switch to the <strong>${newLabel}</strong> group for <strong>${event.title}</strong>. The group is currently full, so your registration is on the <strong>${newLabel} waitlist</strong>.</p>`
                : `<p>Your group switch to <strong>${newLabel}</strong> for <strong>${event.title}</strong> is confirmed.</p>`);
            mail = {
              subject: isDeMail
                ? (isWaitlist ? `Gruppen-Wechsel — auf Warteliste: ${event.title}` : `Gruppen-Wechsel bestätigt: ${event.title}`)
                : (isWaitlist ? `Group switch — added to waitlist: ${event.title}` : `Group switch confirmed: ${event.title}`),
              body: wrapTemplate(isWaitlist ? '#ed8b00' : '#86bc25', isDeMail ? 'Gruppen-Wechsel' : 'Group switch', event.title, innerBody),
            };
          }
          await eventService.queueEmail(mail.subject, currentUserEmail, currentUserName, mail.body, templateType, event.title, eventId)
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
  // v12.12: Re-Seed-Aktion durchreichen.
  async function reseedDefaultEmailTemplates(): Promise<ReseedSummary> {
    return eventService.reseedDefaultEmailTemplates();
  }

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
    // v18.30: Anfrage-Mail mit hoher Wichtigkeit (rotes „!" in Outlook) —
    // der DEX_SEND_MAIL-Flow liest das Importance-Feld aus der Queue aus.
    return eventService.queueEmail(
      subject, adminTo, 'DEX Admin Team', body, 'Info', eventName || 'DEX-Anfrage', '0', requesterEmail, undefined, 'High'
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
  // v13.11: topLevelEvents wird unten direkt aus `eventsForConsumer`
  // gefiltert, weil die Demo-Impersonation pro Event qrScannerEmails
  // shadowen muss.
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

  // v13.11: Demo-Impersonation kann den synthetischen Demo-User per
  // qrScannerEmails einem konkreten Event als Check-In-Helfer
  // zuordnen. Wir injizieren die Demo-Mail in das gewählte Event,
  // damit Header/StartPage/AdminPage die übliche Permission-Logik
  // unverändert nutzen können.
  const eventsForConsumer = React.useMemo(() => {
    try {
      if (typeof window === 'undefined') return events;
      const raw = window.localStorage?.getItem('dex_demo_impersonation');
      if (!raw) return events;
      const payload = JSON.parse(raw);
      // v17.25: Im Demo-Impersonation-Modus das synthetische Showcase-Event
      // (+ Sub-Event) vorne in die Liste haengen, damit der Admin auf der
      // Register-Seite alle Event-Faehigkeiten durchspielen kann. Existiert
      // nur client-seitig — kein SP-Roundtrip. Sprache aus dem persistierten
      // Locale-Key, Fallback DE.
      let withDemo = events;
      try {
        const storedLocale = window.localStorage?.getItem('dex-locale');
        const demoLocale: 'de' | 'en' = storedLocale === 'en' ? 'en' : 'de';
        if (!events.some(e => e.isDemoShowcase)) {
          withDemo = [...buildDemoShowcaseEvents(demoLocale), ...events];
        }
      } catch { /* */ }
      const targetId: string = payload?.checkInEventId || '';
      const demoEmail: string = (payload?.email || '').toLowerCase();
      if (!targetId || !demoEmail) return withDemo;
      return withDemo.map(e => {
        if (e.id !== targetId) return e;
        const current = (e.qrScannerEmails || []).map(x => x.toLowerCase());
        if (current.indexOf(demoEmail) >= 0) return e;
        return { ...e, qrScannerEmails: [...(e.qrScannerEmails || []), demoEmail] };
      });
    } catch { return events; }
  }, [events]);

  return React.createElement(
    EventContext.Provider,
    {
      value: {
        events: eventsForConsumer,
        topLevelEvents: eventsForConsumer.filter(e => !e.parentEventId),
        childEventsOf, isEventsLoading,
        createEvent, registerForEvent, registerTeam,
        getTeamMembers: async (eventId: string, teamId: string): Promise<SPRegistration[]> => {
          const subsiteUrl = subsiteMap.current[eventId];
          if (!subsiteUrl) return [];
          return eventService.getTeamMembers(subsiteUrl, teamId);
        },
        addTeamMember,
        assignTeamlessToTeam,
        joinTeam,
        transferTeamLead,
        createTeamJoinRequest,
        listTeamJoinRequestsForEvent,
        decideTeamJoinRequest,
        listOpenTeamsForEvent,
        cancelRegistration,
        declineEvent,
        cancelTeamMember,
        getMyRegistration, selfCheckIn, checkRegistrationByEmail, getAllRegistrations, deleteEvent, deleteEventItemOnly, updateEvent, updateMyRegistration, switchSplitGroup, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, uploadFieldDocument, listFieldDocuments, deleteFieldDocument, getMyEventNumbers, refreshEvents, refreshParticipantCounts, markExpiredEventsAsCompleted,
        sendAdminInquiry,
        reseedDefaultEmailTemplates,
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
