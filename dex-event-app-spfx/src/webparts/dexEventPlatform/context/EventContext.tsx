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
import { registrationEmail, waitlistEmail, cancellationEmail, buildEmailFromTemplate, loadLogosAsBase64, wrapTemplate, organizerOnboardingEmail, qrCodeEmail } from '../services/EmailTemplates';
import * as QRCode from 'qrcode';

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
  cancelRegistration: (eventId: string) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  checkRegistrationByEmail: (eventId: string, email: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
  deleteEvent: (eventId: string) => Promise<boolean>;
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
  customFields: CustomField[];
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
    // ensure-Aufrufe koennen fehlschlagen wenn User nur Read-Rechte hat
    try { await eventService.ensureEventsList(); } catch { /* */ }
    // v4.2.0 Migration: Audience-Feld von Text auf Note umstellen (falls noch Text).
    // Idempotent: skipt wenn Feld bereits Note ist. Nur Admin hat Write-Rechte,
    // fuer normale User faellt der Call still durch.
    try { await eventService.upgradeAudienceFieldToNote(); } catch { /* */ }
    // Migration: Organizer + OrganizerEmail von Text auf Note umstellen. Bei Events
    // mit 10+ Co-Organizern lief der 255-Zeichen-Cutoff voll und Saves brachen mit
    // HTTP 500 „Invalid text value" ab. Idempotent: skipt wenn Felder schon Note sind.
    try { await eventService.upgradeOrganizerFieldsToNote(); } catch { /* */ }
    try { await eventService.ensureEmailsList(); } catch { /* */ }
    try { await eventService.ensureOutlookList(); } catch { /* */ }
    try { await eventService.ensureParticipantsList(); } catch { /* */ }
    try { await eventService.ensureEmailTemplatesList(); } catch { /* */ }
    try { await eventService.ensureIDReorderList(); } catch { /* */ }
    // v9.0: Audit-Log-Liste fuer Event- und Teilnehmer-Aenderungen
    try { await eventService.ensureChangeLogList(); } catch { /* */ }
    try { await eventService.ensureAssetsFolders(); } catch { /* */ }
    try { await eventService.ensureLogosInConfig(); } catch { /* */ }
    try { await loadLogosAsBase64(props.context.spHttpClient, eventService.siteUrl); } catch { /* */ }
    // v9.21: Test-Team ist jetzt per-Event (kommt aus event.testTeamEmails),
    // kein globaler Refresh mehr noetig.
    // Seed-Migrationen entfernt - erfolgreich abgeschlossen
    await loadEvents();
    setIsEventsLoading(false);
  }

  async function loadEvents(): Promise<void> {
    const spEvents = await eventService.getEvents();
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
    // Teilnehmerzahlen fuer alle Events mit Subsite laden
    const withCounts = await loadParticipantCountsForEvents(mapped);
    // Attachments (Dokumente) fuer alle Events laden
    const withDocs = await Promise.all(withCounts.map(async (evt) => {
      try {
        const attachments = await eventService.getEventAttachments(Number(evt.id));
        return { ...evt, documents: attachments };
      } catch { return evt; }
    }));
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
      emailLanguage: e.EmailLanguage || 'EN',
      emailTemplateOverrides: e.EmailTemplateOverrides || '',
      disableEmails: !!e.DisableEmails,
      disableOutlook: !!e.DisableOutlook,
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
      await loadEvents();
    }
    return success;
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
    const success = await eventService.cancelRegistration(subsiteUrl, myReg.Id, currentUserName, currentUserEmail);
    if (success) {
      const event = events.find(e => e.id === eventId);
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
      await loadEvents();
    }
    return success;
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
    const success = await eventService.deleteEvent(Number(eventId));
    delete subsiteMap.current[eventId];
    // Events immer neu laden, auch wenn Subsite-Loeschung fehlschlug
    await loadEvents();
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
        createEvent, registerForEvent, cancelRegistration,
        getMyRegistration, checkRegistrationByEmail, getAllRegistrations, deleteEvent, updateEvent, updateMyRegistration, switchSplitGroup, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, getMyEventNumbers, refreshEvents, refreshParticipantCounts, markExpiredEventsAsCompleted,
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
