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
import { registrationEmail, waitlistEmail, cancellationEmail, promotionEmail, buildEmailFromTemplate, loadLogosAsBase64, wrapTemplate } from '../services/EmailTemplates';

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
  getMyEventNumbers: () => Promise<{ registered: number[]; waitlisted: number[] }>;
  refreshEvents: () => Promise<void>;
  refreshParticipantCounts: (eventId?: string) => Promise<void>;
  markExpiredEventsAsCompleted: () => Promise<number>;
  sendAdminInquiry: (requesterName: string, requesterEmail: string, eventName: string, message: string) => Promise<boolean>;
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
  isFictive?: boolean;
  durchstarterCapacity?: number;
  funstarterCapacity?: number;
  customFields: CustomField[];
}

const EventContext = React.createContext<EventContextType | undefined>(undefined);

export function EventProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [events, setEvents] = React.useState<DeloitteEvent[]>([]);
  const [isEventsLoading, setIsEventsLoading] = React.useState(true);
  // Map von EventId -> SubsiteUrl fuer schnellen Zugriff
  const subsiteMap = React.useRef<Record<string, string>>({});

  const eventService = React.useMemo(() => new EventService(props.context), []);
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
    try { await eventService.ensureEmailsList(); } catch { /* */ }
    try { await eventService.ensureOutlookList(); } catch { /* */ }
    try { await eventService.ensureParticipantsList(); } catch { /* */ }
    try { await eventService.ensureEmailTemplatesList(); } catch { /* */ }
    try { await eventService.ensureIDReorderList(); } catch { /* */ }
    try { await eventService.ensureAssetsFolders(); } catch { /* */ }
    try { await eventService.ensureLogosInConfig(); } catch { /* */ }
    try { await loadLogosAsBase64(props.context.spHttpClient, eventService.siteUrl); } catch { /* */ }
    // Seed-Migrationen entfernt - erfolgreich abgeschlossen
    await loadEvents();
    setIsEventsLoading(false);
  }

  async function loadEvents(): Promise<void> {
    const spEvents = await eventService.getEvents();
    const mapped = await Promise.all(spEvents.map(e => mapSPEventToDeloitteEvent(e)));
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

    return {
      id: e.Id.toString(),
      eventNumber: e.EventNumber || 0,
      title: e.Title || '',
      // v5.2: EventType-Spalte deprecated. Typ aus CustomFields ableiten
      // (Fallback auf alten SP-Wert wenn noch vorhanden).
      type: (e.EventType as DeloitteEvent['type'])
        || (customFields.some(f => f.id === 'b2run_startblock') ? 'B2Run' : 'Other'),
      status: (e.EventStatus as DeloitteEvent['status']) || 'Under Construction',
      organizers: (e.Organizer || '').split(';').map((s: string) => s.trim()).filter((s: string) => s),
      organizerEmails: (e.OrganizerEmail || '').split(';').map((s: string) => s.trim()).filter((s: string) => s),
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
      currentParticipants,
      waitlistCount,
      imageUrl: e.EventImageUrl || '',
      subsiteUrl: e.SubsiteUrl || '',
      outlookBody: e.OutlookBody || '',
      emailLanguage: e.EmailLanguage || 'EN',
      emailTemplateOverrides: e.EmailTemplateOverrides || '',
      disableEmails: !!e.DisableEmails,
      disableOutlook: !!e.DisableOutlook,
      isFictive: !!e.IsFictive,
      durchstarterCapacity: typeof e.DurchstarterCapacity === 'number' ? e.DurchstarterCapacity : undefined,
      funstarterCapacity: typeof e.FunstarterCapacity === 'number' ? e.FunstarterCapacity : undefined,
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
        helpText: '',
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        spInternalName: (cf as any).spInternalName || '',
      })),
    };
  }

  async function createEvent(input: CreateEventInput): Promise<number | null> {
    const eventId = await eventService.createEvent(input);
    if (eventId) {
      // Events neu laden OHNE Participant Counts (neue Subsite ist noch nicht bereit)
      try {
        const spEvents = await eventService.getEvents();
        const mapped = await Promise.all(spEvents.map(e => mapSPEventToDeloitteEvent(e)));
        setEvents(mapped);
      } catch { /* Events-Refresh fehlgeschlagen, nicht kritisch */ }
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

    // Vorname/Nachname aus displayName extrahieren falls nicht uebergeben
    const nameParts = currentUserName.split(' ');
    const firstNameToUse = participantFirstName || nameParts[0] || '';
    const lastNameToUse = participantLastName || nameParts.slice(1).join(' ') || '';
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
    const isB2runSplit = event && typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    if (event && isB2runSplit && preferredStarterType) {
      try {
        const allRegs = await eventService.getAllRegistrations(subsiteUrl);
        const activeRegs = allRegs.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        const durchCount = activeRegs.filter(r => r.StarterType === 'Durchstarter').length;
        const funCount = activeRegs.filter(r => r.StarterType === 'Funstarter').length;
        const durchFree = (event.durchstarterCapacity || 0) - durchCount;
        const funFree = (event.funstarterCapacity || 0) - funCount;
        if (preferredStarterType === 'Durchstarter' && durchFree > 0) {
          effectiveStarterType = 'Durchstarter';
        } else if (preferredStarterType === 'Funstarter' && funFree > 0) {
          effectiveStarterType = 'Funstarter';
        } else {
          // Wunsch-Typ voll → Warteliste für genau diesen Typ.
          status = 'Warteliste';
          effectiveStarterType = undefined; // wird erst beim Nachrücken gesetzt
        }
      } catch { /* Bei Fehler: normale Logik */ }
    } else if (event && event.maxParticipants > 0 && event.currentParticipants >= event.maxParticipants) {
      status = 'Warteliste';
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
        eventService.queueEmail(
          emailData.subject, emailToUse, nameToUse, emailData.body,
          templateType, event.title, eventId
        ).catch(err => console.warn('[DEX] queueEmail failed:', err));
      }
      // Roommate-Benachrichtigung: alle Custom-Fields vom Typ 'user' durchsuchen,
      // deren Wert eine E-Mail enthaelt (Format "Name <email>"). Fuer jede solche
      // Adresse eine Info-Mail im Deloitte-Template queuen.
      if (!event.disableEmails) {
        for (const f of event.eventSpecificFields) {
          if (f.type !== 'user') continue;
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

    // Typ des Abgemeldeten merken (fuer Nachruecker bei B2Run Split-Capacity)
    const cancelledStarterType = myReg.StarterType || '';

    // Audit: wer klickt gerade 'Abmelden'? = der eingeloggte User. Bei Self-Cancel
    // ist das = der Teilnehmer selbst. Aus der App heraus gibt's aktuell keinen
    // "Abmeldung fuer andere"-Pfad (das macht der Organizer/Admin ueber AdminPage,
    // dort wird eventService.cancelRegistration direkt aufgerufen).
    const success = await eventService.cancelRegistration(subsiteUrl, myReg.Id, currentUserName, currentUserEmail);
    if (success) {
      const event = events.find(e => e.id === eventId);
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
            const emailOk = await eventService.queueEmail(
              emailData.subject, currentUserEmail, currentUserName, emailData.body,
              'Abmeldung', event.title, eventId
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
        // Nachrücken: Ersten Warteliste-Teilnehmer auf den freigewordenen Platz setzen.
        // Bei B2Run-Split-Kapazitäten (getrennte Durchstarter-/Funstarter-Wartelisten,
        // seit v6.5): nur den ersten Warteliste-Teilnehmer nachrücken, dessen
        // PreferredStarterType mit dem freigewordenen Typ übereinstimmt. Damit wird
        // ein abgemeldeter Durchstarter-Platz nicht an einen Funstarter-Warteliste-
        // Teilnehmer vergeben und umgekehrt.
        const isB2runSplitCancel = typeof event.durchstarterCapacity === 'number'
          && typeof event.funstarterCapacity === 'number'
          && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
        try {
          const promoted = await eventService.promoteFirstWaitlistItem(
            subsiteUrl,
            cancelledStarterType || undefined,
            event.maxParticipants,
            (isB2runSplitCancel && cancelledStarterType) ? cancelledStarterType : undefined
          );
          if (promoted && promoted.success && !event.disableEmails) {
            // Nachrueck-E-Mail an den Nachruecker senden
            try {
              const lang = event.emailLanguage || 'EN';
              // {{Name}}: nur Vorname. promoted.name ist "Vorname Nachname" -> erstes Token.
              const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
              const promoteVars = {
                Name: promotedFirstName,
                EventTitle: event.title,
                Organizer: formatOrganizerList(event.organizers, lang),
                AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                WaitlistPosition: '',
              };
              let emailData: { subject: string; body: string };
              // WICHTIG: TemplateType/EmailType-Choice ist ASCII 'Nachruecken'
              // (sowohl im SharePoint-Choice-Feld als auch in DEX_EmailTemplates).
              // Die Variante 'Nachrücken' mit Umlaut existiert nicht in der Liste
              // und wuerde das Queuen der Nachrueck-E-Mail fehlschlagen lassen.
              const spTplRaw = await eventService.getEmailTemplate('Nachruecken', lang).catch(() => null);
              const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, 'Nachruecken');
              if (spTpl) {
                emailData = buildEmailFromTemplate(spTpl, promoteVars);
              } else {
                emailData = promotionEmail(promotedFirstName, event.title);
              }
              if (promoted.email) {
                await eventService.queueEmail(
                  emailData.subject, promoted.email, promoted.name || '', emailData.body,
                  'Nachruecken', event.title, eventId
                );
                // Outlook-Einladung fuer den Nachruecker
                if (!event.disableOutlook) {
                  await eventService.queueOutlookEvent(
                    promoted.email, eventId, event.title, 'Einladen'
                  );
                }
              }
            } catch (err) { console.warn('[DEX] promote email failed:', err); }
          }
        } catch (err) { console.warn('[DEX] promoteFirstWaitlistItem failed:', err); }
        // ID-Reorder in Queue eintragen (nur fuer ID-Neuvergabe, nicht fuer Nachruecken)
        if (subsiteUrl) {
          try {
            const reorderOk = await eventService.queueIDReorder(
              eventId, event.eventNumber || 0, subsiteUrl, event.title
            );
            if (!reorderOk) console.warn('[DEX] queueIDReorder returned false');
          } catch (err) { console.warn('[DEX] queueIDReorder failed:', err); }
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
    if (success) await loadEvents();
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
        getMyRegistration, checkRegistrationByEmail, getAllRegistrations, deleteEvent, updateEvent, updateMyRegistration, getMyEventNumbers, refreshEvents, refreshParticipantCounts, markExpiredEventsAsCompleted,
        sendAdminInquiry,
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
