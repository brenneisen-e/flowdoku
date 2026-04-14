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
import { registrationEmail, waitlistEmail, cancellationEmail, promotionEmail, buildEmailFromTemplate, loadLogosAsBase64 } from '../services/EmailTemplates';

interface EventContextType {
  events: DeloitteEvent[];
  isEventsLoading: boolean;
  createEvent: (event: CreateEventInput) => Promise<number | null>;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string) => Promise<boolean>;
  cancelRegistration: (eventId: string) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
  deleteEvent: (eventId: string) => Promise<boolean>;
  updateEvent: (eventId: string, updates: Record<string, unknown>) => Promise<boolean>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  getMyEventNumbers: () => Promise<{ registered: number[]; waitlisted: number[] }>;
  refreshEvents: () => Promise<void>;
  refreshParticipantCounts: (eventId?: string) => Promise<void>;
  markExpiredEventsAsCompleted: () => Promise<number>;
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
  emailLanguage?: string;
  emailTemplateOverrides?: string;
  disableEmails?: boolean;
  disableOutlook?: boolean;
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

  React.useEffect(() => {
    initEvents().catch(() => setIsEventsLoading(false));
  }, []);

  async function initEvents(): Promise<void> {
    // ensure-Aufrufe koennen fehlschlagen wenn User nur Read-Rechte hat
    try { await eventService.ensureEventsList(); } catch { /* */ }
    try { await eventService.ensureEmailsList(); } catch { /* */ }
    try { await eventService.ensureOutlookList(); } catch { /* */ }
    try { await eventService.ensureParticipantsList(); } catch { /* */ }
    try { await eventService.ensureEmailTemplatesList(); } catch { /* */ }
    try { await eventService.ensureIDReorderList(); } catch { /* */ }
    try { await eventService.ensureAssetsFolders(); } catch { /* */ }
    try { await eventService.ensureLogosInConfig(); } catch { /* */ }
    try { await loadLogosAsBase64(props.context.spHttpClient, eventService.siteUrl); } catch { /* */ }
    // Seed: Events einmalig anlegen + Organizer in DEX_Roles eintragen (idempotent)
    try { await eventService.seedNewEvents(); } catch { /* */ }
    // Seed: Teilnehmer aus den 5 CSV-Migrationen silent einfuegen (idempotent)
    try { await eventService.seedNewEventsParticipants(); } catch { /* */ }
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
      type: (e.EventType as DeloitteEvent['type']) || 'Other',
      status: (e.EventStatus as DeloitteEvent['status']) || 'Under Construction',
      organizers: (e.Organizer || '').split(';').map((s: string) => s.trim()).filter((s: string) => s),
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
      durchstarterCapacity: typeof e.DurchstarterCapacity === 'number' ? e.DurchstarterCapacity : undefined,
      funstarterCapacity: typeof e.FunstarterCapacity === 'number' ? e.FunstarterCapacity : undefined,
      agenda: (() => { try { return e.Agenda ? JSON.parse(e.Agenda) : []; } catch { return []; } })(),
      transferTimes: (() => { try { return e.Transfers ? JSON.parse(e.Transfers) : []; } catch { return []; } })(),
      quiz: (() => { try { return e.FunZone ? JSON.parse(e.FunZone) : []; } catch { return []; } })(),
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

    // B2Run Split-Capacity Logik
    const isB2runSplit = event && typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    if (event && isB2runSplit && preferredStarterType) {
      try {
        // Aktuelle Auslastung fuer beide Typen zaehlen
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
        } else if (preferredStarterType === 'Durchstarter' && funFree > 0) {
          // Durchstarter voll, Fallback auf Funstarter
          effectiveStarterType = 'Funstarter';
        } else if (preferredStarterType === 'Funstarter' && durchFree > 0) {
          // Funstarter voll, Fallback auf Durchstarter
          effectiveStarterType = 'Durchstarter';
        } else {
          // Beide voll -> Warteliste
          status = 'Warteliste';
          effectiveStarterType = undefined; // wird erst beim Nachruecken gesetzt
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

    let success: boolean;
    if (existing && existing.Status === 'Abgemeldet') {
      success = await eventService.reactivateRegistration(subsiteUrl, existing.Id, firstNameToUse, lastNameToUse, customData, status, fieldMap);
    } else {
      success = await eventService.registerForEvent(
        subsiteUrl, firstNameToUse, lastNameToUse, emailToUse, customData, status, fieldMap,
        effectiveStarterType, preferredStarterType
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
      const vars = { Name: nameToUse, EventTitle: event.title, Organizer: event.organizers.join(', '), AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, WaitlistPosition: posText };
      let emailData: { subject: string; body: string };
      const spTemplate = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
      if (spTemplate) {
        emailData = buildEmailFromTemplate(spTemplate, vars);
      } else {
        emailData = status === 'Warteliste'
          ? waitlistEmail(nameToUse, event.title, waitlistPosition)
          : registrationEmail(nameToUse, event.title);
      }
      if (!event.disableEmails) {
        eventService.queueEmail(
          emailData.subject, emailToUse, nameToUse, emailData.body,
          templateType, event.title, eventId
        ).catch(err => console.warn('[DEX] queueEmail failed:', err));
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

    const success = await eventService.cancelRegistration(subsiteUrl, myReg.Id);
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
            const cancelVars = { Name: currentUserName, EventTitle: event.title, AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView` };
            let emailData: { subject: string; body: string };
            const spTpl = await eventService.getEmailTemplate('Abmeldung', lang).catch(() => null);
            if (spTpl) {
              emailData = buildEmailFromTemplate(spTpl, cancelVars);
            } else {
              emailData = cancellationEmail(currentUserName, event.title);
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
        // Nachruecken: Ersten Warteliste-Teilnehmer auf den freigewordenen Platz setzen.
        // Der Nachruecker erbt den StarterType des Abgemeldeten (bei B2Run-Split).
        // Wird hier direkt client-seitig gemacht, damit der Power Automate Flow
        // keinen doppelten Nachrueck-Versuch macht.
        try {
          const promoted = await eventService.promoteFirstWaitlistItem(
            subsiteUrl,
            cancelledStarterType || undefined
          );
          if (promoted && promoted.success && !event.disableEmails) {
            // Nachrueck-E-Mail an den Nachruecker senden
            try {
              const lang = event.emailLanguage || 'EN';
              const promoteVars = {
                Name: promoted.name || '',
                EventTitle: event.title,
                Organizer: event.organizers.join(', '),
                AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                WaitlistPosition: '',
              };
              let emailData: { subject: string; body: string };
              const spTpl = await eventService.getEmailTemplate('Nachrücken', lang).catch(() => null);
              if (spTpl) {
                emailData = buildEmailFromTemplate(spTpl, promoteVars);
              } else {
                emailData = promotionEmail(promoted.name || '', event.title);
              }
              if (promoted.email) {
                await eventService.queueEmail(
                  emailData.subject, promoted.email, promoted.name || '', emailData.body,
                  'Nachrücken', event.title, eventId
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

  return React.createElement(
    EventContext.Provider,
    {
      value: {
        events, isEventsLoading,
        createEvent, registerForEvent, cancelRegistration,
        getMyRegistration, getAllRegistrations, deleteEvent, updateEvent, updateMyRegistration, getMyEventNumbers, refreshEvents, refreshParticipantCounts, markExpiredEventsAsCompleted,
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
