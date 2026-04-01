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
import { registrationEmail, waitlistEmail, cancellationEmail } from '../services/EmailTemplates';

interface EventContextType {
  events: DeloitteEvent[];
  isEventsLoading: boolean;
  createEvent: (event: CreateEventInput) => Promise<number | null>;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string) => Promise<boolean>;
  cancelRegistration: (eventId: string) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
  deleteEvent: (eventId: string) => Promise<boolean>;
  updateEvent: (eventId: string, updates: Record<string, unknown>) => Promise<boolean>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  getMyEventNumbers: () => Promise<{ registered: number[]; waitlisted: number[] }>;
  refreshEvents: () => Promise<void>;
}

export interface CreateEventInput {
  title: string;
  type: string;
  status: string;
  description: string;
  location: string;
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
    await eventService.ensureEventsList();
    await eventService.ensureEmailsList();
    await eventService.ensureOutlookList();
    await eventService.ensureParticipantsList();
    await eventService.ensureIDReorderList();
    await eventService.ensureAssetsFolders();
    await loadEvents();
    setIsEventsLoading(false);
  }

  async function loadEvents(): Promise<void> {
    const spEvents = await eventService.getEvents();
    const mapped = await Promise.all(spEvents.map(e => mapSPEventToDeloitteEvent(e)));
    setEvents(mapped);
  }

  async function mapSPEventToDeloitteEvent(e: SPEvent): Promise<DeloitteEvent> {
    // SubsiteUrl merken
    if (e.SubsiteUrl) {
      subsiteMap.current[e.Id.toString()] = e.SubsiteUrl;
    }

    // Teilnehmeranzahl ermitteln
    let currentParticipants = 0;
    let waitlistCount = 0;
    if (e.SubsiteUrl) {
      try {
        const counts = await eventService.getRegistrationCount(e.SubsiteUrl);
        currentParticipants = counts.registered;
        waitlistCount = counts.waitlist;
      } catch { /* Teilnehmerliste nicht erreichbar */ }
    }

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
      organizers: [e.Organizer || ''],
      location: e.Location || '',
      locationAudience: e.LocationFilter ? e.LocationFilter.split(',').map(s => s.trim()) : [],
      audienceFilter: e.Audience ? e.Audience.split(',').map(s => s.trim()) : [],
      filterMode: (e.FilterMode as 'AND' | 'OR') || 'OR',
      startDate: e.StartDate || '',
      endDate: e.EndDate || '',
      registrationDeadline: e.RegistrationDeadline || '',
      description: e.Description || '',
      maxParticipants: e.MaxParticipants || 0,
      currentParticipants,
      waitlistCount,
      imageUrl: e.EventImageUrl || '',
      subsiteUrl: e.SubsiteUrl || '',
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
      await loadEvents();
    }
    return eventId;
  }

  async function registerForEvent(
    eventId: string,
    customData: Record<string, string>,
    participantFirstName?: string,
    participantLastName?: string,
    participantEmail?: string
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
    if (event && event.maxParticipants > 0 && event.currentParticipants >= event.maxParticipants) {
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
        subsiteUrl, firstNameToUse, lastNameToUse, emailToUse, customData, status, fieldMap
      );
    }

    if (success && event) {
      // Dual-Write: DEX_Participants aktualisieren
      if (event.eventNumber) {
        eventService.upsertParticipant(
          firstNameToUse, lastNameToUse, emailToUse, event.eventNumber, status
        ).catch(() => {});
      }
      // E-Mail in Queue eintragen (Deloitte-Template)
      const emailData = status === 'Warteliste'
        ? waitlistEmail(nameToUse, event.title)
        : registrationEmail(nameToUse, event.title);
      eventService.queueEmail(
        emailData.subject, emailToUse, nameToUse, emailData.body,
        status === 'Warteliste' ? 'Warteliste' : 'Anmeldung',
        event.title, eventId
      ).catch(() => {});
      // Outlook-Termin-Einladung in Queue eintragen
      if (status !== 'Warteliste') {
        eventService.queueOutlookEvent(
          emailToUse, eventId, event.title, 'Einladen'
        ).catch(() => {});
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

    const success = await eventService.cancelRegistration(subsiteUrl, myReg.Id);
    if (success) {
      const event = events.find(e => e.id === eventId);
      if (event) {
        // Dual-Write: DEX_Participants aktualisieren
        if (event.eventNumber) {
          eventService.removeParticipantEvent(currentUserEmail, event.eventNumber).catch(() => {});
        }
        const emailData = cancellationEmail(currentUserName, event.title);
        eventService.queueEmail(
          emailData.subject, currentUserEmail, currentUserName, emailData.body,
          'Abmeldung', event.title, eventId
        ).catch(() => {});
        // Outlook-Termin-Einladung zurückziehen
        eventService.queueOutlookEvent(
          currentUserEmail, eventId, event.title, 'Ausladen'
        ).catch(() => {});
        // ID-Reorder in Queue eintragen
        if (subsiteUrl) {
          eventService.queueIDReorder(
            eventId, event.eventNumber || 0, subsiteUrl, event.title
          ).catch(() => {});
        }
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

  return React.createElement(
    EventContext.Provider,
    {
      value: {
        events, isEventsLoading,
        createEvent, registerForEvent, cancelRegistration,
        getMyRegistration, getAllRegistrations, deleteEvent, updateEvent, updateMyRegistration, getMyEventNumbers, refreshEvents,
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
