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

interface EventContextType {
  events: DeloitteEvent[];
  isEventsLoading: boolean;
  createEvent: (event: CreateEventInput) => Promise<number | null>;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantName?: string, participantEmail?: string) => Promise<boolean>;
  cancelRegistration: (eventId: string) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
  deleteEvent: (eventId: string) => Promise<boolean>;
  updateEvent: (eventId: string, updates: Record<string, unknown>) => Promise<boolean>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
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
    participantName?: string,
    participantEmail?: string
  ): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;

    const nameToUse = participantName || currentUserName;
    const emailToUse = participantEmail || currentUserEmail;

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
      success = await eventService.reactivateRegistration(subsiteUrl, existing.Id, nameToUse, customData, status, fieldMap);
    } else {
      success = await eventService.registerForEvent(
        subsiteUrl, nameToUse, emailToUse, customData, status, fieldMap
      );
    }

    if (success && event) {
      // E-Mail in Queue eintragen
      const emailType = status === 'Warteliste' ? 'Warteliste' : 'Anmeldung';
      eventService.queueEmail(
        `${emailType}: ${event.title}`,
        emailToUse,
        nameToUse,
        `Hallo ${nameToUse},<br><br>du wurdest erfolgreich für "${event.title}" ${status === 'Warteliste' ? 'auf die Warteliste gesetzt' : 'angemeldet'}.<br><br>Viele Grüße,<br>DEX Event Platform`,
        emailType,
        event.title,
        eventId
      ).catch(() => {});
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
        eventService.queueEmail(
          `Abmeldung: ${event.title}`,
          currentUserEmail,
          currentUserName,
          `Hallo ${currentUserName},<br><br>du wurdest erfolgreich von "${event.title}" abgemeldet.<br><br>Viele Grüße,<br>DEX Event Platform`,
          'Abmeldung',
          event.title,
          eventId
        ).catch(() => {});
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

    return eventService.updateRegistrationData(subsiteUrl, myReg.Id, customData, fieldMap);
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
        getMyRegistration, getAllRegistrations, deleteEvent, updateEvent, updateMyRegistration, refreshEvents,
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
