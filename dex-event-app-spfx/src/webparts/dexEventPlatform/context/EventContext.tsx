/**
 * Event Context - zentraler State fuer alle Events
 *
 * Laedt Events aus der SharePoint-Liste DEX_Events.
 * Erstellt die Liste automatisch beim ersten Start.
 * Verwaltet Registrierungen in separaten Teilnehmerlisten.
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
  registerForEvent: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  cancelRegistration: (eventId: string) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
  refreshEvents: () => Promise<void>;
  getRegistrationListUrl: (eventId: string) => string;
}

export interface CreateEventInput {
  title: string;
  type: string;
  status: string;
  description: string;
  location: string;
  locationFilter: string;
  audience: string;
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
  // Map von EventId -> RegistrationListName fuer schnellen Zugriff
  const regListMap = React.useRef<Record<string, string>>({});

  const eventService = React.useMemo(() => new EventService(props.context), []);
  const currentUserEmail = props.context.pageContext.user.email;
  const currentUserName = props.context.pageContext.user.displayName;

  React.useEffect(() => {
    initEvents().catch(() => setIsEventsLoading(false));
  }, []);

  async function initEvents(): Promise<void> {
    await eventService.ensureEventsList();
    await loadEvents();
    setIsEventsLoading(false);
  }

  async function loadEvents(): Promise<void> {
    const spEvents = await eventService.getEvents();
    const mapped = await Promise.all(spEvents.map(e => mapSPEventToDeloitteEvent(e)));
    setEvents(mapped);
  }

  async function mapSPEventToDeloitteEvent(e: SPEvent): Promise<DeloitteEvent> {
    // RegistrationListName merken
    if (e.RegistrationListName) {
      regListMap.current[e.Id.toString()] = e.RegistrationListName;
    }

    // Teilnehmeranzahl ermitteln
    let currentParticipants = 0;
    let waitlistCount = 0;
    if (e.RegistrationListName) {
      try {
        const counts = await eventService.getRegistrationCount(e.RegistrationListName);
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
      locationAudience: e.Audience ? e.Audience.split(',').map(s => s.trim()) : [],
      startDate: e.StartDate || '',
      endDate: e.EndDate || '',
      registrationDeadline: e.RegistrationDeadline || '',
      description: e.Description || '',
      maxParticipants: e.MaxParticipants || 0,
      currentParticipants,
      waitlistCount,
      imageUrl: e.EventImageUrl || '',
      eventSpecificFields: customFields.map(cf => ({
        id: cf.id,
        label: cf.label,
        type: cf.type,
        required: cf.required,
        options: cf.options,
        helpText: '',
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

  async function registerForEvent(eventId: string, customData: Record<string, string>): Promise<boolean> {
    const regListName = regListMap.current[eventId];
    if (!regListName) return false;

    // Pruefen ob schon registriert
    const existing = await eventService.getMyRegistration(regListName, currentUserEmail);
    if (existing && existing.Status !== 'Cancelled') return false;

    // Pruefen ob Platz frei oder Waitlist
    const event = events.find(e => e.id === eventId);
    let status = 'Registered';
    if (event && event.maxParticipants > 0 && event.currentParticipants >= event.maxParticipants) {
      status = 'Waitlist';
    }

    const success = await eventService.registerForEvent(
      regListName, currentUserName, currentUserEmail, customData, status
    );

    if (success) {
      // Teilnehmer-IDs neu vergeben
      await reassignParticipantIds(regListName);
      await loadEvents();
    }
    return success;
  }

  async function cancelRegistration(eventId: string): Promise<boolean> {
    const regListName = regListMap.current[eventId];
    if (!regListName) return false;

    const myReg = await eventService.getMyRegistration(regListName, currentUserEmail);
    if (!myReg) return false;

    const success = await eventService.cancelRegistration(regListName, myReg.Id);
    if (success) {
      // Teilnehmer-IDs neu vergeben nach Stornierung
      await reassignParticipantIds(regListName);
      await loadEvents();
    }
    return success;
  }

  /**
   * Teilnehmer-IDs neu vergeben.
   * Alle aktiven Teilnehmer (Registered, Checked-In) bekommen fortlaufende IDs.
   * Abgemeldete und Waitlist-Teilnehmer bekommen keine ID.
   */
  async function reassignParticipantIds(regListName: string): Promise<void> {
    try {
      const allRegs = await eventService.getAllRegistrations(regListName);
      // Nur aktive Teilnehmer, sortiert nach Registrierungsdatum
      const active = allRegs
        .filter(r => r.Status === 'Registered' || r.Status === 'Checked-In')
        .sort((a, b) => new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());

      // IDs neu vergeben (Title-Feld als ParticipantId nutzen)
      for (let i = 0; i < active.length; i++) {
        const newId = (i + 1).toString();
        if (active[i].Title !== newId) {
          await eventService.updateRegistrationTitle(regListName, active[i].Id, newId);
        }
      }

      // Abgemeldete: ID loeschen
      const cancelled = allRegs.filter(r => r.Status === 'Cancelled');
      for (const reg of cancelled) {
        if (reg.Title && reg.Title !== '') {
          await eventService.updateRegistrationTitle(regListName, reg.Id, '');
        }
      }
    } catch {
      // ID-Neuvergabe fehlgeschlagen
    }
  }

  async function getMyRegistration(eventId: string): Promise<SPRegistration | null> {
    const regListName = regListMap.current[eventId];
    if (!regListName) return null;
    return eventService.getMyRegistration(regListName, currentUserEmail);
  }

  async function getAllRegistrations(eventId: string): Promise<SPRegistration[]> {
    const regListName = regListMap.current[eventId];
    if (!regListName) return [];
    return eventService.getAllRegistrations(regListName);
  }

  async function refreshEvents(): Promise<void> {
    await loadEvents();
  }

  function getRegistrationListUrl(eventId: string): string {
    const regListName = regListMap.current[eventId];
    const base = props.context.pageContext.web.absoluteUrl;
    if (regListName) {
      return `${base}/Lists/${regListName}/AllItems.aspx`;
    }
    return `${base}/Lists`;
  }

  return React.createElement(
    EventContext.Provider,
    {
      value: {
        events, isEventsLoading,
        createEvent, registerForEvent, cancelRegistration,
        getMyRegistration, getAllRegistrations, refreshEvents,
        getRegistrationListUrl,
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
