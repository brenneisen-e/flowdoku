/**
 * Preview-Provider für das Handbuch (v6.27).
 *
 * Wrappt echte App-Komponenten mit isolierten Context-Providern, die
 * Demo-Daten bereitstellen — damit der Nutzer im Handbuch die echte
 * Layout-/Styling-Ansicht sieht, aber keine SharePoint-Aufrufe rausgehen
 * und kein echter State verändert wird. Klicks werden im Wrapper-Modal
 * zusätzlich per pointer-events:none neutralisiert (siehe AppPreview.tsx).
 *
 * Demo-Event: "Office Event Köln" (kein B2Run, um die generische
 * Check-In-Story klarer zu zeigen).
 */
import * as React from 'react';
import { NavigationContext } from '../../../context/NavigationContext';
import { UserContext } from '../../../context/UserContext';
import { RoleContext } from '../../../context/RoleContext';
import { EventContext } from '../../../context/EventContext';
import { DeloitteEvent, User } from '../../../types';

export type PreviewRole = 'Admin' | 'Organizer' | 'User';

const noop = (): void => { /* intentionally empty — preview is read-only */ };
const asyncNoop = async (): Promise<boolean> => false;
const asyncNull = async (): Promise<null> => null;
const asyncArr = async (): Promise<[]> => [];

const demoUser: User = {
  id: 'demo-user',
  firstName: 'Maja',
  surname: 'Musterfrau',
  email: 'maja.musterfrau@deloitte.de',
  isAdmin: false,
  role: 'Organizer',
  location: 'Köln',
  jobTitle: 'Senior',
};

export const DEMO_EVENT_ID = 'demo-event-1';
const demoEvent: DeloitteEvent = {
  id: DEMO_EVENT_ID,
  eventNumber: 42,
  title: 'Office Event Köln',
  type: 'Other',
  status: 'Active',
  organizers: ['Musterfrau, Maja'],
  organizerEmails: ['maja.musterfrau@deloitte.de'],
  qrScannerNames: [],
  qrScannerEmails: [],
  location: 'Deloitte Office Köln',
  locationAudience: ['Köln'],
  audienceFilter: [],
  filterMode: 'OR',
  startDate: '2026-09-15T18:00:00',
  endDate: '2026-09-15T22:00:00',
  registrationDeadline: '2026-09-10T23:59:00',
  lastDeregisterDate: '2026-09-12T23:59:00',
  description: 'Das jährliche Office-Event der Kölner Niederlassung — Networking, Food-Trucks, Musik.',
  maxParticipants: 150,
  currentParticipants: 87,
  waitlistCount: 0,
  imageUrl: '',
  outlookBody: '',
  emailLanguage: 'DE',
  agenda: [],
  transferTimes: [],
  documents: [],
  quiz: [],
  eventSpecificFields: [],
};

export function PreviewContextStack(props: {
  children: React.ReactNode;
  role?: PreviewRole;
  /** Wenn gesetzt, startet NavigationContext mit dieser Page (statt 'landing').
   *  Für Previews die nicht die Root-Page sondern eine konkrete Sub-Page
   *  zeigen (z.B. CheckInPage). */
  page?: string;
  /** Wenn gesetzt, startet NavigationContext mit dieser selectedEventId.
   *  Brauchen wir z.B. für die Scanner-Preview (CheckInPage mit Event
   *  schon ausgewählt, damit das Scanner-UI statt des Event-Pickers rendert). */
  selectedEventId?: string;
  /** Zusätzliche Events über das Default-Demo-Event hinaus. */
  extraEvents?: DeloitteEvent[];
}): React.ReactElement {
  const role: PreviewRole = props.role || 'Organizer';
  const allEvents = props.extraEvents ? [demoEvent, ...props.extraEvents] : [demoEvent];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const nav: any = {
    currentPage: props.page || 'landing',
    selectedEventId: props.selectedEventId || null,
    navIntent: undefined,
    navigate: noop, goBack: noop, clearIntent: noop,
  };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const user: any = { currentUser: demoUser, isLoading: false, photoUrl: '' };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const roles: any = {
    roles: [], currentUserRole: role, isRolesLoading: false,
    isAdmin: role === 'Admin', isOrganizer: role !== 'User',
    canCreateEvents: role !== 'User', siteUrl: '',
    addRole: asyncNoop, updateRole: asyncNoop, updateRoleLocation: asyncNoop,
    removeRole: asyncNoop, refreshRoles: async () => { /* */ },
    searchUser: asyncNull, searchUsers: asyncArr, searchGroups: asyncArr,
    getGroupMembers: asyncNull,
  };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const events: any = {
    events: allEvents, topLevelEvents: allEvents,
    childEventsOf: () => [], isEventsLoading: false,
    createEvent: async () => null, registerForEvent: asyncNoop,
    cancelRegistration: asyncNoop, getMyRegistration: asyncNull,
    checkRegistrationByEmail: asyncNull, getAllRegistrations: asyncArr,
    deleteEvent: asyncNoop, updateEvent: asyncNoop,
    updateMyRegistration: asyncNoop,
    getMyEventNumbers: async () => ({ registered: [], waitlisted: [] }),
    markExpiredEventsAsCompleted: async () => 0,
    sendAdminInquiry: asyncNoop, refreshEvents: async () => { /* */ },
    applyEventTemplateOverride: (raw: string) => raw,
    formatOrganizerList: () => '',
  };
  return (
    <NavigationContext.Provider value={nav}>
      <UserContext.Provider value={user}>
        <RoleContext.Provider value={roles}>
          <EventContext.Provider value={events}>
            {props.children}
          </EventContext.Provider>
        </RoleContext.Provider>
      </UserContext.Provider>
    </NavigationContext.Provider>
  );
}
