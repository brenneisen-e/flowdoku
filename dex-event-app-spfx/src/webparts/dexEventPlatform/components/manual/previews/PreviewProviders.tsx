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
 *
 * v13.7: 30 Demo-Teilnehmer mit gemischten Status, damit die
 * Teilnehmertabelle und die KPI-Kacheln in den Handbuch-Previews
 * realistisch befüllt sind.
 */
import * as React from 'react';
import { NavigationContext } from '../../../context/NavigationContext';
import { UserContext } from '../../../context/UserContext';
import { RoleContext } from '../../../context/RoleContext';
import { EventContext } from '../../../context/EventContext';
import { DeloitteEvent, User } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

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

interface DemoPerson {
  firstName: string;
  surname: string;
  status: 'Angemeldet' | 'QR versendet' | 'Eingecheckt' | 'Warteliste';
  jobTitle: string;
  department: string;
}

// 30 fiktive Teilnehmer für das „Office Event Köln" Demo-Event.
// Statusverteilung: 22 Angemeldet, 4 QR versendet, 2 Eingecheckt, 2 Warteliste.
const demoPeople: DemoPerson[] = [
  { firstName: 'Lena',      surname: 'Brandt',     status: 'Eingecheckt', jobTitle: 'Manager',        department: 'Audit' },
  { firstName: 'Jonas',     surname: 'Hartmann',   status: 'Eingecheckt', jobTitle: 'Senior',         department: 'Consulting' },
  { firstName: 'Sophie',    surname: 'Wagner',     status: 'QR versendet', jobTitle: 'Senior Manager', department: 'Tax' },
  { firstName: 'Maximilian',surname: 'Becker',     status: 'QR versendet', jobTitle: 'Consultant',     department: 'Consulting' },
  { firstName: 'Anna',      surname: 'Schulz',     status: 'QR versendet', jobTitle: 'Director',       department: 'Risk Advisory' },
  { firstName: 'Tim',       surname: 'Klein',      status: 'QR versendet', jobTitle: 'Senior',         department: 'Audit' },
  { firstName: 'Julia',     surname: 'Fischer',    status: 'Angemeldet',  jobTitle: 'Consultant',     department: 'Tax' },
  { firstName: 'Niklas',    surname: 'Weber',      status: 'Angemeldet',  jobTitle: 'Senior',         department: 'Consulting' },
  { firstName: 'Hannah',    surname: 'Meyer',      status: 'Angemeldet',  jobTitle: 'Manager',        department: 'Financial Advisory' },
  { firstName: 'Felix',     surname: 'Schäfer',    status: 'Angemeldet',  jobTitle: 'Senior Manager', department: 'Consulting' },
  { firstName: 'Lara',      surname: 'Koch',       status: 'Angemeldet',  jobTitle: 'Senior',         department: 'Risk Advisory' },
  { firstName: 'Paul',      surname: 'Richter',    status: 'Angemeldet',  jobTitle: 'Partner',        department: 'Audit' },
  { firstName: 'Mia',       surname: 'Wolf',       status: 'Angemeldet',  jobTitle: 'Consultant',     department: 'Consulting' },
  { firstName: 'Luca',      surname: 'Neumann',    status: 'Angemeldet',  jobTitle: 'Manager',        department: 'Tax' },
  { firstName: 'Emilia',    surname: 'Schwarz',    status: 'Angemeldet',  jobTitle: 'Senior',         department: 'Financial Advisory' },
  { firstName: 'Ben',       surname: 'Zimmermann', status: 'Angemeldet',  jobTitle: 'Consultant',     department: 'Audit' },
  { firstName: 'Clara',     surname: 'Braun',      status: 'Angemeldet',  jobTitle: 'Senior Manager', department: 'Risk Advisory' },
  { firstName: 'David',     surname: 'Krüger',     status: 'Angemeldet',  jobTitle: 'Senior',         department: 'Consulting' },
  { firstName: 'Marie',     surname: 'Hoffmann',   status: 'Angemeldet',  jobTitle: 'Manager',        department: 'Tax' },
  { firstName: 'Leon',      surname: 'Schmidt',    status: 'Angemeldet',  jobTitle: 'Consultant',     department: 'Consulting' },
  { firstName: 'Mara',      surname: 'Bauer',      status: 'Angemeldet',  jobTitle: 'Senior',         department: 'Audit' },
  { firstName: 'Henri',     surname: 'Schneider',  status: 'Angemeldet',  jobTitle: 'Director',       department: 'Financial Advisory' },
  { firstName: 'Lina',      surname: 'Lehmann',    status: 'Angemeldet',  jobTitle: 'Consultant',     department: 'Tax' },
  { firstName: 'Elias',     surname: 'Lang',       status: 'Angemeldet',  jobTitle: 'Senior',         department: 'Consulting' },
  { firstName: 'Pia',       surname: 'Vogel',      status: 'Angemeldet',  jobTitle: 'Manager',        department: 'Risk Advisory' },
  { firstName: 'Noah',      surname: 'Huber',      status: 'Angemeldet',  jobTitle: 'Senior Manager', department: 'Audit' },
  { firstName: 'Helena',    surname: 'Peters',     status: 'Angemeldet',  jobTitle: 'Consultant',     department: 'Consulting' },
  { firstName: 'Theo',      surname: 'Frank',      status: 'Angemeldet',  jobTitle: 'Senior',         department: 'Tax' },
  { firstName: 'Sara',      surname: 'Albers',     status: 'Warteliste',  jobTitle: 'Consultant',     department: 'Financial Advisory' },
  { firstName: 'Jakob',     surname: 'Walter',     status: 'Warteliste',  jobTitle: 'Senior',         department: 'Consulting' },
];

const emailFor = (p: DemoPerson): string =>
  `${p.firstName.toLowerCase().replace(/[äöüß]/g, c => ({ ä: 'ae', ö: 'oe', ü: 'ue', ß: 'ss' }[c] || c))}.${p.surname.toLowerCase().replace(/[äöüß]/g, c => ({ ä: 'ae', ö: 'oe', ü: 'ue', ß: 'ss' }[c] || c))}@deloitte.de`;

const demoRegistrations: SPRegistration[] = demoPeople.map((p, idx) => {
  // TeilnehmerIDs müssen lückenlos 1..N (über aktive UND Warteliste)
  // durchlaufen, sonst zeigt AdminPage die orange „IDs sind ggf. nicht
  // korrekt"-Warnbox (siehe `recentCancellation` in AdminPage.tsx).
  const tid = idx + 1;
  const regDate = new Date(2026, 7, 5 + idx).toISOString();
  const checkedIn = p.status === 'Eingecheckt';
  return {
    Id: idx + 1,
    Title: emailFor(p),
    TeilnehmerID: tid,
    Anrede: '',
    Vorname: p.firstName,
    Nachname: p.surname,
    ParticipantName: `${p.surname}, ${p.firstName}`,
    ParticipantEmail: emailFor(p),
    Status: p.status,
    RegistrationDate: regDate,
    CancellationDate: '',
    CheckedInDate: checkedIn ? '2026-09-15T17:48:00' : undefined,
    CheckedInByName: checkedIn ? 'Musterfrau, Maja' : undefined,
    CheckedInByEmail: checkedIn ? 'maja.musterfrau@deloitte.de' : undefined,
    Location: 'Köln',
    CustomData: JSON.stringify({ Department: p.department, JobTitle: p.jobTitle }),
  };
});

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
  currentParticipants: demoRegistrations.filter(r => r.Status !== 'Warteliste').length,
  waitlistCount: demoRegistrations.filter(r => r.Status === 'Warteliste').length,
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
    checkRegistrationByEmail: asyncNull,
    getAllRegistrations: async (eventId: string) =>
      eventId === DEMO_EVENT_ID ? demoRegistrations : [],
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
