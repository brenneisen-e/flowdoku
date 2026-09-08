/* Render-Harness für Screenshots des Event-Wizards — ausserhalb von SharePoint.
 * Alle Provider, die SPFx brauchen, werden durch Mock-Werte ersetzt; Language
 * und Dialog laufen echt. Nur für den Screenshot-Lauf, nicht Teil des Projekts. */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { LanguageProvider } from '../../src/webparts/dexEventPlatform/context/LanguageContext';
import { DialogProvider } from '../../src/webparts/dexEventPlatform/context/DialogContext';
import { UserContext } from '../../src/webparts/dexEventPlatform/context/UserContext';
import { RoleContext } from '../../src/webparts/dexEventPlatform/context/RoleContext';
import { NavigationContext } from '../../src/webparts/dexEventPlatform/context/NavigationContext';
import { EventContext } from '../../src/webparts/dexEventPlatform/context/EventContext';
import EventCreationPage from '../../src/webparts/dexEventPlatform/components/EventCreationPage';

try { localStorage.setItem('dex-locale', 'de'); } catch { /* */ }
try { initializeIcons(); } catch { /* */ }

const params = new URLSearchParams(window.location.search);
const mode = params.get('mode') || 'edit';

const me = { id: '1', firstName: 'Eike', surname: 'Brenneisen', email: 'eike.brenneisen@example.com', isAdmin: true, role: 'Admin', location: 'DE - Koeln', jobTitle: 'Manager' };

const d = (offsetDays: number, hh: number, mm = 0): string => {
  const x = new Date(); x.setDate(x.getDate() + offsetDays); x.setHours(hh, mm, 0, 0);
  return x.toISOString();
};
const day = (offsetDays: number): string => d(offsetDays, 9).slice(0, 10);

const parent: any = {
  id: '1', eventNumber: 4711, title: 'DTP Basics Training October 2026', type: 'Other', status: 'Active',
  organizers: ['Brenneisen, Eike'], organizerEmails: [me.email],
  coOrganizerNames: ['Rettinger, Carolin'], coOrganizerEmails: ['carolin.rettinger@example.com'],
  qrScannerNames: [], qrScannerEmails: [], contactName: 'Eike Brenneisen', contactEmail: me.email,
  location: 'Campus Hackescher Markt (Haus A, Etage 4)',
  locationAddress: { street: 'Litfaß-Platz', houseNo: '2', zip: '10178', city: 'Berlin' },
  locationAudience: [], audienceFilter: [], filterMode: 'OR',
  startDate: d(30, 9), endDate: d(34, 17), registrationDeadline: d(20, 23, 59), lastDeregisterDate: d(25, 23, 59),
  description: '<p>Fünf Tage Grundlagen für neue Kolleginnen und Kollegen: Fachthemen am Vormittag, Fallarbeit am Nachmittag, gemeinsames Abendprogramm.</p>',
  maxParticipants: 40, currentParticipants: 12, waitlistCount: 0, waitlistEnabled: true,
  outlookBody: '', emailLanguage: 'DE', imageUrl: '',
  agenda: [
    { id: 'a1', date: day(30), time: '09:00', endTime: '10:30', icon: 'Calendar', title: 'Welcome & Kick-off', location: 'Plenum', cluster: 'Tag 1' },
    { id: 'a2', date: day(30), time: '11:00', endTime: '12:30', icon: 'Calendar', title: 'Business Chemistry 1/2', location: 'Raum 3.12', cluster: 'Tag 1' },
    { id: 'a3', date: day(31), time: '09:00', endTime: '12:00', icon: 'Calendar', title: 'Joint Venture', location: 'Plenum', cluster: 'Tag 2' },
    { id: 'a4', date: day(31), time: '13:00', endTime: '16:30', icon: 'Calendar', title: 'Post Merger Integration Theory', location: 'Plenum', cluster: 'Tag 2' },
  ],
  transferTimes: [], documents: [], quiz: [],
  eventSpecificFields: [],
  teamRegistrationEnabled: false, teamSize: 4,
  childEventTermSingular: 'Session', childEventTermPlural: 'Sessions',
};
const child = (id: string, title: string, off: number): any => ({
  ...parent, id, eventNumber: 4711 + Number(id), title, parentEventId: '1',
  startDate: d(off, 9), endDate: d(off, 17), agenda: [], coOrganizerNames: [], coOrganizerEmails: [],
  maxParticipants: 20, currentParticipants: 5,
});
const children = [child('2', 'Day 1 · Business Chemistry', 30), child('3', 'Day 2 · Joint Venture', 31), child('4', 'Day 3 · PMI', 32)];
const events = mode === 'edit' ? [parent, ...children] : [];

const asyncNoop = async (): Promise<any> => undefined;
const proxy = <T extends object>(base: T, fallback: (key: string) => any): T =>
  new Proxy(base, { get: (t, k) => (k in t ? (t as any)[k] : fallback(String(k))) });

const eventCtx: any = proxy({
  events, isLoading: false, subsiteMap: {},
  childEventsOf: (pid: string) => events.filter(e => e.parentEventId === pid),
  calDayParentOf: () => null,
  getLastEventUpdateError: () => null,
  ensureEventDocuments: async () => [], refreshEventDocuments: async () => [],
  getMyRegistration: async () => null,
  getRegistrationsFor: async () => [],
  refreshEvents: asyncNoop, createEvent: asyncNoop, updateEvent: asyncNoop, deleteEvent: asyncNoop, deleteEventItemOnly: asyncNoop,
  requestCoOrganizerApprovals: asyncNoop, notifyNewCoOrganizers: asyncNoop, notifyAdminsExternalAudienceAccess: asyncNoop,
  eventDocuments: {}, registrationsByEvent: {},
}, () => asyncNoop);

const roleCtx: any = proxy({
  isAdmin: true, originalIsAdmin: true, isFA: false, isPowerUser: true, canCreateEvents: true, isLoading: false,
  roles: [], currentRole: 'Admin', siteUrl: '', lastRightsAudit: null,
  hadRoleRightsIssue: () => false, lastRoleRightsMissing: () => [],
  searchUser: async () => null, getGroupMembers: async () => null, refreshRoles: asyncNoop,
  searchUsers: async () => [], searchGroups: async () => [], searchUsersByLocation: async () => [],
  getEmployeeData: async () => ({}), getBasicProfiles: async () => ({}),
}, () => async () => []);

const navCtx: any = proxy({
  currentPage: mode === 'edit' ? 'edit-event' : 'create-event',
  selectedEventId: mode === 'edit' ? '1' : null,
  navIntent: undefined,
  navigate: () => undefined, goBack: () => undefined, clearIntent: () => undefined, setNavigationGuard: () => undefined,
  history: [],
}, () => () => undefined);

const userCtx: any = { currentUser: me, isLoading: false, photoUrl: '', groupEmails: [] };

const App: React.FC = () => (
  <div className="dexApp" lang="de-DE">
    <LanguageProvider>
      <DialogProvider>
        <UserContext.Provider value={userCtx}>
          <RoleContext.Provider value={roleCtx}>
            <NavigationContext.Provider value={navCtx}>
              <EventContext.Provider value={eventCtx}>
                <div className="main-content"><EventCreationPage /></div>
              </EventContext.Provider>
            </NavigationContext.Provider>
          </RoleContext.Provider>
        </UserContext.Provider>
      </DialogProvider>
    </LanguageProvider>
  </div>
);

ReactDOM.render(<App />, document.getElementById('root'));
