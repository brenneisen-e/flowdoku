/* Render-Harness für Screenshots — ausserhalb von SharePoint.
 * Alle Provider, die SPFx brauchen, werden durch Mock-Werte ersetzt; Language
 * und Dialog laufen echt. Nur für den Screenshot-Lauf, nicht Teil des Projekts.
 *
 * ?mode=edit|create  → Event-Wizard (wie seit v31.2)
 * ?page=landing|start|list|register|myevents → die Teilnehmer-Seiten (v31.7)
 * ?mobile=1          → Handy-Zweige (window.__dexForceMobile, s. useIsMobile)
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { LanguageProvider } from '../../src/webparts/dexEventPlatform/context/LanguageContext';
import { DialogProvider } from '../../src/webparts/dexEventPlatform/context/DialogContext';
import { UserContext } from '../../src/webparts/dexEventPlatform/context/UserContext';
import { RoleContext } from '../../src/webparts/dexEventPlatform/context/RoleContext';
import { NavigationContext } from '../../src/webparts/dexEventPlatform/context/NavigationContext';
import { EventContext } from '../../src/webparts/dexEventPlatform/context/EventContext';
import { TicketContext } from '../../src/webparts/dexEventPlatform/context/TicketContext';
import EventCreationPage from '../../src/webparts/dexEventPlatform/components/EventCreationPage';
import LandingPage from '../../src/webparts/dexEventPlatform/components/LandingPage';
import StartPage from '../../src/webparts/dexEventPlatform/components/StartPage';
import EventListPage from '../../src/webparts/dexEventPlatform/components/EventListPage';
import RegistrationPage from '../../src/webparts/dexEventPlatform/components/RegistrationPage';
import MyEventsPage from '../../src/webparts/dexEventPlatform/components/MyEventsPage';
import * as sample from './sampleData';

const params = new URLSearchParams(window.location.search);
const mode = params.get('mode') || 'edit';
const page = params.get('page') || 'wizard';

/* Der Handy-Schalter muss VOR dem ersten Render stehen: `useIsMobile` liest ihn
 * im State-Initialisierer. Wer ihn später setzt, sieht bis zum nächsten
 * matchMedia-Ereignis weiter die Desktop-Fassung — also nie. */
if (params.get('mobile') === '1') {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__dexForceMobile = true;
}

/* Merker aus früheren Läufen würden die Seiten verfälschen (Landing zeigt
 * gecachte „Du bist angemeldet"-Kacheln, MyEvents seinen 60-s-Cache). Der
 * Harness startet deshalb immer leer — nur die Sprache bleibt gesetzt. */
try { localStorage.clear(); sessionStorage.clear(); } catch { /* */ }
try { localStorage.setItem('dex-locale', 'de'); } catch { /* */ }
/* ?view=list schaltet die Event-Übersicht auf die Listen-Ansicht. Die merkt
 * sich die Seite selbst im localStorage — und nur in dieser Ansicht gibt es
 * überhaupt einen isMobile-Zweig (EventListView). */
try {
  const v = params.get('view');
  if (v) localStorage.setItem('dex-eventlist-view', v);
} catch { /* */ }
try { initializeIcons(); } catch { /* */ }

const me = sample.me;
const d = sample.d;
const day = sample.day;

/* ------------------------------------------------------- Wizard-Beispiel --- */

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
const wizardChildren = [child('2', 'Day 1 · Business Chemistry', 30), child('3', 'Day 2 · Joint Venture', 31), child('4', 'Day 3 · PMI', 32)];
const wizardEvents = mode === 'edit' ? [parent, ...wizardChildren] : [];

/* ------------------------------------------------------------ Kontexte ----- */

const isWizard = page === 'wizard';
const events: any[] = isWizard ? wizardEvents : sample.allEvents;
const topLevelEvents: any[] = isWizard
  ? wizardEvents.filter(e => !e.parentEventId)
  : sample.topLevelEvents;

const asyncNoop = async (): Promise<any> => undefined;
const proxy = <T extends object>(base: T, fallback: (key: string) => any): T =>
  new Proxy(base, { get: (t, k) => (k in t ? (t as any)[k] : fallback(String(k))) });

/* Der Proxy-Fallback unten liefert `async () => undefined`. Das trägt für alles,
 * was nur „gemacht" wird (Speichern, Mail, Refresh) — aber NICHT für Aufrufe,
 * deren Rückgabe die Seite gleich weiterverarbeitet. `undefined.length`,
 * `undefined.filter` und `undefined.registered` sind der schnellste Weg zu
 * einer weißen Seite. Alles, was ein Array, ein Objekt, ein Set oder eine
 * Aufräum-Funktion zurückgeben MUSS, steht deshalb hier ausdrücklich. */
const eventCtx: any = proxy({
  events, topLevelEvents, isLoading: false, isEventsLoading: false,
  eventsReadStatus: 'ok', subsiteMap: {},
  childEventsOf: (pid: string) => events.filter(e => e.parentEventId === pid),
  calDayParentOf: () => null,
  getLastEventUpdateError: () => '',
  getLastEventDeleteError: () => '',
  eventDocuments: {}, registrationsByEvent: {},

  // Dokumente/Bilder — Rückgabe wird direkt gemappt.
  ensureEventDocuments: async () => undefined,
  refreshEventDocuments: async () => undefined,
  listMyEventAttachments: async () => [],
  listFieldDocuments: async () => [],

  // Eigene Anmeldungen: die Quelle fast aller Teilnehmer-Ansichten.
  getMyRegistration: async (eventId: string) => sample.myRegistrations[eventId] || null,
  checkRegistrationByEmail: async () => null,
  getAllRegistrations: async (eventId: string) => sample.registrationsOf(eventId),
  getRegistrationsFor: async () => [],
  getMyEventNumbers: async () => sample.myEventNumbers,
  getEventNumbersForEmail: async () => ({ registered: [], waitlisted: [] }),
  getMyProxyRegistrations: async () => [],
  getMyAssistantLinks: async () => [],
  getEventComms: async () => [],
  getTeamMembers: async () => [],
  listOpenTeamsForEvent: async () => [],
  listTeamJoinRequestsForEvent: async () => [],

  // Der Platzzähler ist die Quelle, die JEDER lesen darf — und die einzige,
  // aus der Anmeldeseite und Sub-Event-Liste ihre Belegung nehmen. Ohne ihn
  // fallen sie auf die (zeilenweise gesicherte) Teilnehmerliste zurück und
  // zeigen „1/20", weil dort nur die eigene Zeile steht.
  getLiveCounterStats: async (eventId: string) => {
    const ev = events.find(e => e.id === eventId);
    if (!ev) return null;
    return {
      active: ev.currentParticipants || 0,
      waitlist: ev.waitlistCount || 0,
      seatsKnown: true, durch: 0, fun: 0, groupsKnown: false,
    };
  },
  // MUSS eine Funktion liefern — die Anmeldeseite ruft sie beim Unmount auf.
  subscribeEventRealtime: async () => () => { /* kein Socket im Harness */ },
  refreshParticipantCounts: async () => undefined,
  refreshEvents: asyncNoop,

  // Admin-Kästen der Landing Page — alle vier Zählungen laufen beim Mount an.
  getArchivableCount: async () => ({ total: 0, perList: {} }),
  getDeletableArchiveCount: async () => 0,
  getParticipantDeletionWarnings: async () => [],
  getParticipantDeletionDue: async () => [],
  maybeSendParticipantDeletionWarnings: asyncNoop,
  scanInactiveAccounts: async () => [],
  getSentInactiveNotices: async () => new Set<string>(),
  countExternalRegistrations: async () => 0,
  getOrganizerArchivedEventIds: async () => new Set<string>(),

  // Schreibpfade: im Harness bewusst wirkungslos.
  createEvent: asyncNoop, updateEvent: asyncNoop, deleteEvent: asyncNoop, deleteEventItemOnly: asyncNoop,
  registerForEvent: async () => ({ ok: false, status: 'Angemeldet', reason: 'harness' }),
  registerTeam: async () => ({ ok: false }),
  cancelRegistration: async () => false,
  declineEvent: async () => false,
  updateMyRegistration: async () => false,
  uploadFieldDocument: async () => false,
  delegateRegistrationToAssistant: asyncNoop,
  recordProxyDelegation: asyncNoop,
  sendBundledUpdateMail: async () => false,
  requestCoOrganizerApprovals: asyncNoop, notifyNewCoOrganizers: asyncNoop,
  notifyAdminsExternalAudienceAccess: asyncNoop,
}, () => asyncNoop);

const roleCtx: any = proxy({
  isAdmin: true, originalIsAdmin: true, isFA: false, isOrganizer: true, isPowerUser: true,
  canCreateEvents: true, isLoading: false, isRolesLoading: false, isImpersonating: false,
  previewAsUser: false, setPreviewAsUser: () => undefined,
  roles: [], currentRole: 'Admin', currentUserRole: 'Admin', siteUrl: '',
  rolesReadStatus: 'ok', lastRightsAudit: null,
  hadRoleRightsIssue: () => false, lastRoleRightsMissing: () => [],
  searchUser: async () => null, getGroupMembers: async () => null, refreshRoles: asyncNoop,
  searchUsers: async () => [], searchGroups: async () => [], searchUsersByLocation: async () => [],
  getEmployeeData: async () => ({}), getBasicProfiles: async () => ({}),
}, () => async () => []);

/* Welche Seite der NavigationContext meldet — einige Komponenten lesen
 * `currentPage`/`selectedEventId` und nicht nur die Props. */
const PAGE_OF: Record<string, string> = {
  wizard: mode === 'edit' ? 'edit-event' : 'create-event',
  landing: 'landing', start: 'start', list: 'register',
  register: 'registration', myevents: 'my-events',
};
const SELECTED_OF: Record<string, string | null> = {
  wizard: mode === 'edit' ? '1' : null,
  landing: null, start: null, list: null,
  register: 'ev-open', myevents: null,
};

/* ?event=<id> zeigt die Anmeldeseite eines anderen Beispiel-Events — vor allem
 * `ev-umb`, die Klammer mit drei Terminen: das ist die Ansicht mit der
 * Termin-Auswahl, und die sieht man mit `ev-open` nie. */
const selectedEventId = params.get('event')
  || (SELECTED_OF[page] !== undefined ? SELECTED_OF[page] : null);

const navCtx: any = proxy({
  currentPage: PAGE_OF[page] || 'landing',
  selectedEventId,
  navIntent: undefined,
  navigate: () => undefined, goBack: () => undefined, clearIntent: () => undefined, setNavigationGuard: () => undefined,
  history: [],
}, () => () => undefined);

const userCtx: any = { currentUser: me, isLoading: false, photoUrl: '', groupEmails: [] };

/* Die Startseite liest `useTickets()` — der Hook wirft ohne Provider, die
 * Seite käme gar nicht erst zum Rendern. `powerUserQueue` muss ein Array sein
 * (Badge-Zählung mit `.filter`). */
const ticketCtx: any = proxy({
  tickets: [], myTickets: [], powerUserQueue: [], isTicketsLoading: false,
  canAnswerTickets: true,
  ticketsForEvent: () => [], openCountForEvent: () => 0, reminderTargets: () => [],
}, () => asyncNoop);

/* ------------------------------------------------------------- Rendern ----- */

const PageComponent: React.FC = () => {
  switch (page) {
    case 'landing': return <LandingPage />;
    case 'start': return <StartPage />;
    case 'list': return <EventListPage />;
    case 'register': return <RegistrationPage />;
    case 'myevents': return <MyEventsPage />;
    default: return <EventCreationPage />;
  }
};

const App: React.FC = () => (
  <div className="dexApp" lang="de-DE">
    <LanguageProvider>
      <DialogProvider>
        <UserContext.Provider value={userCtx}>
          <RoleContext.Provider value={roleCtx}>
            <NavigationContext.Provider value={navCtx}>
              <EventContext.Provider value={eventCtx}>
                <TicketContext.Provider value={ticketCtx}>
                  <div className="main-content"><PageComponent /></div>
                </TicketContext.Provider>
              </EventContext.Provider>
            </NavigationContext.Provider>
          </RoleContext.Provider>
        </UserContext.Provider>
      </DialogProvider>
    </LanguageProvider>
  </div>
);

ReactDOM.render(<App />, document.getElementById('root'));
