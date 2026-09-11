/* Render-Harness für die APP-HÜLLE — Header + .app-layout + .main-content.
 *
 * Warum eine zweite Einstiegsdatei neben `entry.tsx`: Die dortige App rendert
 * bewusst NUR die Seite (`.dexApp > .main-content > Seite`) und lässt Header
 * und `.app-layout` weg, damit `overflow: hidden` die fullPage-Bilder nicht
 * abschneidet. Für eine Kartierung der Hülle ist aber genau das der
 * Gegenstand — also baut diese Datei den Rahmen so nach, wie ihn
 * `DexEventPlatform.tsx` aufspannt, inklusive der JS-gesetzten Höhe auf
 * `.app-layout` (dort `layoutRef`, Zeile ~808) und der
 * `html, body { overflow: hidden; height: 100vh }`-Injektion.
 *
 * ?view=landing|start   → welche Seite unter dem Header steht (Vorgabe landing)
 * ?role=admin|user      → echte Rolle; `user` zeigt die „Neu hier?"-Pille
 *                         mittig im Header (sie hängt an currentUserRole)
 * ?mobile=1             → Handy-Zweige (window.__dexForceMobile, s. useIsMobile)
 *
 * Nicht Teil des Produkt-Builds.
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
import Header from '../../src/webparts/dexEventPlatform/components/Header';
import LandingPage from '../../src/webparts/dexEventPlatform/components/LandingPage';
import StartPage from '../../src/webparts/dexEventPlatform/components/StartPage';
import * as sample from './sampleData';

const params = new URLSearchParams(window.location.search);
const view = params.get('view') || 'landing';
const roleParam = params.get('role') || 'admin';

if (params.get('mobile') === '1') {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__dexForceMobile = true;
}

try { localStorage.clear(); sessionStorage.clear(); } catch { /* */ }
try { localStorage.setItem('dex-locale', 'de'); } catch { /* */ }
try { initializeIcons(); } catch { /* */ }

const me = sample.me;
const asyncNoop = async (): Promise<any> => undefined;
const proxy = <T extends object>(base: T, fallback: (key: string) => any): T =>
  new Proxy(base, { get: (t, k) => (k in t ? (t as any)[k] : fallback(String(k))) });

/* Dieselben Pflicht-Rückgaben wie in entry.tsx: Alles, was ein Array, ein
 * Objekt, ein Set oder eine Aufräum-Funktion liefern MUSS, steht ausdrücklich
 * hier — der Proxy-Fallback `async () => undefined` trägt nur für Aufrufe,
 * deren Rückgabe niemand weiterverarbeitet. */
const eventCtx: any = proxy({
  events: sample.allEvents, topLevelEvents: sample.topLevelEvents,
  isLoading: false, isEventsLoading: false,
  eventsReadStatus: 'ok', subsiteMap: {},
  childEventsOf: (pid: string) => sample.allEvents.filter((e: any) => e.parentEventId === pid),
  calDayParentOf: () => null,
  getLastEventUpdateError: () => '',
  getLastEventDeleteError: () => '',
  eventDocuments: {}, registrationsByEvent: {},
  ensureEventDocuments: async () => undefined,
  refreshEventDocuments: async () => undefined,
  listMyEventAttachments: async () => [],
  listFieldDocuments: async () => [],
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
  getLiveCounterStats: async (eventId: string) => {
    const ev = sample.allEvents.find((e: any) => e.id === eventId);
    if (!ev) return null;
    return { active: ev.currentParticipants || 0, waitlist: ev.waitlistCount || 0, seatsKnown: true, durch: 0, fun: 0, groupsKnown: false };
  },
  subscribeEventRealtime: async () => () => { /* kein Socket im Harness */ },
  refreshParticipantCounts: async () => undefined,
  refreshEvents: asyncNoop,
  setTutorialDemoActive: () => undefined,
  getArchivableCount: async () => ({ total: 0, perList: {} }),
  getDeletableArchiveCount: async () => 0,
  getParticipantDeletionWarnings: async () => [],
  getParticipantDeletionDue: async () => [],
  maybeSendParticipantDeletionWarnings: asyncNoop,
  scanInactiveAccounts: async () => [],
  getSentInactiveNotices: async () => new Set<string>(),
  countExternalRegistrations: async () => 0,
  getOrganizerArchivedEventIds: async () => new Set<string>(),
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

const isAdminRole = roleParam === 'admin';
const roleCtx: any = proxy({
  isAdmin: isAdminRole, originalIsAdmin: isAdminRole, isFA: false,
  isOrganizer: isAdminRole, isPowerUser: isAdminRole,
  canCreateEvents: isAdminRole, isLoading: false, isRolesLoading: false, isImpersonating: false,
  previewAsUser: false, setPreviewAsUser: () => undefined,
  roles: [], currentRole: isAdminRole ? 'Admin' : 'User',
  currentUserRole: isAdminRole ? 'Admin' : 'User', siteUrl: '',
  rolesReadStatus: 'ok', lastRightsAudit: null,
  hadRoleRightsIssue: () => false, lastRoleRightsMissing: () => [],
  searchUser: async () => null, getGroupMembers: async () => null, refreshRoles: asyncNoop,
  searchUsers: async () => [], searchGroups: async () => [], searchUsersByLocation: async () => [],
  getEmployeeData: async () => ({}), getBasicProfiles: async () => ({}),
}, () => async () => []);

const navCtx: any = proxy({
  currentPage: view === 'start' ? 'start' : 'landing',
  selectedEventId: null,
  navIntent: undefined,
  navigate: () => undefined, goBack: () => undefined, clearIntent: () => undefined, setNavigationGuard: () => undefined,
  history: [],
}, () => () => undefined);

const userCtx: any = { currentUser: me, isLoading: false, photoUrl: '', groupEmails: [] };

const ticketCtx: any = proxy({
  tickets: [], myTickets: [], powerUserQueue: [], isTicketsLoading: false,
  canAnswerTickets: true,
  ticketsForEvent: () => [], openCountForEvent: () => 0, reminderTargets: () => [],
}, () => asyncNoop);

/* ---------------------------------------------------------- Die Hülle ----- */

/** Baut denselben Rahmen wie `AppContent` in DexEventPlatform.tsx:
 *  `.app-layout` (Höhe per JS) → `<Header/>` → `<main class="main-content">`. */
const Shell: React.FC = () => {
  const layoutRef = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    // 1:1 die Wirkung des `dex-no-scroll`-Style-Tags aus DexEventPlatform.tsx
    // (dort mit den SharePoint-Selektoren, die es hier nicht gibt).
    const el = document.createElement('style');
    el.id = 'dex-no-scroll';
    el.textContent = 'html, body { overflow: hidden !important; height: 100vh !important; max-height: 100vh !important; }';
    document.head.appendChild(el);
    function setHeight(): void {
      if (layoutRef.current) {
        const rect = layoutRef.current.getBoundingClientRect();
        const available = window.innerHeight - rect.top;
        layoutRef.current.style.height = `${Math.max(available, 400)}px`;
      }
    }
    setHeight();
    window.addEventListener('resize', setHeight);
    const t1 = window.setTimeout(setHeight, 500);
    return () => {
      window.removeEventListener('resize', setHeight);
      window.clearTimeout(t1);
      const e = document.getElementById('dex-no-scroll');
      if (e) e.remove();
    };
  }, []);

  return (
    <div className="app-layout" ref={layoutRef}>
      <Header />
      <main className="main-content" style={{ position: 'relative', display: 'flex', flexDirection: 'column' }}>
        {view === 'start' ? <StartPage /> : <LandingPage />}
      </main>
    </div>
  );
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
                  <Shell />
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
