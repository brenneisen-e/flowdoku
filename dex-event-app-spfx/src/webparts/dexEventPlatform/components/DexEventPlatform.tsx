/**
 * DEX Event Experience Platform - Hauptkomponente (SPFx)
 *
 * Wrapper für Navigation und Event Context.
 * Rendert je nach currentPage die passende Unterseite.
 *
 * Autor: Eike Brenneisen
 */

import * as React from 'react';
import DexLogo from './DexLogo';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './DexEventPlatform.module.scss';
import { NavigationProvider, useNavigation } from '../context/NavigationContext';
import { LanguageProvider } from '../context/LanguageContext';
// v20.4: Moderne Confirm-/Alert-Modals statt nativer Browser-Dialoge.
import { DialogProvider } from '../context/DialogContext';
import { EventProvider, useEvents } from '../context/EventContext';
import { BOOT_STAGE_EVENT, BootStage, BootStageDetail, lastBootStage } from '../utils/bootProgress';
import { UserProvider } from '../context/UserContext';
import { RoleProvider, useRoles } from '../context/RoleContext';
// v22.21: Geführtes Tutorial (Onboarding-Tour) — Provider + Overlay.
import { TutorialProvider } from './tutorial/TutorialGuide';
import { TicketProvider } from '../context/TicketContext';
import { deepLinkParams } from '../utils/deepLink';
import Header from './Header';
import LandingPage from './LandingPage';
import StartPage from './StartPage';
import EventListPage from './EventListPage';
import RegistrationPage from './RegistrationPage';
import MyEventsPage from './MyEventsPage';
import ProfilePage from './ProfilePage';
import SelfCheckInPage from './SelfCheckInPage';
import OrganizerRequestsBanner from './OrganizerRequestsBanner';
import GrantAccessHandler from './GrantAccessHandler';
import InviteDownloadHandler from './InviteDownloadHandler';
import { KpiRow } from './LandingPage';

// v20.0 (Audit): Route-Level-Code-Splitting. Die schweren Sekundär-Seiten
// (Wizard ~13k Zeilen + react-datepicker, Organizer Center ~9.5k Zeilen + xlsx,
// Check-in + qr-scanner, Handbuch, Flowcharts, Settings, Rollenmatrix,
// Teilnehmer-KPIs, Live-QR-Anzeige) werden erst beim ersten Aufruf der
// jeweiligen Seite als eigener Webpack-Chunk nachgeladen. Das verkleinert
// das Haupt-Bundle für JEDEN App-Start massiv — User, die nur die
// Anmelde-Flows nutzen, laden den Organizer-Code nie. Die Erst-Navigation
// auf eine dieser Seiten zeigt kurz den Suspense-Fallback (Chunk-Load vom
// SharePoint-CDN, typisch < 0,5 s), danach ist der Chunk gecacht.
const EventCreationPage = React.lazy(() => import('./EventCreationPage'));
const SettingsPage = React.lazy(() => import('./SettingsPage'));
const AdminPage = React.lazy(() => import('./AdminPage'));
const RoleMatrixPage = React.lazy(() => import('./RoleMatrixPage'));
const ParticipantsPage = React.lazy(() => import('./ParticipantsPage'));
const FlowchartPage = React.lazy(() => import('./FlowchartPage'));
const AdminHubPage = React.lazy(() => import('./AdminHubPage'));
const EmailTemplatesPage = React.lazy(() => import('./EmailTemplatesPage'));
const CheckInPage = React.lazy(() => import('./CheckInPage'));
const SelfCheckInDisplayPage = React.lazy(() => import('./SelfCheckInDisplayPage'));
const ManualPage = React.lazy(() => import('./manual/ManualPage'));
const AssistantPage = React.lazy(() => import('./AssistantPage'));
const TicketsPage = React.lazy(() => import('./TicketsPage'));
const ArchitecturePage = React.lazy(() => import('./ArchitecturePage'));
const StatsArchivePage = React.lazy(() => import('./StatsArchivePage'));
const IntroOnePagerPage = React.lazy(() => import('./IntroOnePagerPage'));

export interface IDexEventPlatformProps {
  context: WebPartContext;
}

// Innere Komponente, die den NavigationContext nutzen kann
// v12.7: Banner ganz oben — sichtbar wenn Admin Demo-Impersonation
// gestartet hat. Klick auf X beendet die Impersonation (löscht
// localStorage + reload).
function ImpersonationBanner(props: { currentPage?: string }): React.ReactElement | null {
  const [active, setActive] = React.useState<{ firstName?: string; surname?: string; email?: string; location?: string } | null>(null);
  React.useEffect(() => {
    try {
      const raw = window.localStorage.getItem('dex_demo_impersonation');
      if (raw) setActive(JSON.parse(raw));
    } catch { /* */ }
  }, []);
  if (!active) return null;
  const name = `${active.firstName || ''} ${active.surname || ''}`.trim() || active.email || '—';
  // v18.3: Im Admin-/Wizard-Kontext agiert der Demo-Account als Organizer
  // (read-only Demo-Event) — das wird im Banner sichtbar gemacht.
  const asOrganizer = props.currentPage === 'admin' || props.currentPage === 'edit-event' || props.currentPage === 'create-event';
  const exitImpersonation = (): void => {
    try { window.localStorage.removeItem('dex_demo_impersonation'); }
    catch { /* */ }
    window.location.reload();
  };
  return (
    <div style={{
      background: 'var(--dex-orange, #ed8b00)',
      color: '#fff',
      padding: '10px 24px',
      display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
      fontSize: '0.85rem', fontWeight: 600,
      borderBottom: '2px solid rgba(0,0,0,0.08)',
    }}>
      <span>
        DEMO-MODUS · agierst als <strong>{name}</strong>
        {' · Rolle '}<strong>{asOrganizer ? 'Organizer' : 'User'}</strong>
        {active.location ? <> · Standort <strong>{active.location}</strong></> : null}
      </span>
      <button
        type="button"
        onClick={exitImpersonation}
        style={{
          background: '#fff', color: 'var(--dex-orange, #ed8b00)',
          border: 'none', borderRadius: 999,
          padding: '4px 12px', fontWeight: 700, fontSize: '0.78rem',
          cursor: 'pointer', boxShadow: '0 2px 6px rgba(0,0,0,0.18)',
        }}
      >
        Demo-Modus beenden
      </button>
    </div>
  );
}

function AppContent(): React.ReactElement {
  const { currentPage, navigate } = useNavigation();
  const { isAdmin, isRolesLoading } = useRoles();
  const { markExpiredEventsAsCompleted, autoRepairProxyAccess, maybeSendWeeklyReport, maybeSendPostEventOrganizerMails, reconcileCounters, isEventsLoading, events, getKpiCache, recomputeEventKpiOnly } = useEvents();

  // v11.52: KPI-Boxen im Boot-Loader. Live-Zählung über alle Event-
  // Subsites war zu langsam (Counts kommen erst nach mehreren Sekunden) —
  // jetzt lesen wir einen gecachten Wert aus der _Config-Zeile von
  // DEX_EmailTemplates: EIN schneller REST-Call, sofort verfügbar.
  // Im Hintergrund läuft nach vollem Event-Load ein Refresh, der die
  // Cache-Werte aktualisiert — nächster Boot sieht frische Zahlen.
  const [kpiCache, setKpiCache] = React.useState<{ participants: number; events: number } | null>(null);
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const v = await getKpiCache();
        if (!cancelled && v) setKpiCache(v);
      } catch { /* */ }
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // v11.54: Self-Heal-Recompute nur durch Admin oder Organizer auslösen.
  // Normale User lesen den gecachten Wert (1 schneller GET) — kein
  // Recompute-Roundtrip, kein SP-Write. Bei 10k+ aktiven Usern reicht es
  // völlig, wenn die paar Dutzend Admins/Organizer den Cache periodisch
  // konvergieren lassen. Direkt-Increments bei Register/Cancel/Create/
  // Delete laufen weiterhin für alle Rollen — Cache bleibt nah dran.
  const kpiRefreshedRef = React.useRef(false);
  React.useEffect(() => {
    if (isEventsLoading || isRolesLoading) return;
    // v26.64: Neue, schlanke KPI-Architektur (Maintainer-Entscheidung):
    //  • „Events" = direkt aus DEX_Events (billig, exakt — IsFictive/Status/
    //    ParentEventId stehen dort). Nur diese Zahl berechnet ein Admin beim
    //    Boot neu — ein einziger Listen-Read, KEIN Subsite-Scan mehr.
    //  • „Teilnehmer" = der Live-Bump-Counter in _Config, den alle Rollen bei
    //    jeder An-/Abmeldung selbst ±1 fortschreiben (bumpKpiParticipants).
    //    Dieser Live-Wert ist FÜHREND und wird NICHT mehr durch einen teuren
    //    40-Subsite-Scan überschrieben (der zog den Boot in die Länge und war
    //    trotzdem nur ein einmal-pro-Session-Snapshot).
    // Normale User lesen nur den Cache; der Recompute läuft für Admins.
    if (!isAdmin) return;
    if (!events || events.length === 0) return; // erst wenn die App bereit ist
    if (kpiRefreshedRef.current) return;
    kpiRefreshedRef.current = true;
    const SESSION_KEY = 'dex-kpi-cache-refreshed';
    try {
      if (sessionStorage.getItem(SESSION_KEY) === '1') {
        // eslint-disable-next-line no-console
        console.log('[DEX KPI] Events-Zahl in dieser Browser-Session bereits neu berechnet — übersprungen.');
        return;
      }
    } catch { /* */ }
    // eslint-disable-next-line no-console
    console.log('[DEX KPI] Admin-Session — Events-Zahl aus DEX_Events neu berechnen (Teilnehmer bleibt der Live-Counter).');
    recomputeEventKpiOnly()
      .then(events2 => {
        if (events2 === null) return;
        setKpiCache(prev => ({ participants: prev?.participants ?? 0, events: events2 }));
        // eslint-disable-next-line no-console
        console.log(`[DEX KPI] Events-Zahl aktualisiert: ${events2} (in _Config gespeichert).`);
        try { sessionStorage.setItem(SESSION_KEY, '1'); } catch { /* */ }
      })
      .catch(() => { /* */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEventsLoading, isRolesLoading, isAdmin, events]);

  const kpiHostedEvents = kpiCache?.events ?? 0;
  const kpiParticipants = kpiCache?.participants ?? 0;

  // v7.5/7.6: Boot-Progress. Zeitbasiert + asymptotisch: über die ersten ~3-4
  // Sekunden steigt der Wert relativ schnell, danach immer langsamer Richtung
  // 95%. Sobald BEIDE Provider (Rollen + Events) fertig sind, springt der
  // Wert auf 100%. Dadurch ist der Übergang nicht mehr klobig phasenbasiert
  // (vorher 30 → 75 → 100), sondern fühlt sich wie eine echte Lade-Anzeige an.
  const [bootProgress, setBootProgress] = React.useState<number>(8);
  const bootLoadingRef = React.useRef({ roles: true, events: true });
  bootLoadingRef.current = { roles: isRolesLoading, events: isEventsLoading };
  // v29.32: Was gerade lädt — der Balken allein erklärt ein längeres Warten nicht.
  // v29.41: …und WIE WEIT es ist. Der Ladepfad meldet seine Abschnitte jetzt
  // selbst (utils/bootProgress); der Balken kriecht zwischen den Meldungen
  // asymptotisch auf das Ende des laufenden Abschnitts zu. Vorher war er rein
  // zeitgesteuert und stand nach vier Sekunden bei 92 %, egal wie weit die App
  // war — genau die Stelle, an der es „hängt".
  const [bootPhase, setBootPhase] = React.useState<'roles' | 'events' | 'done' | BootStage>('roles');
  const stageRef = React.useRef<{ floor: number; target: number }>({ floor: 8, target: 30 });
  React.useEffect(() => {
    const apply = (d: BootStageDetail): void => {
      // Nie zurückspringen: Ein später gemeldeter Abschnitt zieht die
      // Untergrenze nach oben, ein früherer lässt sie stehen.
      stageRef.current = {
        floor: Math.max(stageRef.current.floor, d.floor),
        target: Math.max(stageRef.current.target, d.target),
      };
      setBootPhase(d.stage);
      setBootProgress(prev => Math.max(prev, d.floor));
    };
    const last = lastBootStage();
    if (last) apply(last);
    const onStage = (e: Event): void => {
      const d = (e as CustomEvent).detail as BootStageDetail | undefined;
      if (d && typeof d.target === 'number') apply(d);
    };
    window.addEventListener(BOOT_STAGE_EVENT, onStage);
    return () => window.removeEventListener(BOOT_STAGE_EVENT, onStage);
  }, []);
  React.useEffect(() => {
    const id = setInterval(() => {
      const { roles, events } = bootLoadingRef.current;
      if (!roles && !events) {
        setBootPhase('done');
        setBootProgress(100);
        clearInterval(id);
        return;
      }
      // Sind die Events durch und fehlen nur noch die Rechte, sagt der Text das
      // auch — sonst steht dort „Anhänge…", während längst etwas anderes läuft.
      if (!events && roles) setBootPhase('roles');
      setBootProgress(prev => {
        const { target } = stageRef.current;
        if (prev >= target) return prev;
        // Ease-out zum Abschnitts-Ende: immer Bewegung, nie ein Sprung, und der
        // Rest bleibt für die echte Fertig-Meldung.
        const step = Math.max(0.06, (target - prev) * 0.022);
        return Math.min(target, prev + step);
      });
    }, 60);
    return () => clearInterval(id);
  }, []);
  const layoutRef = React.useRef<HTMLDivElement>(null);

  // v24.50: Beim Seitenwechsel immer nach oben scrollen — sonst „hängt" die
  // neue Seite an der Scroll-Position der vorigen. Deckt das App-Layout, das
  // Fenster UND den SharePoint-Canvas-Scrollcontainer ab (in SP scrollt nicht
  // window, sondern ein eigener Container).
  React.useEffect(() => {
    const toTop = (el: Element | null): void => { try { if (el && 'scrollTop' in el) (el as HTMLElement).scrollTop = 0; } catch { /* */ } };
    try { window.scrollTo(0, 0); } catch { /* */ }
    toTop(document.scrollingElement);
    toTop(document.documentElement);
    toTop(document.body);
    toTop(layoutRef.current);
    // Bekannte SharePoint-Scrollcontainer (Workbench / moderne Seite).
    ['[data-automation-id="contentScrollRegion"]', '.SPPageChrome-content', '#spPageCanvasContent', '[class*="contentScrollRegion"]']
      .forEach(sel => { try { document.querySelectorAll(sel).forEach(toTop); } catch { /* */ } });
  }, [currentPage]);

  // Deep-Link Handling: Wenn die Seite mit ?action=cancel&event=<eventNumber>
  // aufgerufen wird (z.B. aus einer Outlook-Decline-Reminder-Mail), direkt auf
  // My Events navigieren mit der eventId - MyEventsPage cancelt dann die
  // Registrierung automatisch.
  //
  // Damit der User nicht 5 Sekunden lang auf der LandingPage "hängt" bis die
  // Events geladen sind, zeigen wir statt der LandingPage einen Vollbild-
  // Lade-Spinner, sobald wir erkennen dass ein Cancel-Deep-Link aktiv ist.
  const isCancelDeepLink = React.useMemo(() => {
    try {
      const params = deepLinkParams();
      return params.get('action') === 'cancel' && !!params.get('event');
    } catch { return false; }
  }, []);

  // v18.33: Self-Check-in-Deep-Link (?action=selfcheckin&…). Wird per Handy-
  // Kamera gescannter QR-Code aufgerufen. Rendert direkt die Ergebnisseite,
  // unabhängig von currentPage — kein In-App-Scanner nötig.
  // v20.2 BUG-FIX: vorher ein eingefrorenes useMemo([]) — die Buttons
  // „Meine Events"/„Zur Startseite" auf der Ergebnisseite liefen damit ins
  // Leere (renderPage kurzschloss weiter auf die Ergebnisseite, solange
  // ?action=selfcheckin in der URL stand). Jetzt State + Leave-Handler, der
  // die Deep-Link-Parameter aus der URL entfernt und dann navigiert.
  const [isSelfCheckInDeepLink, setIsSelfCheckInDeepLink] = React.useState<boolean>(() => {
    try {
      return deepLinkParams().get('action') === 'selfcheckin';
    } catch { return false; }
  });
  const leaveSelfCheckIn = React.useCallback((page: 'start' | 'my-events'): void => {
    try {
      const sp = new URLSearchParams(window.location.search);
      ['action', 'token', 'event', 'code', 't'].forEach(k => sp.delete(k));
      const qs = sp.toString();
      window.history.replaceState({}, '', window.location.pathname + (qs ? `?${qs}` : '') + window.location.hash);
    } catch { /* URL-Cleanup best-effort */ }
    setIsSelfCheckInDeepLink(false);
    navigate(page);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // v9.43: Deep-Link Handling für ?action=event-created/event-updated&event=<id>.
  // Wird beim Klick auf 'Zur Übersicht' nach Event-Erstellung/-Update aufgerufen
  // (Hard-Reload). Zeigt eine grüne Erfolgs-Banner-Meldung und navigiert
  // automatisch in die Event-Liste. Die Action wird im useEffect-Block direkt
  // ausgewertet (s.u.) — kein extra useMemo nötig.
  const [successBanner, setSuccessBanner] = React.useState<{ title: string; type: 'create' | 'update' } | null>(null);

  // v9.43: Banner nach 8 Sekunden automatisch ausblenden — User hat genug Zeit zu lesen.
  React.useEffect(() => {
    if (!successBanner) return;
    const t = setTimeout(() => setSuccessBanner(null), 8000);
    return () => clearTimeout(t);
  }, [successBanner]);

  // v9.45: Soft-Refresh-Listener für Event-Submit-Erfolg. EventCreationPage
  // dispatcht diesen Event direkt nach Submit (createEvent oder updateEvent),
  // statt die App per window.location.href neu zu laden. Wir navigieren zur
  // Event-Liste und zeigen den grünen Erfolgs-Banner an — der Wizard unmountet
  // sauber, kein Page-Reload nötig.
  React.useEffect(() => {
    const onSubmitSuccess = (e: Event): void => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const detail = (e as any).detail as { title?: string; eventId?: string; type?: 'create' | 'update' } | undefined;
      if (!detail) return;
      setSuccessBanner({
        title: detail.title || 'Event',
        type: detail.type === 'update' ? 'update' : 'create',
      });
      // v17.4: Beim Update kehrt der User in das Organizer-Menü des Events
      // zurück (AdminPage mit dem soeben gespeicherten Event vorselektiert)
      // statt in die Event-Liste — weiterarbeiten ohne Such-Klick.
      // Beim Create bleibt der bisherige Pfad: Event-Liste, damit der neue
      // Event direkt in der Übersicht auftaucht.
      if (detail.type === 'update' && detail.eventId) {
        navigate('admin', detail.eventId);
      } else {
        navigate('register');
      }
    };
    window.addEventListener('dex-event-submit-success', onSubmitSuccess);
    return () => window.removeEventListener('dex-event-submit-success', onSubmitSuccess);
  }, [navigate]);

  const didHandleDeepLink = React.useRef(false);
  React.useEffect(() => {
    if (didHandleDeepLink.current) return;
    if (isEventsLoading) return;
    try {
      const params = deepLinkParams();
      const action = params.get('action');
      const eventParam = params.get('event');
      if (action === 'cancel' && eventParam) {
        didHandleDeepLink.current = true;
        const eventNumber = parseInt(eventParam, 10);
        if (!isNaN(eventNumber)) {
          const evt = events.find(e => e.eventNumber === eventNumber);
          if (evt) {
            navigate('my-events', evt.id, 'auto-cancel');
          } else {
            navigate('my-events');
          }
        }
      } else if (action === 'assistreq') {
        // v24.42: Deep-Link aus der Anforderungs-Mail. Die verwaltende Person
        // (Owner) landet in „Meine Events" — dort sieht sie die offene
        // Anforderung als Banner und kann sie nach Erledigung quittieren.
        didHandleDeepLink.current = true;
        navigate('my-events');
      } else if ((action === 'admin' || action === 'checkin') && eventParam) {
        // Deep-Link für Organizer/Admin direkt ins jeweilige Event springen —
        // z.B. ?action=admin&event=35 öffnet das Admin Center mit Event mit
        // SharePoint-Item-ID 35 vorselektiert (Teilnehmerliste, Edit, Mails,
        // Spalten fixen). ?action=checkin&event=35 springt direkt ins
        // Check-In-Tool. Der `event`-Parameter ist die SharePoint-Item-ID aus
        // der DEX_Events-Liste — identisch zum Wert, den die Quick-Action-
        // Kachel "Deep-Link kopieren" im Admin Center erzeugt. Fallback: wenn
        // der Wert numerisch ist und keine ID matched, versuchen wir noch das
        // alte Schema (eventNumber), damit alte Links nicht ins Leere laufen.
        didHandleDeepLink.current = true;
        const targetPage = action === 'checkin' ? 'check-in' : 'admin';
        let evt = events.find(e => e.id === eventParam);
        if (!evt) {
          const asNumber = parseInt(eventParam, 10);
          if (!isNaN(asNumber)) {
            evt = events.find(e => e.eventNumber === asNumber);
          }
        }
        if (evt) {
          navigate(targetPage, evt.id);
        } else {
          navigate(targetPage);
        }
      } else if (action === 'register' && eventParam) {
        // v26.67: Deep-Link aus der „Anmeldung abschließen"-Erinnerung — direkt
        // zur Anmeldeseite mit vorgewähltem Event (SharePoint-Item-ID; Fallback
        // eventNumber). Die Person kann dort ihre Programmpunkte auswählen.
        didHandleDeepLink.current = true;
        let evt = events.find(e => e.id === eventParam);
        if (!evt) { const n = parseInt(eventParam, 10); if (!isNaN(n)) evt = events.find(e => e.eventNumber === n); }
        navigate('register', evt ? evt.id : undefined);
      } else if (action === 'event-created' || action === 'event-updated') {
        // v9.43: Deep-Link nach Event-Erstellung/-Update. Wird vom EventCreationPage-
        // Success-Button aufgerufen mit window.location.href = '...?action=event-
        // created&event=<id>'. Hard-Reload räumt den SP-Permission-Cache auf
        // (frische Subsite kann sonst 400/404 werfen) und der App-Bootstrap
        // navigiert dann sauber zur Event-Liste mit grüner Erfolgs-Banner.
        didHandleDeepLink.current = true;
        const evt = events.find(e => e.id === eventParam);
        setSuccessBanner({
          title: evt?.title || 'Event',
          type: action === 'event-updated' ? 'update' : 'create',
        });
        navigate('register');
        // URL aufräumen, damit ein zweiter Reload nicht erneut den Banner triggert
        try {
          const cleanUrl = window.location.pathname + (window.location.hash || '');
          window.history.replaceState({}, '', cleanUrl);
        } catch { /* */ }
      } else if (action === 'selfcheckin-display' && eventParam) {
        // v18.33: Deep-Link auf die rotierende Live-Check-in-Anzeige (Organizer).
        // event = SharePoint-Item-ID oder Event-Nummer.
        didHandleDeepLink.current = true;
        let evt = events.find(e => e.id === eventParam);
        if (!evt) {
          const asNumber = parseInt(eventParam, 10);
          if (!isNaN(asNumber)) evt = events.find(e => e.eventNumber === asNumber);
        }
        navigate('self-checkin-display', evt ? evt.id : undefined);
      } else if (action === 'manual') {
        // v6.23: Deep-Link ins Handbuch. Optionaler `section`-Query-Parameter
        // steuert die initial angezeigte Sektion (wird in ManualPage direkt aus
        // der URL gelesen, damit kein zusätzlicher Navigation-State nötig ist).
        // Beispiel: ?action=manual&section=check-in
        didHandleDeepLink.current = true;
        navigate('manual');
      } else if (action === 'tickets') {
        // v26: Deep-Link aus der Power-User-Benachrichtigung → Tickets-Inbox.
        // Die optionale ?id=<ticketId> liest die TicketsPage selbst aus der URL
        // und öffnet das Ticket direkt zum Beantworten.
        didHandleDeepLink.current = true;
        navigate('tickets');
      } else if (action === 'ask') {
        // v26.7: Deep-Link aus der Antwort-Mail an den Fragesteller → App öffnen
        // und das „Hast du Fragen?"-Fenster auf „Deine Fragen" zeigen.
        didHandleDeepLink.current = true;
        navigate('start');
        try { window.setTimeout(() => { window.dispatchEvent(new CustomEvent('dex-open-questions')); }, 500); } catch { /* */ }
      }
    } catch { /* URL-Parsing fehlgeschlagen, ignorieren */ }
  }, [isEventsLoading, events]);

  // Admin-Cleanup einmal pro App-Session: Abgelaufene Events (EndDate < jetzt)
  // mit Status='Active' werden automatisch auf 'Completed' gesetzt.
  const didExpireCheck = React.useRef(false);
  React.useEffect(() => {
    if (didExpireCheck.current) return;
    if (!isAdmin) return;
    if (isEventsLoading) return; // warten bis Events geladen sind
    didExpireCheck.current = true;
    markExpiredEventsAsCompleted().catch(err => console.warn('[DEX] expire check failed:', err));
  }, [isAdmin, isEventsLoading]);

  // v22.42: Sichtbarkeits-Auto-Fix beim Admin-Start. Repariert im Hintergrund
  // die Zeilen-Autoren von Fremd-Anmeldungen (Assistenz/Organizer-für-andere),
  // damit der betroffene Teilnehmer seine Anmeldung in „Meine Events" sieht
  // und sich selbst abmelden kann. Gedrosselt 1×/24h (localStorage), läuft
  // verzögert + best-effort, blockiert den App-Start nicht.
  const didAutoAccessFix = React.useRef(false);
  React.useEffect(() => {
    if (didAutoAccessFix.current) return;
    if (!isAdmin) return;
    if (isEventsLoading) return;
    if (!events || events.length === 0) return;
    didAutoAccessFix.current = true;
    // Etwas verzögert, damit der erste Render/Boot Vorrang hat.
    const t = window.setTimeout(() => {
      autoRepairProxyAccess().catch(err => console.warn('[DEX] auto access-fix failed:', err));
    }, 4000);
    return () => window.clearTimeout(t);
  }, [isAdmin, isEventsLoading, events]);

  // v24.74: Sitzplatz-Counter beim Admin-Start frischziehen (Admin hat
  // Vollzugriff → echte Aktiv-/Warteliste-Zahlen). So stimmen die für ALLE
  // lesbaren Counter (SeatsTaken/WaitlistTaken), bevor normale Teilnehmer das
  // Anmeldeformular öffnen. Gedrosselt 1×/6h via localStorage, verzögert,
  // best-effort (blockiert den Boot nicht).
  const didReconcileCounters = React.useRef(false);
  React.useEffect(() => {
    if (didReconcileCounters.current) return;
    if (!isAdmin) return;
    if (isEventsLoading) return;
    if (!events || events.length === 0) return;
    let due = true;
    try {
      const last = parseInt(window.localStorage.getItem('dex_counter_reconcile_lastrun') || '0', 10);
      if (last && Date.now() - last < 6 * 60 * 60 * 1000) due = false;
    } catch { /* */ }
    if (!due) return;
    didReconcileCounters.current = true;
    try { window.localStorage.setItem('dex_counter_reconcile_lastrun', String(Date.now())); } catch { /* */ }
    const t = window.setTimeout(() => {
      reconcileCounters().catch(err => console.warn('[DEX] counter reconcile failed:', err));
    }, 6000);
    return () => window.clearTimeout(t);
  }, [isAdmin, isEventsLoading, events]);

  // v23.8: Wöchentlicher Admin-Bericht. Beim Admin-Start (einmal pro App-
  // Session) wird geprüft, ob seit dem letzten Bericht ≥7 Tage vergangen sind —
  // die 7-Tage-Logik + der Doppelversand-Schutz liegen serverseitig in der Liste
  // DEX_WeeklyReports (plus ein localStorage-Backstop in maybeSendWeeklyReport).
  const didWeeklyReport = React.useRef(false);
  React.useEffect(() => {
    if (didWeeklyReport.current) return;
    if (!isAdmin) return;
    if (isEventsLoading) return;
    if (!events || events.length === 0) return;
    didWeeklyReport.current = true;
    // v23.13: Der frühere 12h-localStorage-Throttle ist ENTFERNT — er wurde bei
    // jedem Boot gesetzt und drosselte (bei häufigem Öffnen) den Check dauerhaft,
    // sodass der Bericht nie lief. Der eigentliche Schutz vor Mehrfachversand ist
    // ohnehin serverseitig (DEX_WeeklyReports ≥7 Tage) + localStorage-Backstop.
    // Deutlich verzögert — Boot/erste Interaktion haben Vorrang.
    // v23.12 BUG-FIX: KEIN clearTimeout-Cleanup zurückgeben (siehe unten) — sonst
    // bricht ein Re-Render (events-Änderung) den Timer ab und er feuert nie.
    window.setTimeout(() => {
      maybeSendWeeklyReport().catch(err => console.warn('[DEX] weekly report failed:', err));
    }, 8000);
  }, [isAdmin, isEventsLoading, events]);

  // v24.2: Post-Event-Mail an den Organizer — beim App-Öffnen nach dem
  // Event-Tag. Für JEDEN (nicht nur Admins); die Funktion filtert intern auf
  // die eigenen, abgelaufenen Events und drosselt 1×/Event/Organizer.
  const didPostEventMail = React.useRef(false);
  React.useEffect(() => {
    if (didPostEventMail.current) return;
    if (isEventsLoading) return;
    if (!events || events.length === 0) return;
    didPostEventMail.current = true;
    window.setTimeout(() => {
      maybeSendPostEventOrganizerMails().catch(err => console.warn('[DEX] post-event mail failed:', err));
    }, 9000);
  }, [isEventsLoading, events]);

  // Dynamische Höhe + SharePoint-Scroll unterdrücken
  React.useEffect(() => {
    // Style-Tag injizieren mit !important - SP kann das nicht überschreiben
    const styleEl = document.createElement('style');
    styleEl.id = 'dex-no-scroll';
    styleEl.textContent = `
      html, body {
        overflow: hidden !important;
        overflow-y: hidden !important;
        height: 100vh !important;
        max-height: 100vh !important;
      }
      .SPPageChrome,
      .sp-App-root,
      .CanvasZone,
      .CanvasSection,
      .ControlZone,
      .ControlZone--control,
      [class*="canvasWrapper"],
      [class*="webPartContainer"],
      [data-automation-id="CanvasControl"],
      [data-automation-id="CanvasSection"],
      [data-automation-id="CanvasZone"],
      #spPageCanvasContent,
      #workbenchPageContent,
      .SPCanvas-canvas,
      .Canvas-slideUpIn,
      div[class*="pageContent"],
      div[class*="mainContent"] {
        overflow: hidden !important;
        overflow-y: hidden !important;
      }
      /* SharePoint padding/margin unter dem WebPart entfernen */
      .CanvasZone,
      [data-automation-id="CanvasZone"] {
        padding-bottom: 0 !important;
        margin-bottom: 0 !important;
      }
      .CanvasSection,
      [data-automation-id="CanvasSection"] {
        padding-bottom: 0 !important;
        margin-bottom: 0 !important;
      }
      /* Globale Keyframes für Loading-Spinner (werden von EventListPage/MyEventsPage genutzt,
         unabhängig davon ob die LandingPage vorher gerendert wurde) */
      @keyframes dexOrbSpin { to { transform: rotate(360deg); } }
    `;
    document.head.appendChild(styleEl);

    function setHeight(): void {
      if (layoutRef.current) {
        const rect = layoutRef.current.getBoundingClientRect();
        const available = window.innerHeight - rect.top;
        layoutRef.current.style.height = `${Math.max(available, 400)}px`;
      }
    }

    setHeight();
    window.addEventListener('resize', setHeight);
    const timer = setTimeout(setHeight, 500);
    const timer2 = setTimeout(setHeight, 1500);

    return () => {
      window.removeEventListener('resize', setHeight);
      clearTimeout(timer);
      clearTimeout(timer2);
      const el = document.getElementById('dex-no-scroll');
      if (el) el.remove();
    };
  }, []);

  // Seitenauswahl basierend auf dem aktuellen State
  const renderPage = (): React.ReactElement => {
    // v6.26: Boot-Loader. Auf der LandingPage entscheidet die Bubble "Jetzt
    // einchecken" davon, ob der User Admin / Organizer / QR-Scanner von einem
    // Event ist — das wissen wir erst, sobald BEIDE Context-Provider fertig
    // geladen haben (DEX_Roles + DEX_Events). Vorher einen Vollbild-Spinner
    // zeigen, damit der User nicht kurz die LandingPage ohne Bubble sieht,
    // bevor die Bubble nachrutscht.
    if (currentPage === 'landing' && !isCancelDeepLink && !isSelfCheckInDeepLink && (isEventsLoading || isRolesLoading)) {
      // Denselben Orb-Look wie auf der Landing-Page, damit der Übergang
      // Boot-Loader → LandingPage flüssig wirkt. Keyframes injizieren wir
      // hier selbst, falls LandingPage noch nicht gemountet war.
      if (typeof document !== 'undefined' && !document.getElementById('dex-orb-keyframes')) {
        const style = document.createElement('style');
        style.id = 'dex-orb-keyframes';
        style.textContent = '@keyframes dexOrbSpin { to { transform: rotate(360deg); } }';
        document.head.appendChild(style);
      }
      // v6.29: Indeterminate Progress-Bar-Keyframes. Läuft endlos von links
      // nach rechts durch eine 30%-breite Farbzone — hübscher als ein Text.
      if (typeof document !== 'undefined' && !document.getElementById('dex-progress-keyframes')) {
        const style = document.createElement('style');
        style.id = 'dex-progress-keyframes';
        style.textContent = '@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }';
        document.head.appendChild(style);
      }
      // v29.41: Dezenter Schimmer auf dem gefüllten Teil. Zwischen zwei
      // Abschnitts-Meldungen bewegt sich der Balken nur noch um Bruchteile
      // eines Prozents — ohne dieses Lebenszeichen liest sich das als Hänger.
      // Bewusst schwach (weiß auf Grün, 45 %) und langsam, nicht als Blinken.
      if (typeof document !== 'undefined' && !document.getElementById('dex-progress-pulse-keyframes')) {
        const style = document.createElement('style');
        style.id = 'dex-progress-pulse-keyframes';
        style.textContent = '@keyframes dexProgressShimmer { 0% { transform: translateX(-120%); } 100% { transform: translateX(320%); } }';
        document.head.appendChild(style);
      }
      return (
        <div className="landing" style={{ position: 'relative' }}>
          <div className="landing__hero">
            <div className="landing__card" style={{ position: 'relative', textAlign: 'center' }}>
              {/* v28.34: Der Boot-Loader nutzte dieselbe .landing__orb-Huelle wie
                  die Landing Page — seit v28.33 zeichnet dort aber das DexLogo-
                  Canvas statt des Farbverlaufs, die Huelle allein blieb also leer
                  (große weisse Flaeche über „Welcome to DEX"). Jetzt rendert der
                  Boot-Loader dasselbe Logo, der Übergang bleibt fluessig. */}
              <div className="landing__orb">
                <DexLogo title="DEX" motion="oscillate" style={{ width: '100%' }} />
              </div>
              <div className="landing__text">
                <h1 style={{ lineHeight: 1.25 }}>
                  Welcome to <strong style={{ whiteSpace: 'nowrap' }}>DEX.</strong>
                </h1>
                <p style={{ color: 'var(--dex-gray-500)', marginTop: 12, fontSize: '0.95rem' }}>
                  Just a moment…
                </p>
              </div>
              {/* Determinate-Progress-Bar (v7.5). Die Phasen Rollen-Load
                  und Events-Load werden auf Prozentwerte gemappt; der Wert
                  zählt sanft in Richtung des aktuellen Phasen-Targets,
                  damit der User auf einen Blick sieht "wie weit die App ist". */}
              <div style={{
                width: 'min(320px, 80%)',
                marginTop: 8, display: 'flex', flexDirection: 'column', gap: 6,
              }}>
                <div style={{
                  height: 6, borderRadius: 3,
                  background: 'var(--dex-gray-200, #e5e5e5)',
                  overflow: 'hidden', position: 'relative',
                }}>
                  <div style={{
                    position: 'absolute', top: 0, bottom: 0, left: 0,
                    width: `${bootProgress}%`,
                    background: 'var(--dex-green, #86bc25)',
                    transition: 'width 240ms ease-out',
                    borderRadius: 3,
                    overflow: 'hidden',
                  }}>
                    {bootPhase !== 'done' && (
                      <div style={{
                        position: 'absolute', top: 0, bottom: 0, left: 0, width: '35%',
                        background: 'linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.45) 50%, rgba(255,255,255,0) 100%)',
                        animation: 'dexProgressShimmer 2.1s ease-in-out infinite',
                      }} />
                    )}
                  </div>
                </div>
                {/* v29.32: Links steht, worauf gerade gewartet wird — ein
                    Balken, der bei einem langsamen Tenant lange fast voll
                    steht, erklärt sich sonst nicht. */}
                <div style={{
                  fontSize: '0.78rem', color: 'var(--dex-gray-500)',
                  display: 'flex', justifyContent: 'space-between', gap: 8,
                  fontVariantNumeric: 'tabular-nums',
                }}>
                  {/* Der Loader läuft VOR der Sprachwahl des Nutzers — deshalb
                      die Browsersprache statt des App-Locales. */}
                  <span>
                    {(() => {
                      const de = typeof navigator !== 'undefined' && (navigator.language || '').toLowerCase().indexOf('de') === 0;
                      // v29.41: Die Abschnitte des Ladepfads beim Namen nennen —
                      // „Events werden geladen" stand vorher über allem, was
                      // nach der Rechteprüfung kam (inklusive der langen
                      // Teilnehmerzahlen- und Anhang-Runden).
                      // v29.47: Die Beschriftungen sagen jetzt, was der
                      // NUTZER davon hat — „Datenmodell wird geprüft" war die
                      // Innensicht des Programms und für Teilnehmer nichts als
                      // ein technisches Rätsel.
                      switch (bootPhase) {
                        case 'roles': return de ? 'Deine Events werden gesucht…' : 'Looking up your events…';
                        case 'schema': return de ? 'Einen Moment noch…' : 'Just a moment…';
                        case 'logos': return de ? 'Fast fertig…' : 'Almost there…';
                        case 'events': return de ? 'Events werden geladen…' : 'Loading events…';
                        case 'mapping': return de ? 'Events werden vorbereitet…' : 'Preparing events…';
                        case 'counts': return de ? 'Freie Plätze werden geprüft…' : 'Checking available seats…';
                        case 'documents': return de ? 'Fast fertig…' : 'Almost there…';
                        default: return de ? 'Fertig' : 'Done';
                      }
                    })()}
                  </span>
                  <span>{Math.round(bootProgress)} %</span>
                </div>
              </div>
              {/* v11.48/v11.50/v11.55: KPI-Boxen im Boot-Loader.
                  v12.2/v12.4: Container breiter (520 statt 320), Caption
                  größer (0.95 statt 0.78rem), Block weiter nach oben
                  (marginTop 8 statt 24). */}
              <div style={{ width: 'min(520px, 90%)', marginTop: 8 }}>
                <div style={{
                  fontSize: '0.95rem', color: 'var(--dex-gray-600)',
                  textAlign: 'center', marginBottom: 12, fontStyle: 'italic',
                  fontWeight: 500,
                }}>
                  So far used for…
                </div>
                <KpiRow
                  locale="en"
                  eventsLoading={kpiCache === null}
                  participantsLoading={kpiCache === null}
                  events={kpiHostedEvents}
                  participants={kpiParticipants}
                />
              </div>
            </div>
          </div>
        </div>
      );
    }
    // Deep-Link Cancel aktiv? Dann Lade-Spinner statt LandingPage zeigen,
    // solange wir noch nicht zu MyEventsPage navigiert sind.
    if (isCancelDeepLink && currentPage === 'landing') {
      return (
        <div style={{
          display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
          minHeight: 400, padding: 48, textAlign: 'center',
        }}>
          <div style={{
            width: 56, height: 56, borderRadius: '50%',
            border: '4px solid var(--dex-gray-200, #e5e5e5)',
            borderTopColor: 'var(--dex-green, #86bc25)',
            animation: 'dexOrbSpin 0.8s linear infinite',
            marginBottom: 20,
          }} />
          <div style={{ fontSize: 16, fontWeight: 600, color: 'var(--dex-gray-800, #333)' }}>
            Loading your registration…
          </div>
          <div style={{ fontSize: 13, color: 'var(--dex-gray-500, #888)', marginTop: 6 }}>
            Cancelling registration, please wait.
          </div>
        </div>
      );
    }
    // v18.33: Self-Check-in-Deep-Link rendert direkt die Ergebnisseite.
    // v20.2: onLeave räumt die URL-Parameter weg und gibt die Navigation frei.
    if (isSelfCheckInDeepLink) {
      return <SelfCheckInPage onLeave={leaveSelfCheckIn} />;
    }
    switch (currentPage) {
      case 'landing':
        return <LandingPage />;
      case 'start':
        return <StartPage />;
      case 'register':
        return <EventListPage />;
      case 'registration':
        // v28.8 (Refresh-Fix): Die RegistrationPage darf beim Boot NICHT
        // vor dem Events-Load mounten. Ihr „Event nicht gefunden"-Ausstieg
        // (`if (!event) return`) liegt mitten im Komponentenkörper — mountet
        // die Seite ohne Event (weniger Hooks) und rendert nach dem Laden
        // MIT Event (mehr Hooks), wirft React „Rendered more hooks than
        // during the previous render" → weißer Screen. Beim Refresh-Restore
        // (v27.12) war genau das der Fall. Solange die Events laden, kommt
        // deshalb HIER der Spinner; die Seite mountet erst mit fertigen Daten.
        if (isEventsLoading) {
          return (
            <div className="page-container text-center">
              <div style={{ padding: 48 }}>
                <svg width={48} height={48} viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block', margin: '0 auto 16px' }}>
                  <circle cx={24} cy={24} r={20} fill="none" stroke="rgba(134,188,37,0.20)" strokeWidth={4} />
                  <path d="M 24 4 A 20 20 0 0 1 44 24" fill="none" stroke="#86bc25" strokeWidth={4} strokeLinecap="round">
                    <animateTransform attributeName="transform" type="rotate" from="0 24 24" to="360 24 24" dur="1s" repeatCount="indefinite" />
                  </path>
                </svg>
                <p style={{ color: 'var(--dex-gray-400)' }}>Event wird geladen …</p>
              </div>
            </div>
          );
        }
        return <RegistrationPage />;
      case 'my-events':
        return <MyEventsPage />;
      case 'assistant':
        return <AssistantPage />;
      case 'create-event':
      case 'edit-event':
        // v28.8: gleicher Boot-Guard wie bei 'registration' — der Wizard
        // initialisiert seinen State EINMALIG aus `editEvent` (useState-
        // Initializer). Mountet er beim Refresh-Restore vor dem Events-Load,
        // ist editEvent noch null → der Organizer sähe einen LEEREN Wizard,
        // der beim Speichern ein neues Event anlegen würde.
        if (isEventsLoading) {
          return (
            <div className="page-container text-center">
              <div style={{ padding: 48 }}>
                <svg width={48} height={48} viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block', margin: '0 auto 16px' }}>
                  <circle cx={24} cy={24} r={20} fill="none" stroke="rgba(134,188,37,0.20)" strokeWidth={4} />
                  <path d="M 24 4 A 20 20 0 0 1 44 24" fill="none" stroke="#86bc25" strokeWidth={4} strokeLinecap="round">
                    <animateTransform attributeName="transform" type="rotate" from="0 24 24" to="360 24 24" dur="1s" repeatCount="indefinite" />
                  </path>
                </svg>
                <p style={{ color: 'var(--dex-gray-400)' }}>Event wird geladen …</p>
              </div>
            </div>
          );
        }
        return <EventCreationPage />;
      case 'settings':
        return <SettingsPage />;
      case 'admin':
        return <AdminPage />;
      case 'admin-hub':
        return <AdminHubPage />;
      case 'email-templates':
        return <EmailTemplatesPage />;
      case 'profile':
        return <ProfilePage />;
      case 'role-matrix':
        return <RoleMatrixPage />;
      case 'participants':
        return <ParticipantsPage />;
      case 'flowcharts':
        return <FlowchartPage />;
      case 'check-in':
        return <CheckInPage />;
      case 'self-checkin-display':
        return <SelfCheckInDisplayPage />;
      case 'manual':
        return <ManualPage />;
      case 'tickets':
        return <TicketsPage />;
      case 'architecture':
        return <ArchitecturePage />;
      case 'stats-archive':
        return <StatsArchivePage />;
      case 'intro-onepager':
        return <IntroOnePagerPage />;
      default:
        return <LandingPage />;
    }
  };

  // v6.29: Während der Boot-Loader läuft, Header verstecken. Sonst würde
  // schon die "Jetzt einchecken"-Bubble / QR-Icon blinken bevor der eigentliche
  // Welcome-Screen sichtbar ist.
  const isBootLoading = currentPage === 'landing' && !isCancelDeepLink && !isSelfCheckInDeepLink && (isEventsLoading || isRolesLoading);
  // v9.26: Page-ID jetzt im Header-Avatar-Popup statt unten links.
  return (
    <div className="app-layout" ref={layoutRef}>
      {!isBootLoading && <Header />}
      <ImpersonationBanner currentPage={currentPage} />
      {/* v23.37: Admin-Hinweis auf offene „Organizer werden"-Anträge (+ Deep-Link).
          v24.19: kein eigener Padding-Wrapper mehr — der erzeugte sonst einen
          weißen 12px-Streifen unter dem Header, auch wenn der Banner nichts
          anzeigt (er gibt null zurück, wenn keine Anträge offen sind). Den
          Abstand trägt jetzt der Banner selbst, nur wenn er wirklich rendert. */}
      {!isBootLoading && <OrganizerRequestsBanner />}
      {/* v26.59: grantaccess-Deep-Link aus der „SharePoint-Zugriff benötigt"-Mail
          — vergibt als Admin direkt Leserechte und zeigt das Ergebnis als Modal. */}
      {!isBootLoading && <GrantAccessHandler />}
      {/* v26.73: downloadinvite-Deep-Link aus der externen Instruktions-Mail —
          lädt den an der Teilnehmer-Zeile abgelegten .eml-Entwurf per Button. */}
      {!isBootLoading && <InviteDownloadHandler />}

      {successBanner && (
        <div
          role="status"
          style={{
            position: 'fixed', top: 16, left: '50%', transform: 'translateX(-50%)',
            zIndex: 2500, maxWidth: 'calc(100vw - 32px)',
            padding: '14px 22px', borderRadius: 12,
            background: 'linear-gradient(135deg, #86bc25, #6b9a1e)',
            color: '#fff', boxShadow: '0 8px 24px rgba(107,154,30,0.35)',
            display: 'flex', alignItems: 'center', gap: 14,
            fontSize: '0.95rem', fontWeight: 500,
            animation: 'dexBannerSlideIn 0.4s ease-out',
          }}
        >
          <span style={{ fontSize: '1.4rem', lineHeight: 1 }}>✓</span>
          <span style={{ lineHeight: 1.4 }}>
            <strong>{successBanner.type === 'update' ? 'Event aktualisiert!' : 'Event erstellt!'}</strong>
            {' '}
            <span style={{ opacity: 0.95 }}>{'„'}{successBanner.title}{'“ '}{successBanner.type === 'update' ? 'wurde gespeichert' : 'wurde angelegt'} — {successBanner.type === 'create' ? 'taucht jetzt in der Eventliste auf' : 'die Änderungen sind live'}.</span>
          </span>
          <button
            type="button"
            onClick={() => setSuccessBanner(null)}
            aria-label="Schließen"
            style={{
              background: 'rgba(255,255,255,0.2)', border: 'none', color: '#fff',
              width: 26, height: 26, borderRadius: '50%', cursor: 'pointer',
              fontSize: '1rem', lineHeight: 1, padding: 0,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}
          >×</button>
          <style>{`@keyframes dexBannerSlideIn { from { opacity: 0; transform: translate(-50%, -20px); } to { opacity: 1; transform: translate(-50%, 0); } }`}</style>
        </div>
      )}
      <main className="main-content" style={{ position: 'relative', display: 'flex', flexDirection: 'column' }}>
        {/* v20.0: Suspense-Grenze für die lazy geladenen Sekundär-Seiten-Chunks.
            Sprachneutraler Fallback — der Chunk-Load dauert typisch < 0,5 s. */}
        <React.Suspense fallback={<div style={{ padding: 64, textAlign: 'center', color: 'var(--dex-gray-400, #999)', fontSize: '1.2rem', letterSpacing: 4 }}>…</div>}>
          {renderPage()}
        </React.Suspense>
      </main>
    </div>
  );
}

export default function DexEventPlatform(props: IDexEventPlatformProps): React.ReactElement {
  // Context global verfügbar machen für ProfilePage
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__dexSpfxContext = props.context;

  return (
    <div className={styles.dexApp}>
      <LanguageProvider>
        <DialogProvider>
          <UserProvider context={props.context}>
            <RoleProvider context={props.context}>
              <NavigationProvider>
                <EventProvider context={props.context}>
                  <TicketProvider context={props.context}>
                    <TutorialProvider>
                      <AppContent />
                    </TutorialProvider>
                  </TicketProvider>
                </EventProvider>
              </NavigationProvider>
            </RoleProvider>
          </UserProvider>
        </DialogProvider>
      </LanguageProvider>
    </div>
  );
}
