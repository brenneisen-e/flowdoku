/**
 * DEX Event Experience Platform - Hauptkomponente (SPFx)
 *
 * Wrapper fuer Navigation und Event Context.
 * Rendert je nach currentPage die passende Unterseite.
 *
 * Autor: Eike Brenneisen
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './DexEventPlatform.module.scss';
import { NavigationProvider, useNavigation } from '../context/NavigationContext';
import { LanguageProvider } from '../context/LanguageContext';
import { EventProvider, useEvents } from '../context/EventContext';
import { UserProvider } from '../context/UserContext';
import { RoleProvider, useRoles } from '../context/RoleContext';
import Header from './Header';
import LandingPage from './LandingPage';
import StartPage from './StartPage';
import EventListPage from './EventListPage';
import RegistrationPage from './RegistrationPage';
import MyEventsPage from './MyEventsPage';
import ProfilePage from './ProfilePage';
import SelfCheckInPage from './SelfCheckInPage';
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
const CheckInPage = React.lazy(() => import('./CheckInPage'));
const SelfCheckInDisplayPage = React.lazy(() => import('./SelfCheckInDisplayPage'));
const ManualPage = React.lazy(() => import('./manual/ManualPage'));

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
  const { isAdmin, isOrganizer, isRolesLoading } = useRoles();
  const { markExpiredEventsAsCompleted, isEventsLoading, events, getKpiCache, updateKpiCache } = useEvents();

  // v11.52: KPI-Boxen im Boot-Loader. Live-Zaehlung ueber alle Event-
  // Subsites war zu langsam (Counts kommen erst nach mehreren Sekunden) —
  // jetzt lesen wir einen gecachten Wert aus der _Config-Zeile von
  // DEX_EmailTemplates: EIN schneller REST-Call, sofort verfuegbar.
  // Im Hintergrund laeuft nach vollem Event-Load ein Refresh, der die
  // Cache-Werte aktualisiert — naechster Boot sieht frische Zahlen.
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

  // v11.54: Self-Heal-Recompute nur durch Admin oder Organizer ausloesen.
  // Normale User lesen den gecachten Wert (1 schneller GET) — kein
  // Recompute-Roundtrip, kein SP-Write. Bei 10k+ aktiven Usern reicht es
  // voellig, wenn die paar Dutzend Admins/Organizer den Cache periodisch
  // konvergieren lassen. Direkt-Increments bei Register/Cancel/Create/
  // Delete laufen weiterhin fuer alle Rollen — Cache bleibt nah dran.
  const kpiRefreshedRef = React.useRef(false);
  React.useEffect(() => {
    if (isEventsLoading || isRolesLoading) return;
    if (!isAdmin && !isOrganizer) return;
    if (!events || events.length === 0) return;
    if (kpiRefreshedRef.current) return;
    kpiRefreshedRef.current = true;
    const SESSION_KEY = 'dex-kpi-cache-refreshed';
    let alreadyDoneThisSession = false;
    try { alreadyDoneThisSession = sessionStorage.getItem(SESSION_KEY) === '1'; } catch { /* */ }
    const hosted = events.filter(e => !e.isFictive && e.status !== 'Cancelled');
    const eventsCount = hosted.length;
    const participantsCount = hosted.reduce((s, e) => s + (e.currentParticipants || 0), 0);
    setKpiCache({ participants: participantsCount, events: eventsCount });
    if (alreadyDoneThisSession) return;
    updateKpiCache({ participants: participantsCount, events: eventsCount })
      .then(ok => {
        if (ok) { try { sessionStorage.setItem(SESSION_KEY, '1'); } catch { /* */ } }
      })
      .catch(() => { /* */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEventsLoading, isRolesLoading, isAdmin, isOrganizer, events]);

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
  React.useEffect(() => {
    const start = Date.now();
    const id = setInterval(() => {
      if (!bootLoadingRef.current.roles && !bootLoadingRef.current.events) {
        setBootProgress(100);
        clearInterval(id);
        return;
      }
      const elapsed = Date.now() - start;
      // v7.10/v11.79: Linearer Fortschritt — nach 4 Sekunden bei 100%
      // (vorher 10 s, war seit dem App-Boot-Speedup auf ~1.6 s zu langsam).
      // Wenn das Laden frueher fertig ist, setzt der Done-Branch oben sofort
      // auf 100% und stoppt das Intervall. Sollte Roles/Events laenger als
      // 4 s brauchen, friert der Balken bei 99% ein, bis das Done-Signal
      // kommt — sonst stuende der Balken auf 100%, obwohl die App noch
      // nicht bereit ist (das wirkt verwirrend / "haengend").
      const target = Math.min(99, Math.round((elapsed / 4000) * 100));
      setBootProgress(prev => Math.max(prev, target));
    }, 60);
    return () => clearInterval(id);
  }, []);
  const layoutRef = React.useRef<HTMLDivElement>(null);

  // Deep-Link Handling: Wenn die Seite mit ?action=cancel&event=<eventNumber>
  // aufgerufen wird (z.B. aus einer Outlook-Decline-Reminder-Mail), direkt auf
  // My Events navigieren mit der eventId - MyEventsPage cancelt dann die
  // Registrierung automatisch.
  //
  // Damit der User nicht 5 Sekunden lang auf der LandingPage "haengt" bis die
  // Events geladen sind, zeigen wir statt der LandingPage einen Vollbild-
  // Lade-Spinner, sobald wir erkennen dass ein Cancel-Deep-Link aktiv ist.
  const isCancelDeepLink = React.useMemo(() => {
    try {
      const params = new URLSearchParams(window.location.search);
      return params.get('action') === 'cancel' && !!params.get('event');
    } catch { return false; }
  }, []);

  // v18.33: Self-Check-in-Deep-Link (?action=selfcheckin&…). Wird per Handy-
  // Kamera gescannter QR-Code aufgerufen. Rendert direkt die Ergebnisseite,
  // unabhängig von currentPage — kein In-App-Scanner nötig.
  const isSelfCheckInDeepLink = React.useMemo(() => {
    try {
      return new URLSearchParams(window.location.search).get('action') === 'selfcheckin';
    } catch { return false; }
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
      // v17.4: Beim Update kehrt der User in das Organizer-Menue des Events
      // zurueck (AdminPage mit dem soeben gespeicherten Event vorselektiert)
      // statt in die Event-Liste — weiterarbeiten ohne Such-Klick.
      // Beim Create bleibt der bisherige Pfad: Event-Liste, damit der neue
      // Event direkt in der Uebersicht auftaucht.
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
      const params = new URLSearchParams(window.location.search);
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
      /* Globale Keyframes fuer Loading-Spinner (werden von EventListPage/MyEventsPage genutzt,
         unabhaengig davon ob die LandingPage vorher gerendert wurde) */
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
      return (
        <div className="landing" style={{ position: 'relative' }}>
          <div className="landing__hero">
            <div className="landing__card" style={{ position: 'relative', textAlign: 'center' }}>
              <div className="landing__orb">
                <div className="landing__orb-inner" />
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
                  }} />
                </div>
                <div style={{
                  fontSize: '0.78rem', color: 'var(--dex-gray-500)',
                  textAlign: 'right', fontVariantNumeric: 'tabular-nums',
                }}>
                  {bootProgress} %
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
    if (isSelfCheckInDeepLink) {
      return <SelfCheckInPage />;
    }
    switch (currentPage) {
      case 'landing':
        return <LandingPage />;
      case 'start':
        return <StartPage />;
      case 'register':
        return <EventListPage />;
      case 'registration':
        return <RegistrationPage />;
      case 'my-events':
        return <MyEventsPage />;
      case 'create-event':
      case 'edit-event':
        return <EventCreationPage />;
      case 'settings':
        return <SettingsPage />;
      case 'admin':
        return <AdminPage />;
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
  // Context global verfuegbar machen fuer ProfilePage
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__dexSpfxContext = props.context;

  return (
    <div className={styles.dexApp}>
      <LanguageProvider>
        <UserProvider context={props.context}>
          <RoleProvider context={props.context}>
            <NavigationProvider>
              <EventProvider context={props.context}>
                <AppContent />
              </EventProvider>
            </NavigationProvider>
          </RoleProvider>
        </UserProvider>
      </LanguageProvider>
    </div>
  );
}
