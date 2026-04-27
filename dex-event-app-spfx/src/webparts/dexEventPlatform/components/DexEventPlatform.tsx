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
import EventCreationPage from './EventCreationPage';
import SettingsPage from './SettingsPage';
import ProfilePage from './ProfilePage';
import AdminPage from './AdminPage';
import RoleMatrixPage from './RoleMatrixPage';
import ParticipantsPage from './ParticipantsPage';
import FlowchartPage from './FlowchartPage';
import CheckInPage from './CheckInPage';
import ManualPage from './manual/ManualPage';

export interface IDexEventPlatformProps {
  context: WebPartContext;
}

// Innere Komponente, die den NavigationContext nutzen kann
function AppContent(): React.ReactElement {
  const { currentPage, navigate } = useNavigation();
  const { isAdmin, isRolesLoading } = useRoles();
  const { markExpiredEventsAsCompleted, isEventsLoading, events } = useEvents();

  // v7.5: Boot-Progress. Wir mappen die zwei Lade-Phasen auf Prozentwerte
  // und animieren sanft dazwischen, damit der User auf der Landing-Page
  // den Fortschritt nicht nur als unbestimmten Balken, sondern konkret
  // als "X %" sieht. Die Phasen:
  //   - Rollen + Events laden noch: Target 30%
  //   - Rollen fertig, Events laden: Target 70%
  //   - alles fertig: Target 100% (Loader verschwindet sowieso)
  const [bootProgress, setBootProgress] = React.useState<number>(8);
  React.useEffect(() => {
    let target: number;
    if (isRolesLoading && isEventsLoading) target = 30;
    else if (isRolesLoading) target = 50;
    else if (isEventsLoading) target = 75;
    else target = 100;
    // Sanft hochzählen alle 60ms in Richtung Target — fühlt sich
    // weniger ruckelig an als ein direkter setBootProgress(target).
    const id = setInterval(() => {
      setBootProgress(prev => {
        if (prev >= target) return target;
        const delta = Math.max(1, Math.round((target - prev) * 0.18));
        return Math.min(target, prev + delta);
      });
    }, 60);
    return () => clearInterval(id);
  }, [isRolesLoading, isEventsLoading]);
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
    if (currentPage === 'landing' && !isCancelDeepLink && (isEventsLoading || isRolesLoading)) {
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
                  Willkommen auf der neuen <strong style={{ whiteSpace: 'nowrap' }}>Event Experience Platform.</strong>
                </h1>
                <p style={{ color: 'var(--dex-gray-500)', marginTop: 12, fontSize: '0.95rem' }}>
                  Jeden Moment geht&apos;s los…
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
      case 'manual':
        return <ManualPage />;
      default:
        return <LandingPage />;
    }
  };

  // v6.29: Während der Boot-Loader läuft, Header verstecken. Sonst würde
  // schon die "Jetzt einchecken"-Bubble / QR-Icon blinken bevor der eigentliche
  // Welcome-Screen sichtbar ist.
  const isBootLoading = currentPage === 'landing' && !isCancelDeepLink && (isEventsLoading || isRolesLoading);
  return (
    <div className="app-layout" ref={layoutRef}>
      {!isBootLoading && <Header />}
      <main className="main-content">
        {renderPage()}
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
