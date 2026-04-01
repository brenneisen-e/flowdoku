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
import { EventProvider } from '../context/EventContext';
import { UserProvider } from '../context/UserContext';
import { RoleProvider } from '../context/RoleContext';
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

export interface IDexEventPlatformProps {
  context: WebPartContext;
}

// Innere Komponente, die den NavigationContext nutzen kann
function AppContent(): React.ReactElement {
  const { currentPage } = useNavigation();
  const layoutRef = React.useRef<HTMLDivElement>(null);

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
      default:
        return <LandingPage />;
    }
  };

  return (
    <div className="app-layout" ref={layoutRef}>
      <Header />
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
      <UserProvider context={props.context}>
        <RoleProvider context={props.context}>
          <NavigationProvider>
            <EventProvider context={props.context}>
              <AppContent />
            </EventProvider>
          </NavigationProvider>
        </RoleProvider>
      </UserProvider>
    </div>
  );
}
