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

export interface IDexEventPlatformProps {
  context: WebPartContext;
}

// Innere Komponente, die den NavigationContext nutzen kann
function AppContent(): React.ReactElement {
  const { currentPage } = useNavigation();
  const layoutRef = React.useRef<HTMLDivElement>(null);

  // Dynamische Hoehe: misst wie viel Platz ueber der App ist
  React.useEffect(() => {
    function setHeight(): void {
      if (layoutRef.current) {
        const rect = layoutRef.current.getBoundingClientRect();
        const available = window.innerHeight - rect.top;
        layoutRef.current.style.height = `${Math.max(available, 400)}px`;
      }
    }
    setHeight();
    window.addEventListener('resize', setHeight);
    // Nochmal nach kurzer Verzoegerung (SP laedt Header nach)
    const timer = setTimeout(setHeight, 500);
    return () => {
      window.removeEventListener('resize', setHeight);
      clearTimeout(timer);
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
        return <EventCreationPage />;
      case 'settings':
        return <SettingsPage />;
      case 'profile':
        return <ProfilePage />;
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
