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

export interface IDexEventPlatformProps {
  context: WebPartContext;
}

// Innere Komponente, die den NavigationContext nutzen kann
function AppContent(): React.ReactElement {
  const { currentPage } = useNavigation();
  const layoutRef = React.useRef<HTMLDivElement>(null);

  // Dynamische Höhe + SharePoint-Scroll unterdrücken
  React.useEffect(() => {
    const overflowTargets: HTMLElement[] = [];

    function setHeight(): void {
      if (layoutRef.current) {
        const rect = layoutRef.current.getBoundingClientRect();
        const available = window.innerHeight - rect.top;
        layoutRef.current.style.height = `${Math.max(available, 400)}px`;
      }
    }

    // SharePoint-Container finden und overflow unterdrücken
    function suppressSpScroll(): void {
      if (!layoutRef.current) return;
      let el: HTMLElement | null = layoutRef.current.parentElement;
      while (el && el !== document.body) {
        const style = window.getComputedStyle(el);
        if (style.overflowY === 'scroll' || style.overflowY === 'auto') {
          el.style.overflowY = 'hidden';
          overflowTargets.push(el);
        }
        el = el.parentElement;
      }
      // Auch den Body und HTML absichern
      document.documentElement.style.overflowY = 'hidden';
      document.body.style.overflowY = 'hidden';
    }

    setHeight();
    suppressSpScroll();
    window.addEventListener('resize', setHeight);
    // Nochmal nach Verzögerung (SP lädt Header nach)
    const timer = setTimeout(() => { setHeight(); suppressSpScroll(); }, 500);
    const timer2 = setTimeout(() => { setHeight(); suppressSpScroll(); }, 1500);

    return () => {
      window.removeEventListener('resize', setHeight);
      clearTimeout(timer);
      clearTimeout(timer2);
      // Overflow wiederherstellen beim Unmount
      overflowTargets.forEach(el => { el.style.overflowY = ''; });
      document.documentElement.style.overflowY = '';
      document.body.style.overflowY = '';
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
      case 'admin':
        return <AdminPage />;
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
