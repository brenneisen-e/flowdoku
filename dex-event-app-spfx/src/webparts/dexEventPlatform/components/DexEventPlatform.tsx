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
    <div className="app-layout">
      <Header />
      <main className="main-content">
        {renderPage()}
      </main>
    </div>
  );
}

export default function DexEventPlatform(props: IDexEventPlatformProps): React.ReactElement {
  return (
    <div className={styles.dexApp}>
      <UserProvider context={props.context}>
        <RoleProvider context={props.context}>
          <NavigationProvider>
            <EventProvider>
              <AppContent />
            </EventProvider>
          </NavigationProvider>
        </RoleProvider>
      </UserProvider>
    </div>
  );
}
