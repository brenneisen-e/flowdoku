/**
 * User Context - aktueller SharePoint-Benutzer
 *
 * Liest den eingeloggten User aus dem SPFx WebPartContext
 * und stellt Name, E-Mail und Login-Name bereit.
 * Rolle wird ueber den RoleContext gesteuert.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { User } from '../types';

interface UserContextType {
  currentUser: User;
  isLoading: boolean;
}

const UserContext = React.createContext<UserContextType | undefined>(undefined);

export function UserProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [currentUser, setCurrentUser] = React.useState<User>({
    id: '',
    firstName: '',
    surname: '',
    email: '',
    isAdmin: false,
    role: 'User',
    location: '',
  });
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    // User-Daten aus dem SPFx-Context laden
    const spUser = props.context.pageContext.user;

    // DisplayName aufteilen (z.B. "Brenneisen, Eike" oder "Eike Brenneisen")
    const displayName = spUser.displayName || '';
    let firstName = '';
    let surname = '';

    if (displayName.indexOf(',') > -1) {
      const parts = displayName.split(',');
      surname = parts[0].trim();
      firstName = parts.length > 1 ? parts[1].trim() : '';
    } else {
      const parts = displayName.split(' ');
      firstName = parts[0] || '';
      surname = parts.slice(1).join(' ') || '';
    }

    setCurrentUser({
      id: spUser.loginName || '',
      firstName: firstName,
      surname: surname,
      email: spUser.email || '',
      isAdmin: false,
      role: 'User',
      location: '',
    });
    setIsLoading(false);

    // Standort ueber SP-Profil nachladen
    loadUserLocation(props.context).then(location => {
      if (location) {
        setCurrentUser(prev => ({ ...prev, location }));
      }
    }).catch(() => { /* Standort konnte nicht geladen werden */ });
  }, []);

  return React.createElement(
    UserContext.Provider,
    { value: { currentUser, isLoading } },
    props.children
  );
}

/**
 * Office-Standort aus dem SP User Profile lesen
 */
async function loadUserLocation(context: WebPartContext): Promise<string> {
  try {
    const profileUrl = `${context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
    const response = await context.spHttpClient.get(profileUrl, SPHttpClient.configurations.v1);

    if (response.ok) {
      const data = await response.json();
      if (data.UserProfileProperties) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const officeProp = data.UserProfileProperties.find((p: any) => p.Key === 'Office' || p.Key === 'SPS-Location');
        if (officeProp) return officeProp.Value || '';
      }
    }
  } catch { /* Profil nicht verfuegbar */ }
  return '';
}

export function useCurrentUser(): UserContextType {
  const ctx = React.useContext(UserContext);
  if (!ctx) throw new Error('useCurrentUser must be used within UserProvider');
  return ctx;
}
