/**
 * User Context - aktueller SharePoint-Benutzer
 *
 * Liest den eingeloggten User aus dem SPFx WebPartContext
 * und stellt Name, E-Mail und Login-Name bereit.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
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
      // Format: "Nachname, Vorname"
      const parts = displayName.split(',');
      surname = parts[0].trim();
      firstName = parts.length > 1 ? parts[1].trim() : '';
    } else {
      // Format: "Vorname Nachname"
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
      location: '',
    });
    setIsLoading(false);

    // Admin-Status und Standort ueber SP-Profil nachladen
    loadUserProfile(props.context, spUser.email).then(profile => {
      if (profile) {
        setCurrentUser(prev => ({
          ...prev,
          location: profile.location,
          isAdmin: profile.isAdmin,
        }));
      }
    }).catch(() => { /* Profil konnte nicht geladen werden, Basisdaten reichen */ });
  }, []);

  return React.createElement(
    UserContext.Provider,
    { value: { currentUser, isLoading } },
    props.children
  );
}

/**
 * SP User Profile nachladen (Office-Standort + Gruppencheck)
 */
async function loadUserProfile(
  context: WebPartContext,
  _email: string
): Promise<{ location: string; isAdmin: boolean } | null> {
  try {
    // Office-Standort aus dem User-Profil lesen
    const profileUrl = `${context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
    const profileResponse = await context.spHttpClient.get(
      profileUrl,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (context.spHttpClient as any).configurations?.v1 || { headers: { 'Accept': 'application/json' } }
    );

    let location = '';
    if (profileResponse.ok) {
      const profileData = await profileResponse.json();
      // OfficeLocation aus UserProfileProperties extrahieren
      if (profileData.UserProfileProperties) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const officeProp = profileData.UserProfileProperties.find((p: any) => p.Key === 'Office' || p.Key === 'SPS-Location');
        if (officeProp) {
          location = officeProp.Value || '';
        }
      }
    }

    // Admin-Check: Pruefen ob User in der Site-Owners-Gruppe ist
    let isAdmin = false;
    try {
      const groupUrl = `${context.pageContext.web.absoluteUrl}/_api/web/currentuser/issiteadmin`;
      const groupResponse = await context.spHttpClient.get(
        groupUrl,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (context.spHttpClient as any).configurations?.v1 || { headers: { 'Accept': 'application/json' } }
      );
      if (groupResponse.ok) {
        const groupData = await groupResponse.json();
        isAdmin = groupData.value === true;
      }
    } catch {
      // Kein Admin-Zugriff = kein Admin
    }

    return { location, isAdmin };
  } catch {
    return null;
  }
}

export function useCurrentUser(): UserContextType {
  const ctx = React.useContext(UserContext);
  if (!ctx) throw new Error('useCurrentUser must be used within UserProvider');
  return ctx;
}
