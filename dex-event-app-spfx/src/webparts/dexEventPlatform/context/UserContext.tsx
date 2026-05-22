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
  photoUrl: string;
}

export const UserContext = React.createContext<UserContextType | undefined>(undefined);

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
  const [photoUrl, setPhotoUrl] = React.useState<string>('');

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

    // Standort + JobTitle ueber SP-Profil nachladen.
    // JobTitle brauchen wir fuer die Assistant-Ausnahme in der Registration
    // ("Assistant" darf fuer "Director"/"Partner" registrieren).
    loadUserProfile(props.context).then(profile => {
      if (profile.location || profile.jobTitle || profile.department || profile.mobilePhone) {
        setCurrentUser(prev => ({
          ...prev,
          ...(profile.location ? { location: profile.location } : {}),
          ...(profile.jobTitle ? { jobTitle: profile.jobTitle } : {}),
          ...(profile.department ? { department: profile.department } : {}),
          ...(profile.mobilePhone ? { mobilePhone: profile.mobilePhone } : {}),
        }));
      }
    }).catch(() => { /* Profil konnte nicht geladen werden */ });

    // Profilbild ueber Microsoft Graph laden
    loadUserPhoto(props.context).then(url => {
      if (url) setPhotoUrl(url);
    }).catch(() => { /* Foto nicht verfuegbar */ });

    // Cleanup: Object URLs freigeben um Memory Leaks zu vermeiden
    return () => {
      setPhotoUrl(prev => {
        if (prev && prev.startsWith('blob:')) URL.revokeObjectURL(prev);
        return '';
      });
    };
  }, []);

  return React.createElement(
    UserContext.Provider,
    { value: { currentUser, isLoading, photoUrl } },
    props.children
  );
}

/**
 * Office-Standort + JobTitle aus dem SP User Profile lesen.
 * JobTitle liegt im Property "Title" (bzw. "SPS-JobTitle" im Fallback).
 */
async function loadUserProfile(context: WebPartContext): Promise<{ location: string; jobTitle: string; department: string; mobilePhone: string }> {
  try {
    const profileUrl = `${context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
    const response = await context.spHttpClient.get(profileUrl, SPHttpClient.configurations.v1);

    if (response.ok) {
      const data = await response.json();
      if (data.UserProfileProperties) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties;
        const getProp = (keys: string[]): string => {
          for (const k of keys) {
            const p = props.find(x => x.Key === k);
            if (p && p.Value) return p.Value;
          }
          return '';
        };
        return {
          location: getProp(['Office', 'SPS-Location']),
          jobTitle: getProp(['Title', 'SPS-JobTitle']),
          // v11.94: Department + Mobile zusätzlich aus dem SP-Profil.
          department: getProp(['Department', 'SPS-Department']),
          mobilePhone: getProp(['CellPhone', 'SPS-MobilePhone', 'MobilePhone']),
        };
      }
    }
  } catch { /* Profil nicht verfuegbar */ }
  return { location: '', jobTitle: '', department: '', mobilePhone: '' };
}

/**
 * Profilbild ueber Microsoft Graph laden
 */
async function loadUserPhoto(context: WebPartContext): Promise<string> {
  try {
    const graphClient = await context.msGraphClientFactory.getClient('3');
    const blob = await graphClient.api('/me/photo/$value').get();
    if (blob && blob.size > 0) {
      return URL.createObjectURL(blob);
    }
  } catch {
    // Graph nicht verfuegbar — Fallback auf SharePoint UserPhoto
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const email = context.pageContext.user.email;
      return `${siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
    } catch { /* */ }
  }
  return '';
}

export function useCurrentUser(): UserContextType {
  const ctx = React.useContext(UserContext);
  if (!ctx) throw new Error('useCurrentUser must be used within UserProvider');
  return ctx;
}
