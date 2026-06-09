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
  /** v15.27: lowercase E-Mail-Adressen aller Mailverteiler/Gruppen,
   *  in denen der eingeloggte User Mitglied ist. Wird beim Boot via
   *  Microsoft Graph (/me/memberOf) geladen und in matchesAudience
   *  fuer Sichtbarkeits-Checks gegen Audience-Filter-DL-Adressen
   *  benutzt — vorher war ein literaler E-Mail-Vergleich, der nur
   *  bei direkter Adresse, nie bei DL-Mitgliedschaft matchte. */
  groupEmails: string[];
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
  const [groupEmails, setGroupEmails] = React.useState<string[]>([]);

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

    // v12.7: Falls Demo-Impersonation aktiv ist, überschreiben wir
    // Location / Name / Email mit den Impersonations-Werten. Damit sieht
    // der Admin tatsächlich die Welt durch die Brille des gewählten Users.
    const imp = (() => {
      try {
        const raw = typeof window !== 'undefined' ? window.localStorage?.getItem('dex_demo_impersonation') : null;
        if (!raw) return null;
        const obj = JSON.parse(raw);
        if (!obj || typeof obj !== 'object') return null;
        return obj as { email?: string; firstName?: string; surname?: string; location?: string };
      } catch { return null; }
    })();

    setCurrentUser({
      id: spUser.loginName || '',
      firstName: imp?.firstName || firstName,
      surname: imp?.surname || surname,
      email: imp?.email || spUser.email || '',
      isAdmin: false,
      role: 'User',
      location: imp?.location || '',
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

    // v15.27: Gruppen-/DL-Mitgliedschaften des Users laden, damit der
    // Audience-Filter-Check zur Laufzeit Verteiler-Mitgliedschaften matched.
    // Seit v16.4 ist die Pre-Compiled-Liste in event.audienceResolvedEmails
    // der primaere Pfad — der Runtime-Lookup hier dient nur noch als Fallback,
    // damit Events, die seit dem Schema-Upgrade noch nicht neu gespeichert
    // wurden, weiterhin funktionieren.
    loadUserGroupEmails(props.context).then(list => {
      setGroupEmails(list);
    }).catch(() => { setGroupEmails([]); });

    // Cleanup: Object URLs freigeben um Memory Leaks zu vermeiden
    return () => {
      setPhotoUrl(prev => {
        if (prev && prev.startsWith('blob:')) URL.revokeObjectURL(prev);
        return '';
      });
    };
  }, []);

  // v20.0 (Audit): Value memoizen — reine Datenfelder, verhindert App-weite
  // Re-Renders aller useCurrentUser()-Consumer bei Parent-Re-Renders.
  const value = React.useMemo<UserContextType>(
    () => ({ currentUser, isLoading, photoUrl, groupEmails }),
    [currentUser, isLoading, photoUrl, groupEmails]
  );

  return React.createElement(
    UserContext.Provider,
    { value },
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
 * v15.27: Gruppen-Mitgliedschaften (Mailverteiler + Security-Groups) des
 * eingeloggten Users via Microsoft Graph laden. Wird in matchesAudience
 * benutzt, um zu pruefen ob der User in einer Audience-DL Mitglied ist.
 *
 * Wir nehmen sowohl `mail` als auch `mailNickname@deloitte.com` /
 * `mailNickname@deloitte.de` auf — die Audience-Filter koennen entweder
 * die volle DL-Adresse (z.B. DETTDUESSELDORF@deloitte.com) oder den
 * Nicknames (selten) enthalten.
 */
async function loadUserGroupEmails(context: WebPartContext): Promise<string[]> {
  try {
    const client = await context.msGraphClientFactory.getClient('3');
    // /me/memberOf liefert alle Gruppen + Roles. Wir filtern auf groups
    // mit `mail`. $top=500 reicht in der Regel; wenn ein User in mehr als
    // 500 DLs ist, kommt Pagination ins Spiel — fuer DEALL ist das aber
    // praxisfern.
    // v17.11: `@odata.type` ist in $select nicht erlaubt (Graph wirft 400
    // „Parsing OData Select and Expand failed"). Wir holen nur mail +
    // mailNickname — der Type kommt via standard-odata-metadata wenn
    // benoetigt, aber wir filtern hier nur auf Gruppen mit Mail-Adresse
    // ohnehin.
    const resp = await client.api('/me/memberOf?$select=mail,mailNickname&$top=500').get();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const values: any[] = (resp && resp.value) || [];
    const out = new Set<string>();
    for (const g of values) {
      if (g.mail) out.add(String(g.mail).toLowerCase().trim());
      if (g.mailNickname) {
        const nick = String(g.mailNickname).toLowerCase().trim();
        // Plausible DL-Address-Schreibweisen mit den beiden Deloitte-Domains.
        out.add(`${nick}@deloitte.com`);
        out.add(`${nick}@deloitte.de`);
      }
    }
    return Array.from(out);
  } catch (err) {
    console.warn('[DEX] loadUserGroupEmails failed:', err);
    return [];
  }
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
