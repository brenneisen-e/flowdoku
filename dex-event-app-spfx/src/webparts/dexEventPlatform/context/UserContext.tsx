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
    // Audience-Filter-Check zur Laufzeit nicht nur direkte E-Mails matched,
    // sondern auch Verteiler-Mitgliedschaften (z.B. DETTDUESSELDORF@deloitte.com).
    loadUserGroupEmails(props.context).then(list => {
      setGroupEmails(list);
      // v15.28 Quickfix: explizite Targeted-Lookups fuer bekannte DLs, die
      // ueber /me/memberOf evtl. nicht zurueckkommen (z.B. weil verschachtelt
      // oder Distribution-only ohne Security). Fuer jede DL holen wir die
      // Members und pruefen, ob currentUser drin steckt. Liste hier
      // erweitern, sobald weitere produktive Audience-DLs auftauchen.
      const targetedDLs = [
        'DETTDuesseldorf@deloitte.com',
      ];
      const userEmail = (props.context.pageContext.user.email || '').toLowerCase();
      (async () => {
        const extra: string[] = [];
        for (const dl of targetedDLs) {
          try {
            const inDl = await isUserInGroup(props.context, dl, userEmail);
            if (inDl) extra.push(dl.toLowerCase());
          } catch { /* */ }
        }
        if (extra.length > 0) {
          setGroupEmails(prev => {
            const set = new Set(prev);
            for (const e of extra) set.add(e);
            return Array.from(set);
          });
        }
      })().catch(() => { /* */ });
    }).catch(() => { setGroupEmails([]); });

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
    { value: { currentUser, isLoading, photoUrl, groupEmails } },
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
    const resp = await client.api('/me/memberOf?$select=mail,mailNickname,@odata.type&$top=500').get();
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
 * v15.28: Quickfix — pruefe explizit, ob der eingeloggte User Mitglied
 * einer bestimmten DL ist (per Graph). Wird fuer Audience-Filter-DLs
 * genutzt, die /me/memberOf evtl. nicht zurueckliefert (verschachtelt
 * oder Distribution-only ohne Security-Enablement).
 *
 * Schritt 1: Group-ID per `/groups?$filter=mail eq '<dl>'` aufloesen.
 * Schritt 2: Per `/groups/{id}/members?$filter=...` checken, ob die
 *           UserEmail Mitglied ist. (Wir holen $top=999, paginieren
 *           NICHT — fuer typische Standort-DLs reicht das.)
 */
async function isUserInGroup(context: WebPartContext, groupEmail: string, userEmail: string): Promise<boolean> {
  if (!groupEmail || !userEmail) return false;
  try {
    const client = await context.msGraphClientFactory.getClient('3');
    const lookup = await client
      .api(`/groups?$filter=mail eq '${groupEmail.replace(/'/g, "''")}'&$select=id&$top=1`)
      .get();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const groups: any[] = (lookup && lookup.value) || [];
    if (groups.length === 0) return false;
    const groupId = groups[0].id;
    if (!groupId) return false;
    const userLc = userEmail.toLowerCase();
    // transitiveMembers loest verschachtelte Gruppen automatisch auf.
    const members = await client
      .api(`/groups/${groupId}/transitiveMembers?$select=mail,userPrincipalName&$top=999`)
      .get();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const list: any[] = (members && members.value) || [];
    for (const m of list) {
      const mail = (m.mail || '').toLowerCase();
      const upn = (m.userPrincipalName || '').toLowerCase();
      if (mail === userLc || upn === userLc) return true;
    }
    return false;
  } catch (err) {
    console.warn('[DEX] isUserInGroup failed:', groupEmail, err);
    return false;
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
