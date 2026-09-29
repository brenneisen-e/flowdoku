/**
 * Die angemeldete Person — Name, Adresse und Foto.
 *
 * Aufbau wie `UserContext` in DEX, nur ohne dessen Standort-, Vorschau- und
 * Impersonate-Teil: Diese Plattform braucht den Namen für die Begrüßung und
 * das Profilbild im Kopf, mehr nicht.
 *
 * Der Name kommt in zwei Stufen. Sofort und ohne Netz aus `pageContext` (über
 * `safeDisplayName`, das ein Claims-Token wegräumt); danach, sobald das
 * Benutzerprofil antwortet, der gepflegte Name. Die Begrüßung steht also
 * nie leer da und springt höchstens einmal von der Adresse auf den Namen —
 * besser als eine Seite, die auf einen Profildienst wartet, der in manchen
 * Mandanten Sekunden braucht.
 *
 * Das Foto ist nur eine Adresse: `userphoto.aspx` liefert bei fehlendem Bild
 * ein Standardbild oder 404. Der Aufrufer fängt `onError` und zeigt dann die
 * Initialen — deshalb wird hier nichts vorab geprüft.
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { firstNameOf, resolveMyDisplayName, safeDisplayName } from '../utils/displayName';

export interface CurrentUser {
  firstName: string;
  surname: string;
  email: string;
  /** Vollständiger Name, wie er angezeigt wird. */
  displayName: string;
}

interface UserContextType {
  currentUser: CurrentUser;
  /** Adresse des Profilfotos — kann auf ein Standardbild oder 404 zeigen. */
  photoUrl: string;
}

const UserContext = React.createContext<UserContextType | undefined>(undefined);

/** „Nachname, Vorname" und „Vorname Nachname" → Vorname und Nachname. */
function teileName(raw: string): { firstName: string; surname: string } {
  const v = (raw || '').trim();
  if (!v) return { firstName: '', surname: '' };
  if (v.indexOf(',') > -1) {
    const parts = v.split(',');
    return { firstName: firstNameOf(v), surname: (parts[0] || '').trim() };
  }
  const worte = v.split(/\s+/);
  return { firstName: worte[0] || '', surname: worte.slice(1).join(' ') };
}

export function UserProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const email = props.context.pageContext.user.email || '';
  const anfang = safeDisplayName(props.context.pageContext.user.displayName || '', email);

  const [name, setName] = React.useState(anfang);

  React.useEffect(() => {
    let weg = false;
    resolveMyDisplayName(props.context)
      .then(n => { if (!weg && n) setName(n); })
      // Ohne Profil bleibt es beim Namen aus pageContext — kein Fehler, nur weniger schön.
      .catch(() => undefined);
    return () => { weg = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const value = React.useMemo<UserContextType>(() => {
    const teile = teileName(name);
    // Steht als Name nur die Adresse (Profil ohne Namen), ist der „Vorname"
    // die Adresse — die Begrüßung darf dann nicht „Hallo, a.b@c.de." sagen.
    const istAdresse = name.indexOf('@') > -1;
    return {
      currentUser: {
        firstName: istAdresse ? '' : teile.firstName,
        surname: istAdresse ? '' : teile.surname,
        email,
        displayName: name,
      },
      photoUrl: email
        ? `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`
        : '',
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [name, email]);

  return React.createElement(UserContext.Provider, { value }, props.children);
}

export function useCurrentUser(): UserContextType {
  const ctx = React.useContext(UserContext);
  if (!ctx) throw new Error('useCurrentUser muss innerhalb des UserProvider stehen');
  return ctx;
}
