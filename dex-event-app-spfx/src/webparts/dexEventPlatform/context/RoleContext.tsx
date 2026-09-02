/**
 * Role Context - Rollenverwaltung über SharePoint-Liste
 *
 * Erstellt beim ersten Start automatisch die DEX_Roles-Liste
 * auf dem SharePoint. Admins können andere User berechtigen.
 *
 * Rollen: Admin (ehem. SuperAdmin), Organizer (ehem. EventAdmin), User
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';
import { RoleAssignment, UserRole } from '../types';
import { looksLikeClaimName, resolveMyDisplayName, safeDisplayName } from '../utils/displayName';
import { dlog } from '../utils/debugLog';
import { isAdminRole, roleRank } from '../utils/roleRank';
import { isCurrentUser } from '../utils/sessionIdentities';

interface RoleContextType {
  roles: RoleAssignment[];
  currentUserRole: UserRole;
  isRolesLoading: boolean;
  isAdmin: boolean;
  /** v12.7: Echte Rolle aus DEX_Roles, unabhängig von Demo-Impersonation.
   *  Wird genutzt um in der Header-UI das „Demo: als User testen"-Menü
   *  weiterhin zu zeigen, auch wenn isAdmin durch Impersonation auf
   *  false gesetzt wurde. */
  originalIsAdmin: boolean;
  /** v12.7: aktiv wenn Admin Demo-Impersonation gestartet hat. */
  isImpersonating: boolean;
  /** v30.3: „Übersicht als User sehen" — leichter Vorschau-Modus für
   *  Organizer UND Admins. Anders als die Demo-Impersonation bleibt die
   *  eigene Identität (E-Mail, Standort, Anmeldungen) erhalten; nur die
   *  Rollen-Bypässe fallen weg. Bewusst NICHT in localStorage: ein Reload
   *  beendet die Vorschau, damit niemand versehentlich darin hängen
   *  bleibt. Die Seiten sperren in der Vorschau zusätzlich das Anmelden. */
  previewAsUser: boolean;
  setPreviewAsUser: (on: boolean) => void;
  /** Rolle 'F&A'. Admins sehen das Center ebenfalls.
   *
   *  v30.60: Der Zuschnitt ist ein anderer als in v30.5. Damals war F&A eine
   *  reine Lese-Rolle (Teilnehmer-Rechte plus das Center). Nutzer-Ansage
   *  01.09.2026: „diese Person können zwei Sachen: 1. alles was Organizer
   *  können (aber plus die neue Abrechnungsfunktion) und 2. Zugriff auf das
   *  F&A Center." F&A ist damit ein Organizer mit zwei Zusätzen — deshalb
   *  zählt die Rolle unten in `isOrganizer` mit, und `canEditBilling`
   *  bekommt sie als eigenes Argument. */
  isFA: boolean;
  isOrganizer: boolean;
  canCreateEvents: boolean;
  /** v26: Der aktuelle User ist Power-User (Organizer/Admin mit IsPowerUser-Flag).
   *  Power-User + Admins beantworten Tickets über die „Tickets"-Kachel. */
  isPowerUser: boolean;
  siteUrl: string;
  addRole: (userEmail: string, userName: string, role: UserRole, location: string) => Promise<boolean>;
  updateRole: (itemId: number, newRole: UserRole) => Promise<boolean>;
  setPowerUser: (itemId: number, isPowerUser: boolean) => Promise<boolean>;
  updateRoleLocation: (itemId: number, location: string) => Promise<boolean>;
  removeRole: (itemId: number) => Promise<boolean>;
  refreshRoles: () => Promise<void>;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string } | null>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>;
  /** v30.61: Personalnummer, Kostenstelle, Firma und Land mehrerer Personen
   *  auf einmal aus dem Verzeichnis (Graph, `User.Read.All` nötig). Leere Map
   *  = nicht erlaubt oder nicht gepflegt — nicht „gibt es nicht". */
  getEmployeeData: (emails: string[]) => Promise<Record<string, { employeeId?: string; costCenter?: string; companyName?: string; country?: string; department?: string }>>;
  searchGroups: (query: string) => Promise<Array<{ email: string; displayName: string }>>;
  getGroupMembers: (groupEmail: string) => Promise<{ groupName: string; members: Array<{ email: string; displayName: string; firstName?: string; lastName?: string; jobTitle?: string; location?: string }> } | null>;
  searchUsersByLocation: (location: string) => Promise<Array<{ email: string; displayName: string; firstName: string; lastName: string; location: string; jobTitle: string }>>;
}

// Migration: alte SP-Werte auf neue mappen
function migrateRole(spRole: string): UserRole {
  if (spRole === 'SuperAdmin') return 'Admin';
  if (spRole === 'EventAdmin') return 'Organizer';
  if (spRole === 'Admin' || spRole === 'IT-Admin' || spRole === 'Organizer' || spRole === 'F&A' || spRole === 'User') return spRole as UserRole;
  return 'User';
}

export const RoleContext = React.createContext<RoleContextType | undefined>(undefined);

export function RoleProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [roles, setRoles] = React.useState<RoleAssignment[]>([]);
  const [currentUserRole, setCurrentUserRole] = React.useState<UserRole>('User');
  const [isRolesLoading, setIsRolesLoading] = React.useState<boolean>(true);
  const spService = React.useMemo(() => new SharePointService(props.context), []);

  const currentUserEmail = props.context.pageContext.user.email;
  // v28.64: Siehe utils/displayName.ts — der Anzeigename aus dem pageContext
  // kann das Claims-Login-Token sein. Gleiche Absicherung wie im EventContext.
  const rawUserName = props.context.pageContext.user.displayName;
  const [profileUserName, setProfileUserName] = React.useState('');
  React.useEffect(() => {
    if (!looksLikeClaimName(rawUserName)) return;
    resolveMyDisplayName(props.context)
      .then(n => { if (n) setProfileUserName(n); })
      .catch(() => { /* Fallback bleibt die E-Mail */ });
  }, [rawUserName]);
  const currentUserName = profileUserName || safeDisplayName(rawUserName, currentUserEmail);

  React.useEffect(() => {
    initRoles().catch(() => setIsRolesLoading(false));
  }, []);

  async function initRoles(): Promise<void> {
    // v11.74: Profiling der Roles-Boot-Phase.
    const tBoot = performance.now();
    const tEns = performance.now();
    // v6.34 (Security-Hotfix): ensureRolesList sagt uns, ob die Liste gerade
    // frisch angelegt wurde. Nur dann (echte Erstinstallation) darf der erste
    // User automatisch Admin werden. Sonst bleibt der Default 'User' — sonst
    // gab es einen kritischen Bug: getRoles() liefert bei 403 Forbidden (oder
    // Netzwerk-Fehler) [] zurück, und die alte Logik interpretierte das als
    // "leere Liste → Erstinstallation" und beförderte JEDEN aufrufenden User
    // zum Admin.
    const { isNewlyCreated } = await spService.ensureRolesList();
    // eslint-disable-next-line no-console
    dlog('perf', `[DEX][perf][roles] ensureRolesList = ${Math.round(performance.now() - tEns)} ms`);
    const tGet = performance.now();
    const spRoles = await spService.getRoles();
    // eslint-disable-next-line no-console
    dlog('perf', `[DEX][perf][roles] getRoles = ${Math.round(performance.now() - tGet)} ms (n=${spRoles ? spRoles.length : 'null'})`);
    void tBoot;

    if (spRoles === null) {
      // API-Fehler / Permission-Issue: KEINE Rolle zuordnen, als 'User' lassen.
      // Sonst würde ein vorübergehender Read-Fehler die gesamte Access-Control
      // umgehen. Besser: der User sieht nichts, als dass er plötzlich Admin ist.
      setRoles([]);
      setCurrentUserRole('User');
      setIsRolesLoading(false);
      console.warn('[DEX] RoleContext: DEX_Roles konnte nicht gelesen werden — User-Rolle bleibt auf "User" (keine Admin-Auto-Upgrade).');
      return;
    }

    const myRole = spRoles.find(
      r => r.Title && r.Title.toLowerCase() === currentUserEmail.toLowerCase()
    );

    if (isNewlyCreated && spRoles.length === 0) {
      // Echte Erstinstallation: Liste wurde gerade von ensureRolesList angelegt
      // und ist erwartungsgemäß leer. Den aufrufenden User als Initial-Admin
      // eintragen, damit überhaupt jemand weitere Rollen vergeben kann.
      await spService.addRole(currentUserEmail, currentUserName, 'Admin', '', 'System (Erstinstallation)');
      setCurrentUserRole('Admin');
      setRoles([{
        id: 0, userEmail: currentUserEmail, userName: currentUserName,
        role: 'Admin', location: '', assignedBy: 'System (Erstinstallation)',
        assignedDate: new Date().toISOString(),
      }]);
    } else {
      // Normale Folge-Laufzeit: Rollen aus der Liste mappen, eigene Rolle
      // setzen. Wenn der User NICHT in DEX_Roles steht, bleibt er bei 'User'.
      const mapped: RoleAssignment[] = spRoles.map(r => ({
        id: r.Id,
        userEmail: r.Title || '',
        userName: r.UserName || '',
        role: migrateRole(r.Role),
        location: r.UserLocation || '',
        assignedBy: r.AssignedBy || '',
        assignedDate: r.AssignedDate || '',
        isPowerUser: !!r.IsPowerUser,
      }));
      setRoles(mapped);
      setCurrentUserRole(myRole ? migrateRole(myRole.Role) : 'User');

      // Alte Werte in SP migrieren (im Hintergrund)
      for (const r of spRoles) {
        if (r.Role === 'SuperAdmin' || r.Role === 'EventAdmin') {
          spService.updateRole(r.Id, migrateRole(r.Role)).catch(err => console.warn('[DEX] role migration failed:', err));
        }
      }
    }

    setIsRolesLoading(false);
  }

  async function refreshRoles(): Promise<void> {
    const spRoles = await spService.getRoles();
    // v6.34: Auch hier API-Fehler nicht als leere Liste interpretieren (siehe
    // initRoles) — sonst würde ein Netzwerk-Fehler mitten in der App-Nutzung
    // die Rolle des aktiven Users auf 'User' zurücksetzen (oder in einem
    // anderen Code-Pfad auf Admin, was wir gerade gefixt haben).
    if (spRoles === null) {
      console.warn('[DEX] RoleContext.refresh: DEX_Roles nicht lesbar, bestehender State bleibt.');
      return;
    }
    const mapped: RoleAssignment[] = spRoles.map(r => ({
      id: r.Id, userEmail: r.Title || '', userName: r.UserName || '',
      role: migrateRole(r.Role), location: r.UserLocation || '',
      assignedBy: r.AssignedBy || '', assignedDate: r.AssignedDate || '',
      isPowerUser: !!r.IsPowerUser,
    }));
    setRoles(mapped);
    const myRole = spRoles.find(
      r => r.Title && r.Title.toLowerCase() === currentUserEmail.toLowerCase()
    );
    setCurrentUserRole(myRole ? migrateRole(myRole.Role) : 'User');
  }

  async function addRole(
    userEmail: string, userName: string, role: UserRole, location: string
  ): Promise<boolean> {
    // v29.63: Kein zweiter Eintrag fuer dieselbe Person. `addRole` hat bisher
    // bedingungslos eine Zeile in DEX_Roles angelegt — auch wenn es fuer die
    // Adresse laengst eine gab. Zwei Zeilen sind kein Schoenheitsfehler:
    // `refreshRoles` bestimmt die Rolle des Users per `find`, nimmt also die
    // ERSTE Zeile. Eine spaetere Aenderung an der zweiten bleibt damit
    // wirkungslos, und eine Herabstufung kann von einer alten Zeile
    // ueberstimmt werden.
    const mail = (userEmail || '').trim().toLowerCase();
    const existing = roles.filter(r => (r.userEmail || '').trim().toLowerCase() === mail)[0];
    if (existing) {
      // Gleiche Rolle: nichts zu tun, aber die Rechte noch einmal setzen —
      // sie koennen beim ersten Mal am Throttling gescheitert sein.
      if (existing.role === role) {
        try {
          if (role === 'Admin' || role === 'IT-Admin') {
            await spService.grantFullControlOnRolesList(userEmail);
            await spService.grantFullControlOnEventsList(userEmail);
            await spService.grantOrganizerPermissions(userEmail);
          } else if (role === 'Organizer' || role === 'F&A') {
            // v30.60: F&A arbeitet wie ein Organizer und braucht dieselben
            // SharePoint-Rechte — ohne sie scheitert das Anlegen der Subsite,
            // und zwar erst beim Speichern eines Events.
            await spService.grantReadOnRolesList(userEmail);
            await spService.grantOrganizerPermissions(userEmail);
          }
        } catch (err) { console.warn('[DEX] permission re-grant failed (best-effort):', err); }
        return true;
      }
      // v30.67: Nie implizit herabstufen. `addRole` heißt „diese Rechte
      // sicherstellen" — wer schon MEHR hat (F&A oder Admin, wenn Organizer
      // angefragt ist), verlöre über diesen Weg still das F&A Center bzw. die
      // Admin-Rechte; genau so ist die Freigabe eines „Organizer werden"-
      // Antrags einer F&A-Person passiert. Eine bewusste Herabstufung geht
      // über `updateRole` aus der Rollenverwaltung, nie über `addRole`.
      if (roleRank(existing.role) > roleRank(role)) {
        console.warn(`[DEX] addRole: ${mail} hat bereits die Rolle ${existing.role} — ${role} wäre eine Herabstufung und wird NICHT gesetzt.`);
        return false;
      }
      // Andere Rolle: die bestehende Zeile aendern statt eine zweite anzulegen.
      return updateRole(existing.id, role);
    }
    const success = await spService.addRole(userEmail, userName, role, location, currentUserName);
    if (success) {
      try {
        if (role === 'Admin' || role === 'IT-Admin') {
          await spService.grantFullControlOnRolesList(userEmail);
          await spService.grantFullControlOnEventsList(userEmail);
          await spService.grantOrganizerPermissions(userEmail); // Site-Rechte für Subsite-Erstellung
        } else if (role === 'Organizer' || role === 'F&A') {
          await spService.grantReadOnRolesList(userEmail);
          await spService.grantOrganizerPermissions(userEmail);
        }
      } catch (err) { console.warn('[DEX] permission grant for addRole failed (best-effort):', err); }
      await refreshRoles();
    }
    return success;
  }

  // v30.67: Alle direkten Rechte einer Person entziehen — DEX_Roles,
  // DEX_Events und (NEU) der Web-Root. Rückgabe false heißt „mindestens ein
  // Entzug wirkt nicht"; der Aufrufer reicht das nach oben statt Erfolg zu
  // melden. Jeder Entzug betrifft nur die direkte Zuweisung DIESES Principals
  // (`getbyprincipalid`), nie eine Gruppe.
  async function revokeAllAccess(userEmail: string): Promise<boolean> {
    const rolesOk = await spService.revokeAccessOnRolesList(userEmail);
    const eventsOk = await spService.revokeAccessOnEventsList(userEmail);
    // Selbstschutz: Die angemeldete Person darf sich nicht selbst vom Web
    // aussperren — ein Admin, der die Rollenverwaltung an sich testet, käme
    // sonst nicht mehr an die Site. Die Listenrechte laufen wie bisher; nur
    // der Web-Entzug wird übersprungen, mit Warnung statt still.
    let siteOk = true;
    if (isCurrentUser(props.context, userEmail)) {
      console.warn(`[DEX] RoleContext: Web-Rechte von ${userEmail} werden NICHT entzogen — das ist die angemeldete Person (Selbstschutz).`);
    } else {
      siteOk = await spService.revokeSiteAccess(userEmail);
    }
    return rolesOk && eventsOk && siteOk;
  }

  async function updateRole(itemId: number, newRole: UserRole): Promise<boolean> {
    const oldRole = roles.find(r => r.id === itemId);
    const success = await spService.updateRole(itemId, newRole);
    // v30.67: `rightsOk` trägt, ob die SharePoint-Rechte zur neuen Rolle
    // passen. Vorher wurde jeder Fehler im Rechte-Block als „best-effort"
    // weggeloggt und nur `success` (die DEX_Roles-Zeile) zurückgegeben —
    // die App meldete Erfolg, während die alten Rechte stehen blieben.
    let rightsOk = true;
    if (success && oldRole) {
      try {
        if (newRole === 'Admin' || newRole === 'IT-Admin') {
          await spService.grantFullControlOnRolesList(oldRole.userEmail);
          await spService.grantFullControlOnEventsList(oldRole.userEmail);
          await spService.grantOrganizerPermissions(oldRole.userEmail);
        } else if (newRole === 'Organizer' || newRole === 'F&A') {
          // v30.67: Downgrade Admin → Organizer/F&A. `addroleassignment` ist
          // ADDITIV — Read auf DEX_Roles kam bisher NEBEN das bestehende Full
          // Control, der Ex-Admin konnte die Rollenliste weiter bearbeiten
          // und sich selbst wieder hochstufen. Deshalb ERST die bestehende
          // Zuweisung auf DEX_Roles entfernen, DANN Read vergeben. DEX_Events
          // und Web-Root bleiben: dort haben Organizer dasselbe Full Control.
          if (isAdminRole(oldRole.role)) {
            rightsOk = await spService.revokeAccessOnRolesList(oldRole.userEmail) && rightsOk;
          }
          await spService.grantReadOnRolesList(oldRole.userEmail);
          await spService.grantOrganizerPermissions(oldRole.userEmail);
        } else if (newRole === 'User') {
          rightsOk = await revokeAllAccess(oldRole.userEmail) && rightsOk;
        }
      } catch (err) {
        console.warn('[DEX] permission grant for updateRole failed:', err);
        rightsOk = false;
      }
      await refreshRoles();
    }
    if (success && !rightsOk) {
      console.warn(`[DEX] updateRole: Rolle von ${oldRole?.userEmail} steht in DEX_Roles auf ${newRole}, aber mindestens ein SharePoint-Recht konnte NICHT entzogen werden — siehe Warnungen darüber.`);
    }
    return success && rightsOk;
  }

  // v18.5: Power-User-Flag setzen/entfernen (Zusatz auf einem Organizer).
  async function setPowerUser(itemId: number, isPowerUser: boolean): Promise<boolean> {
    const success = await spService.setPowerUser(itemId, isPowerUser);
    if (success) await refreshRoles();
    return success;
  }

  async function removeRole(itemId: number): Promise<boolean> {
    const roleEntry = roles.find(r => r.id === itemId);
    const success = await spService.deleteRole(itemId);
    // v30.67: Die Entzüge laufen jetzt AWAITED und mit Ergebnis — vorher
    // fire-and-forget mit `.catch(console.warn)`, und das Web-Recht (Full
    // Control auf der Site-Collection aus `grantOrganizerPermissions`) wurde
    // gar nicht angefasst. Die Person war aus DEX_Roles verschwunden und
    // konnte trotzdem jede Teilnehmer-Subsite lesen.
    let rightsOk = true;
    if (success) {
      await refreshRoles();
      if (roleEntry) {
        try {
          rightsOk = await revokeAllAccess(roleEntry.userEmail);
        } catch (err) {
          console.warn('[DEX] removeRole: Rechte-Entzug fehlgeschlagen:', err);
          rightsOk = false;
        }
        if (!rightsOk) {
          console.warn(`[DEX] removeRole: Zeile von ${roleEntry.userEmail} ist aus DEX_Roles entfernt, aber mindestens ein SharePoint-Recht konnte NICHT entzogen werden — siehe Warnungen darüber.`);
        }
      }
    }
    return success && rightsOk;
  }

  async function updateRoleLocation(itemId: number, location: string): Promise<boolean> {
    const success = await spService.updateRoleLocation(itemId, location);
    if (success) await refreshRoles();
    return success;
  }

  async function searchUser(email: string): Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string } | null> {
    return spService.searchUserByEmail(email);
  }

  async function searchUsers(query: string, includeInternational: boolean = false): Promise<Array<{ email: string; displayName: string; location: string; jobTitle: string }>> {
    return spService.searchUsers(query, includeInternational);
  }

  async function searchGroups(query: string): Promise<Array<{ email: string; displayName: string }>> {
    return spService.searchGroups(query);
  }

  async function getGroupMembers(groupEmail: string): Promise<{ groupName: string; members: Array<{ email: string; displayName: string; firstName?: string; lastName?: string; jobTitle?: string; location?: string }> } | null> {
    return spService.getGroupMembers(groupEmail);
  }

  async function getEmployeeData(emails: string[]): Promise<Record<string, { employeeId?: string; costCenter?: string; companyName?: string; country?: string; department?: string }>> {
    return spService.getEmployeeData(emails);
  }

  async function searchUsersByLocation(location: string): Promise<Array<{ email: string; displayName: string; firstName: string; lastName: string; location: string; jobTitle: string }>> {
    return spService.searchUsersByLocation(location);
  }

  // v12.7: Demo-Impersonation. Admins können in den Header-User-Menü auf
  // „Demo: als User testen" klicken, einen Standort wählen, danach
  // agiert die App so als wäre der User ein normaler 'User' am gewählten
  // Standort. Wird in localStorage gehalten, damit das auch über
  // Reload bestehen bleibt. Beenden via Banner-Klick oben.
  const isImpersonating = typeof window !== 'undefined' && !!window.localStorage?.getItem('dex_demo_impersonation');
  // v30.3: Organizer-/Admin-Vorschau „Übersicht als User sehen" — gleiche
  // Rollen-Absenkung wie die Impersonation, aber ohne Identitätswechsel.
  const [previewAsUser, setPreviewAsUser] = React.useState(false);
  const effectiveRole: UserRole = (isImpersonating || previewAsUser) ? 'User' : currentUserRole;
  // v26.33: IT-Admin hat die gleichen App-Rechte wie Admin (nur keine Mails —
  // das regelt die exakte Role='Admin'-Filterung in den Empfänger-Listen).
  const isAdmin = effectiveRole === 'Admin' || effectiveRole === 'IT-Admin';
  // v30.60: F&A zählt als Organizer. Vorher war die Rolle eine Sackgasse —
  // wer sie bekam, verlor die Organizer-Rechte, die dieselbe Person für ihre
  // eigenen Events braucht. Die Abrechnung ist ein ZUSATZ zur Organizer-
  // Arbeit, kein Ersatz dafür.
  const isOrganizer = effectiveRole === 'Organizer' || effectiveRole === 'F&A' || isAdmin;
  const canCreateEvents = isOrganizer;
  // v30.5: F&A sieht das F&A Center — Impersonation/Vorschau senken auch das ab.
  const isFA = effectiveRole === 'F&A';
  // v26: Power-User-Flag des aktuellen Users aus DEX_Roles. Im Demo-/
  // Impersonations-Modus bewusst false (wie isAdmin) — der Demo-User soll
  // exakt das sehen, was ein normaler User sieht.
  const myRoleEntry = roles.find(r => (r.userEmail || '').toLowerCase() === currentUserEmail.toLowerCase());
  const isPowerUser = !isImpersonating && !previewAsUser && !!myRoleEntry?.isPowerUser;
  const originalIsAdmin = currentUserRole === 'Admin' || currentUserRole === 'IT-Admin';
  const siteUrl = props.context.pageContext.web.absoluteUrl;

  // v20.0 (Audit): Context-Value memoizen. Die Rollen-Methoden schließen nur
  // über `roles` (find-Lookups) + stabile Service-/Setter-Referenzen — `roles`
  // ist als Dependency enthalten, daher bleiben die Closures frisch. Die
  // Funktions-Identitäten selbst sind bewusst keine Dependencies (sie werden
  // pro Render neu erzeugt und würden das Memo wirkungslos machen).
  const value = React.useMemo<RoleContextType>(() => ({
    roles, currentUserRole, isRolesLoading,
    isAdmin, isOrganizer, canCreateEvents, isPowerUser, siteUrl,
    originalIsAdmin, isImpersonating, previewAsUser, setPreviewAsUser, isFA,
    addRole, updateRole, setPowerUser, updateRoleLocation, removeRole, refreshRoles, searchUser, searchUsers, searchGroups, getGroupMembers, searchUsersByLocation, getEmployeeData,
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [roles, currentUserRole, isRolesLoading, isImpersonating, previewAsUser, siteUrl]);

  return React.createElement(
    RoleContext.Provider,
    { value },
    props.children
  );
}

export function useRoles(): RoleContextType {
  const ctx = React.useContext(RoleContext);
  if (!ctx) throw new Error('useRoles must be used within RoleProvider');
  return ctx;
}
