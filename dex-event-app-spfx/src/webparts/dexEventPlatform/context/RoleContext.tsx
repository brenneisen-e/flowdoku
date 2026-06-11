/**
 * Role Context - Rollenverwaltung ueber SharePoint-Liste
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
  isOrganizer: boolean;
  canCreateEvents: boolean;
  siteUrl: string;
  addRole: (userEmail: string, userName: string, role: UserRole, location: string) => Promise<boolean>;
  updateRole: (itemId: number, newRole: UserRole) => Promise<boolean>;
  setPowerUser: (itemId: number, isPowerUser: boolean) => Promise<boolean>;
  updateRoleLocation: (itemId: number, location: string) => Promise<boolean>;
  removeRole: (itemId: number) => Promise<boolean>;
  refreshRoles: () => Promise<void>;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string } | null>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>;
  searchGroups: (query: string) => Promise<Array<{ email: string; displayName: string }>>;
  getGroupMembers: (groupEmail: string) => Promise<{ groupName: string; members: Array<{ email: string; displayName: string; firstName?: string; lastName?: string; jobTitle?: string; location?: string }> } | null>;
  searchUsersByLocation: (location: string) => Promise<Array<{ email: string; displayName: string; firstName: string; lastName: string; location: string; jobTitle: string }>>;
}

// Migration: alte SP-Werte auf neue mappen
function migrateRole(spRole: string): UserRole {
  if (spRole === 'SuperAdmin') return 'Admin';
  if (spRole === 'EventAdmin') return 'Organizer';
  if (spRole === 'Admin' || spRole === 'Organizer' || spRole === 'User') return spRole as UserRole;
  return 'User';
}

export const RoleContext = React.createContext<RoleContextType | undefined>(undefined);

export function RoleProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [roles, setRoles] = React.useState<RoleAssignment[]>([]);
  const [currentUserRole, setCurrentUserRole] = React.useState<UserRole>('User');
  const [isRolesLoading, setIsRolesLoading] = React.useState<boolean>(true);
  const spService = React.useMemo(() => new SharePointService(props.context), []);

  const currentUserEmail = props.context.pageContext.user.email;
  const currentUserName = props.context.pageContext.user.displayName;

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
    console.log(`[DEX][perf][roles] ensureRolesList = ${Math.round(performance.now() - tEns)} ms`);
    const tGet = performance.now();
    const spRoles = await spService.getRoles();
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][roles] getRoles = ${Math.round(performance.now() - tGet)} ms (n=${spRoles ? spRoles.length : 'null'})`);
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
    const success = await spService.addRole(userEmail, userName, role, location, currentUserName);
    if (success) {
      try {
        if (role === 'Admin') {
          await spService.grantFullControlOnRolesList(userEmail);
          await spService.grantFullControlOnEventsList(userEmail);
          await spService.grantOrganizerPermissions(userEmail); // Site-Rechte für Subsite-Erstellung
        } else if (role === 'Organizer') {
          await spService.grantReadOnRolesList(userEmail);
          await spService.grantOrganizerPermissions(userEmail);
        }
      } catch (err) { console.warn('[DEX] permission grant for addRole failed (best-effort):', err); }
      await refreshRoles();
    }
    return success;
  }

  async function updateRole(itemId: number, newRole: UserRole): Promise<boolean> {
    const oldRole = roles.find(r => r.id === itemId);
    const success = await spService.updateRole(itemId, newRole);
    if (success && oldRole) {
      try {
        if (newRole === 'Admin') {
          await spService.grantFullControlOnRolesList(oldRole.userEmail);
          await spService.grantFullControlOnEventsList(oldRole.userEmail);
          await spService.grantOrganizerPermissions(oldRole.userEmail);
        } else if (newRole === 'Organizer') {
          await spService.grantReadOnRolesList(oldRole.userEmail);
          await spService.grantOrganizerPermissions(oldRole.userEmail);
        } else if (newRole === 'User') {
          await spService.revokeAccessOnRolesList(oldRole.userEmail);
          await spService.revokeAccessOnEventsList(oldRole.userEmail);
        }
      } catch (err) { console.warn('[DEX] permission grant for updateRole failed (best-effort):', err); }
      await refreshRoles();
    }
    return success;
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
    if (success) {
      await refreshRoles();
      if (roleEntry) {
        spService.revokeAccessOnRolesList(roleEntry.userEmail).catch(err => console.warn('[DEX] revokeAccessOnRolesList failed:', err));
        spService.revokeAccessOnEventsList(roleEntry.userEmail).catch(err => console.warn('[DEX] revokeAccessOnEventsList failed:', err));
      }
    }
    return success;
  }

  async function updateRoleLocation(itemId: number, location: string): Promise<boolean> {
    const success = await spService.updateRoleLocation(itemId, location);
    if (success) await refreshRoles();
    return success;
  }

  async function searchUser(email: string): Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string } | null> {
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

  async function searchUsersByLocation(location: string): Promise<Array<{ email: string; displayName: string; firstName: string; lastName: string; location: string; jobTitle: string }>> {
    return spService.searchUsersByLocation(location);
  }

  // v12.7: Demo-Impersonation. Admins können in den Header-User-Menü auf
  // „Demo: als User testen" klicken, einen Standort wählen, danach
  // agiert die App so als wäre der User ein normaler 'User' am gewählten
  // Standort. Wird in localStorage gehalten, damit das auch über
  // Reload bestehen bleibt. Beenden via Banner-Klick oben.
  const isImpersonating = typeof window !== 'undefined' && !!window.localStorage?.getItem('dex_demo_impersonation');
  const effectiveRole: UserRole = isImpersonating ? 'User' : currentUserRole;
  const isAdmin = effectiveRole === 'Admin';
  const isOrganizer = effectiveRole === 'Organizer' || effectiveRole === 'Admin';
  const canCreateEvents = isOrganizer;
  const originalIsAdmin = currentUserRole === 'Admin';
  const siteUrl = props.context.pageContext.web.absoluteUrl;

  // v20.0 (Audit): Context-Value memoizen. Die Rollen-Methoden schließen nur
  // über `roles` (find-Lookups) + stabile Service-/Setter-Referenzen — `roles`
  // ist als Dependency enthalten, daher bleiben die Closures frisch. Die
  // Funktions-Identitäten selbst sind bewusst keine Dependencies (sie werden
  // pro Render neu erzeugt und würden das Memo wirkungslos machen).
  const value = React.useMemo<RoleContextType>(() => ({
    roles, currentUserRole, isRolesLoading,
    isAdmin, isOrganizer, canCreateEvents, siteUrl,
    originalIsAdmin, isImpersonating,
    addRole, updateRole, setPowerUser, updateRoleLocation, removeRole, refreshRoles, searchUser, searchUsers, searchGroups, getGroupMembers, searchUsersByLocation,
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [roles, currentUserRole, isRolesLoading, isImpersonating, siteUrl]);

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
