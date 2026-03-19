/**
 * Role Context - Rollenverwaltung ueber SharePoint-Liste
 *
 * Erstellt beim ersten Start automatisch die DEX_Roles-Liste
 * auf dem SharePoint. SuperAdmins koennen andere User berechtigen.
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
  isSuperAdmin: boolean;
  isEventAdmin: boolean;
  canCreateEvents: boolean;
  siteUrl: string;
  addRole: (userEmail: string, userName: string, role: UserRole, location: string) => Promise<boolean>;
  updateRole: (itemId: number, newRole: UserRole) => Promise<boolean>;
  removeRole: (itemId: number) => Promise<boolean>;
  refreshRoles: () => Promise<void>;
  searchUser: (email: string) => Promise<{ displayName: string; location: string } | null>;
  searchUsers: (query: string) => Promise<Array<{ email: string; displayName: string; location: string }>>;
}

const RoleContext = React.createContext<RoleContextType | undefined>(undefined);

export function RoleProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [roles, setRoles] = React.useState<RoleAssignment[]>([]);
  const [currentUserRole, setCurrentUserRole] = React.useState<UserRole>('User');
  const [isRolesLoading, setIsRolesLoading] = React.useState<boolean>(true);
  const spService = React.useMemo(() => new SharePointService(props.context), []);

  const currentUserEmail = props.context.pageContext.user.email;
  const currentUserName = props.context.pageContext.user.displayName;

  // Rollen-Liste erstellen und laden beim Start
  React.useEffect(() => {
    initRoles().catch(() => setIsRolesLoading(false));
  }, []);

  async function initRoles(): Promise<void> {
    // Liste automatisch erstellen falls noetig
    await spService.ensureRolesList();

    // Rollen laden
    const spRoles = await spService.getRoles();

    // Pruefen ob aktueller User schon in der Liste ist
    const myRole = spRoles.find(
      r => r.Title && r.Title.toLowerCase() === currentUserEmail.toLowerCase()
    );

    if (spRoles.length === 0) {
      // Erste Ausfuehrung: aktuellen User als SuperAdmin eintragen
      await spService.addRole(
        currentUserEmail,
        currentUserName,
        'SuperAdmin',
        '',
        'System (Erstinstallation)'
      );
      setCurrentUserRole('SuperAdmin');
      setRoles([{
        id: 0,
        userEmail: currentUserEmail,
        userName: currentUserName,
        role: 'SuperAdmin',
        location: '',
        assignedBy: 'System (Erstinstallation)',
        assignedDate: new Date().toISOString(),
      }]);
    } else {
      // Rollen in App-Format umwandeln
      const mapped: RoleAssignment[] = spRoles.map(r => ({
        id: r.Id,
        userEmail: r.Title || '',
        userName: r.UserName || '',
        role: (r.Role as UserRole) || 'User',
        location: r.UserLocation || '',
        assignedBy: r.AssignedBy || '',
        assignedDate: r.AssignedDate || '',
      }));
      setRoles(mapped);
      setCurrentUserRole(myRole ? (myRole.Role as UserRole) : 'User');
    }

    setIsRolesLoading(false);
  }

  async function refreshRoles(): Promise<void> {
    const spRoles = await spService.getRoles();
    const mapped: RoleAssignment[] = spRoles.map(r => ({
      id: r.Id,
      userEmail: r.Title || '',
      userName: r.UserName || '',
      role: (r.Role as UserRole) || 'User',
      location: r.UserLocation || '',
      assignedBy: r.AssignedBy || '',
      assignedDate: r.AssignedDate || '',
    }));
    setRoles(mapped);

    const myRole = spRoles.find(
      r => r.Title && r.Title.toLowerCase() === currentUserEmail.toLowerCase()
    );
    setCurrentUserRole(myRole ? (myRole.Role as UserRole) : 'User');
  }

  async function addRole(
    userEmail: string, userName: string, role: UserRole, location: string
  ): Promise<boolean> {
    const success = await spService.addRole(userEmail, userName, role, location, currentUserName);
    if (success) {
      // SP-Berechtigung setzen: EventAdmin bekommt Read auf die Liste
      if (role === 'EventAdmin') {
        await spService.grantReadOnRolesList(userEmail);
      }
      await refreshRoles();
    }
    return success;
  }

  async function updateRole(itemId: number, newRole: UserRole): Promise<boolean> {
    // Alten User finden fuer Berechtigungsaenderung
    const oldRole = roles.find(r => r.id === itemId);
    const success = await spService.updateRole(itemId, newRole);
    if (success && oldRole) {
      if (newRole === 'EventAdmin') {
        // Read-Berechtigung geben
        await spService.grantReadOnRolesList(oldRole.userEmail);
      } else if (newRole === 'User') {
        // Berechtigung entziehen
        await spService.revokeAccessOnRolesList(oldRole.userEmail);
      }
      await refreshRoles();
    }
    return success;
  }

  async function removeRole(itemId: number): Promise<boolean> {
    // User-Email merken bevor der Eintrag geloescht wird
    const roleEntry = roles.find(r => r.id === itemId);
    const success = await spService.deleteRole(itemId);
    if (success) {
      // Berechtigung entziehen
      if (roleEntry) {
        await spService.revokeAccessOnRolesList(roleEntry.userEmail);
      }
      await refreshRoles();
    }
    return success;
  }

  async function searchUser(email: string): Promise<{ displayName: string; location: string } | null> {
    return spService.searchUserByEmail(email);
  }

  async function searchUsers(query: string): Promise<Array<{ email: string; displayName: string; location: string }>> {
    return spService.searchUsers(query);
  }

  const isSuperAdmin = currentUserRole === 'SuperAdmin';
  const isEventAdmin = currentUserRole === 'EventAdmin' || currentUserRole === 'SuperAdmin';
  const canCreateEvents = isEventAdmin;
  const siteUrl = props.context.pageContext.web.absoluteUrl;

  return React.createElement(
    RoleContext.Provider,
    {
      value: {
        roles, currentUserRole, isRolesLoading,
        isSuperAdmin, isEventAdmin, canCreateEvents, siteUrl,
        addRole, updateRole, removeRole, refreshRoles, searchUser, searchUsers,
      },
    },
    props.children
  );
}

export function useRoles(): RoleContextType {
  const ctx = React.useContext(RoleContext);
  if (!ctx) throw new Error('useRoles must be used within RoleProvider');
  return ctx;
}
