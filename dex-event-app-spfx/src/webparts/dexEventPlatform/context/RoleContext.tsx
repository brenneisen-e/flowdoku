/**
 * Role Context - Rollenverwaltung ueber SharePoint-Liste
 *
 * Erstellt beim ersten Start automatisch die DEX_Roles-Liste
 * auf dem SharePoint. Admins koennen andere User berechtigen.
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
  isOrganizer: boolean;
  canCreateEvents: boolean;
  siteUrl: string;
  addRole: (userEmail: string, userName: string, role: UserRole, location: string) => Promise<boolean>;
  updateRole: (itemId: number, newRole: UserRole) => Promise<boolean>;
  removeRole: (itemId: number) => Promise<boolean>;
  refreshRoles: () => Promise<void>;
  searchUser: (email: string) => Promise<{ displayName: string; location: string } | null>;
  searchUsers: (query: string) => Promise<Array<{ email: string; displayName: string; location: string }>>;
}

// Migration: alte SP-Werte auf neue mappen
function migrateRole(spRole: string): UserRole {
  if (spRole === 'SuperAdmin') return 'Admin';
  if (spRole === 'EventAdmin') return 'Organizer';
  if (spRole === 'Admin' || spRole === 'Organizer' || spRole === 'User') return spRole as UserRole;
  return 'User';
}

const RoleContext = React.createContext<RoleContextType | undefined>(undefined);

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
    await spService.ensureRolesList();
    const spRoles = await spService.getRoles();

    const myRole = spRoles.find(
      r => r.Title && r.Title.toLowerCase() === currentUserEmail.toLowerCase()
    );

    if (spRoles.length === 0) {
      // Erste Ausfuehrung: aktuellen User als Admin eintragen
      await spService.addRole(currentUserEmail, currentUserName, 'Admin', '', 'System (Erstinstallation)');
      setCurrentUserRole('Admin');
      setRoles([{
        id: 0, userEmail: currentUserEmail, userName: currentUserName,
        role: 'Admin', location: '', assignedBy: 'System (Erstinstallation)',
        assignedDate: new Date().toISOString(),
      }]);
    } else {
      const mapped: RoleAssignment[] = spRoles.map(r => ({
        id: r.Id,
        userEmail: r.Title || '',
        userName: r.UserName || '',
        role: migrateRole(r.Role),
        location: r.UserLocation || '',
        assignedBy: r.AssignedBy || '',
        assignedDate: r.AssignedDate || '',
      }));
      setRoles(mapped);
      setCurrentUserRole(myRole ? migrateRole(myRole.Role) : 'User');

      // Alte Werte in SP migrieren (im Hintergrund)
      for (const r of spRoles) {
        if (r.Role === 'SuperAdmin' || r.Role === 'EventAdmin') {
          spService.updateRole(r.Id, migrateRole(r.Role)).catch(() => {});
        }
      }
    }

    setIsRolesLoading(false);
  }

  async function refreshRoles(): Promise<void> {
    const spRoles = await spService.getRoles();
    const mapped: RoleAssignment[] = spRoles.map(r => ({
      id: r.Id, userEmail: r.Title || '', userName: r.UserName || '',
      role: migrateRole(r.Role), location: r.UserLocation || '',
      assignedBy: r.AssignedBy || '', assignedDate: r.AssignedDate || '',
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
      if (role === 'Admin') {
        await spService.grantFullControlOnRolesList(userEmail);
        await spService.grantFullControlOnEventsList(userEmail);
      } else if (role === 'Organizer') {
        await spService.grantReadOnRolesList(userEmail);
        await spService.grantFullControlOnEventsList(userEmail);
      }
      await refreshRoles();
    }
    return success;
  }

  async function updateRole(itemId: number, newRole: UserRole): Promise<boolean> {
    const oldRole = roles.find(r => r.id === itemId);
    const success = await spService.updateRole(itemId, newRole);
    if (success && oldRole) {
      if (newRole === 'Admin') {
        await spService.grantFullControlOnRolesList(oldRole.userEmail);
        await spService.grantFullControlOnEventsList(oldRole.userEmail);
      } else if (newRole === 'Organizer') {
        await spService.grantReadOnRolesList(oldRole.userEmail);
        await spService.grantFullControlOnEventsList(oldRole.userEmail);
      } else if (newRole === 'User') {
        await spService.revokeAccessOnRolesList(oldRole.userEmail);
        await spService.revokeAccessOnEventsList(oldRole.userEmail);
      }
      await refreshRoles();
    }
    return success;
  }

  async function removeRole(itemId: number): Promise<boolean> {
    const roleEntry = roles.find(r => r.id === itemId);
    const success = await spService.deleteRole(itemId);
    if (success) {
      await refreshRoles();
      if (roleEntry) {
        spService.revokeAccessOnRolesList(roleEntry.userEmail).catch(() => {});
        spService.revokeAccessOnEventsList(roleEntry.userEmail).catch(() => {});
      }
    }
    return success;
  }

  async function searchUser(email: string): Promise<{ displayName: string; location: string } | null> {
    return spService.searchUserByEmail(email);
  }

  async function searchUsers(query: string): Promise<Array<{ email: string; displayName: string; location: string }>> {
    return spService.searchUsers(query);
  }

  const isAdmin = currentUserRole === 'Admin';
  const isOrganizer = currentUserRole === 'Organizer' || currentUserRole === 'Admin';
  const canCreateEvents = isOrganizer;
  const siteUrl = props.context.pageContext.web.absoluteUrl;

  return React.createElement(
    RoleContext.Provider,
    {
      value: {
        roles, currentUserRole, isRolesLoading,
        isAdmin, isOrganizer, canCreateEvents, siteUrl,
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
