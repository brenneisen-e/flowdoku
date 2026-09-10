/**
 * Rollen der AI Use Case Platform.
 *
 * Drei Rollen, Aufbau wie DEX:
 *   Admin   — alles, inklusive Rollenverwaltung
 *   Kurator — Use Cases anlegen und pflegen
 *   User    — sehen und aufrufen
 *
 * Die zwei Regeln, die DEX teuer gelernt hat, gelten hier von Anfang an:
 *
 * 1. Ein Lesefehler befoerdert niemanden. `getRoles` liefert bei 403 `null`,
 *    nicht `[]` — und nur eine FRISCH angelegte Liste macht die erste Person
 *    zum Admin. In DEX war „getRoles ist leer" einmal gleichbedeutend mit
 *    „Erstinstallation", und ein 403 machte damit jeden Aufrufer zum Admin
 *    (v6.34).
 *
 * 2. Eine Rolle wirkt nur mit Leserecht auf der Rollenliste. Deshalb setzt
 *    jede Zuweisung Read nach — und meldet, wenn das nicht geklappt hat,
 *    statt Erfolg zu behaupten (v30.80/v30.85).
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';
import { RoleAssignment, UserRole } from '../types';
import { isCurrentUser } from '../utils/sessionIdentities';

interface RoleContextType {
  roles: RoleAssignment[];
  currentUserRole: UserRole;
  isRolesLoading: boolean;
  /** `forbidden` = die Rollenliste ist nicht lesbar. Wer darin steht, ist trotzdem „User". */
  rolesReadStatus: 'loading' | 'ok' | 'forbidden' | 'error';
  isAdmin: boolean;
  isKurator: boolean;
  /** Vorschau „als Nutzer sehen" — ohne Identitaetswechsel, endet beim Neuladen. */
  previewAsUser: boolean;
  setPreviewAsUser: (on: boolean) => void;
  addRole: (email: string, name: string, role: UserRole) => Promise<boolean>;
  updateRole: (itemId: number, role: UserRole) => Promise<boolean>;
  removeRole: (itemId: number) => Promise<boolean>;
  refreshRoles: () => Promise<void>;
  /** Welche Rechte beim letzten Schreiben NICHT gesetzt werden konnten. */
  lastRightsMissing: () => string[];
  siteUrl: string;
  service: SharePointService;
}

const RoleContext = React.createContext<RoleContextType | undefined>(undefined);

function normalizeRole(v: string): UserRole {
  if (v === 'Admin' || v === 'Kurator' || v === 'User') return v;
  return 'User';
}

/** Hoehere Zahl = mehr Rechte. Fuer „nie implizit herabstufen". */
function roleRank(r: UserRole): number {
  return r === 'Admin' ? 3 : r === 'Kurator' ? 2 : 1;
}

export function RoleProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [roles, setRoles] = React.useState<RoleAssignment[]>([]);
  const [currentUserRole, setCurrentUserRole] = React.useState<UserRole>('User');
  const [isRolesLoading, setIsRolesLoading] = React.useState(true);
  const [rolesReadStatus, setRolesReadStatus] = React.useState<'loading' | 'ok' | 'forbidden' | 'error'>('loading');
  const [previewAsUser, setPreviewAsUser] = React.useState(false);
  const rightsMissingRef = React.useRef<string[]>([]);

  const service = React.useMemo(() => new SharePointService(props.context), []);
  const myEmail = props.context.pageContext.user.email;
  const myName = props.context.pageContext.user.displayName || myEmail;

  const mapRows = (rows: Array<{ Id: number; Title: string; UserName: string; Role: string; AssignedBy: string; AssignedDate: string }>): RoleAssignment[] =>
    rows.map(r => ({
      id: r.Id,
      userEmail: r.Title || '',
      userName: r.UserName || '',
      role: normalizeRole(r.Role),
      assignedBy: r.AssignedBy || '',
      assignedDate: r.AssignedDate || '',
    }));

  React.useEffect(() => {
    init().catch(() => setIsRolesLoading(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function init(): Promise<void> {
    await service.ensureRolesList();
    const rows = await service.getRoles();

    if (rows === null) {
      // Nicht lesbar. KEINE Rolle zuordnen — sonst umgeht ein
      // voruebergehender Fehler die gesamte Zugriffssteuerung. Lieber sieht
      // jemand zu wenig als versehentlich alles.
      setRoles([]);
      setCurrentUserRole('User');
      setRolesReadStatus(service.lastRolesReadStatus === 403 ? 'forbidden' : 'error');
      setIsRolesLoading(false);
      console.warn('[AIUC] Rollenliste nicht lesbar — Rolle bleibt „User".');
      return;
    }

    setRolesReadStatus('ok');

    if (rows.length === 0) {
      // Erstinstallation: Die Rollenliste ist LEER und WAR LESBAR.
      //
      // In DEX haengt dieser Zweig zusaetzlich an `isNewlyCreated`, und das
      // aus gutem Grund: Dort lieferte `getRoles` bei einem 403 ein leeres
      // Array, „leer" hiess also auch „darf nicht lesen", und jeder Aufrufer
      // wurde Admin (v6.34).
      //
      // Hier kann das nicht passieren: `getRoles` liefert bei JEDEM Fehler
      // `null`, und dieser Zweig ist erst hinter der `rows === null`-Pruefung
      // erreichbar. „Leer" heisst hier also wirklich leer.
      //
      // Warum das nicht mehr nur an `isNewlyCreated` haengt: Beim ersten
      // Live-Versuch am 10.09.2026 wurde die Liste angelegt, das Anlegen der
      // Spalten scheiterte still, und der Aufrufer blieb „User" — ohne
      // Zugang zur Rollenverwaltung, mit der er sich haette helfen koennen.
      // Eine leere Rollenliste ohne Admin ist eine Sackgasse.
      console.warn('[AIUC] Rollenliste ist leer — die angemeldete Person wird Admin (Erstinstallation).');
      await service.addRole(myEmail, myName, 'Admin', 'System (Erstinstallation)');
      await service.grantFullControlOnRolesList(myEmail);
      setCurrentUserRole('Admin');
      setRoles([{
        id: 0, userEmail: myEmail, userName: myName, role: 'Admin',
        assignedBy: 'System (Erstinstallation)', assignedDate: new Date().toISOString(),
      }]);
      setIsRolesLoading(false);
      return;
    }

    const mapped = mapRows(rows);
    setRoles(mapped);
    // Ueber ALLE Schreibweisen der angemeldeten Person, nicht nur
    // `pageContext.user.email` — steht die Zeile unter dem Alias, waere die
    // Person sonst „User" (DEX v30.81).
    const mine = mapped.filter(r => isCurrentUser(props.context, r.userEmail))[0];
    setCurrentUserRole(mine ? mine.role : 'User');
    setIsRolesLoading(false);
  }

  const refreshRoles = React.useCallback(async (): Promise<void> => {
    const rows = await service.getRoles();
    if (rows === null) {
      // Bestehenden Stand STEHEN LASSEN. Ein Netzwerkfehler mitten in der
      // Nutzung darf niemanden herabstufen.
      console.warn('[AIUC] Rollen nicht lesbar — bestehender Stand bleibt.');
      return;
    }
    setRolesReadStatus('ok');
    const mapped = mapRows(rows);
    setRoles(mapped);
    const mine = mapped.filter(r => isCurrentUser(props.context, r.userEmail))[0];
    setCurrentUserRole(mine ? mine.role : 'User');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service]);

  /** Rechte zur Rolle setzen. Rueckgabe: was NICHT gesetzt werden konnte. */
  async function applyRights(email: string, role: UserRole): Promise<string[]> {
    const missing: string[] = [];
    if (role === 'Admin') {
      if (!(await service.grantFullControlOnRolesList(email))) missing.push('Rollenliste (Vollzugriff)');
      missing.push(...await service.grantCuratorPermissions(email));
    } else if (role === 'Kurator') {
      if (!(await service.grantReadOnRolesList(email))) missing.push('Rollenliste (Lesen)');
      missing.push(...await service.grantCuratorPermissions(email));
    } else {
      // Herabstufung auf User: Rechte spiegelbildlich entziehen. Wer das
      // weglaesst, laesst einen Ex-Kurator weiter schreiben — in DEX war das
      // der Fall „einmal Organizer, immer Zugriff" (v30.67).
      if (!(await service.revokeAllAccess(email))) missing.push('Entzug unvollstaendig');
    }
    return missing;
  }

  const addRole = React.useCallback(async (email: string, name: string, role: UserRole): Promise<boolean> => {
    rightsMissingRef.current = [];
    const lc = (email || '').trim().toLowerCase();
    const existing = roles.filter(r => (r.userEmail || '').trim().toLowerCase() === lc)[0];

    if (existing) {
      // Nie implizit herabstufen: `addRole` heisst „diese Rechte
      // sicherstellen". Wer schon mehr hat, verloere sie sonst still
      // (DEX v30.67).
      if (roleRank(existing.role) > roleRank(role)) {
        console.warn(`[AIUC] ${lc} hat bereits ${existing.role} — ${role} waere eine Herabstufung und wird NICHT gesetzt.`);
        return false;
      }
      if (existing.role === role) {
        // Rechte trotzdem nachsetzen: Sie koennen beim ersten Mal an der
        // Drosselung gescheitert sein.
        rightsMissingRef.current = await applyRights(email, role);
        return true;
      }
      return updateRole(existing.id, role);
    }

    const ok = await service.addRole(email, name, role, myName);
    if (ok) {
      rightsMissingRef.current = await applyRights(email, role);
      await refreshRoles();
    }
    return ok;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roles, service, refreshRoles, myName]);

  const updateRole = React.useCallback(async (itemId: number, role: UserRole): Promise<boolean> => {
    rightsMissingRef.current = [];
    const before = roles.filter(r => r.id === itemId)[0];
    const ok = await service.updateRole(itemId, role);
    if (ok && before) {
      rightsMissingRef.current = await applyRights(before.userEmail, role);
      await refreshRoles();
    }
    return ok;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roles, service, refreshRoles]);

  const removeRole = React.useCallback(async (itemId: number): Promise<boolean> => {
    rightsMissingRef.current = [];
    const entry = roles.filter(r => r.id === itemId)[0];
    const ok = await service.deleteRole(itemId);
    if (ok) {
      await refreshRoles();
      if (entry && !(await service.revokeAllAccess(entry.userEmail))) {
        rightsMissingRef.current = ['Entzug unvollstaendig'];
      }
    }
    return ok;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roles, service, refreshRoles]);

  const effectiveRole: UserRole = previewAsUser ? 'User' : currentUserRole;
  const isAdmin = effectiveRole === 'Admin';
  const isKurator = effectiveRole === 'Kurator' || isAdmin;

  const value = React.useMemo<RoleContextType>(() => ({
    roles, currentUserRole, isRolesLoading, rolesReadStatus,
    isAdmin, isKurator, previewAsUser, setPreviewAsUser,
    addRole, updateRole, removeRole, refreshRoles,
    lastRightsMissing: () => rightsMissingRef.current,
    siteUrl: props.context.pageContext.web.absoluteUrl,
    service,
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [roles, currentUserRole, isRolesLoading, rolesReadStatus, previewAsUser, addRole, updateRole, removeRole, refreshRoles, service]);

  return React.createElement(RoleContext.Provider, { value }, props.children);
}

export function useRoles(): RoleContextType {
  const ctx = React.useContext(RoleContext);
  if (!ctx) throw new Error('useRoles muss innerhalb des RoleProvider stehen');
  return ctx;
}
