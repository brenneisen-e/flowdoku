/**
 * v30.67: EINE Ableitung für „welche Rolle trägt mehr Rechte?" und „hat die
 * Rolle bereits Organizer-Rechte?".
 *
 * Warum: Der Personenkreis „hat schon Organizer-Rechte" war an drei Stellen
 * unabhängig voneinander ausformuliert (`roleFilter` in
 * services/events/organizer.ts, `elevated` in context/actions/organizerRoles.ts,
 * `alreadyEntitled` im OrganizerRequestsBanner). Als v30.60 F&A zu einem
 * Organizer-Superset machte, wurde nur `isOrganizer` im RoleContext
 * nachgezogen — die Freigabe eines „Organizer werden"-Antrags stufte eine
 * F&A-Person deshalb still auf Organizer herab. Wer einen neuen Rollenwert
 * einführt, ergänzt ihn HIER und nirgends sonst.
 */
import type { UserRole } from '../types';

/**
 * Rang einer Rolle — je höher, desto mehr Rechte. Legacy-Werte (`SuperAdmin`,
 * `EventAdmin`) zählen wie ihre heutigen Entsprechungen: Sie können in
 * DEX_Roles stehen bleiben, weil die Hintergrund-Migration in `initRoles`
 * still scheitern kann (siehe roleFilter in organizer.ts).
 *
 * F&A steht ÜBER Organizer: „alles, was Organizer können, plus die
 * Abrechnungsfunktion und das F&A Center" (Nutzer-Ansage 01.09.2026).
 */
export function roleRank(role: UserRole | string | undefined | null): number {
  switch (role) {
    case 'Admin':
    case 'IT-Admin':
    case 'SuperAdmin':
      return 3;
    case 'F&A':
      return 2;
    case 'Organizer':
    case 'EventAdmin':
      return 1;
    default:
      return 0;
  }
}

/** Admin und IT-Admin: gleiche App-Rechte (v26.33), gleiche SharePoint-Rechte. */
export function isAdminRole(role: UserRole | string | undefined | null): boolean {
  return roleRank(role) >= 3;
}

/** Alles, was Events anlegen und bearbeiten darf — dieselbe Menge wie
 *  `RoleContext.isOrganizer` (Organizer, F&A, Admin, IT-Admin). */
export function hasOrganizerRights(role: UserRole | string | undefined | null): boolean {
  return roleRank(role) >= 1;
}
