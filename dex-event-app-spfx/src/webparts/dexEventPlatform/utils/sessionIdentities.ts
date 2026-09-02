/**
 * v30.67: Alle E-Mail-Schreibweisen der angemeldeten Person.
 *
 * `pageContext.user.email` und die Adresse im `loginName`
 * (`i:0#.f|membership|user@domain`) können auseinanderfallen — SMTP-Adresse
 * gegen UPN/Alias (CLAUDE.md: „Die E-Mail-Adresse ist der einzige Schlüssel —
 * und sie ist nicht eindeutig"). `canRegisterForOthers` in
 * services/events/profileData.ts sammelt seit v19.6 bewusst beide; hier steht
 * dieselbe Ableitung für die Stellen, die sonst an EINER Schreibweise
 * scheitern: die Frist-Ausnahme für Organizer und der Selbstschutz beim
 * Rechte-Entzug (ein Admin darf sich beim Testen nicht selbst aussperren —
 * auch dann nicht, wenn seine DEX_Roles-Zeile unter dem Alias steht).
 */

/** Minimaler Ausschnitt des SPFx-Kontexts — so lässt sich der Helfer aus
 *  `WebPartContext` UND aus `EventService.context` füttern. */
export interface SessionContextLike {
  pageContext: { user: { email?: string; loginName?: string } };
}

/** Kleingeschriebene, getrimmte Menge aller bekannten Adressen des Users.
 *  Leer, wenn der Kontext keine Adresse hergibt. */
export function sessionIdentities(ctx: SessionContextLike): Set<string> {
  const out = new Set<string>();
  const raw = (ctx.pageContext.user.email || '').toLowerCase().trim();
  if (raw) out.add(raw);
  const loginName = (ctx.pageContext.user.loginName || '').toLowerCase();
  const m = loginName.match(/[^|]+@[^|\s]+$/);
  if (m) out.add(m[0].trim());
  return out;
}

/** Ist `email` (in irgendeiner Schreibweise) die angemeldete Person? Eine
 *  leere Adresse ist nie die eigene. */
export function isCurrentUser(ctx: SessionContextLike, email: string | undefined | null): boolean {
  const e = (email || '').toLowerCase().trim();
  if (!e) return false;
  return sessionIdentities(ctx).has(e);
}
