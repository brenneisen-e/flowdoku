/**
 * v32.29: Ist die angemeldete Person im Test-Team dieses Events?
 *
 * Das Test-Team (`_testTeam`, Wizard Schritt „Details") darf einen Entwurf
 * sehen und sich probeweise anmelden. Bis v32.28 prüfte das nur der
 * Entwurfs-Filter der Startseite — und nur gegen die SMTP-Adresse. Danach
 * warf der Verteiler-Filter die Person wieder hinaus, die Anmeldeseite
 * sperrte vor „Aktiv ab", und die Termine eines Entwurfs blieben unsichtbar
 * (sie tragen kein eigenes `_testTeam`). Deshalb EIN Helfer für alle Stellen:
 * beide Schreibweisen der Person (SMTP + Adresse im loginName, s.
 * utils/sessionIdentities) und bei Terminen die Liste der Klammer.
 */
import { DeloitteEvent } from '../types';

export interface TestTeamUserLike {
  email?: string;
  /** loginName (`i:0#.f|membership|user@domain`) — `currentUser.id`. */
  id?: string;
}

function identitaeten(user: TestTeamUserLike): Set<string> {
  const out = new Set<string>();
  const e = (user.email || '').toLowerCase().trim();
  if (e) out.add(e);
  const m = (user.id || '').toLowerCase().match(/[^|]+@[^|\s]+$/);
  if (m) out.add(m[0].trim());
  return out;
}

export function istImTestTeam(
  user: TestTeamUserLike,
  ev: Pick<DeloitteEvent, 'testTeamEmails'> | null | undefined,
  klammer?: Pick<DeloitteEvent, 'testTeamEmails'> | null
): boolean {
  const ids = identitaeten(user);
  if (ids.size === 0) return false;
  const liste = [...((ev && ev.testTeamEmails) || []), ...((klammer && klammer.testTeamEmails) || [])];
  return liste.some(x => ids.has((x || '').toLowerCase().trim()));
}
