/**
 * v29.25: Sperre der Selbst-Abmeldung — zweistufige Organizer-Option.
 *
 * Der Wizard fragt zweistufig: „Abmeldung durch User ermöglichen?" (Default
 * ja; bei Nein gibt es keine Abmeldefrist, Piggyback `_noSelfCancel`) und —
 * nur bei Ja mit gesetzter Frist — „auch nach der Abmeldefrist erlauben?"
 * (Default ja = Late-Cancel mit Organizer-Mail; bei Nein Piggyback
 * `_noCancelAfterDeadline`).
 *
 * Die Auswertung steht bewusst an EINER Stelle, weil vier Ansichten sie
 * treffen müssen und jede Abweichung ein Widerspruch für den Teilnehmer
 * wäre: „Meine Events" (Abmelde-Knopf + Deep-Link aus der Mail), die
 * Sub-Event-Liste darunter, die Anmeldeseite (Abwählen gebuchter Sub-Events
 * beim Absenden) und das Organizer Center (dort schaltet dieselbe Regel den
 * No-Show-Knopf frei).
 */
import { DeloitteEvent } from '../types';

/**
 * Warum die Person sich nicht selbst abmelden darf — oder null, wenn sie es
 * darf. Beide Flags sind event-weit (liegen auf dem Haupt-/Klammer-Event);
 * ein Sub-Event erbt sie über `parent`.
 *  - 'always': Selbst-Abmeldung ist für dieses Event komplett deaktiviert.
 *  - 'afterDeadline': nur noch bis zur Abmeldefrist — maßgeblich ist die
 *    eigene Frist des (Sub-)Events, ohne eigene Frist die des Parents.
 *    Ohne gesetzte Frist sperrt dieses Flag nichts.
 */
export function selfCancelLockReason(
  ev: DeloitteEvent | undefined | null,
  parent?: DeloitteEvent | undefined | null,
): 'always' | 'afterDeadline' | null {
  if (!ev) return null;
  if (ev.noSelfCancel || (parent && parent.noSelfCancel)) return 'always';
  const flag = !!(ev.noCancelAfterDeadline || (parent && parent.noCancelAfterDeadline));
  if (!flag) return null;
  const deadline = ev.lastDeregisterDate || (parent ? parent.lastDeregisterDate : '') || '';
  if (!deadline) return null;
  const d = new Date(deadline);
  return (!isNaN(d.getTime()) && d.getTime() < Date.now()) ? 'afterDeadline' : null;
}

/** true = die Person darf sich bei diesem Event NICHT mehr selbst abmelden. */
export function selfCancelLocked(
  ev: DeloitteEvent | undefined | null,
  parent?: DeloitteEvent | undefined | null,
): boolean {
  return selfCancelLockReason(ev, parent) !== null;
}
