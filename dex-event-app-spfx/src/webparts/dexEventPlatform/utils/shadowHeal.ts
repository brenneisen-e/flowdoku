/**
 * v30.16 — Selbstheilung für die Klammer-Schatten-Zeile.
 *
 * Die Klammer-Zeile ist immer der LETZTE Schreibvorgang einer Anmeldung —
 * genau der, der unter Drosselung (429) am ehesten scheitert. Bis v30.15
 * bekam der Teilnehmer dann einen Fehler-Dialog („SharePoint rejected the
 * last write … bitte später erneut speichern") und musste selbst handeln.
 *
 * Jetzt: Der fehlgeschlagene Klammer-Schreibvorgang wird hier als Merker
 * in localStorage abgelegt — MIT den Formular-Antworten (customData), damit
 * der Nachzug verlustfrei ist. Abgearbeitet wird er zweifach: sofort über
 * eine Hintergrund-Wiederholungskette auf der Erfolgsseite und (falls der
 * Browser vorher zugeht) still beim nächsten App-Start (EventContext).
 * registerForEvent ist für die Klammer idempotent — ein doppelter Nachzug
 * fügt nichts doppelt ein.
 */

export interface PendingShadowParent {
  eventId: string;
  customData: Record<string, string>;
  firstName: string;
  lastName: string;
  email: string;
  /** true = stellvertretende Anmeldung (Nachzug braucht die Proxy-Flags). */
  proxy?: boolean;
  ts: number;
}

const KEY = 'dex_pending_shadow_parent_v1';
const MAX_AGE_MS = 14 * 24 * 60 * 60 * 1000;

export function readPendingShadowParents(): PendingShadowParent[] {
  try {
    const raw = window.localStorage.getItem(KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return [];
    return arr.filter((e: PendingShadowParent) =>
      !!e && !!e.eventId && !!e.email && (Date.now() - (e.ts || 0)) < MAX_AGE_MS);
  } catch { return []; }
}

function writePendingShadowParents(list: PendingShadowParent[]): void {
  try {
    if (list.length === 0) window.localStorage.removeItem(KEY);
    // Sicherheits-Deckel: mehr als 20 offene Nachzüge sind ein anderes Problem.
    else window.localStorage.setItem(KEY, JSON.stringify(list.slice(-20)));
  } catch { /* localStorage gesperrt → Heilung läuft dann nur über das Organizer-Panel */ }
}

export function addPendingShadowParent(entry: PendingShadowParent): void {
  const rest = readPendingShadowParents().filter(e =>
    !(e.eventId === entry.eventId && e.email.toLowerCase() === entry.email.toLowerCase()));
  rest.push(entry);
  writePendingShadowParents(rest);
}

export function removePendingShadowParent(eventId: string, email: string): void {
  writePendingShadowParents(readPendingShadowParents().filter(e =>
    !(e.eventId === eventId && e.email.toLowerCase() === email.toLowerCase())));
}
