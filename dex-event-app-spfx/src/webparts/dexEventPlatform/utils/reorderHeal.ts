/**
 * v30.80 — Merker für nicht geschriebene DEX_IDReorder-Aufträge.
 *
 * Der Reorder-Auftrag ist bei jeder Abmeldung der EINZIGE Anstoß für
 * Neu-Nummerierung und Nachrücken (seit v6.7 macht die App beides nicht
 * selbst) — und er ist der fünfte bis siebte Schreibvorgang derselben
 * Abmeldung, also der, den die Drosselung (429) am ehesten trifft. Bis
 * v30.79 hieß ein gescheiterter POST an sechs von zehn Stellen nur
 * console.warn: kein Lauf in der Run history, keine Zeile in der Queue,
 * niemand rückte nach (Befunde 07.09.2026, zweimal).
 *
 * Jetzt: Scheitert der Auftrag trotz mehrerer Versuche, landet er hier als
 * Merker in localStorage. Abgearbeitet wird er (a) beim nächsten
 * erfolgreichen Reorder-Schreibvorgang derselben Sitzung und (b) still beim
 * nächsten App-Start (EventContext, 25 s nach dem Boot). Ein doppelter
 * Auftrag ist harmlos — der Flow nummeriert idempotent und rückt nur nach,
 * wenn ein Platz frei ist.
 */

export interface PendingReorder {
  eventId: string;
  eventNumber: number;
  subsiteUrl: string;
  eventTitle: string;
  cancelledName?: string;
  cancelledEmail?: string;
  /** Aufrufstelle (self-cancel, organizer-cancel, …) — für das Event-Log. */
  via: string;
  ts: number;
}

const KEY = 'dex_pending_id_reorder_v1';
const MAX_AGE_MS = 14 * 24 * 60 * 60 * 1000;

export function readPendingReorders(): PendingReorder[] {
  try {
    const raw = window.localStorage.getItem(KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return [];
    return arr.filter((e: PendingReorder) =>
      !!e && !!e.eventId && !!e.subsiteUrl && (Date.now() - (e.ts || 0)) < MAX_AGE_MS);
  } catch { return []; }
}

function writePendingReorders(list: PendingReorder[]): void {
  try {
    if (list.length === 0) window.localStorage.removeItem(KEY);
    // Deckel: mehr als 30 offene Aufträge sind ein anderes Problem (Rechte).
    else window.localStorage.setItem(KEY, JSON.stringify(list.slice(-30)));
  } catch { /* localStorage gesperrt → nur Event-Log + Admin-Aktion */ }
}

/** Ein Auftrag je Subsite reicht — der Flow arbeitet die ganze Liste ab. */
export function addPendingReorder(entry: PendingReorder): void {
  const rest = readPendingReorders().filter(e => e.subsiteUrl.toLowerCase() !== entry.subsiteUrl.toLowerCase());
  rest.push(entry);
  writePendingReorders(rest);
}

export function removePendingReorder(subsiteUrl: string): void {
  writePendingReorders(readPendingReorders().filter(e => e.subsiteUrl.toLowerCase() !== subsiteUrl.toLowerCase()));
}
