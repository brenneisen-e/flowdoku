/**
 * v29.31: Cache-Schlüssel und Invalidierung der „Konto inaktiv"-Prüfung.
 *
 * Die Prüfung (Graph: existiert zu dieser Adresse noch ein aktives Konto?)
 * läuft pro Event höchstens einmal am Tag und liegt in localStorage — zwei
 * Ansichten lesen sie:
 *   - die Box IM Event (AdminPage, je Event)
 *   - die Sammel-Box auf der Startseite (LandingPage, über alle Events)
 *
 * Die Startseiten-Box zeigte abgemeldete Personen bis zu 24 Stunden weiter an:
 * Sie kennt nur den gecachten Scan, nicht den aktuellen Anmeldestand. Deshalb
 * gehören die Schlüssel hierher (statt doppelt als Literal in beide Dateien)
 * und werden nach jeder Abmeldung verworfen. Die Anzeige prüft zusätzlich
 * live gegen die Teilnehmerliste — eine Abmeldung in einem anderen Browser
 * räumt diesen Cache ja nicht mit auf.
 */

/** Sammel-Scan der Startseite (alle relevanten Events). */
export const INACTIVE_SUMMARY_CACHE_KEY = 'dex_inactivesummary_v2';

/** Prüfergebnis eines einzelnen Events (Organizer Center). */
export function accountCheckCacheKey(eventId: string): string {
  return `dex_acctcheck_v2_${eventId}`;
}

/* ------------------------------------------------------------------ *
 * v31.22 — derselbe Cache, aber je ADRESSE statt je Event.
 *
 * Der Befund (11.09.2026: „dauert extrem lange mit Laden nach Klick auf ein
 * anderes Sub-Event … weiter optimieren"): Die Prüfung lag je EVENT im
 * Cache. Bei einem Klammer-Event mit drei Terminen sind es aber weitgehend
 * DIESELBEN Personen — 65, 417 und 427 Anmeldungen, und die 417 aus DINNER
 * stehen fast alle auch in MEETING. Jeder Reiterwechsel war ein Cache-Miss
 * und prüfte alle Adressen erneut: 417 / 7 = 60 Graph-Abfragen, plus bis zu
 * 60 weitere im Alias-Durchgang, und das je Reiter.
 *
 * Je Adresse gecacht kostet der zweite Reiter fast nichts — die Schnittmenge
 * ist schon geprüft. Das ist die eigentliche Ersparnis; die Parallelisierung
 * in `checkAccountsActive` macht nur den ersten Lauf schneller.
 *
 * Bewusst EIN Eintrag für alle Adressen statt einer je Adresse: 400 einzelne
 * localStorage-Schlüssel zu schreiben ist selbst eine spürbare Bremse (jeder
 * Schreibvorgang ist synchron), und der Eintrag bleibt mit ~50 Byte je
 * Adresse auch bei ein paar tausend Personen klein.
 * ------------------------------------------------------------------ */

/** Ein Eintrag je Adresse: `a` = aktiv, `t` = Zeitpunkt der Prüfung (ms). */
export interface AccountCheckEntry { a: boolean; t: number }

const BY_EMAIL_KEY = 'dex_acctcheck_by_email_v1';
const MAX_AGE_MS = 24 * 60 * 60 * 1000;

/** Liest die noch gültigen Einträge. Unlesbar/abgelaufen = nicht vorhanden. */
export function readAccountChecks(): Record<string, AccountCheckEntry> {
  try {
    const raw = window.localStorage.getItem(BY_EMAIL_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw) as Record<string, AccountCheckEntry>;
    if (!parsed || typeof parsed !== 'object') return {};
    const now = Date.now();
    const out: Record<string, AccountCheckEntry> = {};
    for (const k of Object.keys(parsed)) {
      const e = parsed[k];
      if (e && typeof e.t === 'number' && (now - e.t) < MAX_AGE_MS) out[k] = e;
    }
    return out;
  } catch { return {}; }
}

/**
 * Schreibt frisch geprüfte Adressen dazu.
 *
 * `aktiv` und `inaktiv` sind getrennte Listen, weil „nicht in der
 * Inaktiv-Liste" NICHT heißt „geprüft und aktiv": `checkAccountsActive`
 * meldet nur Adressen aus ERFOLGREICH abgefragten Batches. Wer einen
 * fehlgeschlagenen Batch als „aktiv" wegschreibt, cached einen Lesefehler
 * für 24 Stunden — dieselbe Falle wie „ein Lesefehler ist keine Null".
 */
export function writeAccountChecks(aktiv: string[], inaktiv: string[]): void {
  try {
    const cur = readAccountChecks();
    const now = Date.now();
    aktiv.forEach(e => { if (e) cur[e] = { a: true, t: now }; });
    inaktiv.forEach(e => { if (e) cur[e] = { a: false, t: now }; });
    window.localStorage.setItem(BY_EMAIL_KEY, JSON.stringify(cur));
  } catch { /* localStorage evtl. blockiert — best effort */ }
}

/**
 * Verwirft den Sammel-Scan und – falls Event-IDs übergeben werden – deren
 * Einzel-Ergebnisse. Nach einer Abmeldung wird beim nächsten Öffnen frisch
 * geprüft, statt die Person weiter als „hat womöglich Deloitte verlassen"
 * zu melden.
 */
export function invalidateInactiveAccountCache(eventIds?: string[]): void {
  try {
    window.localStorage.removeItem(INACTIVE_SUMMARY_CACHE_KEY);
    (eventIds || []).forEach(id => {
      if (id) window.localStorage.removeItem(accountCheckCacheKey(id));
    });
    // v31.22: Der Adress-Cache bleibt bewusst STEHEN. Er beantwortet „hat
    // diese Person noch ein Konto" — das ändert eine Abmeldung nicht. Die
    // beiden Event-Caches oben beantworten „wer aus DIESEM Event ist
    // betroffen", und genau das ändert sie.
  } catch { /* localStorage evtl. blockiert — best effort */ }
}
