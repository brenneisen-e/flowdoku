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
  } catch { /* localStorage evtl. blockiert — best effort */ }
}
