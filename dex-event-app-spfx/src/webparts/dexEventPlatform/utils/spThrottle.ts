import { SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * v29.48 — Drosselung von SharePoint abfangen (HTTP 429/503).
 *
 * SharePoint zählt Requests pro Nutzer in einem gleitenden Fenster. Wer viele
 * Sub-Events in einem Rutsch speichert, überschreitet das Kontingent: die
 * nächsten Requests kommen mit **429 (Too Many Requests)** zurück, oft mit
 * einem `Retry-After`-Header, der sagt, wie lange man warten soll.
 *
 * Bis hierher hat die App das als normalen Fehlschlag behandelt. Beim
 * Kalender-Event mit 21 Tagen war das Ergebnis genau das, was der Organizer
 * gemeldet hat: „The event could not be updated. Reason: HTTP 429" beim
 * Speichern der Anmeldefrist und „3 sub-events could not be saved …" beim
 * Speichern des Orts — und weil der Schreibvorgang nie ankam, standen
 * entfernte Tage danach immer noch in der Liste. Es war also nie ein
 * Datenfehler, sondern eine verworfene Antwort.
 *
 * Deshalb hier zwei Dinge:
 *
 *  1. **Wiederholen statt Aufgeben.** Ein 429/503 heißt „abgelehnt, nicht
 *     ausgeführt" — der Request ist nie beim Listen-Schreiben angekommen.
 *     Ihn zu wiederholen ist deshalb auch bei POST/MERGE ungefährlich; es
 *     entsteht kein zweiter Eintrag. Wartezeit: `Retry-After`, sonst 2/4/8/16 s.
 *
 *  2. **Eine gemeinsame Schranke.** Ohne sie laufen die 80 gerade fliegenden
 *     Requests alle in dieselbe Wand, warten alle gleich lang und schlagen
 *     danach gleichzeitig wieder auf — die Drosselung verlängert sich selbst.
 *     `gateUntil` ist deshalb modulweit: sieht EIN Request ein 429, warten
 *     ALLE bis zum selben Zeitpunkt, auch die, die gerade erst starten.
 *
 * Nicht wiederholt werden andere Statuscodes. 403/404/400 sind Aussagen über
 * den Request selbst; sie werden durch Warten nicht wahr.
 */

/** Maximale Zahl der Wiederholungen je Request. */
const MAX_RETRIES = 3;

/**
 * Obergrenze je Wartezeit. v29.50 von 30 s auf 6 s gesenkt: Mit vier Versuchen
 * à 30 s konnte EIN Request zwei Minuten warten — das ist keine Verzögerung
 * mehr, das ist ein Hänger.
 */
const MAX_WAIT_MS = 6000;

/** Obergrenze für das Warten an der gemeinsamen Schranke (v29.50). */
const MAX_GATE_WAIT_MS = 8000;

/**
 * Zeitpunkt, bis zu dem ALLE Requests warten. Wird gesetzt, sobald irgendein
 * Request gedrosselt wurde; das hält die nachfolgenden aus der Wand heraus.
 */
let gateUntil = 0;

const sleep = (ms: number): Promise<void> => new Promise<void>(r => setTimeout(r, ms));

/**
 * Wartet, bis die gemeinsame Schranke offen ist — höchstens MAX_GATE_WAIT_MS,
 * danach läuft der Request trotzdem los.
 *
 * v29.50: Vorher waren es 60 Runden à 2 s. Das war als „lieber warten als
 * scheitern" gedacht und ist in der Praxis das Gegenteil: ein einziges 429
 * legte jeden folgenden Request für bis zu zwei Minuten still. Lieber ein
 * abgelehnter Request als eine Anwendung, die sich tot stellt.
 */
async function waitForGate(): Promise<void> {
  const deadline = Date.now() + MAX_GATE_WAIT_MS;
  for (;;) {
    const now = Date.now();
    const rest = Math.min(gateUntil, deadline) - now;
    if (rest <= 0) return;
    await sleep(Math.min(rest, 1000));
  }
}

/** True, solange die App wegen Drosselung wartet — für Fortschrittstexte. */
export function isThrottled(): boolean {
  return gateUntil > Date.now();
}

/**
 * Führt einen SharePoint-Request aus und wiederholt ihn bei Drosselung.
 *
 * `call` muss den Request JEDES MAL neu aufbauen (also eine Funktion, kein
 * bereits gestartetes Promise) — sonst wird beim zweiten Versuch dieselbe,
 * schon abgelehnte Antwort zurückgegeben.
 */
export async function withThrottleRetry(
  call: () => Promise<SPHttpClientResponse>,
  label?: string,
): Promise<SPHttpClientResponse> {
  let attempt = 0;
  for (;;) {
    await waitForGate();
    const resp = await call();
    if (resp.status !== 429 && resp.status !== 503) return resp;
    if (attempt >= MAX_RETRIES) {
      console.warn('[DEX Throttle] aufgegeben nach', attempt, 'Versuchen —', resp.status, label || '');
      return resp;
    }
    // Retry-After kommt als Sekunden (SharePoint) oder gar nicht.
    let waitMs = 0;
    try {
      const ra = parseInt(resp.headers.get('Retry-After') || '', 10);
      if (ra > 0) waitMs = ra * 1000;
    } catch { /* Header nicht lesbar — Backoff greift */ }
    if (waitMs <= 0) waitMs = 2000 * Math.pow(2, attempt);
    if (waitMs > MAX_WAIT_MS) waitMs = MAX_WAIT_MS;
    gateUntil = Math.max(gateUntil, Date.now() + waitMs);
    attempt++;
    console.warn('[DEX Throttle] HTTP', resp.status, '— warte', waitMs, 'ms, Versuch', attempt, label || '');
  }
}
