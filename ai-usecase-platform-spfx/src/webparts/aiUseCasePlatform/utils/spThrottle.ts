import { SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * v29.48/v29.74 — Drosselung von SharePoint abfangen (HTTP 429/503).
 *
 * Geschichte dieses Moduls, weil jede Regel hier aus einem echten Vorfall
 * stammt:
 *  - v29.48: Retry mit Backoff, damit das Speichern von 21 Sub-Events nicht
 *    an einem einzelnen 429 scheitert.
 *  - v29.50: Retry von den GETs genommen — ein 429 im Boot hielt sonst die
 *    ganze App an. Dabei wurden die Wartezeiten auf 6 s gedeckelt.
 *  - v29.74: Der 6-s-Deckel war der naechste Vorfall. Er galt AUCH fuer den
 *    Retry-After-Header — sagte SharePoint „warte 120 s", klopften wir
 *    nach 6 s wieder an. Gedrosselte Anfragen zaehlen bei Microsoft aber
 *    WEITER aufs Kontingent; wer Retry-After ignoriert, wird vom 429 zur
 *    NUTZER-SPERRE eskaliert (Throttle.htm, „Something\u2019s not right") —
 *    genau das hat einen Organizer nach dem Bearbeiten von 19 Terminen
 *    komplett aus SharePoint ausgesperrt.
 *
 * Die Regeln seither:
 *  1. Retry-After wird IMMER voll respektiert — nie frueher wiederholen.
 *  2. Ist die verlangte Wartezeit laenger, als ein Nutzer vor einem
 *     Speichern-Balken sinnvoll wartet (RETRY_HONOR_MAX), wird GAR NICHT
 *     wiederholt: Der Aufrufer bekommt den 429 und meldet den Fehlschlag.
 *     Ein ehrlicher Fehler ist billiger als eine Kontosperre.
 *  3. Die gemeinsame Schranke (gateUntil) gilt fuer die volle Dauer — kein
 *     Durchlassen nach 8 s mehr. Nur Schreibzugriffe laufen durch dieses
 *     Modul (v29.50), ein wartender Save blockiert also nicht die App.
 *  4. Nach mehreren Drossel-Antworten in Folge schaltet ein Schutzschalter
 *     die Wiederholungen ganz ab, bis die Schranke abgelaufen ist.
 */

/** Maximale Zahl der Wiederholungen je Request. */
const MAX_RETRIES = 2;

/** Backoff ohne Retry-After-Header: 4 s, dann 8 s. */
const FALLBACK_WAIT_MS = [4000, 8000];

/**
 * Laengste Wartezeit, die wir fuer einen Retry noch aussitzen. Verlangt
 * SharePoint mehr, geben wir den 429 sofort an den Aufrufer weiter — der
 * meldet „nicht gespeichert" und der Nutzer versucht es spaeter.
 */
const RETRY_HONOR_MAX_MS = 45000;

/** Obergrenze der gemeinsamen Schranke (Schutz gegen absurde Header-Werte). */
const GATE_MAX_MS = 180000;

/** Ab so vielen Drossel-Antworten in Folge: keine Retries mehr bis Gate-Ablauf. */
const BREAKER_THRESHOLD = 4;

let gateUntil = 0;
let consecutiveThrottles = 0;

const sleep = (ms: number): Promise<void> => new Promise<void>(r => setTimeout(r, ms));

/** Wartet, bis die gemeinsame Schranke offen ist — die VOLLE Dauer. */
async function waitForGate(): Promise<void> {
  for (;;) {
    const rest = Math.min(gateUntil, Date.now() + GATE_MAX_MS) - Date.now();
    if (rest <= 0) return;
    await sleep(Math.min(rest, 2000));
  }
}

/** True, solange die App wegen Drosselung wartet — fuer Fortschrittstexte. */
export function isThrottled(): boolean {
  return gateUntil > Date.now();
}

/**
 * Fuehrt einen SharePoint-Request aus und wiederholt ihn bei Drosselung.
 *
 * `call` muss den Request JEDES MAL neu aufbauen (Funktion, kein bereits
 * gestartetes Promise) — sonst kaeme beim zweiten Versuch dieselbe, schon
 * abgelehnte Antwort zurueck.
 */
export async function withThrottleRetry(
  call: () => Promise<SPHttpClientResponse>,
  label?: string,
): Promise<SPHttpClientResponse> {
  let attempt = 0;
  for (;;) {
    await waitForGate();
    const resp = await call();
    if (resp.status !== 429 && resp.status !== 503) {
      consecutiveThrottles = 0;
      return resp;
    }
    consecutiveThrottles++;
    // Retry-After lesen — Sekunden oder gar nicht.
    let waitMs = 0;
    try {
      const ra = parseInt(resp.headers.get('Retry-After') || '', 10);
      if (ra > 0) waitMs = ra * 1000;
    } catch { /* Header nicht lesbar */ }
    if (waitMs <= 0) waitMs = FALLBACK_WAIT_MS[Math.min(attempt, FALLBACK_WAIT_MS.length - 1)];
    // Schranke IMMER auf die volle verlangte Dauer setzen — auch wenn wir
    // selbst nicht mehr wiederholen. Nachfolgende Schreibzugriffe warten dann.
    gateUntil = Math.max(gateUntil, Date.now() + Math.min(waitMs, GATE_MAX_MS));
    const breakerOpen = consecutiveThrottles >= BREAKER_THRESHOLD;
    if (attempt >= MAX_RETRIES || waitMs > RETRY_HONOR_MAX_MS || breakerOpen) {
      console.warn(
        '[DEX Throttle] gebe auf —', resp.status,
        breakerOpen ? '(Schutzschalter offen)' : (waitMs > RETRY_HONOR_MAX_MS ? `(Retry-After ${Math.round(waitMs / 1000)} s zu lang)` : `(nach ${attempt} Versuchen)`),
        label || '',
      );
      return resp;
    }
    attempt++;
    console.warn('[DEX Throttle] HTTP', resp.status, '— warte', Math.round(waitMs / 1000), 's (Retry-After respektiert), Versuch', attempt, label || '');
    // Nicht nur bis gateUntil warten — waitMs kann durch parallele Requests
    // schon teilweise verstrichen sein; waitForGate am Schleifenkopf deckt das.
  }
}
