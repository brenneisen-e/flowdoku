/**
 * v29.41 — Echte Boot-Fortschritts-Meldungen.
 *
 * Der Start-Balken war bis v29.40 zeitgesteuert mit Phasen-Deckel: Nach ein
 * paar Sekunden stand er bei 92 % und wartete stumm, bis die Events fertig
 * waren — bei einem langsamen Tenant minutenlang. Der Deckel war richtig
 * (nicht 100 % zeigen, solange die App nicht bedienbar ist), die fehlende
 * Information darunter nicht.
 *
 * `initEvents` kennt seine Abschnitte aber genau (Schema-Prüfung, Logos,
 * Events lesen, mappen, Teilnehmerzahlen, Anhänge). Diese Abschnitte meldet es
 * jetzt — über ein window-Event, damit der Boot-Loader nichts vom
 * EventContext wissen muss und die Meldung auch dann ankommt, wenn der Loader
 * noch gar nicht gemountet ist (der letzte Stand wird gemerkt).
 *
 * Die Prozentwerte sind Erfahrungswerte aus dem Perf-Log, keine Messung: Sie
 * sagen, WIE WEIT der Start ist, nicht wie lange er noch dauert.
 */

export type BootStage =
  | 'schema'      // ensure*/upgrade*-Calls (nur beim ersten Boot je Version)
  | 'logos'       // Logo-Base64-Cache für Mail-/Outlook-Vorlagen
  | 'events'      // DEX_Events lesen
  | 'mapping'     // SP-Items → App-Objekte
  | 'counts'      // Teilnehmerzahlen je Subsite
  | 'documents';  // Anhänge der Events

export interface BootStageDetail {
  stage: BootStage;
  /** Oberes Ende dieses Abschnitts in Prozent — der Balken kriecht dorthin. */
  target: number;
  /** Untergrenze beim Betreten des Abschnitts (der Balken springt nie zurück). */
  floor: number;
}

export const BOOT_STAGE_EVENT = 'dex-boot-stage';

/** Abschnitts-Grenzen an EINER Stelle — Loader und Melder dürfen nicht auseinanderlaufen. */
export const BOOT_STAGES: Record<BootStage, { floor: number; target: number }> = {
  schema: { floor: 10, target: 30 },
  logos: { floor: 30, target: 40 },
  events: { floor: 40, target: 62 },
  mapping: { floor: 62, target: 74 },
  counts: { floor: 74, target: 88 },
  documents: { floor: 88, target: 96 },
};

/** Letzter gemeldeter Stand — der Loader mountet evtl. nach der ersten Meldung. */
let lastStage: BootStageDetail | null = null;
export function lastBootStage(): BootStageDetail | null { return lastStage; }

export function emitBootStage(stage: BootStage): void {
  const { floor, target } = BOOT_STAGES[stage];
  lastStage = { stage, floor, target };
  try {
    if (typeof window !== 'undefined') {
      window.dispatchEvent(new CustomEvent(BOOT_STAGE_EVENT, { detail: lastStage }));
    }
  } catch { /* CustomEvent nicht verfügbar → Balken bleibt zeitgesteuert */ }
}
