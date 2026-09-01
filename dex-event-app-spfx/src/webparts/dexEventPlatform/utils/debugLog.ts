/**
 * v30.64: Diagnose-Ausgaben, die standardmäßig STILL sind.
 *
 * Nutzer-Ansage 01.09.2026: „Wenn alles passt, kannst du die Konsole wieder
 * leer machen." Berechtigt — die App schrieb bei jedem Start ein Dutzend
 * `[DEX][perf]`-Zeilen, dazu je Termin eine `[DEX][seats]`-Zeile. Zum Suchen
 * eines Fehlers ist das Gold wert, im Alltag ist es Rauschen, in dem die
 * echten Meldungen untergehen.
 *
 * Deshalb nicht gelöscht, sondern **geschaltet**. Wer etwas untersucht, macht
 * es mit einer Zeile wieder sichtbar:
 *
 *     localStorage.setItem('dex_debug', 'seats')   // nur die Platzzähler
 *     localStorage.setItem('dex_debug', 'all')     // alles
 *     localStorage.removeItem('dex_debug')         // wieder still
 *
 * Danach die Seite neu laden. Mehrere Bereiche mit Komma: `'seats,perf'`.
 *
 * **`console.warn` und `console.error` laufen bewusst weiter.** Sie melden
 * Dinge, die schiefgegangen sind — die soll niemand erst einschalten müssen.
 * Still wird nur, was im Normalbetrieb ohnehin niemand liest.
 */

export type DebugTopic = 'perf' | 'seats' | 'realtime' | 'kpi';

/** Einmal je Seitenaufruf gelesen — localStorage bei jedem Log anzufassen
 *  wäre bei einer Schleife über 21 Termine unnötig teuer. */
let cached: string[] | null = null;

function topics(): string[] {
  if (cached) return cached;
  let raw = '';
  try { raw = window.localStorage.getItem('dex_debug') || ''; } catch { raw = ''; }
  cached = raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  return cached;
}

/** Ist dieser Bereich gerade eingeschaltet? */
export function isDebug(topic: DebugTopic): boolean {
  const t = topics();
  return t.indexOf('all') >= 0 || t.indexOf(topic) >= 0;
}

/**
 * Wie `console.log`, aber nur wenn der Bereich eingeschaltet ist.
 *
 * Das Präfix bleibt im Aufrufer stehen (`[DEX][seats] …`), damit sich der
 * Konsolen-Filter weiterhin darauf anwenden lässt und die Zeilen beim Suchen
 * im Code auffindbar bleiben.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function dlog(topic: DebugTopic, ...args: any[]): void {
  if (!isDebug(topic)) return;
  // eslint-disable-next-line no-console
  console.log(...args);
}
