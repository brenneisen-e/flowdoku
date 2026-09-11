/**
 * v31.25 — Eine eingebaute Stoppuhr für die Konsole.
 *
 * Nutzer-Bitte 11.09.2026: „kannst du in der Konsole mal messen, wie lange
 * das dauert — dann ist es für dich noch einfacher zu optimieren."
 *
 * Genau das Richtige, und es fehlte: An den Tenant komme ich nicht, also habe
 * ich bei den letzten drei Releases aus dem Code heraus GESCHLOSSEN, was
 * teuer ist. Zweimal lag ich daneben (die Anhänge betrafen dieses Event gar
 * nicht; `content-visibility` brachte gemessene 3 ms). Eine Zahl aus dem
 * echten Browser beendet das Raten.
 *
 * ## Was gemessen wird
 *
 * Jede Messung ist ein Abschnitt mit Namen, Dauer und optionalen Angaben
 * (gelesene Zeilen, Größe der Antwort in KB). Geloggt wird EINE Zeile je
 * Abschnitt — keine Gruppen, keine Tabellen, nichts, was man aufklappen muss.
 * Wer die Konsole offen hat, soll mitlesen können, ohne sie zu bedienen.
 *
 * ## Warum das dauerhaft drin bleibt und nicht hinter einem Schalter
 *
 * Ein Schalter, den man erst einschalten muss, ist genau dann aus, wenn das
 * Problem auftritt — und das Problem tritt beim Nutzer auf, nicht bei mir.
 * Die Kosten sind vernachlässigbar: `performance.now()` und ein
 * `console.info` je Ladevorgang, nicht je Zeile.
 *
 * ## Wie man drankommt
 *
 * In der Konsole `dexPerf()` eingeben — das druckt die letzten Messungen als
 * Tabelle und liefert sie zusätzlich als Text zum Kopieren zurück.
 */

/*
 * `no-console` ist hier ausgeschaltet, und nur hier: Die Konsole IST das
 * Ausgabemedium dieser Datei. Anderswo bleibt die Regel scharf — sonst wäre
 * die nächste vergessene Debug-Zeile in einem Release wieder unauffällig.
 */
/* eslint-disable no-console */

export interface PerfEintrag {
  /** Was gemessen wurde, z.B. „Teilnehmer laden". */
  was: string;
  /** Dauer in Millisekunden, auf 0,1 ms gerundet. */
  ms: number;
  /** Gelesene Zeilen, falls es um eine Liste ging. */
  zeilen?: number;
  /** Größe der Antwort in KB, falls messbar. */
  kb?: number;
  /** Freier Zusatz, z.B. „aus dem Speicher" oder der Event-Titel. */
  hinweis?: string;
  /** Uhrzeit der Messung (lokale Zeit, hh:mm:ss). */
  zeit: string;
}

const VERLAUF: PerfEintrag[] = [];
const MAX = 60;

/** Auf 0,1 ms runden — mehr Stellen täuschen eine Genauigkeit vor, die es nicht gibt. */
const r1 = (n: number): number => Math.round(n * 10) / 10;

function uhr(): string {
  const d = new Date();
  const p = (n: number): string => String(n).padStart(2, '0');
  return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/**
 * Eine Messung festhalten und in die Konsole schreiben.
 *
 * Die Zeile ist bewusst so gebaut, dass sie sich kopieren und weiterschicken
 * lässt: ein Präfix zum Filtern, dann Klartext.
 */
export function perfLog(was: string, ms: number, extra?: { zeilen?: number; kb?: number; hinweis?: string }): void {
  const e: PerfEintrag = { was, ms: r1(ms), zeilen: extra?.zeilen, kb: extra?.kb, hinweis: extra?.hinweis, zeit: uhr() };
  VERLAUF.push(e);
  if (VERLAUF.length > MAX) VERLAUF.shift();
  const teile = [`${e.ms} ms`];
  if (typeof e.zeilen === 'number') teile.push(`${e.zeilen} Zeilen`);
  if (typeof e.kb === 'number') teile.push(`${e.kb} KB`);
  if (e.hinweis) teile.push(e.hinweis);
  // `info` statt `log`: In der Konsole lässt sich damit auf „Info" filtern,
  // ohne die Warnungen der App zu verlieren.
  console.info(`[DEX ⏱] ${was} — ${teile.join(' · ')}`);
}

/**
 * Misst eine `Promise`-Funktion und gibt deren Ergebnis unverändert zurück.
 *
 * `zusatz` bekommt das Ergebnis und darf daraus Zeilen/KB ableiten — so
 * bleibt die Messstelle eine Zeile und der gemessene Code unberührt.
 */
export async function perfMiss<T>(
  was: string,
  fn: () => Promise<T>,
  zusatz?: (_ergebnis: T) => { zeilen?: number; kb?: number; hinweis?: string },
): Promise<T> {
  const t0 = performance.now();
  try {
    const out = await fn();
    perfLog(was, performance.now() - t0, zusatz ? zusatz(out) : undefined);
    return out;
  } catch (err) {
    // Auch ein Fehlschlag hat gedauert — und gerade der ist interessant
    // (Drosselung mit Backoff sieht in der Zeit anders aus als ein 403).
    perfLog(`${was} (fehlgeschlagen)`, performance.now() - t0);
    throw err;
  }
}

/** Die letzten Messungen. */
export function perfVerlauf(): PerfEintrag[] {
  return VERLAUF.slice();
}

/**
 * `dexPerf()` in der Konsole: druckt die Tabelle und liefert denselben Stand
 * als Text zurück — zum Markieren und Verschicken.
 */
export function installPerfConsole(): void {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).dexPerf = (): string => {
      const v = perfVerlauf();
      if (v.length === 0) {
        console.info('[DEX ⏱] Noch nichts gemessen — einmal ein Event bzw. einen Termin öffnen.');
        return '';
      }
      console.table(v);
      const text = v.map(e => {
        const teile = [`${e.ms} ms`];
        if (typeof e.zeilen === 'number') teile.push(`${e.zeilen} Zeilen`);
        if (typeof e.kb === 'number') teile.push(`${e.kb} KB`);
        if (e.hinweis) teile.push(e.hinweis);
        return `${e.zeit}  ${e.was} — ${teile.join(' · ')}`;
      }).join('\n');
      console.info('[DEX ⏱] Zum Kopieren:\n' + text);
      return text;
    };
  } catch { /* best-effort — die Messung darf die App nie stören */ }
}
