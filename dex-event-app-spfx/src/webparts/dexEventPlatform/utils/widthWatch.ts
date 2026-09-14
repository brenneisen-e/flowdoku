/**
 * v31.30 — Der Breiten-Wächter: WER macht die Seite breiter als das Fenster?
 *
 * Nutzer-Befund 11.09.2026: „warum wird die Breite verändert, wenn man Dinge
 * aufklappt? Das will ich nicht" — dieselbe Beobachtung wie beim Wechsel
 * zwischen zwei Terminen.
 *
 * ## Warum hier gemessen und nicht geraten wird
 *
 * Die naheliegende Erklärung wäre „die breite Tabelle sprengt die Seite". Die
 * stimmt nachweislich NICHT: Beide Tabellen liegen in eigenen Scroll-Behältern
 * (`dex-ui-table-wrap` mit `overflow-x: auto`, darin `overflow: auto`), und ein
 * Scroll-Behälter trägt zur Mindestbreite seines Elternteils null bei. Auch
 * `.app-layout` steht auf `overflow: hidden`. Nach dem Code KANN unser Inhalt
 * die Seite also gar nicht breiter machen — trotzdem tut es sichtbar etwas.
 *
 * Genau an dieser Stelle hat Raten in diesem Projekt schon zweimal ein Release
 * gekostet (die Anhänge, die das gemeldete Event gar nicht hatten;
 * `content-visibility` mit gemessenen 3 ms). Also dasselbe Vorgehen wie bei der
 * Stoppuhr in `perfLog`: die App sagt selbst, welches Element übersteht.
 *
 * ## Was gemeldet wird
 *
 * Sobald die Seite breiter ist als das Fenster UND sich diese Breite ändert,
 * steht eine Zeile in der Konsole: alte und neue Breite, der Überhang und das
 * Element, das ihn verursacht — mit CSS-Pfad und dem `overflow-x` seiner
 * Vorfahren. Letzteres ist die eigentliche Antwort: Ein Überhang ist nur dann
 * ein Fehler, wenn kein Vorfahr ihn einfängt.
 *
 * Gesucht wird das AUSLÖSENDE Element, nicht jedes überstehende: gemeldet wird
 * nur, wessen Elternteil noch innerhalb liegt. Sonst listet die Ausgabe die
 * ganze Kette vom Körper bis zur Zelle.
 *
 * ## Kosten
 *
 * `scrollWidth` erzwingt ein Layout. Deshalb wird nicht in einem Intervall
 * gemessen, sondern nur, wenn sich im DOM etwas geändert hat — und dann
 * frühestens 400 ms nach der letzten Änderung. Der teure Teil (jedes Element
 * abfragen) läuft erst, wenn die Breite sich WIRKLICH geändert hat.
 */

/* eslint-disable no-console */

/** Ein kurzer, lesbarer Name für ein Element — Tag, Id, erste zwei Klassen. */
function kurz(el: Element): string {
  const tag = el.tagName.toLowerCase();
  const id = el.id ? `#${el.id}` : '';
  const cls = typeof el.className === 'string'
    ? el.className.trim().split(/\s+/).filter(Boolean).slice(0, 2).map(c => `.${c}`).join('')
    : '';
  return `${tag}${id}${cls}`;
}

/** Der Weg von oben bis zum Element — damit man es im Inspektor wiederfindet. */
function pfad(el: Element): string {
  const teile: string[] = [];
  let n: Element | null = el;
  while (n && n !== document.body && teile.length < 8) {
    teile.unshift(kurz(n));
    n = n.parentElement;
  }
  return teile.join(' > ');
}

/**
 * Fängt ein Vorfahr den Überhang ein? Wenn ja, ist das Überstehen harmlos —
 * dann scrollt man innerhalb dieses Behälters, und die Seite bleibt schmal.
 */
function faengtEinVorfahrEin(el: Element): string {
  let n: Element | null = el.parentElement;
  while (n && n !== document.documentElement) {
    try {
      const ov = window.getComputedStyle(n).overflowX;
      if (ov === 'auto' || ov === 'scroll' || ov === 'hidden' || ov === 'clip') {
        return `${kurz(n)} (overflow-x: ${ov})`;
      }
    } catch { /* an entfernten Knoten kann das werfen */ }
    n = n.parentElement;
  }
  return '';
}

export interface Ueberhang {
  /** Kurzname des Elements. */
  element: string;
  /** Wie weit es über den rechten Fensterrand hinausragt, in Pixeln. */
  ueber: number;
  /** Breite des Elements. */
  breite: number;
  /** Der Weg dorthin. */
  pfad: string;
  /** Welcher Vorfahr den Überhang einfängt — leer heißt: keiner. */
  eingefangen: string;
}

/**
 * Alle Elemente, die rechts aus dem Fenster ragen, ohne dass ihr Elternteil
 * das schon täte. Sortiert nach Überhang, die drei größten reichen.
 */
export function ueberhaenge(max = 3): Ueberhang[] {
  const fenster = document.documentElement.clientWidth;
  const treffer: Ueberhang[] = [];
  const alle = document.querySelectorAll('*');
  for (let i = 0; i < alle.length; i++) {
    const el = alle[i];
    let r: DOMRect;
    try { r = el.getBoundingClientRect(); } catch { continue; }
    if (r.width === 0 || r.right <= fenster + 2) continue;
    // Nur der Auslöser, nicht die ganze Kette: Steht schon das Elternteil
    // über, ist dieses Element nur mitgerissen.
    const p = el.parentElement;
    if (p && p !== document.body) {
      try { if (p.getBoundingClientRect().right > fenster + 2) continue; } catch { /* weiter */ }
    }
    treffer.push({
      element: kurz(el),
      ueber: Math.round(r.right - fenster),
      breite: Math.round(r.width),
      pfad: pfad(el),
      eingefangen: faengtEinVorfahrEin(el),
    });
  }
  treffer.sort((a, b) => b.ueber - a.ueber);
  return treffer.slice(0, max);
}

function melde(alt: number, neu: number): void {
  const fenster = document.documentElement.clientWidth;
  const liste = ueberhaenge();
  const kopf = alt > 0
    ? `Seitenbreite geändert: ${alt} → ${neu} px (Fenster ${fenster}, Überhang ${neu - fenster})`
    : `Seite ist breiter als das Fenster: ${neu} px gegen ${fenster} (Überhang ${neu - fenster})`;
  if (liste.length === 0) {
    console.info(`[DEX ↔] ${kopf} — kein einzelnes Element steht über (dann kommt es von außerhalb der App).`);
    return;
  }
  const t = liste[0];
  console.info(`[DEX ↔] ${kopf} — Auslöser: ${t.element}, ${t.ueber} px über dem Rand`
    + (t.eingefangen ? ` · eingefangen von ${t.eingefangen}` : ' · KEIN Vorfahr fängt das ein')
    + `\n         Pfad: ${t.pfad}`);
  if (liste.length > 1) console.info('[DEX ↔] Weitere:', liste.slice(1));
}

/**
 * Beobachtung starten und `dexBreite()` in der Konsole bereitstellen.
 *
 * Meldet von selbst, sobald sich die Seitenbreite ändert — man muss also nicht
 * daran denken, vorher etwas einzutippen.
 */
export function installWidthConsole(): void {
  try {
    let letzte = 0;
    let timer: number | undefined;

    const pruefen = (): void => {
      const breit = document.documentElement.scrollWidth;
      const fenster = document.documentElement.clientWidth;
      // Ein Pixel Unterschied ist Rundung, kein Befund.
      if (breit <= fenster + 2) { letzte = 0; return; }
      if (breit === letzte) return;
      const alt = letzte;
      letzte = breit;
      melde(alt, breit);
    };

    const angestossen = (): void => {
      if (timer) window.clearTimeout(timer);
      // Erst wenn 400 ms lang Ruhe ist — sonst misst man mitten im Rendern.
      timer = window.setTimeout(pruefen, 400);
    };

    const mo = new MutationObserver(angestossen);
    mo.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['style', 'class'] });
    window.addEventListener('resize', angestossen);
    angestossen();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).dexBreite = (): Ueberhang[] => {
      const breit = document.documentElement.scrollWidth;
      const fenster = document.documentElement.clientWidth;
      console.info(`[DEX ↔] Seite ${breit} px, Fenster ${fenster} px, Überhang ${Math.max(0, breit - fenster)} px.`);
      const liste = ueberhaenge(10);
      if (liste.length === 0) console.info('[DEX ↔] Kein Element steht über den rechten Rand hinaus.');
      else console.table(liste);
      return liste;
    };
  } catch { /* best-effort — eine Messung darf die App nie stören */ }
}
