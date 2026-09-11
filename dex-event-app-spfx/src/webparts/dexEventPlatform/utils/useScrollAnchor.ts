/**
 * v31.29 — Beim Reiterwechsel bleibt der Blick, wo er ist.
 *
 * Nutzer-Befund 11.09.2026: „leichter Bildschirmsprung" beim Umschalten
 * zwischen den Terminen. Die Ursache ist keine Animation und kein Nachladen,
 * sondern schlichte Höhenarithmetik: Über der Teilnehmerliste stehen Kästen,
 * die je Termin unterschiedlich hoch sind (Überbuchung, Dubletten, ID-Lücken,
 * unvollständige Anmeldungen, Team-Abschnitt). Wechselt der Termin, ändert
 * sich die Höhe ÜBER dem Punkt, auf den der Nutzer gerade schaut — der
 * Scroll-Versatz bleibt aber gleich, also rutscht der Inhalt unter dem
 * Mauszeiger weg.
 *
 * Der Browser hat dafür „scroll anchoring" (`overflow-anchor`), aber es greift
 * hier nicht: React tauscht beim Wechsel den halben Teilbaum aus, und ein
 * Anker, dessen Element verschwindet, wird verworfen.
 *
 * Also selbst verankern, nach demselben Prinzip: vor dem Wechsel den Abstand
 * eines stehenbleibenden Elements zum oberen Fensterrand merken, nach dem
 * Zeichnen die Differenz aus dem Scroll-Versatz herausrechnen. Das Element ist
 * die Reiter-Leiste über der Liste — genau das, was der Nutzer angeklickt hat
 * und im Blick behält.
 *
 * Fehlschläge sind still und folgenlos: Kein Element, kein Scroll-Container
 * oder eine Differenz unter einem Pixel heißt „nichts tun".
 */

import * as React from 'react';

/**
 * Der nächste scrollbare Vorfahr. In SharePoint ist das NICHT das Fenster,
 * sondern ein Canvas-Container — deshalb wird gesucht statt angenommen.
 */
function scrollEltern(el: HTMLElement | null): HTMLElement | null {
  let p: HTMLElement | null = el ? el.parentElement : null;
  while (p) {
    try {
      const ov = window.getComputedStyle(p).overflowY;
      if ((ov === 'auto' || ov === 'scroll') && p.scrollHeight > p.clientHeight + 1) return p;
    } catch { /* getComputedStyle kann bei entfernten Knoten werfen */ }
    p = p.parentElement;
  }
  return (document.scrollingElement as HTMLElement) || document.documentElement;
}

export interface ScrollAnker<T extends HTMLElement> {
  /** An das Element hängen, das seine Position behalten soll. */
  ref: React.RefObject<T>;
  /** Im Klick-Handler rufen — VOR der Zustandsänderung. */
  merken: () => void;
}

/**
 * @param schluessel Wechselt mit dem Inhalt (z.B. die Event-Id). Nach jedem
 *   Wechsel wird der gemerkte Abstand einmal wiederhergestellt.
 */
export function useScrollAnchor<T extends HTMLElement>(schluessel: string): ScrollAnker<T> {
  const ref = React.useRef<T>(null);
  const merkRef = React.useRef<{ top: number; scroller: HTMLElement | null } | null>(null);

  const merken = React.useCallback(() => {
    const el = ref.current;
    if (!el) { merkRef.current = null; return; }
    try {
      merkRef.current = { top: el.getBoundingClientRect().top, scroller: scrollEltern(el) };
    } catch { merkRef.current = null; }
  }, []);

  // `useLayoutEffect`, nicht `useEffect`: Die Korrektur muss VOR dem nächsten
  // Bild sitzen, sonst sieht man genau den Sprung, den sie verhindern soll.
  React.useLayoutEffect(() => {
    const merk = merkRef.current;
    merkRef.current = null;
    const el = ref.current;
    if (!merk || !el || !merk.scroller) return;
    try {
      const d = el.getBoundingClientRect().top - merk.top;
      if (Math.abs(d) > 1) merk.scroller.scrollTop += d;
    } catch { /* best-effort — eine verschobene Ansicht ist kein Absturz */ }
  }, [schluessel]);

  return { ref, merken };
}
