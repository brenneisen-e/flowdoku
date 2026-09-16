/**
 * v31.64: Scroll-Helfer des Event-Wizards.
 *
 * Zwei Nutzer-Ansagen vom 16.09.2026:
 *  - „wenn ich in einem Schritt unten gescrollt habe und dann auf Weiter
 *    klicke, dann soll er im nächsten Schritt wieder oben anfangen"
 *  - „wenn ich ein Sub-Event anlege und auf Bearbeiten klicke, dann soll es
 *    bei Grundlagen nach oben scrollen zum Namen des Events"
 *
 * Beide brauchen dasselbe: den Wizard so scrollen, dass seine Oberkante
 * direkt unter dem Kopf der App sitzt. Zwei Fallen dabei — in SharePoint
 * scrollt NICHT das Fenster, sondern ein Canvas-Container (deshalb
 * `scrollEltern` aus useScrollAnchor statt `window.scrollTo`), und der
 * App-Header ist per JS gepinnt (Header.tsx v22.28): Seine Unterkante liegt
 * immer bei „Oberkante der SP-Chrome + eigene Höhe", egal ob er gerade in
 * der Fläche steht oder fixiert ist.
 */

import { scrollEltern } from './useScrollAnchor';

/** Oberkante des sichtbaren Bereichs unterhalb der SharePoint-Leiste. */
function chromeTop(): number {
  const candidates = ['[data-automation-id="contentScrollRegion"]', '.SPPageChromeAppDiv', '#spPageCanvasContent'];
  for (const sel of candidates) {
    const el = document.querySelector(sel);
    if (el) return Math.max(0, el.getBoundingClientRect().top);
  }
  return 0;
}

/**
 * Die Kante, unter der etwas „oben" ist: Unterkante des App-Headers (gepinnt
 * oder nicht), sonst die SP-Chrome. Hier hängen der gepinnte Wizard-Kopf
 * (WizardFormShell) und das Scrollen zum Schrittanfang.
 */
export function stickyTopEdge(): number {
  const top = chromeTop();
  const header = document.querySelector('.header') as HTMLElement | null;
  return top + (header ? header.offsetHeight : 0);
}

/**
 * Scrollt so, dass `root` direkt unter der Kante beginnt — aber nur, wenn
 * es gerade DARÜBER liegt (also weggescrollt ist). Steht der Wizard schon
 * sichtbar unter dem Kopf, passiert nichts; sonst würde ein Klick auf
 * „Weiter" die Seite um die Kartenränder nach unten rücken.
 */
export function scrollWizardTop(root: HTMLElement | null): void {
  if (!root) return;
  try {
    const scroller = scrollEltern(root);
    const delta = root.getBoundingClientRect().top - stickyTopEdge() - 8;
    if (delta >= -2) return;
    const isDoc = !scroller || scroller === document.scrollingElement || scroller === document.documentElement || scroller === document.body;
    if (isDoc) window.scrollBy(0, delta);
    else scroller.scrollTop += delta;
  } catch { /* best-effort — eine stehengebliebene Ansicht ist kein Absturz */ }
}
