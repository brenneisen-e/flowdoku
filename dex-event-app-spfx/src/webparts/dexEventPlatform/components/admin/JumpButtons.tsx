/* JumpButtons — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 17217-17316 des
 * Stands vor dem Schnitt). Zeichengleich uebernommen, nur `function` -> `export function`.
 */
import * as React from 'react';

/**
 * v17.8: Floating Jump-Buttons rechts unten. Erscheinen sobald der User
 * den Viewport >300 px nach unten gescrollt hat. Bietet:
 *  - Nach oben springen (window.scrollTo {top:0})
 *  - Zur Warteliste springen (scrollIntoView auf #admin-waitlist-anchor)
 *
 * Nur sichtbar wenn die Teilnehmer-Tabelle >10 Einträge hat (kurze Listen
 * brauchen keine Sprung-Hilfe).
 */
/**
 * v17.13: Floating Jump-Buttons. Im SPFx-Webpart-Kontext scrollt nicht
 * window, sondern ein SP-interner Container — deshalb sind die Buttons
 * jetzt IMMER sichtbar (kein scrollY-Gating), und der Click sucht den
 * tatsächlich scrollenden Vorfahren des Targets statt window.scrollTo.
 */
export default function JumpButtons(props: { hasWaitlist: boolean }): React.ReactElement {
  const { hasWaitlist } = props;
  /** Sucht den ersten scroll-baren Vorfahren — typischerweise der
   *  SP-Page-Body. Fallback auf document.scrollingElement / window. */
  const findScrollParent = (el: HTMLElement | null): HTMLElement | Window => {
    let cur: HTMLElement | null = el ? el.parentElement : null;
    while (cur && cur !== document.body && cur !== document.documentElement) {
      const cs = window.getComputedStyle(cur);
      const overflowY = cs.overflowY;
      const isScrollable = (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay');
      if (isScrollable && cur.scrollHeight > cur.clientHeight) return cur;
      cur = cur.parentElement;
    }
    return (document.scrollingElement as HTMLElement) || document.documentElement || window;
  };
  const scrollToTop = (): void => {
    // Versuche window, documentElement, body und alle scrollbaren Vorfahren.
    try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* */ }
    try { if (document.scrollingElement) (document.scrollingElement as HTMLElement).scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* */ }
    try { if (document.documentElement) document.documentElement.scrollTop = 0; } catch { /* */ }
    try { if (document.body) document.body.scrollTop = 0; } catch { /* */ }
    // SP-Page-Container — typischer Klassenname in modernen SP-Pages.
    const candidates = ['.SPPageChromeAppDiv', '#spPageCanvasContent', '[data-automation-id="contentScrollRegion"]', '.spAppAriaRegion', 'main'];
    for (const sel of candidates) {
      const el = document.querySelector(sel) as HTMLElement | null;
      if (el && el.scrollHeight > el.clientHeight) {
        try { el.scrollTo({ top: 0, behavior: 'smooth' }); } catch { el.scrollTop = 0; }
      }
    }
  };
  const scrollToWaitlist = (): void => {
    const el = document.getElementById('admin-waitlist-anchor');
    if (!el) return;
    try { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch { el.scrollIntoView(); }
    // Zusätzlich: scroll-Parent suchen und manuell scrollen (Fallback
    // wenn scrollIntoView durch den SP-Container gefangen wird).
    const parent = findScrollParent(el);
    if (parent !== window && parent instanceof HTMLElement) {
      const rect = el.getBoundingClientRect();
      const parentRect = parent.getBoundingClientRect();
      try { parent.scrollTo({ top: parent.scrollTop + (rect.top - parentRect.top) - 20, behavior: 'smooth' }); } catch { /* */ }
    }
  };
  return (
    <div style={{
      position: 'fixed',
      // v18: Buttons horizontal ZENTRIERT über dem Content (vorher links am
      // Rand „im Raum hängend"). Als Zeile nebeneinander, mittig unten —
      // liegt damit in der horizontalen Mitte der Teilnehmer-Tabelle /
      // Spaltenüberschriften statt am Seitenrand.
      left: '50%', transform: 'translateX(-50%)',
      bottom: 20, zIndex: 900,
      display: 'flex', flexDirection: 'row', gap: 8,
    }}>
      {hasWaitlist && (
        <button
          type="button"
          onClick={scrollToWaitlist}
          style={{
            background: 'var(--dex-orange, #ed8b00)', color: '#fff',
            border: 'none', padding: '10px 16px', borderRadius: 999,
            cursor: 'pointer', fontSize: '0.85rem', fontWeight: 600,
            boxShadow: '0 4px 12px rgba(0,0,0,0.18)',
            display: 'inline-flex', alignItems: 'center', gap: 6,
          }}
        >
          ↓ Zur Warteliste
        </button>
      )}
      <button
        type="button"
        onClick={scrollToTop}
        style={{
          background: 'var(--dex-green, #86bc25)', color: '#fff',
          border: 'none', padding: '10px 16px', borderRadius: 999,
          cursor: 'pointer', fontSize: '0.85rem', fontWeight: 600,
          boxShadow: '0 4px 12px rgba(0,0,0,0.18)',
          display: 'inline-flex', alignItems: 'center', gap: 6,
        }}
      >
        ↑ Nach oben
      </button>
    </div>
  );
}
