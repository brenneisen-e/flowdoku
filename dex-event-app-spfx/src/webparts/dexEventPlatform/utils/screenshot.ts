/**
 * v26.0.0 — Screenshot des aktuellen Bildschirms für das Ticketsystem.
 *
 * Nutzt html2canvas, das LAZY (`await import`) geladen wird (Konvention: schwere
 * Libs nur dynamisch, nie im Boot-Pfad). Liefert eine PNG-`File`, die als
 * Item-Attachment an ein Ticket gehängt werden kann.
 *
 * Hinweis: Cross-Origin-Bilder ohne CORS-Header erscheinen ggf. leer — das ist
 * für einen Hilfe-Screenshot akzeptabel. SharePoint-eigene Inhalte (gleiche
 * Origin) werden korrekt erfasst.
 */
export async function captureScreen(target?: HTMLElement): Promise<File | null> {
  try {
    if (typeof document === 'undefined') return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod: any = await import('html2canvas');
    const html2canvas = mod.default || mod;
    const el = target || document.body;
    const canvas: HTMLCanvasElement = await html2canvas(el, {
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff',
      scale: 1,
      // Auf den sichtbaren Bereich begrenzen, damit die Datei nicht riesig wird.
      width: Math.min(el.scrollWidth || document.documentElement.clientWidth, 2200),
      windowWidth: document.documentElement.clientWidth,
      windowHeight: document.documentElement.clientHeight,
    });
    const blob: Blob | null = await new Promise((resolve) => {
      try { canvas.toBlob((b) => resolve(b), 'image/png', 0.92); } catch { resolve(null); }
    });
    if (!blob) return null;
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    return new File([blob], `screenshot_${ts}.png`, { type: 'image/png' });
  } catch (e) {
    // eslint-disable-next-line no-console
    console.warn('[DEX] captureScreen fehlgeschlagen (best-effort):', e);
    return null;
  }
}
