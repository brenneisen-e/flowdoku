import * as React from 'react';

/**
 * v26.36: Zentraler Mobile-Breakpoint-Hook. Spiegelt das bisher inline in
 * Header.tsx genutzte Muster (matchMedia 768px + `window.__dexForceMobile` für
 * die Handbuch-Preview) und macht es appweit wiederverwendbar, damit alle
 * Komponenten dieselbe Handy-Grenze nutzen.
 *
 * true, sobald der Viewport ≤ `maxWidth` (Default 768px) ist ODER die
 * Handbuch-Preview `window.__dexForceMobile` gesetzt hat.
 */
export function useIsMobile(maxWidth: number = 768): boolean {
  const query = `(max-width: ${maxWidth}px)`;
  const read = React.useCallback((): boolean => {
    if (typeof window === 'undefined') return false;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if ((window as any).__dexForceMobile) return true;
    return !!(window.matchMedia && window.matchMedia(query).matches);
  }, [query]);

  const [isMobile, setIsMobile] = React.useState<boolean>(read);

  React.useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia(query);
    const handler = (): void => setIsMobile(read());
    // Sofort einmal synchronisieren (falls sich seit dem State-Init etwas änderte).
    handler();
    if (mq.addEventListener) mq.addEventListener('change', handler);
    // eslint-disable-next-line @typescript-eslint/no-deprecated
    else mq.addListener(handler);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener('change', handler);
      // eslint-disable-next-line @typescript-eslint/no-deprecated
      else mq.removeListener(handler);
    };
  }, [query, read]);

  return isMobile;
}
