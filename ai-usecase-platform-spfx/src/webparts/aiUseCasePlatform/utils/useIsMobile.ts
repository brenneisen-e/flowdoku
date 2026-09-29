/**
 * Handy-Breite — eine Grenze für die ganze App.
 *
 * Dieselbe Schwelle wie im SCSS (`@media (max-width: 768px)`); zwei
 * verschiedene Grenzen in CSS und JS ließen die Kopfzeile bei 770 px halb
 * mobil und halb Desktop aussehen.
 */

import * as React from 'react';

const ABFRAGE = '(max-width: 768px)';

export function useIsMobile(): boolean {
  const lese = (): boolean => {
    try { return typeof window !== 'undefined' && !!window.matchMedia && window.matchMedia(ABFRAGE).matches; } catch { return false; }
  };
  const [mobil, setMobil] = React.useState<boolean>(lese);

  React.useEffect(() => {
    let mql: MediaQueryList;
    try { mql = window.matchMedia(ABFRAGE); } catch { return undefined; }
    const auf = (): void => setMobil(mql.matches);
    auf();
    // Ältere Browser kennen nur addListener.
    if (mql.addEventListener) mql.addEventListener('change', auf);
    else mql.addListener(auf);
    return () => {
      if (mql.removeEventListener) mql.removeEventListener('change', auf);
      else mql.removeListener(auf);
    };
  }, []);

  return mobil;
}
