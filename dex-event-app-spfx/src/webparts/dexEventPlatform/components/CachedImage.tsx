// v19.22 — kleine Wrapper-Komponenten, die Bilder über den IndexedDB-Bild-Cache
// (utils/imageCache) anzeigen. Nutzbar AUCH innerhalb von .map()-Schleifen
// (z.B. Listen-Zeilen), weil jede Instanz den Hook an ihrem eigenen Top-Level
// aufruft — anders als ein Hook direkt in der Schleife (verstößt gegen die
// Rules-of-Hooks).

import * as React from 'react';
import { useCachedImage } from '../utils/imageCache';

/**
 * Hintergrund-Bild-Div über den Cache. Rendert nichts, wenn keine URL gesetzt
 * ist. Der Hook wird IMMER aufgerufen (vor dem Early-Return), damit die
 * Rules-of-Hooks gewahrt bleiben.
 */
export function CachedBg({ url, style, position }: {
  url?: string;
  style: React.CSSProperties;
  /** background-position/size — Default 'center/cover no-repeat'. */
  position?: string;
}): React.ReactElement | null {
  const resolved = useCachedImage(url);
  if (!url) return null;
  return <div style={{ ...style, background: `url(${resolved}) ${position || 'center/cover no-repeat'}` }} />;
}

/**
 * <img> über den Cache. Reicht alle übrigen img-Props durch.
 */
export function CachedImg(props: React.ImgHTMLAttributes<HTMLImageElement>): React.ReactElement {
  const resolved = useCachedImage(props.src);
  return <img {...props} src={resolved || props.src} />;
}
