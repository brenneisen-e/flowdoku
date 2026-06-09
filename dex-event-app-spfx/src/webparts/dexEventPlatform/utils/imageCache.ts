// v19.22 — Browser-seitiger Bild-Cache (IndexedDB)
//
// Event-Bilder werden in SharePoint als Item-Attachments gespeichert; ihre URL
// ist pro Bild-Version stabil (der Dateiname enthält einen Zeitstempel, der sich
// nur beim Ersetzen ändert). Das macht sie ideal cachebar.
//
// Ziel: Beim ZWEITEN (und jedem weiteren) Aufruf der App sollen Event-Bilder
// SOFORT erscheinen — ohne erneuten Netzwerk-Roundtrip zu SharePoint, unabhängig
// davon, welche Cache-Header SharePoint mitschickt. Dazu speichern wir pro
// Bild-URL einen Data-URL-String in IndexedDB. Beim nächsten Aufruf liefert der
// Cache den Data-URL direkt aus der lokalen Datenbank.
//
// Robustheit: JEDER Fehlerpfad fällt auf die Original-URL zurück — die Anzeige
// bricht nie. IndexedDB nicht verfügbar, fetch schlägt fehl, Bild zu groß: in
// allen Fällen wird einfach die Original-URL verwendet (= Verhalten wie vor
// v19.22).

import * as React from 'react';

const DB_NAME = 'dex-image-cache';
const STORE = 'images';
const DB_VERSION = 1;
// Grobe Obergrenze an Einträgen — älteste werden geprunt. Event-Bilder sind nach
// Komprimierung klein (~50–300 KB), 150 Einträge bleiben deutlich unter dem
// IndexedDB-Quota.
const MAX_ENTRIES = 150;
// Bilder über dieser Größe NICHT cachen (Data-URL würde das Quota unnötig füllen).
const MAX_BYTES = 4 * 1024 * 1024;

interface CacheRec {
  url: string;
  data: string; // Data-URL (base64)
  ts: number;   // letzter Zugriff (für LRU-Pruning)
}

// In-Memory-Cache pro Session — wiederholte Aufrufe für dieselbe URL gehen nicht
// jedes Mal in IndexedDB.
const memCache = new Map<string, string>();
const inflight = new Map<string, Promise<string>>();

let dbPromise: Promise<IDBDatabase | null> | null = null;

function openDb(): Promise<IDBDatabase | null> {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise<IDBDatabase | null>((resolve) => {
    try {
      if (typeof indexedDB === 'undefined') { resolve(null); return; }
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          const os = db.createObjectStore(STORE, { keyPath: 'url' });
          os.createIndex('ts', 'ts');
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch { resolve(null); }
  });
  return dbPromise;
}

function idbGet(db: IDBDatabase, url: string): Promise<CacheRec | null> {
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).get(url);
      req.onsuccess = () => resolve((req.result as CacheRec) || null);
      req.onerror = () => resolve(null);
    } catch { resolve(null); }
  });
}

function idbPut(db: IDBDatabase, rec: CacheRec): void {
  try {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(rec);
  } catch { /* ignore */ }
}

// Älteste Einträge entfernen, wenn der Store über MAX_ENTRIES wächst. Best-effort,
// blockiert nichts.
function pruneIfNeeded(db: IDBDatabase): void {
  try {
    const tx = db.transaction(STORE, 'readwrite');
    const os = tx.objectStore(STORE);
    const countReq = os.count();
    countReq.onsuccess = () => {
      const over = countReq.result - MAX_ENTRIES;
      if (over <= 0) return;
      let removed = 0;
      // Über den ts-Index aufsteigend (älteste zuerst) löschen.
      const idx = os.index('ts');
      const cursorReq = idx.openCursor();
      cursorReq.onsuccess = () => {
        const cursor = cursorReq.result;
        if (!cursor || removed >= over) return;
        cursor.delete();
        removed++;
        cursor.continue();
      };
    };
  } catch { /* ignore */ }
}

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result as string);
    fr.onerror = () => reject(fr.error);
    fr.readAsDataURL(blob);
  });
}

/**
 * Liefert für eine Bild-URL die gecachte Data-URL (falls vorhanden), sonst lädt
 * sie das Bild, speichert es im Cache und gibt die Data-URL zurück. Bei JEDEM
 * Fehler wird die Original-URL zurückgegeben — die Anzeige bricht nie.
 */
export async function getCachedImage(url: string): Promise<string> {
  if (!url || url.indexOf('data:') === 0) return url || '';
  const mem = memCache.get(url);
  if (mem) return mem;
  const existing = inflight.get(url);
  if (existing) return existing;

  const p = (async (): Promise<string> => {
    try {
      const db = await openDb();
      if (db) {
        const rec = await idbGet(db, url);
        if (rec && rec.data) {
          memCache.set(url, rec.data);
          // Zugriffszeit aktualisieren (LRU) — fire-and-forget.
          idbPut(db, { url, data: rec.data, ts: Date.now() });
          return rec.data;
        }
      }
      // Nicht im Cache → laden. force-cache nutzt den Browser-HTTP-Cache, falls
      // das <img> dieselbe URL parallel schon geladen hat (vermeidet Doppel-
      // Download beim ersten Mal).
      const resp = await fetch(url, { credentials: 'include', cache: 'force-cache' });
      if (!resp.ok) return url;
      const blob = await resp.blob();
      if (blob.size > MAX_BYTES || blob.type.indexOf('image/') !== 0) return url;
      const dataUrl = await blobToDataUrl(blob);
      memCache.set(url, dataUrl);
      if (db) { idbPut(db, { url, data: dataUrl, ts: Date.now() }); pruneIfNeeded(db); }
      return dataUrl;
    } catch {
      return url;
    } finally {
      inflight.delete(url);
    }
  })();
  inflight.set(url, p);
  return p;
}

/**
 * React-Hook: gibt die bestmögliche Bild-URL zurück. Initial die Original-URL
 * (bzw. einen bereits im Memory-Cache liegenden Data-URL → sofort), danach wird
 * im Hintergrund der Data-URL aus IndexedDB aufgelöst und eingetauscht. Da beide
 * dasselbe Bild zeigen, ist der Tausch optisch nahtlos. Beim nächsten App-Aufruf
 * liefert der Cache das Bild ohne SharePoint-Roundtrip.
 */
export function useCachedImage(url: string | undefined): string {
  const initial = url ? (memCache.get(url) || url) : '';
  const [resolved, setResolved] = React.useState<string>(initial);

  React.useEffect(() => {
    let cancelled = false;
    if (!url) { setResolved(''); return undefined; }
    const mem = memCache.get(url);
    setResolved(mem || url);
    if (!mem) {
      getCachedImage(url).then(d => { if (!cancelled && d) setResolved(d); }).catch(() => { /* Original-URL bleibt */ });
    }
    return () => { cancelled = true; };
  }, [url]);

  return resolved;
}

/**
 * Bilder im Hintergrund vorwärmen (z.B. direkt nach dem Laden der Event-Liste),
 * damit sie beim Öffnen eines Events bereits im Cache liegen. Fire-and-forget,
 * sequenziell mit Mini-Yield, um den Main-Thread nicht zu blockieren.
 */
export function prewarmImages(urls: Array<string | undefined>): void {
  const list = urls.filter((u): u is string => !!u && u.indexOf('data:') !== 0 && !memCache.has(u));
  if (list.length === 0) return;
  let i = 0;
  const step = (): void => {
    if (i >= list.length) return;
    const u = list[i++];
    getCachedImage(u).catch(() => { /* ignore */ }).then(() => {
      // Nächstes Bild im nächsten Tick — kein Thread-Block bei vielen Bildern.
      window.setTimeout(step, 0);
    });
  };
  step();
}
