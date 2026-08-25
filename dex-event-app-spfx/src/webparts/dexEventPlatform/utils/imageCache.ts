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

// v29.34: URLs, die der Server mit einem HTTP-Fehler beantwortet hat (404/403 —
// Attachment gelöscht oder nie angekommen). Bewusst NUR bei einer echten
// Fehlerantwort gefüllt: Ein abgebrochener fetch (Netz, CORS) heißt nicht, dass
// die Datei fehlt — ein <img> auf dieselbe URL lädt dann oft trotzdem.
const failedUrls = new Set<string>();

/** v29.34: true, wenn diese URL in dieser Sitzung mit HTTP-Fehler kam. */
export function imageFailed(url: string | undefined): boolean {
  return !!url && failedUrls.has(url);
}

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
      // v29.62: NUR endgültige Fehler merken. Bis hierher landete JEDE
      // Fehlerantwort auf der Sperrliste — auch ein 429 (SharePoint drosselt)
      // oder ein 5xx. Genau die kommen aber gehäuft beim App-Start vor, und
      // die Sperrliste gilt für die ganze Sitzung: Wer die App öffnet und
      // sofort ein Event aufruft, sah das Bild dann nicht mehr, bis er neu
      // lud. 404/403/410 heißen „gibt es nicht (mehr)" — alles andere heißt
      // „gerade nicht", und das darf kein Dauerzustand werden.
      if (!resp.ok) {
        if (resp.status === 404 || resp.status === 403 || resp.status === 410) failedUrls.add(url);
        return url;
      }
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
/**
 * v29.62: Ein Abruf, der nicht am fehlenden Bild lag (429/5xx/Netz), wird EINMAL
 * wiederholt. Ohne das bleibt die Kachel bzw. der Kopf der Anmeldeseite bis zum
 * Neuladen leer, obwohl das Bild existiert — der gemeldete Fall „App geöffnet,
 * sofort ins Event, kein Bild".
 */
const RETRY_MS = 1500;

export function useCachedImage(url: string | undefined): string {
  const initial = url ? (memCache.get(url) || url) : '';
  const [resolved, setResolved] = React.useState<string>(initial);

  React.useEffect(() => {
    let cancelled = false;
    if (!url) { setResolved(''); return undefined; }
    const mem = memCache.get(url);
    setResolved(mem || url);
    let timer: ReturnType<typeof setTimeout> | null = null;
    if (!mem) {
      // v29.62: Zweiter Versuch, wenn der erste nichts Gecachtes brachte und
      // die URL nicht als endgueltig fehlend gilt. `getCachedImage` liefert bei
      // einem Fehlschlag die Original-URL zurueck — daran erkennt man ihn.
      getCachedImage(url)
        .then(d => {
          if (cancelled) return;
          if (d) setResolved(d);
          if (d === url && !imageFailed(url)) {
            timer = setTimeout(() => {
              getCachedImage(url)
                .then(d2 => { if (!cancelled && d2) setResolved(d2); })
                .catch(() => { /* Original-URL bleibt stehen */ });
            }, RETRY_MS);
          }
        })
        .catch(() => { /* Original-URL bleibt */ });
    }
    return () => { cancelled = true; if (timer) clearTimeout(timer); };
  }, [url]);

  return resolved;
}

/**
 * v29.34: Wie `useCachedImage`, aber mit zweiter Quelle. Liefert `primary`,
 * solange die lädt, und schaltet auf `fallback` um, sobald der Server für
 * `primary` einen HTTP-Fehler meldet.
 *
 * Hintergrund: Kachel und Anmeldeseite bevorzugen das unbeschnittene
 * Querformat-Original (`imageOrigUrl`, Piggyback `_imageOrigUrl`) vor dem
 * eigentlichen Event-Bild. Diese URL kann ins Leere zeigen — sie wurde bis
 * v29.33 auch dann geschrieben, wenn der Upload des Originals fehlschlug
 * (der Service riet die Adresse), und sie überlebt einen späteren Bildwechsel.
 * Ohne Rückfall blieb die Kachel dann WEISS, während dasselbe Event im
 * Organizer Center (der `imageUrl` zeigt) ein Bild hatte — genau der Fall, der
 * wie ein Anzeigefehler aussieht, aber zwei verschiedene Dateien sind.
 */
export function useCachedImageWithFallback(primary: string | undefined, fallback: string | undefined): string {
  const [primaryDead, setPrimaryDead] = React.useState<boolean>(() => imageFailed(primary));

  React.useEffect(() => {
    let cancelled = false;
    setPrimaryDead(imageFailed(primary));
    if (!primary || !fallback || primary === fallback || primary.indexOf('data:') === 0) return undefined;
    getCachedImage(primary)
      .then(() => { if (!cancelled && imageFailed(primary)) setPrimaryDead(true); })
      .catch(() => { /* Netzwerkfehler heißt nicht „Datei weg" — primary bleibt */ });
    return () => { cancelled = true; };
  }, [primary, fallback]);

  const effective = (primaryDead ? fallback : primary) || fallback || primary || '';
  return useCachedImage(effective);
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
