// v18.33 — Self-Check-in Hilfsfunktionen
//
// Zwei QR-Modi teilen sich denselben geheimen Token pro Event:
//
//  1. Statischer QR (druckbares PDF): die URL enthält den Token direkt
//     (?action=selfcheckin&token=<token>). Bequem und druckbar, aber per
//     Foto teilbar — daher nur in Kombination mit dem Event-Tag-/Zeitfenster
//     wirklich sinnvoll.
//
//  2. Rotierender Live-QR (Bildschirm am Eingang): die URL enthält KEINEN
//     Token, sondern die Event-Nummer plus einen zeitbasierten HMAC-Code und
//     den Fenster-Index (?action=selfcheckin&event=<Nr>&code=<hmac>&t=<window>).
//     Der Code wechselt alle SELF_CHECKIN_STEP_SECONDS Sekunden. Beim Scannen
//     prüft die App, ob der Code aus dem aktuellen (oder dem direkt
//     vorhergehenden) Zeitfenster stammt — ein abfotografierter Code ist
//     dadurch nach spätestens ~2 Fenstern wertlos.
//
// Die HMAC-Validierung läuft komplett clientseitig über Web Crypto
// (crypto.subtle) — kein Server, keine Power-Automate-Änderung nötig.

/** Länge eines Rotations-Fensters in Sekunden. 45s = guter Kompromiss aus
 *  Sicherheit (kurz genug, dass ein Foto schnell veraltet) und Komfort
 *  (lang genug, dass der Scan im aktuellen Fenster gelingt). */
export const SELF_CHECKIN_STEP_SECONDS = 45;

/** Aktueller Fenster-Index (ganzzahlig) auf Basis der Unix-Zeit. */
export function currentWindowIndex(nowMs: number = Date.now()): number {
  return Math.floor(nowMs / 1000 / SELF_CHECKIN_STEP_SECONDS);
}

/** Kryptografisch zufälligen Token erzeugen (Schlüssel + statischer Link). */
export function generateSelfCheckInToken(): string {
  try {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
      return crypto.randomUUID().replace(/-/g, '');
    }
  } catch { /* fallthrough */ }
  // Fallback ohne randomUUID
  return (
    Date.now().toString(36) +
    Math.random().toString(36).slice(2) +
    Math.random().toString(36).slice(2)
  ).slice(0, 32);
}

/** HMAC-SHA256(secret, windowIndex), gekürzt auf 12 Hex-Zeichen für einen
 *  kompakten QR-Code. Ausreichend gegen zufälliges Raten. */
export async function computeRotatingCode(secret: string, windowIndex: number): Promise<string> {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw',
    enc.encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(String(windowIndex)));
  const bytes = new Uint8Array(sig);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    hex += bytes[i].toString(16).padStart(2, '0');
  }
  return hex.substring(0, 12);
}

/** Prüft, ob ein rotierender Code für das übergebene Fenster gültig UND frisch
 *  ist. Akzeptiert das aktuelle und das direkt vorhergehende Fenster (Toleranz
 *  für Uhren-Drift und langsame Scans), nicht jedoch ältere. */
export async function verifyRotatingCode(
  secret: string,
  code: string,
  windowIndex: number,
  nowMs: number = Date.now()
): Promise<boolean> {
  if (!secret || !code || !Number.isFinite(windowIndex)) return false;
  const current = currentWindowIndex(nowMs);
  // Nur aktuelles oder vorhergehendes Fenster gilt als frisch.
  if (windowIndex !== current && windowIndex !== current - 1) return false;
  const expected = await computeRotatingCode(secret, windowIndex);
  // konstante-Zeit-Vergleich ist hier overkill, einfacher Vergleich genügt
  return expected === code;
}

/** Basis-URL der App-Seite (ohne Query/Hash) — Grundlage aller Deep-Links. */
export function appPageBaseUrl(): string {
  return window.location.origin + window.location.pathname;
}

/** Statischer Check-in-Link für das druckbare PDF. */
export function buildStaticCheckInUrl(token: string): string {
  return `${appPageBaseUrl()}?action=selfcheckin&token=${encodeURIComponent(token)}`;
}

/** Rotierender Check-in-Link für die Live-Anzeige (frischen Code berechnen). */
export async function buildRotatingCheckInUrl(
  secret: string,
  eventNumber: number,
  nowMs: number = Date.now()
): Promise<{ url: string; windowIndex: number; expiresInSeconds: number }> {
  const windowIndex = currentWindowIndex(nowMs);
  const code = await computeRotatingCode(secret, windowIndex);
  const url =
    `${appPageBaseUrl()}?action=selfcheckin&event=${eventNumber}` +
    `&code=${code}&t=${windowIndex}`;
  const elapsed = (nowMs / 1000) % SELF_CHECKIN_STEP_SECONDS;
  const expiresInSeconds = Math.max(1, Math.ceil(SELF_CHECKIN_STEP_SECONDS - elapsed));
  return { url, windowIndex, expiresInSeconds };
}

/** Default-Fensterprüfung: ist „jetzt" innerhalb des erlaubten Check-in-
 *  Zeitraums? Wenn from/to gesetzt sind, gilt [from, to]. Sonst gilt der
 *  Event-Tag: vom Tagesbeginn des Start-Datums bis Tagesende des End-Datums. */
export function isWithinCheckInWindow(
  startDateIso: string | undefined,
  endDateIso: string | undefined,
  fromIso: string | undefined,
  toIso: string | undefined,
  nowMs: number = Date.now()
): { ok: boolean; opensAt?: Date; closesAt?: Date } {
  if (fromIso || toIso) {
    const from = fromIso ? new Date(fromIso).getTime() : -Infinity;
    const to = toIso ? new Date(toIso).getTime() : Infinity;
    return { ok: nowMs >= from && nowMs <= to, opensAt: fromIso ? new Date(fromIso) : undefined, closesAt: toIso ? new Date(toIso) : undefined };
  }
  // Default: Event-Tag(e)
  if (!startDateIso) return { ok: true };
  const start = new Date(startDateIso);
  const opensAt = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0, 0, 0, 0);
  const endRef = endDateIso ? new Date(endDateIso) : start;
  const closesAt = new Date(endRef.getFullYear(), endRef.getMonth(), endRef.getDate(), 23, 59, 59, 999);
  return { ok: nowMs >= opensAt.getTime() && nowMs <= closesAt.getTime(), opensAt, closesAt };
}
