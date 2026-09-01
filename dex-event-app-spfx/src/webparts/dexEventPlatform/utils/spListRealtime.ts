/**
 * v24.75 — Echtzeit-Listen-Push für SharePoint Online (ohne eigenes Backend).
 *
 * Nutzt den (von Elio Struyf öffentlich beschriebenen) nativen SharePoint-
 * Socket.IO-Endpunkt: SharePoint liefert pro Liste eine `notificationUrl`, auf
 * die man sich per socket.io-client verbindet und „notification"-Events
 * empfängt, sobald sich die Liste ändert. Damit bekommt der Browser echte
 * Push-Benachrichtigungen über Listen-Änderungen — kein Polling, kein Azure,
 * keine zusätzlichen Graph-Rechte.
 *
 * WICHTIG:
 * - Inoffizieller Endpunkt → strikt best-effort. Schlägt irgendetwas fehl
 *   (Permissions/Endpoint/Bundling/Verbindung), passiert NICHTS Schlimmes: kein
 *   Live-Update, aber auch kein Crash. Aufrufer hat weiterhin seinen
 *   Lade-/Fokus-Refresh.
 * - `socket.io-client` wird LAZY (`await import`) geladen, damit das Haupt-Bundle
 *   schlank bleibt (Konvention: schwere Libs nur dynamisch).
 * - Eine Verbindung pro überwachter Liste. Aufrufer MUSS die zurückgegebene
 *   Cleanup-Funktion beim Unmount aufrufen.
 */
import { SPHttpClient } from '@microsoft/sp-http';
import { dlog } from './debugLog';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractId(data: any): string {
  if (!data) return '';
  if (typeof data === 'string') return data.replace(/[{}]/g, '');
  const v = data.value ?? data.d?.value ?? data.d?.Id ?? data.Id ?? '';
  return String(v).replace(/[{}]/g, '');
}

/**
 * Auf Änderungen einer SharePoint-Liste lauschen. Liefert eine Cleanup-Funktion.
 * @param spHttpClient SPFx-HTTP-Client (mit User-Kontext).
 * @param webUrl       Absolute URL der (Sub-)Site, in der die Liste liegt.
 * @param listTitle    Listentitel (z.B. „Teilnehmer" oder „DEX_TeilnehmerCounter").
 * @param onChange     Callback bei jeder gemeldeten Listen-Änderung.
 */
export async function subscribeListChanges(
  spHttpClient: SPHttpClient,
  webUrl: string,
  listTitle: string,
  onChange: () => void
): Promise<() => void> {
  // v24.75: Einheitlich geprefixte Diagnose-Logs, damit man im Browser-Console
  // sofort sieht, ob der Echtzeit-Push funktioniert (der Endpunkt ist
  // inoffiziell + wird nur im Tenant getestet).
  const LOG = '[DEX Realtime]';
  let disposed = false;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let socket: any = null;
  const cleanup = (): void => {
    disposed = true;
    try {
      if (socket) {
        // eslint-disable-next-line no-console
        dlog('realtime', `${LOG} Verbindung zu „${listTitle}" wird getrennt.`);
        socket.removeAllListeners?.(); socket.disconnect?.();
      }
    } catch { /* */ }
    socket = null;
  };
  try {
    // eslint-disable-next-line no-console
    dlog('realtime', `${LOG} Abonniere Änderungen an Liste „${listTitle}" …`);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const getJson = async (url: string): Promise<any> => {
      const r = await spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!r.ok) throw new Error(`HTTP ${r.status} for ${url}`);
      return r.json();
    };
    const hostname = new URL(webUrl).hostname;
    const [siteRaw, webRaw, listRaw] = await Promise.all([
      getJson(`${webUrl}/_api/site/id`),
      getJson(`${webUrl}/_api/web/id`),
      getJson(`${webUrl}/_api/web/lists/getbytitle('${listTitle}')/id`),
    ]);
    const siteId = extractId(siteRaw);
    const webId = extractId(webRaw);
    const listId = extractId(listRaw);
    if (!siteId || !webId || !listId) throw new Error('Listen-/Site-IDs nicht auflösbar');
    if (disposed) return cleanup;
    const composite = `${hostname},${siteId},${webId}`;
    const subResp = await getJson(
      `${webUrl}/_api/v2.1/sites('${composite}')/lists('${listId}')/subscriptions/socketIo`
    );
    const notificationUrl: string = subResp?.notificationUrl || subResp?.value?.notificationUrl || subResp?.d?.notificationUrl;
    if (!notificationUrl) throw new Error('keine notificationUrl erhalten');
    if (disposed) return cleanup;
    // eslint-disable-next-line no-console
    dlog('realtime', `${LOG} notificationUrl für „${listTitle}" erhalten — verbinde Socket …`);
    // socket.io-client lazy laden.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod: any = await import('socket.io-client');
    if (disposed) return cleanup;
    const io = mod.io || mod.default || mod;
    socket = io(notificationUrl, { transports: ['websocket'], reconnectionAttempts: 5 });
    socket.on('connect', () => {
      // eslint-disable-next-line no-console
      dlog('realtime', `${LOG} ✅ Verbunden — Live-Updates für „${listTitle}" sind aktiv.`);
    });
    socket.on('notification', () => {
      // eslint-disable-next-line no-console
      dlog('realtime', `${LOG} 🔔 Änderung an „${listTitle}" empfangen → aktualisiere die Anzeige.`);
      if (!disposed) { try { onChange(); } catch { /* */ } }
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    socket.on('connect_error', (err: any) => {
      // eslint-disable-next-line no-console
      console.warn(`${LOG} ⚠️ Verbindungsfehler für „${listTitle}" — kein Live-Update (Lade-/Fokus-Refresh greift weiter):`, err?.message || err);
    });
    socket.on('disconnect', () => {
      // eslint-disable-next-line no-console
      dlog('realtime', `${LOG} Socket für „${listTitle}" getrennt.`);
    });
  } catch (e) {
    // eslint-disable-next-line no-console
    console.warn(`${LOG} ❌ Echtzeit-Push für „${listTitle}" nicht möglich (best-effort — Anzeige bleibt korrekt, nur nicht live):`, e);
  }
  return cleanup;
}
