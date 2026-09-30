/**
 * v32.45 — Teams-Link aus dem EIGENEN Kalender der Person.
 *
 * Im Modus „DEX erzeugt die Teams-Besprechung" (`outlookIsOnlineMeeting`)
 * entsteht der Link erst im Flow, im Gruppenpostfach — DEX speichert ihn
 * nirgends, `eventTeamsLink(event)` ist deshalb leer, und „Meine Events"
 * zeigte nur „Microsoft Teams" als Ort ohne Knopf (Nutzer-Befund 30.09.2026).
 *
 * Die Einladung liegt aber als Kopie im Kalender der Person, mit derselben
 * `iCalUId` wie der Termin im Gruppenpostfach (= `CalendarLink`). Die Kopie
 * trägt `onlineMeeting.joinUrl`. Gelesen wird über `/me/events` — dafür
 * reicht das bereits freigegebene `Calendars.Read.Shared`, und der Zugriff
 * auf das Gruppenpostfach bleibt, wie er ist.
 *
 * Findet sich nichts (Einladung gelöscht, abgelehnt, Graph gesperrt), steht
 * statt eines toten Knopfs der Hinweis, wo der Link liegt.
 */
import * as React from 'react';
import { TeamsJoinButton } from '../TeamsJoinButton';
import { Video } from '../Icons';

// Ergebnis je iCalUId für die Sitzung merken — die Karte rendert oft neu, und
// jeder Graph-Aufruf kostet einen Roundtrip. '' = gesucht, nichts gefunden.
const cache = new Map<string, Promise<string>>();

function ladeJoinUrl(icalUid: string): Promise<string> {
  const hit = cache.get(icalUid);
  if (hit) return hit;
  const p = (async (): Promise<string> => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (!ctx || !ctx.msGraphClientFactory) return '';
    try {
      const client = await ctx.msGraphClientFactory.getClient('3');
      const resp = await client.api('/me/events')
        .filter(`iCalUId eq '${icalUid.replace(/'/g, "''")}'`)
        .select('id,isOnlineMeeting,onlineMeeting')
        .top(1)
        .get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ev: any = (resp && resp.value || [])[0];
      const url = String((ev && ev.onlineMeeting && ev.onlineMeeting.joinUrl) || '');
      return /^https?:\/\//i.test(url) ? url : '';
    } catch (err) {
      console.warn('[DEX] Teams-Link aus eigenem Kalender nicht lesbar:', err);
      // Fehler nicht dauerhaft merken — beim nächsten Öffnen neu versuchen.
      cache.delete(icalUid);
      return '';
    }
  })();
  cache.set(icalUid, p);
  return p;
}

export interface OwnCalendarTeamsLinkProps {
  calendarLink?: string;
  isDe: boolean;
  /** Nur wenn sicher ein Link existiert (DEX hat die Besprechung erzeugt),
   *  steht bei leerem Ergebnis der Hinweis — sonst gar nichts. */
  hinweisWennLeer?: boolean;
}

export function OwnCalendarTeamsLink({ calendarLink, isDe, hinweisWennLeer = true }: OwnCalendarTeamsLinkProps): React.ReactElement | null {
  const uid = (calendarLink || '').trim();
  const [url, setUrl] = React.useState<string | null>(null);
  React.useEffect(() => {
    let aktiv = true;
    if (!uid) { setUrl(''); return undefined; }
    ladeJoinUrl(uid).then(u => { if (aktiv) setUrl(u); }).catch(() => { if (aktiv) setUrl(''); });
    return () => { aktiv = false; };
  }, [uid]);

  if (url === null) return null;
  if (url) return <TeamsJoinButton url={url} isDe={isDe} variant="link" />;
  if (!hinweisWennLeer) return null;
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
      <Video size={14} /> {isDe ? 'Teams-Link steht im Outlook-Termin' : 'Teams link is in the Outlook invite'}
    </span>
  );
}

export default OwnCalendarTeamsLink;
