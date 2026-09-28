/* v32.0: Was der Ladebildschirm bei einem Deep-Link ansagt.
 *
 * Nutzer-Ansage 28.09.2026: Bei einem Link wie `#action=admin&event=310&open=fa`
 * stand „Welcome to DEX. Just a moment…" — dabei weiß die App schon beim
 * Start, WOHIN es geht. Jetzt steht dort „Deine F&A-Teilnehmerliste wird
 * geladen…". Gilt für alle Deep-Links; ohne Deep-Link bleibt der allgemeine
 * Satz. Die Zuordnung spiegelt die Verzweigung in DexEventPlatform
 * (`didHandleDeepLink`) — wer dort eine Aktion ergänzt, ergänzt sie hier. */
import { deepLinkParams } from './deepLink';

export function deepLinkLadeText(isDe: boolean): string | null {
  let action = '';
  let open = '';
  try {
    const p = deepLinkParams();
    action = p.get('action') || '';
    open = p.get('open') || '';
  } catch { return null; }
  if (!action) return null;
  const t = (de: string, en: string): string => (isDe ? de : en);
  switch (action) {
    case 'admin':
      if (open === 'fa') return t('Deine F&A-Teilnehmerliste wird geladen…', 'Loading your F&A attendee list…');
      if (open === 'concur') return t('Deine Concur-Teilnehmerliste wird geladen…', 'Loading your Concur attendee list…');
      if (open === 'teilnehmer') return t('Deine Teilnehmerliste wird geladen…', 'Loading your attendee list…');
      return t('Dein Event im Organizer Center wird geladen…', 'Loading your event in the Organizer Center…');
    case 'checkin': return t('Dein Check-in wird geladen…', 'Loading your check-in…');
    case 'selfcheckin-display': return t('Deine Check-in-Anzeige wird geladen…', 'Loading your check-in display…');
    case 'register': return t('Deine Anmeldung wird geladen…', 'Loading your registration…');
    case 'cancel': return t('Deine Anmeldung wird geladen…', 'Loading your registration…');
    case 'comms': return t('Deine Nachrichten zum Event werden geladen…', 'Loading your event messages…');
    case 'assistreq': return t('Deine Anforderung wird geladen…', 'Loading your request…');
    case 'event-created':
    case 'event-updated': return t('Dein Event wird geladen…', 'Loading your event…');
    case 'manual': return t('Das Handbuch wird geladen…', 'Loading the manual…');
    case 'tickets': return t('Deine Tickets werden geladen…', 'Loading your tickets…');
    case 'fa': return t('Dein F&A Center wird geladen…', 'Loading your F&A Center…');
    case 'ask': return t('Deine Fragen werden geladen…', 'Loading your questions…');
    default: return null;
  }
}
