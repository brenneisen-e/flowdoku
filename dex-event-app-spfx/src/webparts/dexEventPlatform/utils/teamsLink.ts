/**
 * v29.39 — Teams-Besprechungslink eines Events finden.
 *
 * Zwei Quellen, weil es zwei Wege gibt, wie der Link ans Event kommt:
 *
 *  1. Das Feld aus v29.38 (`teamsLink`, Piggyback `_teamsLink`) — der gedachte
 *     Weg, seit es das Feld gibt.
 *  2. Der ORT. Bevor es das Feld gab, haben Organizer den Link schlicht in „Ort"
 *     geschrieben („Teams Meeting https://teams.microsoft.com/meet/…"). Dort
 *     stand er als roher Text: nicht klickbar, und in der Detail-Zeile lief er
 *     aus der Karte. Diese Events gibt es weiter, und sie sollen ohne Nacharbeit
 *     einen Teilnahme-Knopf bekommen.
 *
 * Erkannt werden Teams-URLs (teams.microsoft.com / teams.live.com), NICHT
 * beliebige Links — ein Anfahrtsplan im Ort-Feld ist kein Meeting.
 */

const TEAMS_URL_RE = /https?:\/\/(?:[a-z0-9-]+\.)*(?:teams\.microsoft\.com|teams\.live\.com)\/[^\s"'<>)\]]+/i;

/** Erste Teams-URL in einem Text, sonst ''. Trailing-Satzzeichen fallen weg. */
export function extractTeamsUrl(text?: string): string {
  if (!text) return '';
  const m = text.match(TEAMS_URL_RE);
  if (!m) return '';
  return m[0].replace(/[.,;:]+$/, '');
}

/** Der Teams-Link eines Events: gepflegtes Feld zuerst, sonst aus dem Ort. */
export function eventTeamsLink(ev?: { teamsLink?: string; location?: string; outlookLocation?: string } | null): string {
  if (!ev) return '';
  const own = (ev.teamsLink || '').trim();
  if (/^https?:\/\//i.test(own)) return own;
  return extractTeamsUrl(ev.location) || extractTeamsUrl(ev.outlookLocation);
}

/**
 * Der Ort OHNE die Teams-URL — für die Anzeige, wenn der Link daneben schon als
 * Knopf steht. Bleibt nichts übrig („Teams Meeting <URL>" → „Teams Meeting"),
 * liefert die Funktion den Rest bzw. einen leeren String; der Aufrufer
 * entscheidet dann, ob er die Zeile überhaupt zeigt.
 */
export function locationWithoutTeamsUrl(location?: string): string {
  if (!location) return '';
  return location.replace(TEAMS_URL_RE, '').replace(/\s{2,}/g, ' ').replace(/[\s|,;–-]+$/, '').trim();
}
