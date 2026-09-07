/**
 * v30.94: Der Standard-Text des Outlook-Termins — an EINER Stelle.
 *
 * Nutzer-Befund 07.09.2026: „warum wird der Outlook-Termin im Event Creation
 * zweimal unterschiedlich gerendert?" Der Editor zeigte bei leerem Body einen
 * Platzhalter-Satz, die Vorschau-Karte einen erfundenen Fallback („hier ist
 * dein Kalendereintrag"), und gespeichert wurde ein DRITTER Text — der echte
 * Standard aus wizardSubmit/persistSubEvents (v7.4: Anmeldung + Abmelde-
 * Hinweis + Organizer-Kontakt). Vier Stellen, drei Texte.
 *
 * Jetzt liefert `outlookDefaultBodyTemplate(lang)` die Vorlage MIT
 * Platzhaltern ({{EventTitle}}, {{Organizer}}); wer den fertigen Text braucht,
 * löst sie mit `replacePlaceholders` auf — Editor („Standardtext laden" und
 * Vorschau bei leerem Body), Vorschau-Karte und beide Speicherpfade gleich.
 */

export const OUTLOOK_APP_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';

export function outlookDefaultBodyTemplate(lang: string): string {
  return (lang || '').toUpperCase() === 'EN'
    ? '<p>You are registered for the event <strong>{{EventTitle}}</strong>.</p>'
      + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${OUTLOOK_APP_URL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>`
      + '<p>For organizational questions please contact <strong>{{Organizer}}</strong>.</p>'
    : '<p>Ihr seid für das Event <strong>{{EventTitle}}</strong> angemeldet.</p>'
      + `<p>Falls ihr nicht teilnehmen könnt, meldet euch bitte rechtzeitig über die <a href="${OUTLOOK_APP_URL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
      + '<p>Bei organisatorischen Fragen wendet euch bitte an <strong>{{Organizer}}</strong>.</p>';
}

/** Fallback für {{Organizer}}, wenn (noch) niemand eingetragen ist. */
export function outlookOrganizerFallback(lang: string): string {
  return (lang || '').toUpperCase() === 'EN' ? 'the organizer' : 'den Organizer';
}
