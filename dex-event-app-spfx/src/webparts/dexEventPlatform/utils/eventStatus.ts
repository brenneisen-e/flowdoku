/**
 * v28.94: Aus `AdminPage` herausgeloest. Formatierung und Uebersetzung von
 * Datum und Event-Status — reine Funktionen ohne Bezug auf den Seiten-State.
 */
export function formatDate(iso: string): string {
  if (!iso) return '-';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '-';
  return d.toLocaleDateString('de-DE', {
    day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  });
}
export function getStatusColor(status: string): string {
  switch (status) {
    case 'Active': return 'var(--dex-green)';
    case 'Completed': return 'var(--dex-gray-400)';
    case 'Cancelled': return 'var(--dex-red)';
    default: return 'var(--dex-orange)';
  }
}
export // v9.20: EventStatus-Labels lokalisieren (DE).
// v11.89: 'Under Construction' wird transparent als 'Entwurf' angezeigt,
// solange noch Legacy-Daten existieren — neue Events nutzen IsFictive.
function localizeStatus(status: string): string {
  switch (status) {
    case 'Active': return 'Aktiv';
    case 'Under Construction': return 'Entwurf';
    case 'Completed': return 'Abgeschlossen';
    case 'Cancelled': return 'Abgesagt';
    default: return status;
  }
}
export // v22.16: Heuristik für die „Hinweise"-Box bei aktiven Events — erkennt
// englischsprachigen Event-Inhalt (Beschreibung + Felder), damit die App
// empfehlen kann, die Anmeldesprache fest auf Englisch zu stellen (sonst
// mischt das Formular je nach App-Sprache des Teilnehmers Deutsch/Englisch).
function stripHtmlToText(html: string): string {
  return (html || '').replace(/<[^>]+>/g, ' ').replace(/&[a-z]+;/gi, ' ').replace(/\s+/g, ' ').trim();
}
export function looksEnglishText(text: string): boolean {
  const t = ' ' + (text || '').toLowerCase().replace(/\s+/g, ' ') + ' ';
  if (t.trim().length < 40) return false; // zu wenig Text für ein Urteil
  const enWords = [' the ', ' and ', ' you ', ' your ', ' please ', ' with ', ' for ', ' our ', ' this ', ' are ', ' join ', ' we ', ' from ', ' all '];
  const deWords = [' der ', ' die ', ' das ', ' und ', ' nicht ', ' bitte ', ' wir ', ' euch ', ' dich ', ' ihr ', ' eine ', ' einen ', ' zur ', ' zum ', ' bei ', ' auf '];
  let en = 0;
  let de = 0;
  for (const w of enWords) if (t.indexOf(w) >= 0) en++;
  for (const w of deWords) if (t.indexOf(w) >= 0) de++;
  // Umlaute/ß sind ein starkes Deutsch-Signal.
  if (/[äöüß]/.test(t)) de += 2;
  return en >= 4 && en >= de * 2;
}
export // Status-Werte sind in SP als deutsche Strings gespeichert ('Angemeldet',
// 'QR versendet', 'Warteliste', 'Eingecheckt', 'Abgemeldet'). Die App
// rendert sie hier in der UI-Sprache des Users, ohne den Datenbankwert
// zu ändern.
function translateStatus(status: string, isDe: boolean): string {
  if (isDe || !status) return status;
  switch (status) {
    case 'Angemeldet': return 'Registered';
    case 'QR versendet': return 'QR sent';
    case 'Warteliste': return 'Waitlist';
    case 'Eingecheckt': return 'Checked in';
    case 'No-Show': return 'No-show';
    case 'Abgemeldet': return 'Cancelled';
    default: return status;
  }
}
