/* Mail-HTML für die ANZEIGE in der App aufbereiten.
 *
 * Warum es das braucht: Der Body, den DEX in die Queue bzw. ins
 * Kommunikations-Log schreibt, enthält das Kopfbild absichtlich NUR als
 * Platzhalter — `<img src="{{ORB_URL}}" alt="DEX Event Experience Platform">`.
 * Ersetzt wird er erst vom Power-Automate-Flow beim Versand, mit dem
 * event-spezifischen Bild (EmailTemplates.ts, Kommentar an der Hero-Zeile).
 *
 * Für die Mail im Postfach ist das richtig. Wer denselben Body aber in der App
 * anzeigt — „Nachrichten zum Event" für Teilnehmer, das Kommunikations-Log im
 * Organizer Center — bekommt ein kaputtes Bild und daneben den `alt`-Text.
 * Genau so gemeldet am 09.09.2026.
 *
 * ZIEL IST DIE OUTLOOK-ANSICHT. Die App soll dasselbe zeigen wie das Postfach,
 * also dasselbe Bild EINSETZEN — nicht die Zeile wegräumen. Deshalb bildet
 * `resolveMailPreviewHtml` die Reihenfolge des Flows nach:
 *   1. das Mail-Logo des Events (`EmailImageBase64` → `event.mailImageBase64`),
 *   2. sonst der DEX-Standard-Orb aus dem Session-Cache
 *      (`getCachedOrbBase64()`, gefüllt von `loadLogosAsBase64`).
 * Erst wenn BEIDES fehlt, fliegt das `<img>` raus — dann gäbe es auch in
 * Outlook nichts zu sehen, und ein rotes Kreuz wäre die schlechtere Lüge.
 *
 * Diese Datei ist die EINE Stelle, die einen gespeicherten Mail-Body
 * anzeigefertig macht. Wer eine dritte Vorschau baut, ruft sie ebenfalls —
 * sonst zeigt die nächste Ansicht wieder ein rotes Kreuz.
 */

/** Platzhalter, die der Flow beim Versand ersetzt und die App nicht kennt. */
const FLOW_PLACEHOLDER = /\{\{[A-Z_]+\}\}/;

/**
 * Macht einen gespeicherten Mail-Body anzeigefertig.
 *
 * @param body       Roher HTML-Body aus DEX_Emails / dem Kommunikations-Log.
 * @param eventImage Mail-Logo DIESES Events (`event.mailImageBase64`). Das ist
 *                   das Bild, das der Flow einsetzt — hat Vorrang.
 * @param orbFallback Der DEX-Standard-Orb (`getCachedOrbBase64()`), wie ihn
 *                   der Flow nimmt, wenn das Event kein eigenes Logo hat.
 */
export function resolveMailPreviewHtml(body: string, eventImage?: string, orbFallback?: string): string {
  if (!body) return '';
  let html = body;

  // 1) Dieselbe Reihenfolge wie im Flow: Event-Logo vor Standard-Orb.
  const img = (eventImage && eventImage.trim()) || (orbFallback && orbFallback.trim()) || '';
  if (img) {
    html = html.replace(/\{\{ORB_URL\}\}/g, img);
    html = html.replace(/\{\{LOGO_URL\}\}/g, img);
  }

  // 2) Nur wenn WEDER Event-Logo NOCH Orb da waren, steht hier noch ein
  //    Platzhalter. Ein leeres `src` wäre nicht besser (Chrome zeigt dann
  //    dasselbe kaputte Symbol) — deshalb fliegt das ganze <img> raus. Das ist
  //    der Notausgang, nicht der Normalfall.
  html = html.replace(/<img\b[^>]*>/gi, (tag) => {
    const src = /\ssrc\s*=\s*["']([^"']*)["']/i.exec(tag);
    if (!src) return tag;
    return FLOW_PLACEHOLDER.test(src[1]) ? '' : tag;
  });

  // 3) Die Mail ist eine 600-px-Tabelle. In einem schmalen Rahmen entsteht
  //    daraus ein waagerechter Rollbalken, und man liest jede Zeile zweimal
  //    an. Das injizierte Stylesheet lässt die Tabelle schrumpfen, statt sie
  //    zu beschneiden — die Mail bleibt lesbar, auch auf dem Handy.
  const fit = '<style>'
    + 'html,body{margin:0;padding:0;}'
    + 'table{max-width:100%!important;}'
    + 'img{max-width:100%!important;height:auto!important;}'
    + 'td{word-break:break-word;}'
    + '</style>';
  return /<\/head>/i.test(html) ? html.replace(/<\/head>/i, fit + '</head>') : fit + html;
}
