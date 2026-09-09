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
 * Diese Datei ist die EINE Stelle, die einen gespeicherten Mail-Body
 * anzeigefertig macht. Wer eine dritte Vorschau baut, ruft sie ebenfalls —
 * sonst zeigt die nächste Ansicht wieder ein rotes Kreuz.
 */

/** Platzhalter, die der Flow beim Versand ersetzt und die App nicht kennt. */
const FLOW_PLACEHOLDER = /\{\{[A-Z_]+\}\}/;

/**
 * Macht einen gespeicherten Mail-Body anzeigefertig.
 *
 * @param body      Roher HTML-Body aus DEX_Emails / dem Kommunikations-Log.
 * @param orbBase64 Bild für `{{ORB_URL}}` (Event-Mail-Logo oder der
 *                  zwischengespeicherte DEX-Orb). Leer lassen ist erlaubt —
 *                  dann fliegt das Bild raus, statt kaputt dazustehen.
 */
export function resolveMailPreviewHtml(body: string, orbBase64?: string): string {
  if (!body) return '';
  let html = body;

  // 1) Bekannte Bild-Platzhalter auflösen, soweit wir einen Wert haben.
  if (orbBase64) {
    html = html.replace(/\{\{ORB_URL\}\}/g, orbBase64);
    html = html.replace(/\{\{LOGO_URL\}\}/g, orbBase64);
  }

  // 2) Was danach noch als Platzhalter im `src` steht, kann der Browser nicht
  //    laden. Ein leeres `src` wäre nicht besser (Chrome zeigt dann dasselbe
  //    kaputte Symbol) — deshalb fliegt das ganze <img> raus. Der Text der
  //    Mail bleibt vollständig, es fehlt nur die Dekoration, die es in der
  //    gespeicherten Fassung ohnehin nie gab.
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
