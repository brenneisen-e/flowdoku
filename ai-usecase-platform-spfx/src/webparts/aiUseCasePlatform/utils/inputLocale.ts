/**
 * v30.60: Deutsches Datums- und Uhrzeitformat in den nativen Eingabefeldern.
 *
 * Nutzer-Befund 01.09.2026: Im Agenda-Block standen „09/09/2026" und ein
 * Zeit-Feld mit „04 : 41 PM". Das ist kein Fehler im Code, der die Werte
 * schreibt — gespeichert wird ohnehin ISO. Es ist die **Anzeige** von
 * `<input type="date">` und `<input type="time">`, und die richtet sich nach
 * der Sprache, in der der Browser die Seite sieht. Steht dort Englisch (US),
 * zeigt Chrome MM/DD/YYYY und die 12-Stunden-Uhr — auch wenn die gesamte
 * Oberfläche deutsch ist.
 *
 * **Beeinflussbar ist das ausschließlich über das `lang`-Attribut**, nicht
 * über CSS und nicht über den Wert. Chrome und Edge lesen das `lang` des
 * Elements bzw. des nächsten Vorfahren, der eines trägt.
 *
 * Deshalb EIN Ort statt eines Attributs an jedem Feld: `APP_LANG_ATTR` wird
 * am äußersten Container der App gesetzt und vererbt sich auf jedes
 * Eingabefeld darin — auch auf die, die es noch gar nicht gibt. Ein Helfer,
 * den man an jedem `<input>` vergessen kann, hätte genau dieselbe Lücke
 * wieder aufgerissen.
 *
 * Für Englisch steht bewusst `en-GB` und nicht `en-US`: Britisches Englisch
 * schreibt DD/MM/YYYY und rechnet in 24 Stunden — das ist beides das, was in
 * Deutschland erwartet wird, auch von denen, die die App auf Englisch nutzen.
 *
 * Firefox richtet sich nach der Spracheinstellung des Browsers und ignoriert
 * `lang`; dort bleibt die Anzeige, wie der Nutzer sie eingestellt hat. Das
 * lässt sich von einer Webseite aus nicht ändern — offen gesagt, statt es zu
 * behaupten.
 */

/** Sprach-Tag für die native Datums-/Zeit-Darstellung. */
export function inputLocaleTag(isDe: boolean): string {
  return isDe ? 'de-DE' : 'en-GB';
}
