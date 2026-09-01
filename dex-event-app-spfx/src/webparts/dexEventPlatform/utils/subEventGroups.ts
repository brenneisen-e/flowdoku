/**
 * v30.60: Sub-Event-Reiter nach einem gemeinsamen Präfix bündeln.
 *
 * Anlass (Nutzer-Befund 01.09.2026): Ein Training mit 25 Terminen, alle
 * benannt nach dem Muster „Day 1 - Welcome & Intro", „Day 1 - PMO",
 * „Day 2 - Proposals" … Die Reiter-Leiste war dadurch sechs Zeilen hoch, und
 * die Struktur, die in den Namen steckt (fünf Tage), war nur beim Lesen jeder
 * einzelnen Kachel zu erkennen. Gewünscht: erst die Tage, ein Klick zeigt die
 * Termine dieses Tages.
 *
 * **Die Gruppierung kommt aus den Titeln, nicht aus den Daten.** Naheliegender
 * wäre das Startdatum — aber es trägt nicht: Termine ohne eigene Zeiten erben
 * seit v28.66 den Zeitraum des Hauptevents, mehrere Sessions eines Tages haben
 * dasselbe Datum wie die des Vortags, und bei Kalender-Events heißt der Titel
 * ohnehin schon „Di. 08.09.". Der Organizer hat die Struktur in die Namen
 * geschrieben; genau die wird hier gelesen.
 *
 * **Wann NICHT gruppiert wird** — beides bewusst, denn eine Gruppierung, die
 * nichts spart, ist nur eine zusätzliche Ebene:
 *  - weniger als `MIN_SUBS` Termine (eine kurze Leiste liest man am Stück),
 *  - weniger als zwei Gruppen, oder eine Gruppe, in der praktisch alles liegt
 *    (dann ist das Präfix kein Ordnungsmerkmal, sondern nur der Eventname).
 */

/** Ab so vielen Terminen lohnt sich die Bündelung überhaupt. */
const MIN_SUBS = 8;
/** So viele Termine muss eine Gruppe mindestens haben, um eine zu sein. */
const MIN_PER_GROUP = 2;

/**
 * Trennzeichen zwischen Präfix und Bezeichnung, absteigend nach Deutlichkeit.
 * Ein Bindestrich MIT Leerzeichen drumherum trennt; „Carve-out" darf nicht
 * zerfallen, deshalb nie am nackten `-`.
 */
const SEPARATORS = [' - ', ' – ', ' — ', ': ', ' | ', ' · '];

export interface SubEventGroup {
  /** Anzeigename der Gruppe, z.B. „Day 1". */
  label: string;
  /** Indizes in der ursprünglichen Reihenfolge (0-basiert auf die Sub-Events). */
  idxs: number[];
}

export interface SubEventGrouping {
  /** true = die Leiste soll zweistufig rendern. */
  grouped: boolean;
  groups: SubEventGroup[];
}

/** Das Präfix eines Titels, oder '' wenn keines erkennbar ist. */
export function subEventGroupKey(title: string): string {
  const t = (title || '').trim();
  if (!t) return '';
  for (const sep of SEPARATORS) {
    const i = t.indexOf(sep);
    // Ein Präfix ganz am Anfang (i === 0) gibt es nicht, und ein sehr langes
    // ist keines mehr — dann steht der Trenner mitten im Satz.
    if (i > 0 && i <= 28) return t.slice(0, i).trim();
  }
  return '';
}

export function groupSubEventTabs(titles: string[]): SubEventGrouping {
  const flat: SubEventGrouping = { grouped: false, groups: [] };
  if (!titles || titles.length < MIN_SUBS) return flat;

  const byKey: Record<string, number[]> = {};
  const order: string[] = [];
  let ungrouped = 0;
  titles.forEach((t, i) => {
    const key = subEventGroupKey(t);
    if (!key) { ungrouped++; return; }
    if (!byKey[key]) { byKey[key] = []; order.push(key); }
    byKey[key].push(i);
  });

  // Zu kleine Gruppen sind keine — ihre Termine zählen als „ohne Gruppe".
  const groups: SubEventGroup[] = [];
  for (const key of order) {
    if (byKey[key].length < MIN_PER_GROUP) { ungrouped += byKey[key].length; continue; }
    groups.push({ label: key, idxs: byKey[key] });
  }
  if (groups.length < 2) return flat;
  // Mehr als die Hälfte ohne Gruppe: Das Präfix ordnet dann nicht, es
  // beschriftet nur einen Teil — die flache Leiste bleibt ehrlicher.
  if (ungrouped * 2 > titles.length) return flat;

  // Was übrig bleibt, kommt in eine eigene Gruppe am Ende. Es verschwindet
  // nicht: Ein Termin, der in keiner Gruppe auftaucht, wäre über die Leiste
  // nicht mehr erreichbar.
  const rest: number[] = [];
  const claimed: Record<number, true> = {};
  groups.forEach(g => g.idxs.forEach(i => { claimed[i] = true; }));
  titles.forEach((_, i) => { if (!claimed[i]) rest.push(i); });
  if (rest.length > 0) groups.push({ label: 'Weitere', idxs: rest });

  return { grouped: true, groups };
}

/**
 * Die Bezeichnung eines Termins OHNE das Gruppen-Präfix.
 *
 * Innerhalb der Gruppe „Day 1" ist „Day 1 - PMO" doppelt gemoppelt; die
 * Kachel heißt dort einfach „PMO". Passt das Präfix nicht, bleibt der Titel
 * unverändert — lieber einmal zu lang als einmal falsch beschnitten.
 */
export function stripGroupPrefix(title: string, groupLabel: string): string {
  const t = (title || '').trim();
  if (!groupLabel || t.indexOf(groupLabel) !== 0) return t;
  const restWithSep = t.slice(groupLabel.length);
  for (const sep of SEPARATORS) {
    if (restWithSep.indexOf(sep) === 0) {
      const out = restWithSep.slice(sep.length).trim();
      return out || t;
    }
  }
  return t;
}
