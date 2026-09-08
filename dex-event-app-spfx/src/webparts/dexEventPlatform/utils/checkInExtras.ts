/**
 * v30.53: Zusatz-Angaben, die am Check-in-Tisch gebraucht werden.
 *
 * Beim B2Run gibt der Check-in nicht nur „drin/nicht drin" aus — dort wird
 * gleichzeitig das Trikot ausgegeben und die Startnummer zugeordnet. Wer beides
 * nicht auf dem Schirm hat, muss daneben in einer Excel nachsehen, und genau
 * dabei entsteht die Schlange.
 *
 * **Warum eine Muster-Suche und keine feste Feld-Id:** Die Trikotgröße ist kein
 * Feld der B2Run-Vorlage (`data/b2runKoeln.ts`), sondern ein vom Organizer
 * selbst angelegtes Abfragefeld — es heißt an jedem Event etwas anders
 * („T-Shirt Größe", „Trikotgröße", „Shirt size"). Eine feste Id gäbe es also
 * gar nicht zu treffen.
 *
 * **v31.3: Das Muster war zu eng — und zwar genau für das Feld, das die App
 * selbst vorschlägt.** `data/suggestedFields.ts` bietet `b2run_laufshirt` mit
 * dem Label „Deloitte-Laufshirt" bzw. „Deloitte running shirt" an; beide
 * trafen `t-?\s?shirt|shirt\s*size` NICHT (kein „t" vor „shirt", kein „size"
 * dahinter). Wer sein Event genau so anlegte, wie die App es anbietet, hatte
 * die ganze Trikot-Funktion nicht: keine Bestellliste, keinen Bestand, keinen
 * Gegenvorschlag, nichts am Tisch. Seither reicht das blanke Wort „shirt"
 * (deckt Laufshirt, Sweatshirt, running shirt, T-Shirt, Shirtgröße ab), dazu
 * Trikot/Jersey und die Konfektions-Wörter. „Größe" allein bleibt bewusst
 * draußen — daran hing der alte Fehlgriff „Größe des Gepäckstücks".
 *
 * **v31.4: Aus dem Plan wird ein Kassenbuch.** `shirtAllocate` rechnet bei
 * jedem Aufruf frisch aus Wünschen, Bestand und Teilnehmer-ID-Reihenfolge —
 * das ist eine Planung, kein Protokoll. Am Lauftag reicht das nicht: Person 41
 * bekommt um 8:05 ein XL statt des gewünschten M und nimmt es mit; um 16:40
 * markiert das Team fünf Nicht-Erschienene als No-Show, deren Wünsche fallen
 * aus der Rechnung, und die App schlägt Person 41 wieder M vor — für ein
 * Trikot, das sie längst trägt. Seither hält die Spalte `ShirtIssued`
 * (s. `ShirtIssue`) fest, was wirklich ausgegeben wurde; die Ausgabe wird
 * ZUERST vom Bestand abgezogen und schlägt jede Berechnung.
 *
 * **v31.4.3: Eine Antwort steht an ZWEI Stellen — und der Organizer schrieb
 * nur eine davon.** Jede Formularantwort liegt sowohl im JSON `CustomData`
 * (Schlüssel = `f.id`) als auch in der echten SP-Spalte (`f.spInternalName`);
 * die Anmeldung schreibt beide. Der Bearbeiten-Dialog des Organizer Centers
 * schrieb bis v31.4.3 nur die Spalte — dieses Modul las nur `CustomData`.
 * Folge im Live-Fall (08.09.2026, Lauftag am nächsten Tag): Die
 * Teilnehmertabelle zeigte die korrigierte Größe, Bestellliste und
 * Check-in-Tisch dauerhaft die alte. Gelesen wird jetzt überall über
 * `shirtAnswerOf` (Spalte zuerst, `CustomData` als Rückfall — dieselbe
 * Reihenfolge wie `ParticipantTable`), damit auch die vor v31.4.3
 * auseinandergelaufenen Zeilen ohne erneutes Speichern stimmen.
 *
 * **Was das weiterhin kostet, offen gesagt:** Ein Muster trifft irgendwann
 * etwas Falsches. Das ist am Check-in-Tisch harmlos (eine Zeile zu viel), aber
 * es ist geraten und nicht gesagt. Der saubere Weg wäre ein Haken „am Check-in
 * anzeigen" am Abfragefeld selbst; dann kommt diese Datei weg. Bis dahin steht
 * die Regel hier an EINER Stelle statt verteilt im Render-Baum.
 */

/** Feld-Label → gehört ans Check-in? Aktuell: alles, was nach Trikot/Größe klingt.
 *  v30.60 exportiert: Die Auswertung „benötigte T-Shirts" muss DIESELBE Regel
 *  benutzen, sonst zählt die Bestellung andere Felder als der Check-in-Tisch.
 *  v31.3: „shirt" ohne Vorsilbe, dazu jersey/kleidungsgr/clothing size — s.
 *  Dateikopf. */
export const SHIRT_PATTERN = /(shirt|trikot|jersey|kleidergr|kleidungsgr|konfektionsgr|clothing\s*size)/i;

export interface CheckInExtra {
  label: string;
  value: string;
  /** Hervorgehoben darstellen (Startnummer — die wird am Tisch vorgelesen). */
  strong?: boolean;
  /** v31.3: Kein Chip, sondern ein Satz in Warnfarbe — die Ausweichgröße am
   *  Ausgabetisch ist eine Handlungsanweisung, keine Angabe zum Nachschlagen. */
  tone?: 'warn' | 'danger';
}

export interface FieldDef {
  id: string;
  label: string;
  /**
   * v31.4.3: Die SharePoint-Spalte des Feldes. Sie ist NICHT dasselbe wie
   * `CustomData[id]` — s. `shirtAnswerOf`. Optional, weil sie fehlen kann
   * (Feld nachträglich angelegt, „Custom Fields prüfen" nie gelaufen).
   */
  spInternalName?: string;
  /**
   * v31.4.3: Weitere SP-Spaltennamen DESSELBEN Feldes auf anderen Ebenen.
   *
   * Dasselbe Abfragefeld (gleiche `id`) hat auf der Klammer und auf jedem
   * Termin eine eigene SharePoint-Spalte, und die Namen müssen nicht gleich
   * sein (SharePoint hängt bei Namenskollisionen eine Ziffer an). Wer wie die
   * Bestellliste Zeilen MEHRERER Ebenen in einem Topf zählt, kennt sonst je
   * Zeile die falsche Spalte und fällt still auf den alten `CustomData`-Stand
   * zurück. Aufrufer, die nur eine Ebene lesen (Check-in-Tisch,
   * Bearbeiten-Dialog), lassen das Feld weg.
   */
  spInternalNames?: string[];
}

/**
 * v31.4.3: Die Antwort EINER Person auf das Größenfeld — aus der Spalte, sonst
 * aus `CustomData`.
 *
 * Warum es diese Funktion überhaupt gibt (Live-Fall 08.09.2026, B2Run Köln,
 * Lauftag am nächsten Tag): Der Organizer korrigierte im Organizer Center zwei
 * Größen von „L"/„XL" auf „Herrengröße L"/„Herrengröße XL". Die
 * Teilnehmertabelle zeigte danach den neuen Wert, die Bestellliste weiter den
 * alten — und zwar dauerhaft, jedes Neuladen wieder.
 *
 * Ursache ist kein Zähl-, sondern ein SPEICHER-Unterschied: Eine Antwort steht
 * an ZWEI Stellen derselben Zeile — im JSON `CustomData` (Schlüssel = `f.id`)
 * und in der echten SP-Spalte (`f.spInternalName`). Die Anmeldung schreibt
 * beide (`registration.ts`), der Bearbeiten-Dialog des Organizer Centers bis
 * v31.4.3 nur die Spalte (`useEditModalHandlers.saveEdit` →
 * `adminUpdateRegistration`). Alles, was über `CustomData` liest, sah deshalb
 * ewig den Anmelde-Stand.
 *
 * Die Reihenfolge ist bewusst dieselbe wie in `ParticipantTable`: Spalte zuerst,
 * `CustomData` nur als Rückfall. Damit sagen Tabelle, Check-in-Tisch und
 * Bestellliste dasselbe — auch für Zeilen, die vor v31.4.3 auseinandergelaufen
 * sind. Sie brauchen kein erneutes Speichern.
 */
export function shirtAnswerOf(
  field: FieldDef | undefined | null,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  row: any,
): string {
  if (!field || !row) return '';
  const names = [field.spInternalName || ''].concat(field.spInternalNames || []);
  for (const sp of names) {
    if (!sp) continue;
    const v = row[sp];
    if (v === undefined || v === null) continue;
    const s = typeof v === 'boolean' ? (v ? 'true' : 'false') : String(v).trim();
    if (s) return s;
  }
  const raw = parseCustomData(row.CustomData)[field.id];
  return (raw === undefined || raw === null) ? '' : String(raw).trim();
}

/**
 * v31.3: Das Größenfeld eines Events — an EINER Stelle, damit Bestellliste,
 * Bestandspflege und Check-in-Tisch garantiert dasselbe Feld meinen.
 *
 * Warum das nicht einfach der erste Treffer ist: Seit das Muster auch das
 * blanke Wort „shirt" kennt (sonst fiele „Deloitte-Laufshirt" durch, das die
 * App selbst vorschlägt), trifft es auch Nachbarfelder wie „T-Shirt
 * Wunschfarbe". Welches davon gewinnt, entschied bis dahin allein die
 * Reihenfolge in `customFields` — die der Organizer per Drag frei ändert.
 * Deshalb entscheiden jetzt die ANTWORTEN: Wo Größen drinstehen, ist das
 * Größenfeld. Ohne Antworten (frisches Event) bleibt es beim ersten Treffer,
 * also beim bisherigen Verhalten.
 */
export function shirtFieldOf(
  fields: FieldDef[] | undefined | null,
  regs?: Array<{ CustomData?: string }> | null,
): FieldDef | undefined {
  const hits = (fields || []).filter(f => SHIRT_PATTERN.test(f.label || ''));
  if (hits.length <= 1 || !regs || !regs.length) return hits[0];
  let best = hits[0];
  let bestScore = -1;
  for (const f of hits) {
    let score = 0;
    for (const r of regs) {
      // v31.4.3: über `shirtAnswerOf` — sonst bewertet die Wahl den
      // Anmelde-Stand und nicht den korrigierten.
      const val = shirtAnswerOf(f, r);
      if (!val) continue;
      if (splitShirtSize(val).isSize) score++;
    }
    if (score > bestScore) { bestScore = score; best = f; }
  }
  return best;
}

/**
 * Baut die Zusatz-Zeile für eine Person.
 *
 * `reg` wird strukturell getypt, damit dieses Modul nicht vom EventService
 * abhängt. `customData` ist das bereits geparste `CustomData`-JSON der Zeile.
 *
 * **v31.4 — `bibNote`: der Zettel trägt noch den alten Namen.** Eine in DEX
 * frei gewordene Startnummer ist immer eine GEBRAUCHTE: Sie wurde beim
 * Veranstalter für jemand anderen gedruckt, und die Ummeldung dort ändert den
 * Aufdruck nicht. Am Ausgabetisch muss also jemand den Namen überkleben —
 * das ist eine Handlungsanweisung und kein Nachschlagewert, deshalb Warnton
 * statt Chip (Muster der Trikot-Ausweichgröße, v31.3).
 *
 * Der WORTLAUT kommt vom Aufrufer (wie `labels`): Dieses Modul kennt die
 * Sprache nicht. Die REGEL steht hier — angezeigt wird der Hinweis nur, wenn
 * die Zeile überhaupt eine Nummer trägt, und direkt unter ihr.
 */
export function checkInExtras(
  fields: FieldDef[] | undefined | null,
  customData: Record<string, unknown> | undefined | null,
  reg: { Startnummer?: string; StarterType?: string; PreferredStarterType?: string } | undefined | null,
  labels: { bib: string; group: string },
  bibNote?: { label: string; text: string } | null,
): CheckInExtra[] {
  const out: CheckInExtra[] = [];
  const bib = String((reg && reg.Startnummer) || '').trim();
  if (bib) out.push({ label: labels.bib, value: bib, strong: true });
  if (bib && bibNote && bibNote.text) {
    out.push({ label: bibNote.label, value: bibNote.text, tone: 'warn' });
  }

  const cd = customData || {};
  for (const f of (fields || [])) {
    if (!SHIRT_PATTERN.test(f.label || '')) continue;
    // v31.4.3: Erst die SP-Spalte der Zeile, dann das übergebene `CustomData`.
    // Eine im Organizer Center korrigierte Größe stand bis dahin NUR in der
    // Spalte — der Tisch las die alte (s. `shirtAnswerOf`). `reg` liegt hier
    // ohnehin vor; ohne Zeile bleibt es beim bisherigen Weg über `customData`.
    let v = shirtAnswerOf(f, reg);
    if (!v) {
      const raw = cd[f.id];
      if (raw === undefined || raw === null) continue;
      v = typeof raw === 'boolean' ? (raw ? 'Ja' : 'Nein') : String(raw).trim();
    }
    if (!v) continue;
    out.push({ label: f.label, value: v });
  }

  // Startblock/Gruppe: bei geteilten Kapazitäten steht dort Durchstarter bzw.
  // Funstarter — am Lauftag die zweite Frage nach der Startnummer.
  const grp = String((reg && (reg.StarterType || reg.PreferredStarterType)) || '').trim();
  if (grp) out.push({ label: labels.group, value: grp });

  return out;
}

/** Das geparste CustomData einer Zeile — defensiv, das Feld ist Freitext. */
export function parseCustomData(raw: string | undefined | null): Record<string, unknown> {
  try {
    const o = JSON.parse(raw || '{}');
    return (o && typeof o === 'object') ? o as Record<string, unknown> : {};
  } catch { return {}; }
}

/**
 * v30.60: Wie viele Trikots in welcher Größe? — Grundlage der Bestellung.
 *
 * Gezählt wird über dieselbe Feld-Erkennung wie am Check-in (`SHIRT_PATTERN`).
 * Das ist Absicht: Zwei Regeln für dieselbe Frage würden am Lauftag eine Größe
 * ausgeben, die nie bestellt wurde.
 *
 * Drei Dinge, die die Zahl ehrlich halten:
 *  - Nur ANGEMELDETE Personen zählen. Wartelisten- und abgemeldete Zeilen
 *    stehen weiter in der Liste; wer sie mitzählt, bestellt für Leute, die
 *    nicht kommen.
 *  - Wer keine Größe angegeben hat, verschwindet nicht, sondern erscheint als
 *    eigene Zeile „ohne Angabe" MIT Namen. Eine Bestellsumme, die stillschweigend
 *    kleiner ist als die Teilnehmerzahl, ist der teurere Fehler.
 *  - v31.3: Wer „T-Shirt bereits aus Vorjahren vorhanden" angeklickt hat,
 *    bekommt ebenfalls eine eigene Zeile — aber KEIN Shirt in der Bestellung
 *    (s. `splitShirtSize`).
 *
 * Die Größen werden zum Zählen normalisiert (Groß-/Kleinschreibung, Leerzeichen),
 * angezeigt wird die zuerst gesehene Schreibweise — „M" und „m" sind dieselbe
 * Bestellposition, aber die Liste soll aussehen wie die Eingabe.
 */
export interface ShirtTallyRow {
  /** Angezeigte Größe, oder '' für „ohne Angabe". */
  size: string;
  count: number;
  /** Namen — für die Nachfrage bei fehlender Angabe und zur Kontrolle. */
  names: string[];
  /**
   * v31.4: Dieselben Personen MIT E-Mail. Die Bestellliste springt von hier in
   * die Teilnehmerliste, und dafür braucht sie die Adresse — der Name ist
   * nicht eindeutig (CLAUDE.md: „Die E-Mail-Adresse ist der einzige
   * Schlüssel"). `names` bleibt unverändert, daran hängt der Excel-Export.
   */
  people: Array<{ name: string; email: string }>;
  /** v31.3: Der Wert ist keine Größe, sondern eine Abwahl („habe schon eins"). */
  optOut?: boolean;
}

export interface ShirtTallyResult {
  /** Das erkannte Feld; leer, wenn es an diesem Event gar keines gibt. */
  fieldLabel: string;
  rows: ShirtTallyRow[];
  /** Angemeldete Personen insgesamt (= Summe über alle Zeilen). */
  total: number;
  /** Davon ohne Angabe. */
  missing: number;
  /** v31.3: Davon mit einer Antwort, die keine Größe ist („habe schon eins"). */
  optOut: number;
  /** v31.3: Tatsächlich zu bestellende Shirts = total − missing − optOut. */
  sizeTotal: number;
}

const SHIRT_ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

/** Größe → Zählschlüssel (Groß-/Kleinschreibung, Leerzeichen egal). */
export function shirtSizeKey(size: string): string {
  return (size || '').toLowerCase().replace(/\s+/g, '');
}

/** Bekannte Konfektionsgrößen in natürlicher Reihenfolge (für Sortierung UND Nachbarschaft). */
export const SHIRT_SIZE_ORDER = ['xxs', 'xs', 's', 'm', 'l', 'xl', 'xxl', 'xxxl', '3xl', '4xl'];

// Längste Kürzel zuerst, damit „xxl" nicht als „l" mit Vorsatz „xx" gelesen wird.
const SIZE_ALTS = SHIRT_SIZE_ORDER.slice().sort((a, b) => b.length - a.length).join('|');
// Wert → Vorsatz + Kürzel. Der Trenner ist PFLICHT, sobald ein Vorsatz da ist:
// „Ich brauche keins" endet zwar auf „s", aber ohne Trenner davor — und genau
// solche Sätze dürfen nie als Größe S durchgehen.
const SIZE_TAIL_VALUE = new RegExp(`^(?:(.*?)[\\s\\-–—/_.]+)?(${SIZE_ALTS})$`, 'i');
// Schlüssel → Vorsatz + Kürzel. Schlüssel haben keine Leerzeichen mehr
// („herrengrößexl"), hier muss der Trenner deshalb entfallen. Angewandt wird
// das NUR auf Schlüssel, die aus einem geprüften Größenwert entstanden sind
// (bzw. auf Bestandszeilen, wo ein Fehlgriff nur die Sortierung beträfe).
const SIZE_TAIL_KEY = new RegExp(`^(.*?)(${SIZE_ALTS})$`, 'i');

export interface ShirtSizeParts {
  /** Ist der Wert überhaupt eine Größe? */
  isSize: boolean;
  /** Normalisierter Vorsatz, z.B. „herrengröße" — '' bei blankem „L". */
  prefix: string;
  /** Normalisiertes Kürzel aus SHIRT_SIZE_ORDER, z.B. „xl". */
  size: string;
}

/**
 * v31.3: Ist dieser Antwortwert eine Größe — und welche?
 *
 * Die Regel steht bewusst ANDERSHERUM als der erste Entwurf: Nicht eine Liste
 * bekannter Abwahl-Texte („Habe bereits ein Laufshirt", „Ohne T-Shirt", „I
 * already have one") entscheidet, sondern die Größe muss sich als solche
 * ausweisen. Grund ist der echte Wert aus dem B2Run-Köln-Event: „T-Shirt
 * bereits aus Vorjahren vorhanden" — vom Organizer selbst getippt, in keiner
 * Vorschlagsliste, von keiner Wortliste zu treffen. 30 von 90 Personen standen
 * damit als Bestellposition in der Liste.
 *
 * Eine Größe ist also: ein bekanntes Kürzel am ENDE des Werts, optional mit
 * Vorsatz („Herrengröße L", „Damen XS", „Unisex M"). Alles andere ist keine
 * Größe — es wird gezählt und benannt, aber nicht bestellt, verbraucht keinen
 * Bestand und bekommt keinen Gegenvorschlag. Bei „XS" kann die Regel nicht
 * danebengreifen, weil sie nur das Kürzel selbst prüft.
 */
export function splitShirtSize(value: string): ShirtSizeParts {
  const v = String(value || '').trim().replace(/\s+/g, ' ');
  const m = v ? SIZE_TAIL_VALUE.exec(v) : null;
  if (!m) return { isSize: false, prefix: '', size: '' };
  return { isSize: true, prefix: shirtSizeKey(m[1] || ''), size: m[2].toLowerCase() };
}

/** Zerlegt einen bereits normalisierten Schlüssel (s. SIZE_TAIL_KEY). */
function splitShirtKey(key: string): { prefix: string; size: string } {
  const k = String(key || '');
  const m = k ? SIZE_TAIL_KEY.exec(k) : null;
  if (!m) return { prefix: k, size: '' };
  return { prefix: (m[1] || '').toLowerCase(), size: m[2].toLowerCase() };
}

/** v31.3: Steht hinter diesem SCHLÜSSEL eine Größe? Der gespeicherte Bestand
 *  kennt nur Schlüssel ohne Leerzeichen („herrengrößexl"); `splitShirtSize`
 *  verlangt dort zu Recht einen Trenner und würde die Zeile verwerfen. */
export function isShirtSizeKey(key: string): boolean {
  return !!splitShirtKey(key).size;
}

/**
 * v31.3: Anzeigename eines Schlüssels, wenn niemand ihn geschrieben hat.
 *
 * Der Bestand kennt nur normalisierte Schlüssel („herrengrößexl") — ein
 * `toUpperCase()` daraus („HERRENGRÖSSEXL") ist am Ausgabetisch nicht lesbar.
 */
export function shirtSizeLabel(key: string): string {
  const k = String(key || '');
  if (!k) return '';
  const p = splitShirtKey(k);
  if (!p.size) return k.toUpperCase();
  if (!p.prefix) return p.size.toUpperCase();
  /**
   * v31.4 (Review): Der Trenner am Ende des Vorsatzes wird NICHT mehr
   * weggeworfen, sondern behalten — „herren-l" heißt „Herren-L", nicht
   * „Herren L".
   *
   * Grund ist der Rückweg: Seit v31.4 wandert ein Anzeigename wieder in einen
   * Zählschlüssel (die Auswahl am Tisch und im Bearbeiten-Dialog besteht aus
   * `alloc.rows[].size`, der gewählte Name landet in `ShirtIssued` und wird
   * hier mit `shirtSizeKey` zurückgerechnet). `shirtSizeKey` streicht nur
   * Leerzeichen; aus „Herren L" würde also „herrenl" statt „herren-l", und die
   * Ausgabe liefe auf einen zweiten, gleich aussehenden Schlüssel — der Karton
   * leert sich, die App zeigt ihn voll. Dieselbe Roundtrip-Falle wie bei
   * `{{Organizer}}` (v30.74): Wer einen Wert „backt", schreibt den Umkehrweg
   * gegen dieselbe Funktion. Für Vorsätze ohne Trenner („herrengröße") bleibt
   * es beim Leerzeichen — dort schließt `shirtSizeKey` den Kreis von selbst.
   */
  const pre = p.prefix.replace(/\s+$/, '');
  if (!pre) return p.size.toUpperCase();
  const hasSep = /[-–—/_.]$/.test(pre);
  return pre.charAt(0).toUpperCase() + pre.substring(1) + (hasSep ? '' : ' ') + p.size.toUpperCase();
}

/** Sortierung: erst nach Vorsatz (ohne Vorsatz zuerst), dann nach Größe. */
function compareShirtKeys(a: string, b: string): number {
  const pa = splitShirtKey(a);
  const pb = splitShirtKey(b);
  if (pa.prefix !== pb.prefix) return pa.prefix.localeCompare(pb.prefix, 'de');
  const ra = SHIRT_SIZE_ORDER.indexOf(pa.size);
  const rb = SHIRT_SIZE_ORDER.indexOf(pb.size);
  return ((ra < 0 ? 500 : ra) - (rb < 0 ? 500 : rb)) || a.localeCompare(b, 'de');
}

/**
 * v30.88: Ist-Bestand je Größe (Piggyback `_shirtStock` in EmailTemplateOverrides
 * des Hauptevents), Schlüssel = `shirtSizeKey`, Wert = Stückzahl. Leer, wenn
 * der Organizer nie einen Bestand eingetragen hat — dann gibt es keine
 * Gegenvorschläge, nur die Bestellliste wie bisher.
 */
export type ShirtStock = Record<string, number>;

export function parseShirtStock(overridesJson: string | undefined | null): ShirtStock {
  try {
    const o = JSON.parse(overridesJson || '{}');
    const raw = o && o._shirtStock;
    if (!raw || typeof raw !== 'object') return {};
    const out: ShirtStock = {};
    Object.keys(raw).forEach(k => {
      const n = Number(raw[k]);
      const key = shirtSizeKey(k);
      if (key && isFinite(n) && n >= 0) out[key] = Math.floor(n);
    });
    return out;
  } catch { return {}; }
}

/**
 * v31.4: Was eine Person WIRKLICH bekommen hat — Spalte `ShirtIssued` auf der
 * Teilnehmerzeile (Note, JSON):
 *
 *   { "size": "Herrengröße XL", "at": "2026-09-08T06:05:11Z", "by": "helfer@…" }
 *
 * Ein Objekt, keine Liste: Eine Person bekommt ein Shirt. `size` steht in
 * ANZEIGE-Schreibweise, so wie sie am Ausgabetisch gewählt wurde — der
 * Zählschlüssel entsteht daraus über `shirtSizeKey`.
 *
 * Warum es diese Spalte überhaupt gibt: `shirtAllocate` ist ein PLAN und
 * rechnet bei jedem Aufruf neu. Ohne festgehaltene Ausgabe schlägt sie
 * derselben Person am Nachmittag eine andere Größe vor als am Morgen, sobald
 * sich die Wunschlage ändert (jemand wird No-Show, seine Wünsche fallen aus
 * der Rechnung) — für ein Shirt, das die Person längst trägt.
 */
export interface ShirtIssue {
  /** Ausgegebene Größe in Anzeige-Schreibweise. */
  size: string;
  /** Zeitpunkt der Ausgabe (ISO). */
  at: string;
  /** Wer ausgegeben hat (Helfer-Kennung, meist die E-Mail). */
  by: string;
}

/**
 * v31.4 (Nachtrag): Schlüssel der ausgegebenen Größe im Bearbeiten-Formular
 * des Organizer Centers (`editForm`).
 *
 * Das Formular ist sonst „SP-Spaltenname → Wert"; die Ausgabe ist aber KEIN
 * Feld des Zeilen-Patches, sondern ein eigener Schreibvorgang
 * (`setShirtIssued`). Der doppelte Unterstrich hält den Schlüssel deshalb
 * garantiert aus der Custom-Field-Schleife heraus — und die Konstante hält
 * Dialog und Save-Pfad auf demselben Namen.
 */
export const SHIRT_ISSUED_FORM_KEY = '__shirtIssued';

/** Defensiv wie `parseShirtStock`: kaputtes JSON oder keine Größe → null. */
export function parseShirtIssue(raw: string | undefined | null): ShirtIssue | null {
  const s = (raw || '').trim();
  // Der Normalfall ist die leere Spalte — der darf keine Ausnahme werfen, die
  // Funktion läuft je Zeile und Aufruf durch die ganze Teilnehmerliste.
  if (!s) return null;
  try {
    const o = JSON.parse(s);
    if (!o || typeof o !== 'object' || Array.isArray(o)) return null;
    const size = typeof o.size === 'string' ? o.size.trim() : '';
    if (!size) return null;
    return { size, at: typeof o.at === 'string' ? o.at : '', by: typeof o.by === 'string' ? o.by : '' };
  } catch { return null; }
}

export interface ShirtAllocation {
  /** Gewünschte Größe (Anzeige-Schreibweise). */
  wish: string;
  /** Gegenvorschlag, wenn die Wunschgröße nicht mehr reicht; null = auch keine Nachbargröße mehr da. */
  proposal: string | null;
  /** v31.3: Schlüssel des Gegenvorschlags — für den Rest-Bestand am Tisch. */
  proposalKey?: string;
  /** true = Wunschgröße reicht NICHT für diese Person. */
  short: boolean;
  /** v31.3: Die Antwort ist gar keine Größe („habe schon eins") — kein Shirt nötig. */
  optOut?: boolean;
  /**
   * v31.4: Diese Person hat ihr Shirt bereits bekommen — Anzeige-Schreibweise
   * der AUSGEGEBENEN Größe. Ist das gesetzt, gibt es keinen Wunsch-Abgleich
   * und keinen Gegenvorschlag mehr (`short` ist immer false): Die Tatsache
   * schlägt den Plan.
   */
  issued?: string;
  /**
   * v31.4 (Nachtrag): Diese Ausgabe steht NICHT in der Spalte — sie ist
   * angenommen, weil die Person eingecheckt ist (s. `shirtAllocate`). Der
   * Wert in `issued` ist dann die Wunschgröße, nicht ein festgehaltener
   * Eintrag. Wer die Zahl anzeigt, muss den Unterschied benennen: Eine
   * Annahme ist eine Aussage über den Lauftag, keine über die Daten.
   */
  issuedAssumed?: boolean;
}

export interface ShirtAllocationRow {
  size: string;
  key: string;
  /** Wünsche (angemeldete Personen mit dieser Größe). */
  need: number;
  /** Eingetragener Bestand. */
  stock: number;
  /** Personen, die diese Größe wollten und sie NICHT bekommen. */
  missing: number;
  /** Nach der Verteilung noch übrig (inkl. an Ausweichende vergebener Stücke). */
  spare: number;
  /** v31.4: Wie viele Stück dieser Größe sind schon ausgegeben? Sie sind aus
   *  dem Karton — „Bestand" ist damit nicht mehr dasselbe wie „verfügbar". */
  issued: number;
  /** v31.4 (Nachtrag): Teilmenge von `issued` — davon sind angenommen, weil
   *  die Person eingecheckt ist, ohne dass jemand die Ausgabe festgehalten
   *  hat. Wer „ausgegeben" anzeigt, nennt diese Zahl dazu. */
  issuedAssumed: number;
  /**
   * v31.4 (Review): Wie viele Stück MEHR ausgegeben wurden, als im Bestand
   * stehen. `spare` klemmt bei 0 (eine negative Kartonzahl ist keine
   * Aussage), und der Ausgabe-Vorlauf bricht die betroffenen Personen vor
   * `missing` ab — ohne diese Zahl verschwände die Über-Ausgabe restlos:
   * „fehlt 5" am Morgen kippt am Lauftag auf „reicht", genau in dem Moment,
   * in dem der Mangel real wird. Wer „reicht es?" beantwortet, muss `over`
   * wie `missing` behandeln.
   */
  over: number;
}

export interface ShirtAllocationResult {
  /** Gibt es überhaupt einen Bestand? Ohne Bestand keine Aussage. */
  hasStock: boolean;
  byEmail: Record<string, ShirtAllocation>;
  rows: ShirtAllocationRow[];
  /** Personen mit Wunsch, aber ohne jede passende Größe. */
  noneLeft: string[];
  /** v31.3: Personen, deren Antwort keine Größe ist — sie brauchen keins. */
  optOut: number;
}

/**
 * v30.88: Verteilt den Ist-Bestand auf die Wünsche und macht Gegenvorschläge.
 *
 * Nutzer-Ansage (07.09.2026, B2Run Köln): „neben der Anzahl an Shirts auch
 * ermöglichen, dass man angibt, wie viele Shirts man wirklich hat, und dann
 * wird geguckt, ob es überhaupt passt — und bei jeder Person, wo es nicht
 * mehr passt, ein Gegenvorschlag." Gebraucht am Check-in bzw. bei der Abholung.
 *
 * Regeln — bewusst einfach, damit Tisch und Bestellliste dasselbe sagen:
 *  - Reihenfolge = Teilnehmer-ID aufsteigend (wer zuerst angemeldet war,
 *    bekommt seine Größe). Dieselbe Reihenfolge wie die Warteliste nutzt.
 *  - Reicht die Wunschgröße nicht, kommt die NÄCHSTE Größe mit Restbestand:
 *    erst eine größer, dann eine kleiner, dann zwei größer, zwei kleiner …
 *    Ein Trikot eine Nummer zu groß trägt man; zu klein passt nicht.
 *  - Der Vorschlag verbraucht den Bestand der Ausweichgröße — zwei Personen
 *    bekommen nicht dasselbe letzte L.
 *  - v31.3: Die Nachbarn werden INNERHALB desselben Vorsatzes gesucht.
 *    Seit v30.95 kann das Größenfeld Kategorien haben („Herrengröße"/
 *    „Damengröße"); die Anmeldeseite speichert dann „Herrengröße L". Der
 *    Schlüssel „herrengrößel" steht in keiner SHIRT_SIZE_ORDER — es gab
 *    deshalb NIE einen Gegenvorschlag, die Person hörte „nichts mehr da",
 *    obwohl Herrengröße XL im Karton lag. Der Bestand bleibt je vollständiger
 *    Bezeichnung gepflegt; gesucht wird „Herrengröße XL", nie „Damengröße XL".
 *  - v31.3: Antworten, die keine Größe sind, verbrauchen nichts und bekommen
 *    keinen Vorschlag (s. `splitShirtSize`).
 *  - v31.4: Was schon AUSGEGEBEN ist (`ShirtIssued`), wird zuerst abgezogen
 *    und schlägt jeden Vorschlag (s. den Block unten).
 *  - v31.4 (Nachtrag): Wer `Eingecheckt` ist und keinen Eintrag hat, gilt mit
 *    seiner Wunschgröße als versorgt — als ANNAHME (`issuedAssumed`), nie in
 *    der Spalte (s. den zweiten Vorlauf unten).
 */
export function shirtAllocate(
  fields: FieldDef[] | undefined | null,
  regs: Array<{ Status?: string; CustomData?: string; ParticipantName?: string; ParticipantEmail?: string; TeilnehmerID?: number | string | null; Id?: number; ShirtIssued?: string }> | undefined | null,
  stock: ShirtStock,
  /**
   * v31.4 (Nachtrag): E-Mails (kleingeschrieben), die IRGENDWO eingecheckt
   * sind — für Aufrufer, die je Person nur EINE Zeile übergeben.
   *
   * Bei einem Klammer-Event steht die Größenfrage oft auf der Klammer-Zeile,
   * eingecheckt wird aber auf der Termin-Zeile (dieselbe Trennung, die
   * `ShirtSizeModal` schon für `ShirtIssued` überbrücken musste). Wer die
   * Zeilen vorher zusammenführt, hat den Check-in-Status der verworfenen
   * Zeile nicht mehr — und die Annahme unten liefe für das halbe Event ins
   * Leere. Aufrufer mit vollständiger Zeilenliste (Check-in-Seite) lassen den
   * Parameter weg.
   */
  checkedInEmails?: Record<string, true> | null,
): ShirtAllocationResult {
  const result: ShirtAllocationResult = { hasStock: Object.keys(stock || {}).length > 0, byEmail: {}, rows: [], noneLeft: [], optOut: 0 };
  // v31.3: Die Zeilen mitgeben — bei mehreren treffenden Feldern entscheiden
  // die Antworten, nicht die Feld-Reihenfolge (s. `shirtFieldOf`).
  const field = shirtFieldOf(fields, regs);
  if (!field) return result;
  const remaining: ShirtStock = {};
  Object.keys(stock || {}).forEach(k => { remaining[k] = stock[k]; });
  const display: Record<string, string> = {};
  const parts: Record<string, ShirtSizeParts> = {};
  const need: Record<string, number> = {};
  const missing: Record<string, number> = {};
  const idOf = (r: { TeilnehmerID?: number | string | null; Id?: number }): number => {
    const n = Number(r.TeilnehmerID);
    return isFinite(n) && n > 0 ? n : 1e9 + (r.Id || 0);
  };
  const active = (regs || [])
    .filter(r => SHIRT_ACTIVE_STATI.indexOf(r.Status || '') >= 0)
    .slice()
    .sort((a, b) => idOf(a) - idOf(b));
  const neighbours = (key: string): string[] => {
    const p = parts[key] || splitShirtKey(key);
    const i = SHIRT_SIZE_ORDER.indexOf(p.size);
    if (i < 0) return [];
    const out: string[] = [];
    for (let d = 1; d < SHIRT_SIZE_ORDER.length; d++) {
      if (i + d < SHIRT_SIZE_ORDER.length) out.push(p.prefix + SHIRT_SIZE_ORDER[i + d]);
      if (i - d >= 0) out.push(p.prefix + SHIRT_SIZE_ORDER[i - d]);
    }
    return out;
  };
  const labelOf = (key: string): string => display[key] || shirtSizeLabel(key);
  /**
   * v31.4: ERST die Ausgaben abziehen, dann die Wünsche verteilen.
   *
   * Zwei Entscheidungen stecken hier drin, beide bewusst:
   *
   *  1. **Der Status ist egal.** Gezählt wird über ALLE Zeilen, nicht nur über
   *     `SHIRT_ACTIVE_STATI`. Ein ausgegebenes Shirt ist aus dem Karton —
   *     wer sich abends abmeldet oder als No-Show markiert wird, legt es nicht
   *     zurück. Würden die Ausgaben mit dem Status verschwinden, wüchse der
   *     rechnerische Bestand über Nacht wieder an, und die App verspräche
   *     Shirts, die es nicht mehr gibt.
   *  2. **Vor der Verteilung.** Sonst verplant die Wunsch-Runde Stücke, die
   *     längst weg sind: Person 41 nimmt morgens ein XL mit, nachmittags
   *     rechnet die App das XL erneut jemand anderem zu.
   */
  const issuedOf = new Map<object, ShirtIssue>();
  /**
   * v31.4 (Nachzug): zusätzlich je E-Mail. Die Ausgabe steht auf der Zeile, an
   * der das Team eingecheckt hat (meist der Termin), der Wunsch oft auf der
   * Klammer — beides ist DIESELBE Person. Nur über die Zeilen-Identität zu
   * gehen hieße: kommt die Wunsch-Zeile in `active` nach der Ausgabe-Zeile,
   * überschreibt sie den Eintrag wieder mit einem Gegenvorschlag. Die
   * Reihenfolge der Liste darf über so etwas nicht entscheiden.
   */
  const issuedByEmail: Record<string, ShirtIssue> = {};
  const issuedCount: Record<string, number> = {};
  const issuedAssumedCount: Record<string, number> = {};
  const issuedSeen: Record<string, true> = {};
  /** Zeilen bzw. E-Mails, deren „Ausgabe" nur angenommen ist (s. unten). */
  const assumedOf = new Map<object, true>();
  const assumedEmail: Record<string, true> = {};
  for (const r of (regs || [])) {
    const iss = parseShirtIssue(r.ShirtIssued);
    if (!iss) continue;
    const em = (r.ParticipantEmail || '').toLowerCase().trim();
    // Dieselbe Person kann auf zwei Zeilen stehen (Klammer + Termin) — sie hat
    // trotzdem EIN Shirt bekommen und darf den Bestand nicht zweimal belasten.
    //
    // v31.4 (Review): Die Grenze gehört dazu, statt Eindeutigkeit zu
    // behaupten. Die E-Mail ist der einzige Schlüssel, den es gibt, aber sie
    // ist NICHT eindeutig (CLAUDE.md: SMTP-Adresse vs. UPN/Alias) — steht
    // dieselbe Person unter zwei Schreibweisen, greift dieser Wächter nicht
    // und sie wird zweimal vom Bestand abgezogen. Dasselbe bei einer Zeile
    // ohne Adresse: `em` ist '', der Wächter ist wirkungslos, und
    // `pickShirtAnswerRows` lässt solche Zeilen bewusst einzeln stehen. Beides
    // ist hier nicht heilbar (die zweite Adresse ist real) — wer eine Zahl aus
    // dieser Rechnung meldet, muss die Doppelung kennen.
    if (em && issuedSeen[em]) continue;
    if (em) { issuedSeen[em] = true; issuedByEmail[em] = iss; }
    issuedOf.set(r, iss);
    const ik = shirtSizeKey(iss.size);
    if (!display[ik]) display[ik] = iss.size;
    issuedCount[ik] = (issuedCount[ik] || 0) + 1;
    remaining[ik] = (remaining[ik] || 0) - 1;
  }
  /**
   * v31.4 (Nachtrag): Wer eingecheckt ist, hat sein Trikot bekommen — auch
   * ohne Eintrag in der Spalte.
   *
   * Der Ausgabe-Knopf am Check-in-Tisch gibt es erst seit v31.4; die Leute,
   * die vorher durch den Tisch gegangen sind, tragen ihr Shirt trotzdem. Ohne
   * diese Annahme rechnet die App einen Karton voll, der real halb leer ist —
   * Ansage des Organizers, der am Tisch stand (08.09.2026, B2Run Köln).
   *
   * Drei Grenzen, die die Annahme ehrlich halten:
   *  1. **Ein echter Eintrag schlägt sie immer.** Der zweite Durchlauf läuft
   *     NACH dem ersten und überspringt jede E-Mail, die dort schon gezählt
   *     wurde — die Annahme füllt nur Lücken.
   *  2. **Sie wird nie geschrieben.** Sie lebt in dieser Rechnung, nicht in
   *     der Spalte; sobald jemand am Tisch eine Größe festhält, gilt die.
   *  3. **`No-Show` zählt nicht.** Die Person war nicht da, also hat sie auch
   *     nichts mitgenommen. Nur `Eingecheckt` heißt „stand am Tisch" — auf
   *     dieser Zeile oder (bei zusammengeführten Zeilen) auf einer anderen
   *     Zeile derselben Person (`checkedInEmails`).
   *
   * Ohne echte Wunschgröße gibt es nichts anzunehmen: Wer „ich habe schon
   * eins" angekreuzt oder gar nichts geantwortet hat, hat auch keins bekommen.
   */
  for (const r of (regs || [])) {
    if (parseShirtIssue(r.ShirtIssued)) continue;
    const em = (r.ParticipantEmail || '').toLowerCase().trim();
    if ((r.Status || '') !== 'Eingecheckt' && !(em && checkedInEmails && checkedInEmails[em])) continue;
    if (em && issuedSeen[em]) continue;
    // v31.4.3: Spalte vor `CustomData` (s. `shirtAnswerOf`).
    const wish = shirtAnswerOf(field, r);
    if (!wish || !splitShirtSize(wish).isSize) continue;
    // `at`/`by` bleiben leer — das ist der Unterschied zu einer festgehaltenen
    // Ausgabe und macht sie auch im Datenobjekt erkennbar.
    const iss: ShirtIssue = { size: wish, at: '', by: '' };
    if (em) { issuedSeen[em] = true; issuedByEmail[em] = iss; assumedEmail[em] = true; }
    issuedOf.set(r, iss);
    assumedOf.set(r, true);
    const ik = shirtSizeKey(wish);
    if (!display[ik]) display[ik] = wish;
    issuedCount[ik] = (issuedCount[ik] || 0) + 1;
    issuedAssumedCount[ik] = (issuedAssumedCount[ik] || 0) + 1;
    remaining[ik] = (remaining[ik] || 0) - 1;
  }
  for (const r of active) {
    // v31.4.3: Spalte vor `CustomData` (s. `shirtAnswerOf`).
    const wish = shirtAnswerOf(field, r);
    const email = (r.ParticipantEmail || '').toLowerCase().trim();
    const name = (r.ParticipantName || r.ParticipantEmail || '—').trim();
    const issued = (email && issuedByEmail[email]) || issuedOf.get(r);
    if (issued) {
      // v31.4: Für diese Person gibt es nichts mehr zu rechnen — kein
      // Wunsch-Abgleich, kein `short`, kein Gegenvorschlag. Der Wunsch wird
      // trotzdem gezählt: `need` ist die Wunsch-Spalte (was bestellt wurde),
      // nicht die Ausgabe-Spalte.
      // v31.4 (Nachtrag): `issuedAssumed` sagt, ob das eine festgehaltene
      // Ausgabe war oder die Annahme aus dem Check-in. Aufgelöst wird sie auf
      // demselben Weg wie `issued` — erst über die E-Mail, sonst über die
      // Zeile; sonst behauptete die Anzeige „festgehalten", wo nichts steht.
      const assumed = (email && issuedByEmail[email]) ? !!assumedEmail[email] : !!assumedOf.get(r);
      if (email) result.byEmail[email] = { wish, proposal: null, short: false, issued: issued.size, issuedAssumed: assumed || undefined };
      const wp = wish ? splitShirtSize(wish) : null;
      if (wp && wp.isSize) {
        const wk = shirtSizeKey(wish);
        parts[wk] = wp;
        if (!display[wk]) display[wk] = wish;
        need[wk] = (need[wk] || 0) + 1;
      }
      continue;
    }
    if (!wish) continue;
    const p = splitShirtSize(wish);
    if (!p.isSize) {
      // v31.3: Abwahl statt Größe — kein Bestandsverbrauch, kein
      // Gegenvorschlag, nicht in `noneLeft`. Sonst schlägt die App am Tisch
      // ausgerechnet den Leuten eine Ausweichgröße vor, die gar keins wollen.
      result.optOut++;
      if (email) result.byEmail[email] = { wish, proposal: null, short: false, optOut: true };
      continue;
    }
    const key = shirtSizeKey(wish);
    parts[key] = p;
    if (!display[key]) display[key] = wish;
    need[key] = (need[key] || 0) + 1;
    if (!result.hasStock) continue;
    if ((remaining[key] || 0) > 0) {
      remaining[key]--;
      if (email) result.byEmail[email] = { wish, proposal: null, short: false };
      continue;
    }
    missing[key] = (missing[key] || 0) + 1;
    let proposal: string | null = null;
    let proposalKey = '';
    for (const nk of neighbours(key)) {
      if ((remaining[nk] || 0) > 0) { remaining[nk]--; proposalKey = nk; proposal = labelOf(nk); break; }
    }
    if (!proposal) result.noneLeft.push(name);
    if (email) result.byEmail[email] = { wish, proposal, proposalKey: proposalKey || undefined, short: true };
  }
  // v31.4: Auch Größen, die NUR ausgegeben wurden (Reserve aus einem anderen
  // Karton, niemand hat sie gewünscht, kein Bestand eingetragen), brauchen
  // ihre Zeile — sonst verschwindet die Ausgabe aus der Liste.
  const keys = Array.from(new Set(Object.keys(need).concat(Object.keys(stock || {})).concat(Object.keys(issuedCount))));
  result.rows = keys
    .sort(compareShirtKeys)
    .map(k => ({
      size: labelOf(k),
      key: k,
      need: need[k] || 0,
      stock: (stock && stock[k]) || 0,
      missing: missing[k] || 0,
      // v31.4: Nie negativ. Mehr ausgegeben als eingetragen heißt „nichts mehr
      // da" — eine Zahl unter null wäre keine Aussage über den Karton.
      spare: result.hasStock ? Math.max(0, remaining[k] || 0) : 0,
      // v31.4 (Review): der Gegenwert zu `spare` — was unter null steht, ist
      // die Über-Ausgabe. Nur mit Bestand aussagekräftig; ohne eingetragenen
      // Bestand ist „mehr als da" keine Aussage.
      over: result.hasStock ? Math.max(0, -(remaining[k] || 0)) : 0,
      issued: issuedCount[k] || 0,
      issuedAssumed: issuedAssumedCount[k] || 0,
    }));
  return result;
}

/** v31.3: Zwei Zeilen derselben Person mit zwei verschiedenen Größen. */
export interface ShirtAnswerConflict {
  name: string;
  /** Die widersprüchlichen Werte; der erste ist der, mit dem gerechnet wird. */
  values: string[];
}

export interface ShirtRowPick<T> {
  rows: T[];
  conflicts: ShirtAnswerConflict[];
  /**
   * v31.4: Die NICHT gewählten Zeilen je E-Mail (lowercase).
   *
   * Der Live-Fall vom 08.09.2026: Eine korrigierte Größe kam in der
   * Bestellliste nicht an, und der Konflikt-Kasten schwieg — weil die zweite
   * Zeile GAR KEINE Antwort trug und damit kein Widerspruch war. Sichtbar war
   * nur „L", nicht, aus welcher Zeile das „L" stammt. Wer die Herkunft
   * anzeigen will, braucht deshalb auch die verworfenen Zeilen.
   */
  othersByEmail: Record<string, T[]>;
}

/** Strukturelle Sicht auf eine Anmeldezeile — dieses Modul kennt den
 *  EventService nicht. `Modified` kommt über `$select=*` immer mit, steht aber
 *  (noch) nicht im `SPRegistration`-Typ; deshalb hier optional. */
export interface ShirtRowLike {
  Status?: string;
  CustomData?: string;
  ParticipantName?: string;
  ParticipantEmail?: string;
  Modified?: string;
}

/**
 * v31.3: Welche Zeile einer Person beantwortet die Größenfrage?
 *
 * Der Fall aus dem Live-Event: Der Organizer korrigierte zwei Antworten von
 * „L"/„XL" auf „Herrengröße L"/„Herrengröße XL" — die Bestellliste zeigte
 * unverändert die alten Werte. Ursache war nicht das Speichern, sondern die
 * Zusammenführung: `ShirtSizeModal` lief über `[Klammer, ...Termine]` und
 * behielt je E-Mail die ZUERST gelesene Zeile. Wer mehrere Zeilen hat
 * (Klammer-Schattenzeile plus Termin-Zeile), bei dem gewann damit die Ebene,
 * die zufällig vorne stand — und keine noch so oft wiederholte Korrektur auf
 * der anderen Ebene kam je in der Liste an.
 *
 * Die Reihenfolge der Kriterien (CLAUDE.md: „Antworten stehen dort, wo
 * angemeldet wurde"; Vorbild `HotelPlanningPanel.answerRowsOf`):
 *  1. Eine Zeile MIT Antwort schlägt eine Zeile ohne — das ist der
 *     Schattenzeilen-Fall (die Klammer hat oft gar kein CustomData).
 *  2. Aktiver Status schlägt abgemeldet/No-Show.
 *  3. Bei gleichem Stand gewinnt die zuletzt geänderte Zeile (`Modified`).
 *     Fehlt der Zeitstempel, bleibt es bei der bisherigen Reihenfolge —
 *     kein NaN-Vergleich, kein stiller Sprung.
 *
 * Zeilen ohne E-Mail lassen sich nicht zusammenführen; sie bleiben einzeln
 * stehen (das ist ehrlicher, als sie wegzulassen).
 */
export function pickShirtAnswerRows<T extends ShirtRowLike>(
  fields: FieldDef[] | undefined | null,
  regs: T[] | undefined | null,
): ShirtRowPick<T> {
  const list = regs || [];
  const out: ShirtRowPick<T> = { rows: [], conflicts: [], othersByEmail: {} };
  const field = shirtFieldOf(fields, list);
  const answers: string[] = [];
  const scores: number[] = [];
  const times: number[] = [];
  list.forEach((r, i) => {
    // v31.4.3: Spalte vor `CustomData` (s. `shirtAnswerOf`) — sonst gewinnt
    // hier eine Zeile „mit Antwort", deren Antwort der Organizer längst
    // überschrieben hat.
    const val = field ? shirtAnswerOf(field, r) : '';
    answers[i] = val;
    scores[i] = (val ? 2 : 0) + (SHIRT_ACTIVE_STATI.indexOf(r.Status || '') >= 0 ? 1 : 0);
    const t = Date.parse(String(r.Modified || ''));
    times[i] = isFinite(t) ? t : NaN;
  });
  // Die E-Mail einmal je Zeile normalisieren — bei 5000 Zeilen ist das der
  // Unterschied zwischen einmal durchlaufen und quadratisch suchen.
  const mails: string[] = list.map(r => (r.ParticipantEmail || '').toLowerCase().trim());
  const winner: Record<string, number> = {};
  const emailOrder: string[] = [];
  const rowsByEmail: Record<string, number[]> = {};
  mails.forEach((em, i) => {
    if (!em) return;
    (rowsByEmail[em] = rowsByEmail[em] || []).push(i);
    const cur = winner[em];
    if (cur === undefined) { winner[em] = i; emailOrder.push(em); return; }
    if (scores[i] > scores[cur]) { winner[em] = i; return; }
    if (scores[i] < scores[cur]) return;
    if (isFinite(times[i]) && isFinite(times[cur]) && times[i] > times[cur]) winner[em] = i;
  });
  out.rows = list.filter((r, i) => !mails[i] || winner[mails[i]] === i);
  // Widerspruch benennen statt still zu entscheiden: Zwei verschiedene Größen
  // auf zwei Zeilen heißen meistens, dass nur eine von beiden bearbeitet wurde.
  emailOrder.forEach(em => {
    // v31.4: Verworfene Zeilen mitgeben — auch die ohne Antwort (die lösen
    // keinen Konflikt aus und blieben deshalb bisher unsichtbar).
    const losers = (rowsByEmail[em] || []).filter(i => i !== winner[em]).map(i => list[i]);
    if (losers.length > 0) out.othersByEmail[em] = losers;
    const idxs = (rowsByEmail[em] || []).filter(i => !!answers[i]);
    const seen: string[] = [];
    idxs.forEach(i => { if (seen.some(v => v.toLowerCase() === answers[i].toLowerCase())) return; seen.push(answers[i]); });
    if (seen.length < 2) return;
    const win = winner[em];
    const name = (list[win].ParticipantName || list[win].ParticipantEmail || em).trim();
    const winVal = answers[win];
    out.conflicts.push({ name, values: [winVal].concat(seen.filter(v => v.toLowerCase() !== (winVal || '').toLowerCase())) });
  });
  return out;
}

export function shirtTally(
  fields: FieldDef[] | undefined | null,
  regs: Array<{ Status?: string; CustomData?: string; ParticipantName?: string; ParticipantEmail?: string }> | undefined | null
): ShirtTallyResult {
  const field = shirtFieldOf(fields, regs);
  const out: ShirtTallyResult = { fieldLabel: field ? field.label : '', rows: [], total: 0, missing: 0, optOut: 0, sizeTotal: 0 };
  if (!field) return out;
  const byKey: Record<string, ShirtTallyRow> = {};
  const order: string[] = [];
  for (const r of (regs || [])) {
    if (SHIRT_ACTIVE_STATI.indexOf(r.Status || '') < 0) continue;
    out.total++;
    // v31.4.3: Spalte vor `CustomData` (s. `shirtAnswerOf`) — das war der
    // Grund, warum eine korrigierte Größe hier nie ankam.
    const val = shirtAnswerOf(field, r);
    const key = shirtSizeKey(val);
    const optOut = !!val && !splitShirtSize(val).isSize;
    if (!byKey[key]) { byKey[key] = { size: val, count: 0, names: [], people: [], optOut }; order.push(key); }
    byKey[key].count++;
    const pname = (r.ParticipantName || r.ParticipantEmail || '—').trim();
    byKey[key].names.push(pname);
    // v31.4: dieselbe Person nochmal mit Adresse — für den Sprung in die
    // Teilnehmerliste und die Herkunfts-Zeile darunter.
    byKey[key].people.push({ name: pname, email: (r.ParticipantEmail || '').toLowerCase().trim() });
    if (!val) out.missing++;
    else if (optOut) out.optOut++;
    else out.sizeTotal++;
  }
  // Bekannte Konfektionsgrößen in ihrer natürlichen Reihenfolge (v31.3: je
  // Vorsatz gruppiert, damit „Damengröße S/M/L" zusammen steht), danach die
  // Abwahl-Zeilen, „ohne Angabe" ganz zuletzt — eine Bestellliste, die mit
  // XXL beginnt, liest niemand gern.
  const bucket = (k: string): number => {
    if (!k) return 2;
    return byKey[k] && byKey[k].optOut ? 1 : 0;
  };
  out.rows = order
    .slice()
    .sort((a, b) => (bucket(a) - bucket(b)) || (bucket(a) === 0 ? compareShirtKeys(a, b) : a.localeCompare(b, 'de')))
    .map(k => byKey[k]);
  return out;
}
