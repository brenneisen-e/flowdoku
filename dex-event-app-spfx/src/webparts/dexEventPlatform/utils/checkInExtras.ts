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

export interface FieldDef { id: string; label: string }

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
      const raw = parseCustomData(r.CustomData)[f.id];
      if (raw === undefined || raw === null) continue;
      if (splitShirtSize(String(raw)).isSize) score++;
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
 */
export function checkInExtras(
  fields: FieldDef[] | undefined | null,
  customData: Record<string, unknown> | undefined | null,
  reg: { Startnummer?: string; StarterType?: string; PreferredStarterType?: string } | undefined | null,
  labels: { bib: string; group: string }
): CheckInExtra[] {
  const out: CheckInExtra[] = [];
  const bib = String((reg && reg.Startnummer) || '').trim();
  if (bib) out.push({ label: labels.bib, value: bib, strong: true });

  const cd = customData || {};
  for (const f of (fields || [])) {
    if (!SHIRT_PATTERN.test(f.label || '')) continue;
    const raw = cd[f.id];
    if (raw === undefined || raw === null) continue;
    const v = typeof raw === 'boolean' ? (raw ? 'Ja' : 'Nein') : String(raw).trim();
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
  // Trenner am Ende des Vorsatzes wegnehmen: „herren-l" soll „Herren L"
  // heißen, nicht „Herren- L".
  const pre = p.prefix.replace(/[\s\-–—/_.]+$/, '');
  if (!pre) return p.size.toUpperCase();
  return pre.charAt(0).toUpperCase() + pre.substring(1) + ' ' + p.size.toUpperCase();
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
 */
export function shirtAllocate(
  fields: FieldDef[] | undefined | null,
  regs: Array<{ Status?: string; CustomData?: string; ParticipantName?: string; ParticipantEmail?: string; TeilnehmerID?: number | string | null; Id?: number }> | undefined | null,
  stock: ShirtStock,
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
  for (const r of active) {
    const cd = parseCustomData(r.CustomData);
    const raw = cd[field.id];
    const wish = (raw === undefined || raw === null) ? '' : String(raw).trim();
    if (!wish) continue;
    const email = (r.ParticipantEmail || '').toLowerCase().trim();
    const name = (r.ParticipantName || r.ParticipantEmail || '—').trim();
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
  const keys = Array.from(new Set(Object.keys(need).concat(Object.keys(stock || {}))));
  result.rows = keys
    .sort(compareShirtKeys)
    .map(k => ({
      size: labelOf(k),
      key: k,
      need: need[k] || 0,
      stock: (stock && stock[k]) || 0,
      missing: missing[k] || 0,
      spare: result.hasStock ? (remaining[k] || 0) : 0,
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
  const out: ShirtRowPick<T> = { rows: [], conflicts: [] };
  const field = shirtFieldOf(fields, list);
  const answers: string[] = [];
  const scores: number[] = [];
  const times: number[] = [];
  list.forEach((r, i) => {
    let val = '';
    if (field) {
      const raw = parseCustomData(r.CustomData)[field.id];
      val = (raw === undefined || raw === null) ? '' : String(raw).trim();
    }
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
    const cd = parseCustomData(r.CustomData);
    const raw = cd[field.id];
    const val = (raw === undefined || raw === null) ? '' : String(raw).trim();
    const key = shirtSizeKey(val);
    const optOut = !!val && !splitShirtSize(val).isSize;
    if (!byKey[key]) { byKey[key] = { size: val, count: 0, names: [], optOut }; order.push(key); }
    byKey[key].count++;
    byKey[key].names.push((r.ParticipantName || r.ParticipantEmail || '—').trim());
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
