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
 * **Was das kostet, offen gesagt:** Ein Muster trifft irgendwann etwas
 * Falsches — ein Feld „Größe des Gepäckstücks" landete hier ebenfalls. Das ist
 * am Check-in-Tisch harmlos (eine Zeile zu viel), aber es ist geraten und
 * nicht gesagt. Der saubere Weg wäre ein Haken „am Check-in anzeigen" am
 * Abfragefeld selbst; dann kommt diese Datei weg. Bis dahin steht die Regel
 * hier an EINER Stelle statt verteilt im Render-Baum.
 */

/** Feld-Label → gehört ans Check-in? Aktuell: alles, was nach Trikot/Größe klingt.
 *  v30.60 exportiert: Die Auswertung „benötigte T-Shirts" muss DIESELBE Regel
 *  benutzen, sonst zählt die Bestellung andere Felder als der Check-in-Tisch. */
export const SHIRT_PATTERN = /(t-?\s?shirt|trikot|shirt\s*size|kleidergr|konfektionsgr)/i;

export interface CheckInExtra {
  label: string;
  value: string;
  /** Hervorgehoben darstellen (Startnummer — die wird am Tisch vorgelesen). */
  strong?: boolean;
}

interface FieldDef { id: string; label: string }

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
 * Zwei Dinge, die die Zahl ehrlich halten:
 *  - Nur ANGEMELDETE Personen zählen. Wartelisten- und abgemeldete Zeilen
 *    stehen weiter in der Liste; wer sie mitzählt, bestellt für Leute, die
 *    nicht kommen.
 *  - Wer keine Größe angegeben hat, verschwindet nicht, sondern erscheint als
 *    eigene Zeile „ohne Angabe" MIT Namen. Eine Bestellsumme, die stillschweigend
 *    kleiner ist als die Teilnehmerzahl, ist der teurere Fehler.
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
}

export interface ShirtTallyResult {
  /** Das erkannte Feld; leer, wenn es an diesem Event gar keines gibt. */
  fieldLabel: string;
  rows: ShirtTallyRow[];
  /** Angemeldete Personen insgesamt (= Summe über alle Zeilen). */
  total: number;
  /** Davon ohne Angabe. */
  missing: number;
}

const SHIRT_ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

/** Größe → Zählschlüssel (Groß-/Kleinschreibung, Leerzeichen egal). */
export function shirtSizeKey(size: string): string {
  return (size || '').toLowerCase().replace(/\s+/g, '');
}

/** Bekannte Konfektionsgrößen in natürlicher Reihenfolge (für Sortierung UND Nachbarschaft). */
export const SHIRT_SIZE_ORDER = ['xxs', 'xs', 's', 'm', 'l', 'xl', 'xxl', 'xxxl', '3xl', '4xl'];

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
  /** true = Wunschgröße reicht NICHT für diese Person. */
  short: boolean;
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
 *  - Unbekannte Größenbezeichnungen (z.B. „Damen M") haben keine Nachbarn in
 *    SHIRT_SIZE_ORDER; für sie gibt es nur „reicht / reicht nicht".
 */
export function shirtAllocate(
  fields: FieldDef[] | undefined | null,
  regs: Array<{ Status?: string; CustomData?: string; ParticipantName?: string; ParticipantEmail?: string; TeilnehmerID?: number | string | null; Id?: number }> | undefined | null,
  stock: ShirtStock,
): ShirtAllocationResult {
  const result: ShirtAllocationResult = { hasStock: Object.keys(stock || {}).length > 0, byEmail: {}, rows: [], noneLeft: [] };
  const field = (fields || []).filter(f => SHIRT_PATTERN.test(f.label || ''))[0];
  if (!field) return result;
  const remaining: ShirtStock = {};
  Object.keys(stock || {}).forEach(k => { remaining[k] = stock[k]; });
  const display: Record<string, string> = {};
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
    const i = SHIRT_SIZE_ORDER.indexOf(key);
    if (i < 0) return [];
    const out: string[] = [];
    for (let d = 1; d < SHIRT_SIZE_ORDER.length; d++) {
      if (i + d < SHIRT_SIZE_ORDER.length) out.push(SHIRT_SIZE_ORDER[i + d]);
      if (i - d >= 0) out.push(SHIRT_SIZE_ORDER[i - d]);
    }
    return out;
  };
  for (const r of active) {
    const cd = parseCustomData(r.CustomData);
    const raw = cd[field.id];
    const wish = (raw === undefined || raw === null) ? '' : String(raw).trim();
    if (!wish) continue;
    const key = shirtSizeKey(wish);
    if (!display[key]) display[key] = wish;
    need[key] = (need[key] || 0) + 1;
    const email = (r.ParticipantEmail || '').toLowerCase().trim();
    const name = (r.ParticipantName || r.ParticipantEmail || '—').trim();
    if (!result.hasStock) continue;
    if ((remaining[key] || 0) > 0) {
      remaining[key]--;
      if (email) result.byEmail[email] = { wish, proposal: null, short: false };
      continue;
    }
    missing[key] = (missing[key] || 0) + 1;
    let proposal: string | null = null;
    for (const nk of neighbours(key)) {
      if ((remaining[nk] || 0) > 0) { remaining[nk]--; proposal = nk.toUpperCase(); break; }
    }
    if (!proposal) result.noneLeft.push(name);
    if (email) result.byEmail[email] = { wish, proposal, short: true };
  }
  const keys = Array.from(new Set(Object.keys(need).concat(Object.keys(stock || {}))));
  const rank = (k: string): number => { const i = SHIRT_SIZE_ORDER.indexOf(k); return i >= 0 ? i : 500; };
  result.rows = keys
    .sort((a, b) => (rank(a) - rank(b)) || a.localeCompare(b, 'de'))
    .map(k => ({
      size: display[k] || k.toUpperCase(),
      key: k,
      need: need[k] || 0,
      stock: (stock && stock[k]) || 0,
      missing: missing[k] || 0,
      spare: result.hasStock ? (remaining[k] || 0) : 0,
    }));
  return result;
}

export function shirtTally(
  fields: FieldDef[] | undefined | null,
  regs: Array<{ Status?: string; CustomData?: string; ParticipantName?: string; ParticipantEmail?: string }> | undefined | null
): ShirtTallyResult {
  const field = (fields || []).filter(f => SHIRT_PATTERN.test(f.label || ''))[0];
  const out: ShirtTallyResult = { fieldLabel: field ? field.label : '', rows: [], total: 0, missing: 0 };
  if (!field) return out;
  const byKey: Record<string, ShirtTallyRow> = {};
  const order: string[] = [];
  for (const r of (regs || [])) {
    if (SHIRT_ACTIVE_STATI.indexOf(r.Status || '') < 0) continue;
    out.total++;
    const cd = parseCustomData(r.CustomData);
    const raw = cd[field.id];
    const val = (raw === undefined || raw === null) ? '' : String(raw).trim();
    const key = val.toLowerCase().replace(/\s+/g, '');
    if (!byKey[key]) { byKey[key] = { size: val, count: 0, names: [] }; order.push(key); }
    byKey[key].count++;
    byKey[key].names.push((r.ParticipantName || r.ParticipantEmail || '—').trim());
    if (!val) out.missing++;
  }
  // Bekannte Konfektionsgrößen in ihrer natürlichen Reihenfolge, alles andere
  // alphabetisch dahinter, „ohne Angabe" ganz zuletzt — eine Bestellliste, die
  // mit XXL beginnt, liest niemand gern.
  const SIZE_ORDER = ['xxs', 'xs', 's', 'm', 'l', 'xl', 'xxl', 'xxxl', '3xl', '4xl'];
  const rank = (k: string): number => {
    if (!k) return 9999;
    const i = SIZE_ORDER.indexOf(k);
    return i >= 0 ? i : 500;
  };
  out.rows = order
    .map(k => ({ k, row: byKey[k] }))
    .sort((a, b) => (rank(a.k) - rank(b.k)) || a.k.localeCompare(b.k, 'de'))
    .map(x => x.row);
  return out;
}
