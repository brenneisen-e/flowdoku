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

/** Feld-Label → gehört ans Check-in? Aktuell: alles, was nach Trikot/Größe klingt. */
const SHIRT_PATTERN = /(t-?\s?shirt|trikot|shirt\s*size|kleidergr|konfektionsgr)/i;

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
