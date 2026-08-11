/**
 * saveGuard (v28.98)
 *
 * Merker dafür, dass gerade ein Speichervorgang läuft, der nicht unterbrochen
 * werden darf.
 *
 * Hintergrund: Das Speichern eines Events mit vielen Sub-Events dauert —
 * jeder Termin ist ein eigenes DEX_Events-Item mit eigener Teilnehmerliste und
 * eigenem Outlook-Termin. Wer in dieser Zeit auf „Zurück" klickt, verlässt die
 * Seite mitten im Lauf: Was schon geschrieben ist, bleibt; der Rest nicht. Das
 * Ergebnis ist ein halb angelegtes Event, dem man das nicht ansieht.
 *
 * Bewusst ein Modul-Singleton mit Abonnenten statt eines React-Contexts: Der
 * Header steht ausserhalb der Seiten-Komponente, und ein Context nur für dieses
 * eine Flag würde beide aneinanderbinden. Die Abonnenten sind nötig, weil der
 * Header sich neu zeichnen muss — anders als beim wizardStepContext (v26.30),
 * der nur imperativ gelesen wird.
 */

let saving = false;
const listeners = new Set<(_v: boolean) => void>();

/** Speichervorgang an-/abmelden. */
export function setSaveInProgress(v: boolean): void {
  if (saving === v) return;
  saving = v;
  listeners.forEach(l => { try { l(v); } catch { /* ein defekter Listener darf den Save nicht stoeren */ } });
}

export function isSaveInProgress(): boolean { return saving; }

/** Abonnieren; gibt die Abmelde-Funktion zurück. */
export function subscribeSaveInProgress(l: (_v: boolean) => void): () => void {
  listeners.add(l);
  return () => { listeners.delete(l); };
}
