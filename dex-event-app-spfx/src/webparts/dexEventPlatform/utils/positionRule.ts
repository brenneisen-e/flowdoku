/**
 * v32.2.2: Frage nur für bestimmte Positionen — Nutzer-Ansage 28.09.2026:
 * „dass die Frage nur angezeigt wird, wenn man eine bestimmte Position hat
 * (z.B. alle, die nicht Partner oder Director sind)".
 *
 * Gespeichert am Feld als `showForPositions: { mode, values }`:
 *  - mode 'only'   → nur Personen, deren Position einen der Werte enthält
 *  - mode 'except' → alle außer diesen
 * Verglichen wird als Teil-Text ohne Groß-/Kleinschreibung gegen die
 * Position aus dem Microsoft-Profil (`jobTitle`) — „Partner" trifft damit
 * auch „Associate Partner", „Manager" auch „Senior Manager". Die Vorschläge
 * im Assistenten sagen das dazu.
 *
 * Unbekannte Position (Profil ohne Titel, stellvertretende Anmeldung ohne
 * geladenes Profil): Die Frage wird GEZEIGT. Eine fälschlich gezeigte Frage
 * ist harmlos, eine fälschlich versteckte Pflichtfrage fehlt dem Organizer.
 *
 * EINE Stelle für alle Auswertungen (Anmeldeseite, Termin-Fragen, Pflicht-
 * Prüfung, Meine Events) — wer eine neue Stelle baut, die Felder filtert,
 * ruft diese Funktion zusätzlich zu showIf.
 */
export interface PositionRule {
  mode: 'only' | 'except';
  values: string[];
}

/** Standard-Positionen für die Auswahl-Chips (Deloitte-Stufen). */
export const POSITION_PRESETS: string[] = [
  'Partner', 'Director', 'Senior Manager', 'Manager', 'Senior Consultant', 'Consultant', 'Analyst', 'Assistant',
];

export function positionRuleAllows(rule: PositionRule | undefined | null, jobTitle: string | undefined | null): boolean {
  if (!rule || !Array.isArray(rule.values) || rule.values.length === 0) return true;
  const jt = (jobTitle || '').trim().toLowerCase();
  if (!jt) return true;
  const hit = rule.values.some(v => {
    const x = (v || '').trim().toLowerCase();
    return !!x && jt.indexOf(x) >= 0;
  });
  return rule.mode === 'except' ? !hit : hit;
}

/** Bereinigt eine Regel fürs Speichern — leere Regel = keine Regel. */
export function cleanPositionRule(rule: PositionRule | undefined | null): PositionRule | undefined {
  if (!rule) return undefined;
  const values = (rule.values || []).map(v => (v || '').trim()).filter(Boolean);
  if (values.length === 0) return undefined;
  return { mode: rule.mode === 'except' ? 'except' : 'only', values };
}
