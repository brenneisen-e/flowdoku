/**
 * v27.10 (Refactor): Heuristiken für die Feldart-Empfehlung im Feld-Editor —
 * unverändert aus EventCreationPage.tsx extrahiert (reine Regex-Helfer ohne
 * Komponenten-Abhängigkeiten).
 *
 * v24.25: wenn das Feld-Label nach einem Datum bzw. nach einer Person/einem
 * Namen klingt, schlägt der Wizard die passende Feldart vor (Kalender bzw.
 * People-Picker).
 */

export function labelLooksLikeDate(label: string): boolean {
  // v26.91: „date" nur als eigenes Wort (\b) — sonst matchte es „Date"n in
  // „Datenschutz(hinweise)" und schlug fälschlich eine Datums-Umstellung vor.
  // „datum" bleibt ohne Grenze (deutsche Komposita wie „Geburtsdatum").
  return /(datum|\bdate\b|check[\s-]?in|check[\s-]?out|anreise|abreise|geburtstag|birthday|deadline|frist|termin|ankunft|abfahrt|arrival|departure)/i.test(label || '');
}

export function labelLooksLikeName(label: string): boolean {
  return /(\bname\b|vorname|nachname|ansprechpartner|counselor|kolleg|mitarbeiter|\bmentor\b|\bpate\b|\bbuddy\b|begleitung|\bgast\b)/i.test(label || '');
}

// v24.28: Felder, die i.d.R. schon automatisch aus dem Deloitte-Profil kommen
// (Standort, Abteilung, Unternehmenszugehörigkeit/Rechtsträger, Name, E-Mail)
// — die muss der Organizer nicht extra abfragen. „name" allein bewusst NICHT
// (zu mehrdeutig, z.B. „Name of counselor").
export function labelLooksLikeProfile(label: string): boolean {
  // v26.91: Telefon/Mobil/Handy bewusst NICHT mehr — eine (private) Mobilnummer
  // z.B. für den B2Run-Infoservice steht i.d.R. NICHT im Deloitte-Profil und ist
  // eine legitime Abfrage; der „schon automatisch erfasst"-Hinweis passte da nicht.
  return /(vorname|nachname|first ?name|last ?name|e-?mail|abteilung|department|standort|location|\boffice\b|\bbüro\b|firma|company|unternehmen|arbeitgeber|gesellschaft|\bgmbh\b|legal ?entity|\bentity\b|rechtsträger|member ?firm|adresse|address|job ?title)/i.test(label || '');
}
