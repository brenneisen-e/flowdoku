/**
 * v28.94: Aus `EventCreationPage` herausgeloest. Heuristiken, die anhand der
 * BESCHRIFTUNG eines Abfragefelds einen passenderen Feldtyp vorschlagen —
 * reine Textpruefungen ohne Bezug zum Wizard-State.
 */
export // v24.25: Heuristiken für die Feldart-Empfehlung im Feld-Editor — wenn das
// Feld-Label nach einem Datum bzw. nach einer Person/einem Namen klingt,
// schlägt der Wizard die passende Feldart vor (Kalender bzw. People-Picker).
function labelLooksLikeDate(label: string): boolean {
  // v26.91: „date" nur als eigenes Wort (\b) — sonst matchte es „Date"n in
  // „Datenschutz(hinweise)" und schlug fälschlich eine Datums-Umstellung vor.
  // „datum" bleibt ohne Grenze (deutsche Komposita wie „Geburtsdatum").
  return /(datum|\bdate\b|check[\s-]?in|check[\s-]?out|anreise|abreise|geburtstag|birthday|deadline|frist|termin|ankunft|abfahrt|arrival|departure)/i.test(label || '');
}
export function labelLooksLikeName(label: string): boolean {
  return /(\bname\b|vorname|nachname|ansprechpartner|counselor|kolleg|mitarbeiter|\bmentor\b|\bpate\b|\bbuddy\b|begleitung|\bgast\b)/i.test(label || '');
}
export // v24.28: Felder, die i.d.R. schon automatisch aus dem Deloitte-Profil kommen
// (Standort, Abteilung, Unternehmenszugehörigkeit/Rechtsträger, Telefon, Name,
// E-Mail) — die muss der Organizer nicht extra abfragen. „name" allein bewusst
// NICHT (zu mehrdeutig, z.B. „Name of counselor").
function labelLooksLikeProfile(label: string): boolean {
  // v26.91: Telefon/Mobil/Handy bewusst NICHT mehr — eine (private) Mobilnummer
  // z.B. für den B2Run-Infoservice steht i.d.R. NICHT im Deloitte-Profil und ist
  // eine legitime Abfrage; der „schon automatisch erfasst"-Hinweis passte da nicht.
  // v31.45: Sechs Begriffe dazu (Nutzer-Befund 15.09.2026: Ein Event fragte
  // „Level", „Office" und „Bereich" ab, obwohl Position und Geschäftsbereich in
  // der Profilkarte darüber stehen). Erkannt wurden davon nur „Office".
  // Neu: Geschäftsbereich, Business Unit, Service Line, Level, Grade, Position.
  //
  // Bewusst NICHT „Bereich" allein: „In welchem Bereich möchtest du mitmachen?"
  // ist eine legitime Frage, und ein Fehlalarm macht den Hinweis unglaubwürdig —
  // dieselbe Lehre wie v26.83 (Telefon) und v26.91 (Mobilnummer).
  return /(vorname|nachname|first ?name|last ?name|e-?mail|abteilung|department|geschäftsbereich|geschaeftsbereich|business ?unit|service ?line|standort|location|\boffice\b|\bbüro\b|firma|company|unternehmen|arbeitgeber|gesellschaft|\bgmbh\b|legal ?entity|\bentity\b|rechtsträger|member ?firm|adresse|address|job ?title|\blevel\b|\bgrade\b|\bposition\b)/i.test(label || '');
}

/**
 * v31.45: Ein Auswahlfeld, das in Wahrheit ein Übernachtungs-Zeitraum ist.
 *
 * Nutzer-Befund 15.09.2026 mit Bild: Ein Event fragte „Hotelzimmer benötigt?"
 * als Dropdown mit den Optionen „Ja - beide Nächte", „Ja - nur den 19.01.",
 * „Ja - nur den 20.01.", „Nein". Dafür gibt es seit langem die Feldart
 * `daterange` („Übernachtungs-Zeitraum", Kalender + Nächte).
 *
 * Das ist nicht nur umständlicher, es KOSTET Funktion: Die Hotelplanung im
 * Organizer Center (Zimmerverteilung, `autoDistribute`, `formStayOf`) liest die
 * Nächte aus dem Zeitraum-Feld. Bei „Ja - nur den 19.01." als Auswahltext steht
 * dort nichts Auswertbares — die Person faellt aus der Verteilung.
 *
 * Erkannt wird bewusst nur die KOMBINATION aus Hotel-Label UND datumsartigen
 * Optionen. „Hotelzimmer benötigt? Ja/Nein" ohne Daten ist eine saubere
 * Ja/Nein-Frage und soll keinen Hinweis bekommen.
 */
export function selectLooksLikeStayRange(label: string, options: string[] | undefined): boolean {
  const l = (label || '').trim();
  if (!l) return false;
  if (!/(hotel|zimmer|übernacht|uebernacht|\broom\b|accommodation|nächte|naechte|\bnights?\b)/i.test(l)) return false;
  const opts = (options || []).filter(Boolean);
  if (opts.length < 2) return false;
  // Datumsartig = ein Tag.Monat, ein ISO-Datum oder ein Monatsname.
  const datumsartig = (o: string): boolean =>
    /\b\d{1,2}\.\s?\d{1,2}\.?(\s?\d{2,4})?\b/.test(o)
    || /\b\d{4}-\d{2}-\d{2}\b/.test(o)
    || /(januar|februar|märz|maerz|april|juni|juli|august|september|oktober|november|dezember|january|february|march|may|june|july|august|october|december)/i.test(o);
  return opts.filter(datumsartig).length >= 2;
}
