/**
 * Deutung der Hotel-Fragen aus dem Anmeldeformular.
 *
 * v29.6: Aus `HotelSetupWizard` herausgezogen, weil die Hotelplanung (Spalte
 * „Nächte gewünscht") dieselbe Deutung braucht. Zwei Kopien derselben Regel
 * laufen erfahrungsgemäß auseinander, und dann widersprechen sich Assistent
 * und Tabelle — genau die Sorte Widerspruch, die in v29.2/29.3 zwei Releases
 * gekostet hat.
 *
 * Die Formulare fragen typischerweise zwei Dinge getrennt ab:
 *   „Hotel (24-25 Sept): Do you require accommodation?"      → Bedarfsfrage
 *   „Hotel (additional nights): … beforehand?"               → Zusatznächte
 * Die Bedarfsfrage beantwortet „braucht ein Zimmer ja/nein", die zweite sagt,
 * wie viele Nächte ÜBER den Standard-Zeitraum hinaus.
 */

/** Feld-Label/Hilfetext gehört zum Thema Hotel. */
export const HOTEL_RE = /hotel|unterkunft|übernacht|übernacht|accommodation|lodging|zimmer/i;
/** …und meint die ZUSATZnächte, nicht den Bedarf. */
export const EXTRA_RE = /additional|extra|zusätzlich|zusätzlich|weitere|vorab|beforehand|früher|frueher|verlänger|verlänger|longer/i;
/** Antwort verneint den Bedarf. */
export const NO_RE = /^\s*(nein|no|kein|none|nicht)\b|^\s*-\s*$/i;
const AFTER_RE = /after|danach|länger|länger|abreis|departure|extend|following/i;
const BEFORE_RE = /before|beforehand|vorab|vorher|früher|frueher|anreis|prior/i;
const WORD_NUMS: Record<string, number> = {
  one: 1, two: 2, three: 3, four: 4, five: 5,
  ein: 1, eine: 1, zwei: 2, drei: 3, vier: 4, fünf: 5, fuenf: 5,
};

/**
 * Antwort auf die Zusatznächte-Frage deuten.
 * `null` = nicht deutbar (der Organizer wählt dann selbst).
 */
export const parseExtraAnswer = (ansIn: string): { nights: number; after: boolean } | null => {
  const ans = (ansIn || '').trim();
  if (!ans) return { nights: 0, after: false };
  if (NO_RE.test(ans)) return { nights: 0, after: false };
  const s = ans.toLowerCase();
  let n = 0;
  // „2 additional nights" / „2 Nächte" — nur Zahlen direkt vor dem Nacht-Wort,
  // damit Datumsangaben wie „from 23 - 24 Sept." nicht mitgezählt werden.
  const digit = s.match(/(\d+)\s*(additional\s+|extra\s+|weitere\s+|zusätzliche\s+|zusätzliche\s+)?(night|nacht|nächte|naechte)/);
  if (digit) n = parseInt(digit[1], 10);
  else {
    const keys = Object.keys(WORD_NUMS);
    for (const w of keys) {
      if (new RegExp(`\\b${w}\\b`, 'i').test(s)) { n = WORD_NUMS[w]; break; }
    }
  }
  if (!n) return null;
  const after = AFTER_RE.test(s) && !BEFORE_RE.test(s);
  return { nights: n, after };
};
