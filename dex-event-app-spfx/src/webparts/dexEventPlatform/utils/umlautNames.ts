/**
 * v31.97: Umlaute in Namen zurückgewinnen — für die Concur-Teilnehmerliste.
 *
 * Befund 25.09.2026 (erster Concur-Import B2Run Köln): 14 von 44 Zeilen
 * „Cannot import", alle mit Umschrift im Namen („Benoehr", „Gaessler",
 * „Huesing"). Concur gleicht Mitarbeitende über den EXAKTEN Namen ab und
 * führt sie mit Umlaut („Rütten"); DEX hat Vor- und Nachname aus dem Profil,
 * und dort steht bei diesen Personen die Umschrift.
 *
 * Zwei Wege, in dieser Reihenfolge:
 *  1. `profileUmlautName`: Hat das Profil (Anzeigename, Vor-/Nachname) eine
 *     Schreibweise MIT Umlaut, die umgeschrieben genau dem DEX-Namen
 *     entspricht, gilt sie. Das ist ein Beleg, keine Vermutung.
 *  2. `suggestUmlauts`: sonst ein Vorschlag nach Regeln. Blind ersetzen geht
 *     nicht — „Michael", „Samuel", „Bauer", „Goethe" haben echte ae/ue/oe.
 *     Die Regeln schließen die häufigen Fälle aus, der Rest ist eine
 *     Vermutung und wird im Dialog je Person bestätigt.
 * „ss" → „ß" wird nie geraten („Gaessler" kann Gäßler oder Gässler sein).
 */

const UMLAUT = /[äöüÄÖÜß]/;

/** Die Umschrift, auf die beide Seiten verglichen werden. */
export function transliterate(s: string): string {
  return (s || '')
    .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue')
    .replace(/Ä/g, 'Ae').replace(/Ö/g, 'Oe').replace(/Ü/g, 'Ue')
    .replace(/ß/g, 'ss')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Enthält der Name eine Stelle, die eine Umschrift sein KÖNNTE? */
export function mayBeTransliterated(s: string): boolean {
  return /ae|oe|ue|ss/i.test(s || '');
}

export interface NameParts { first: string; last: string; }

/**
 * Sucht unter den Profil-Kandidaten eine Schreibweise mit Umlaut, die zum
 * DEX-Namen passt. `null` = keine gefunden (dann bleibt der DEX-Name).
 */
export function profileUmlautName(dex: NameParts, candidates: NameParts[]): NameParts | null {
  const f = transliterate(dex.first);
  const l = transliterate(dex.last);
  for (const c of candidates) {
    if (!c || !(UMLAUT.test(c.first) || UMLAUT.test(c.last))) continue;
    if (transliterate(c.first) === f && transliterate(c.last) === l) return { first: c.first.trim(), last: c.last.trim() };
  }
  return null;
}

// Namen mit echtem ae/oe/ue. Kleingeschrieben, ganzes Wort.
const KEEP = [
  'michael', 'michaela', 'raphael', 'raphaela', 'rafael', 'rafaela', 'mikael', 'gael', 'israel', 'ismael',
  'samuel', 'manuel', 'manuela', 'emanuel', 'emanuela', 'immanuel', 'miguel', 'joel', 'joelle', 'noel', 'noemi',
  'zoe', 'chloe', 'joe', 'aeneas', 'goethe', 'hoeness', 'coelho', 'boeing', 'phoebe', 'raoul', 'suen',
];

const VOWEL = /[aeiouyäöü]/i;

/** Ein Wort: ae/oe/ue → ä/ö/ü, außer dort, wo es fast immer echt ist. */
function suggestWord(w: string): string {
  if (!w || KEEP.indexOf(w.toLowerCase()) >= 0) return w;
  let out = '';
  for (let i = 0; i < w.length; i++) {
    const a = w[i];
    const b = w[i + 1] || '';
    const pair = (a + b).toLowerCase();
    if (pair === 'ae' || pair === 'oe' || pair === 'ue') {
      const prev = w[i - 1] || '';
      const rest = w.slice(i + 2);
      // au/eu/ou + e („Bauer", „Neuer"), qu („Quentin"), vorangehender Vokal
      // allgemein, Wortende („Rue", „Zoe") und -el am Ende („Samuel").
      const echt = (prev && (VOWEL.test(prev) || prev.toLowerCase() === 'q'))
        || rest.length === 0
        || /^l$/i.test(rest);
      if (!echt) {
        const upper = a === a.toUpperCase();
        const m = pair === 'ae' ? 'ä' : pair === 'oe' ? 'ö' : 'ü';
        out += upper ? m.toUpperCase() : m;
        i++;
        continue;
      }
    }
    out += a;
  }
  return out;
}

/** Vorschlag für einen ganzen Namen (Wörter und Bindestrich-Teile einzeln). */
export function suggestUmlauts(name: string): string {
  return (name || '').split(/(\s+|-)/).map(p => (/^(\s+|-)$/.test(p) ? p : suggestWord(p))).join('');
}
