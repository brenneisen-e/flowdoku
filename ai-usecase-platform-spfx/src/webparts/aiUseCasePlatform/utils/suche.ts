/**
 * Die Use-Case-Suche — EINE Regel für Kopfzeile und Kachelwand.
 *
 * Bis v1.2 stand der Vergleich in der Kachelwand. Seit die Suche in der
 * Kopfzeile sitzt (wie bei DEX, `GlobalSearch`), brauchen ihn zwei Stellen:
 * das Ergebnis-Fenster unter dem Suchfeld und die Kachelwand. Stünde er zweimal
 * da, fände das Fenster „credit memo" und die Wand nicht — dieselbe Falle wie
 * zwei Bedienwege für dieselbe Auswahl.
 *
 * Der Vergleich ignoriert Bindestriche, Leerzeichen und Umlaute. Genau die
 * Falle aus DEX v31.10: Dort fand die Aktionssuche „email" nicht, weil die
 * Titel „E-Mail" heißen und roh verglichen wurde. Hier heißen Use Cases
 * „KI-Kreditanalyse" und „VST-Cockpit" — wer „ki kredit" tippt, muss sie finden.
 */

import { UseCase } from '../types';

/** Klein, ohne Umlaute, ohne alles außer Buchstaben und Ziffern. */
export function normSuche(s: string): string {
  return (s || '')
    .toLowerCase()
    .replace(/ä/g, 'a').replace(/ö/g, 'o').replace(/ü/g, 'u').replace(/ß/g, 'ss')
    .replace(/[^a-z0-9]/g, '');
}

/** Die Suchbegriffe: ein Wort je Begriff, alle müssen passen. */
export function suchTokens(suche: string): string[] {
  return (suche || '').trim().split(/\s+/).map(normSuche).filter(Boolean);
}

/** Passt der Use Case zu ALLEN Begriffen? Kein Begriff = passt immer. */
export function passtZurSuche(uc: UseCase, tokens: string[]): boolean {
  if (tokens.length === 0) return true;
  const heu = normSuche(uc.titel) + ' ' + normSuche(uc.kurzbeschreibung) + ' ' + normSuche(uc.bereich) + ' ' + normSuche(uc.schlagworte.join(' '));
  return tokens.every(tok => heu.indexOf(tok) >= 0);
}
