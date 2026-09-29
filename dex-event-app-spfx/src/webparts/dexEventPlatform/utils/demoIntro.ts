/**
 * v32.18: Admin-Demo „Einführungs-Hinweis ansehen" (Burger-Menü).
 *
 * Nutzer-Ansage 29.09.2026: Als Admin will ich sehen, wie die Startseite für
 * Leute aussieht, denen DEX den Hinweis auf die Einführungsveranstaltung
 * zeigt. Echt sieht man ihn nie: Er steht nur in der Organizer-Werbekarte
 * (also nicht für Organizer) und verschwindet, sobald man zur Einführung
 * angemeldet ist. Der Schalter lebt nur im Speicher dieser Sitzung — ein
 * Neuladen beendet die Demo, damit niemand darin hängen bleibt.
 */
import * as React from 'react';

let aktiv = false;
const hoerer = new Set<() => void>();

export const setDemoEinfuehrung = (v: boolean): void => {
  aktiv = v;
  hoerer.forEach(f => f());
};

export const useDemoEinfuehrung = (): boolean => {
  const [, neu] = React.useReducer((x: number) => x + 1, 0);
  React.useEffect(() => {
    hoerer.add(neu);
    return () => { hoerer.delete(neu); };
  }, []);
  return aktiv;
};
