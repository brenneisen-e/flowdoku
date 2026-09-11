/**
 * v31.29 — Lange Teilnehmerlisten erst als Ausschnitt zeichnen, den Rest im
 * nächsten Frame.
 *
 * ## Warum das nötig ist
 *
 * Seit v31.24/v31.27 liegen die Zeilen beim Reiterwechsel schon im Speicher —
 * die Konsole meldet „Liste stand sofort (aus dem Speicher)". Trotzdem hakt
 * der Klick. Der Grund ist nicht mehr das Laden, sondern das ZEICHNEN: Bei 439
 * Zeilen mal rund zwanzig Spalten baut React in EINEM Durchlauf etwa 9.000
 * Zellen, jede mit Pille, Foto, Knöpfen. Das ist ein einzelner, nicht
 * unterbrechbarer Block auf dem Hauptthread — React 17 kann ihn nicht
 * aufteilen (`useTransition`/`useDeferredValue` gibt es erst ab 18, und ein
 * Versions-Sprung hängt am SPFx-Gerüst, nicht an uns).
 *
 * ## Warum nicht virtualisieren
 *
 * Der Lehrbuch-Weg für lange Listen ist Virtualisierung (react-window &Co.:
 * nur der sichtbare Ausschnitt im DOM). Hier trägt er nicht: Die Tabelle lebt
 * von `min-width: max-content` — die Spaltenbreiten rechnet der Browser aus
 * ALLEN Zellen. Virtualisierung setzt absolut positionierte Zeilen fester
 * Höhe voraus; damit wären automatische Spaltenbreiten, der klebende
 * Tabellenkopf und der waagerechte Gleichlauf von Kopf und Körper weg. Das
 * wäre ein Umbau der gesamten Tabelle — in einer Codebasis ohne Tests
 * (CLAUDE.md) der teuerste denkbare Weg für ein Ruckeln beim Umschalten.
 *
 * ## Was stattdessen passiert
 *
 * Genau das, was der Nutzer vorgeschlagen hat (11.09.2026: „nur die ersten 10
 * TN der anderen Liste laden, anzeigen, schnell, und den Rest noch
 * nachladen"): Der erste Durchlauf zeichnet {@link ERSTE_ZEILEN} Zeilen — mehr
 * als auf den Bildschirm passt, also sieht niemand einen Ausschnitt — und der
 * nächste Frame zeichnet den Rest. Der Klick wird dadurch sichtbar beantwortet,
 * bevor die teure Arbeit läuft.
 *
 * Zwei Stufen, nicht zehn: Ohne `React.memo` auf der Zeile baut JEDE Stufe die
 * bereits gezeichneten Zellen erneut — bei zehn Stufen wäre die Gesamtarbeit
 * ein Vielfaches. Zwei Stufen kosten rund 9 % mehr Gesamtarbeit und verschieben
 * die spürbare Wartezeit hinter das erste Bild. (Eine memoisierte Zeile wäre
 * die lineare Variante, braucht aber alle zwanzig Zell-Abhängigkeiten als
 * Props — der nächste Schritt, falls zwei Stufen nicht reichen.)
 *
 * Kurze Listen werden NICHT gestückelt: Unterhalb von
 * {@link OHNE_STUECKELN_BIS} ist der zweite Durchlauf teurer als der Gewinn.
 */

import * as React from 'react';
import { perfLog } from './perfLog';

/** Erste Stufe — deutlich mehr, als in den 70vh hohen Scrollbereich passt. */
export const ERSTE_ZEILEN = 40;

/** Bis hierhin in einem Zug zeichnen — Stückeln lohnt erst darüber. */
export const OHNE_STUECKELN_BIS = 60;

interface Stand {
  /** Wofür dieser Stand gilt — wechselt er, beginnt die Stückelung von vorn. */
  schluessel: string;
  /** Wie viele Zeilen dieser Durchlauf zeichnen darf. */
  n: number;
}

/**
 * Liefert die Zahl der Zeilen, die JETZT gezeichnet werden dürfen.
 *
 * @param gesamt Wie viele Zeilen es insgesamt gibt.
 * @param schluessel Wechselt, wenn eine ANDERE Liste gezeigt wird (Termin,
 *   Suchbegriff). Nur dann wird neu gestückelt — beim bloßen Umsortieren oder
 *   beim Nachladen im Hintergrund bleibt die Tabelle vollständig stehen, sonst
 *   blinkt sie bei jeder Aktualisierung.
 */
export function useProgressiveRows(gesamt: number, schluessel: string): number {
  const ersteStufe = gesamt <= OHNE_STUECKELN_BIS ? gesamt : ERSTE_ZEILEN;
  const [stand, setStand] = React.useState<Stand>(() => ({ schluessel, n: ersteStufe }));
  // Zustand während des Renderns anpassen statt im Effect: Ein Effect liefe
  // erst NACH dem Zeichnen — die teure Arbeit wäre dann schon getan und die
  // Tabelle spränge von 439 auf 40 zurück.
  const passt = stand.schluessel === schluessel;
  if (!passt) setStand({ schluessel, n: ersteStufe });
  const n = passt ? Math.min(stand.n, gesamt) : ersteStufe;

  // Startzeit der laufenden Stückelung — als Ref, damit das Protokollieren
  // keinen weiteren Durchlauf auslöst.
  const startRef = React.useRef(performance.now());
  const letzterSchluessel = React.useRef(schluessel);
  if (letzterSchluessel.current !== schluessel) {
    letzterSchluessel.current = schluessel;
    startRef.current = performance.now();
  }

  React.useEffect(() => {
    if (n >= gesamt) {
      if (startRef.current > 0) {
        perfLog('Teilnehmertabelle fertig gezeichnet', performance.now() - startRef.current,
          { zeilen: gesamt, hinweis: gesamt <= OHNE_STUECKELN_BIS ? 'in einem Zug' : 'zweite Stufe' });
        startRef.current = 0;
      }
      return undefined;
    }
    const id = requestAnimationFrame(() => {
      if (startRef.current > 0) {
        perfLog('Teilnehmertabelle — erste Zeilen sichtbar', performance.now() - startRef.current, { zeilen: n });
      }
      setStand(s => (s.schluessel === schluessel ? { schluessel, n: gesamt } : s));
    });
    return () => cancelAnimationFrame(id);
  }, [n, gesamt, schluessel]);

  return n;
}
