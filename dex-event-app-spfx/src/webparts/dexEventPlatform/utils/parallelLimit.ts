/**
 * v31.24 — Mehrere asynchrone Aufgaben gleichzeitig, aber gedeckelt.
 *
 * Dieselbe Hilfe, die in `services/events/outlookQueue` seit v31.22 als
 * `mitGrenze` steckt — hier auf Modul-Ebene, weil die zweite Stelle dazukam
 * (die Termin-Listen einer Klammer). Zwei Kopien wären die Konstruktion, bei
 * der die eine irgendwann anders zählt als die andere.
 *
 * **Warum nicht einfach `Promise.all`:** Das startet ALLES gleichzeitig. Bei
 * neunzehn Terminen sind das neunzehn parallele Abfragen gegen dieselbe
 * SharePoint-Site — SharePoint antwortet darauf mit 429 (Drosselung), der
 * Client wartet den Backoff ab, und am Ende dauert es länger als
 * nacheinander. Der Deckel ist der Punkt, nicht die Parallelität.
 *
 * **Warum vier und nicht zehn:** Browser öffnen je Host rund sechs
 * Verbindungen. Mehr als das steht ohnehin in der Warteschlange des Browsers;
 * vier lässt zusätzlich Luft für alles andere, was die Seite gerade lädt
 * (Bilder, die Event-Liste, der Konten-Check).
 *
 * Fehler einzelner Aufgaben werden NICHT geschluckt — sie gehören dem
 * Aufrufer, der allein weiß, ob ein Fehlschlag „übersprungen" oder „Abbruch"
 * heißt. (In DEX ist dieser Unterschied teuer gelernt: Ein Lesefehler ist
 * keine Null, v30.67.)
 */
export async function parallelLimit(
  aufgaben: Array<() => Promise<void>>,
  grenze = 4,
): Promise<void> {
  if (aufgaben.length === 0) return;
  let next = 0;
  const laeufer = new Array(Math.min(Math.max(1, grenze), aufgaben.length))
    .fill(0)
    .map(async () => {
      for (;;) {
        const i = next++;
        if (i >= aufgaben.length) return;
        await aufgaben[i]();
      }
    });
  await Promise.all(laeufer);
}
