/**
 * v28.94: `shortSubEventTitle` stand dreimal im Code — in `EventCreationPage`,
 * `AdminPage` und (in einer verkürzten Variante ohne `parentTitle`) in
 * `CheckInPage`. Drei Kopien derselben Regel heißen: Wer die Regel ändert,
 * ändert sie an einer Stelle und wundert sich, dass die Reiter im Assistenten
 * anders heißen als im Organizer Center.
 *
 * Die Regel selbst: Sub-Event-Titel werden als „<Hauptevent> | <Teil>"
 * gespeichert, damit sie in SharePoint-Listen und Mails für sich stehen. In
 * der Oberfläche interessiert nur der hintere Teil.
 */
export function shortSubEventTitle(title: string | undefined, parentTitle?: string): string {
  const t = (title || '').trim();
  if (!t) return t;
  const pipe = t.lastIndexOf('|');
  if (pipe >= 0) {
    const after = t.substring(pipe + 1).trim();
    if (after) return after;
  }
  // Ohne Pipe: Ein vorangestellter Hauptevent-Name wird abgeschnitten, samt
  // der üblichen Trennzeichen dahinter.
  const p = (parentTitle || '').trim();
  if (p && t.toLowerCase().startsWith(p.toLowerCase())) {
    const rest = t.substring(p.length).replace(/^[\s|:\-–—·•]+/, '').trim();
    if (rest) return rest;
  }
  return t;
}
