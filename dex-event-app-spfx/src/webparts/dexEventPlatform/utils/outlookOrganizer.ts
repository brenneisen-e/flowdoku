/**
 * v27.10 (Refactor): unverändert aus EventCreationPage.tsx extrahiert
 * (reiner String-Helfer ohne Komponenten-Abhängigkeiten).
 *
 * v27.3: Der Outlook-Body wird beim Speichern mit fest aufgelöstem {{Organizer}}
 * gespeichert. Kommen später Organizer dazu, blieb der alte Name eingebacken.
 * Beim Edit-Laden mappen wir den eingebackenen Organizer-Namen wieder auf
 * {{Organizer}} zurück — dann löst der nächste Save mit ALLEN aktuellen
 * Organizern neu auf. Sicher: findet sich nichts, bleibt der Body unverändert.
 */
export function reinsertOrganizerPlaceholder(body: string, organizers: string[]): string {
  if (!body || !organizers || organizers.length === 0) return body;
  if (body.indexOf('{{Organizer}}') >= 0) return body; // schon Platzhalter
  const names = organizers.map(n => (n || '').trim()).filter(Boolean);
  const full = names.join('; ');
  if (full && body.indexOf(full) >= 0) return body.split(full).join('{{Organizer}}');
  // Bereits „kaputte"/veraltete Bodies: den ersten enthaltenen Organizer-Namen
  // (längster zuerst, um Teil-Treffer zu vermeiden) auf den Platzhalter mappen.
  for (const n of [...names].sort((a, b) => b.length - a.length)) {
    if (n.length >= 3 && body.indexOf(n) >= 0) return body.split(n).join('{{Organizer}}');
  }
  return body;
}
