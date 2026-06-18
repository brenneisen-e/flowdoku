/**
 * v23.44: Strukturierte Release Notes — Laufzeit-Quelle für den Wochenbericht.
 *
 * WICHTIG (Konvention, siehe CLAUDE.md): Bei JEDEM Release sowohl
 * `docs/release-notes.md` ALS AUCH diese Liste pflegen — der automatische
 * Wochenbericht liest diese Liste und zeigt die Einträge der Berichtswoche an.
 * `date` = ISO-Datum (YYYY-MM-DD), neueste Version oben. `type` = 'Feature'
 * (neue Funktion) oder 'Bugfix' (Fehlerbehebung). `text` = Klartext, nicht
 * technisch (aus Sicht des Nutzers).
 */
export interface ReleaseNote {
  version: string;
  date: string; // YYYY-MM-DD
  type: 'Feature' | 'Bugfix';
  text: string;
}

export const RELEASE_NOTES: ReleaseNote[] = [
  { version: '23.45.0', date: '2026-06-18', type: 'Feature', text: 'Sub-Events werden im Check-in-Auswahlbildschirm jetzt unter ihrem Hauptevent gruppiert und eingeklappt angezeigt (Aufklappen per Pfeil). Die „Admin"-Kachel auf der Startseite hat eine Kapitänsmütze als Symbol. Die Organizer-Eventübersicht heißt jetzt klar „Organizer – Eventübersicht".' },
  { version: '23.44.0', date: '2026-06-18', type: 'Feature', text: 'Prozessübersicht in saubere, verständliche Ablaufdiagramme (BPMN) überführt: Kreise für Start/Ende, abgerundete Rechtecke für Schritte, Rauten für Entscheidungen — durchgehend ohne Technik-Begriffe erklärt. Im Admin-Bereich gibt es jetzt eine eigene Kachel zum direkten Öffnen einer SharePoint-Liste; die Event-Liste der Organizer zeigt nur noch „Neues Event erstellen" (alle anderen Admin-Funktionen liegen gebündelt im Admin-Bereich). Der Wochenbericht enthält ab sofort die Neuerungen der Woche.' },
  { version: '23.43.0', date: '2026-06-18', type: 'Feature', text: 'Prozessübersicht verständlicher: einladende Einleitung und Legende in Klartext.' },
  { version: '23.42.0', date: '2026-06-18', type: 'Feature', text: 'Event-Übersicht der Organizer/Admins neu gestaltet: höhere Zeilen, Datum und Ort mit Symbolen, Organizer mit Foto und Name (Detail beim Drüberfahren), Warteliste unter der Teilnehmerzahl, Status als farbiges Eck-Etikett oben links (grün = aktiv, orange = Entwurf).' },
  { version: '23.41.0', date: '2026-06-18', type: 'Feature', text: 'Neue „Admin"-Kachel auf der Startseite als zentrale Anlaufstelle: Prozessübersicht, Rollenverwaltung, Einstellungen, Handbuch, Organizer Center, Archivieren/Löschen und eine verständliche Erklärung aller Hintergrund-Listen.' },
  { version: '23.40.0', date: '2026-06-18', type: 'Feature', text: 'Löschkonzept fürs Archiv: Einträge älter als ein halbes Jahr können von Admins endgültig gelöscht werden; beim Öffnen der App sieht ein Admin einen Hinweis dazu.' },
  { version: '23.39.0', date: '2026-06-18', type: 'Feature', text: 'Die Archivierung räumt jetzt auch alte Einträge von bereits gelöschten Events weg, damit die Hintergrund-Listen schlank bleiben.' },
  { version: '23.38.0', date: '2026-06-18', type: 'Bugfix', text: 'Der automatische Wochenbericht wird wieder zuverlässig ausgelöst und geht als eine Mail mit euren echten Namen an alle Admins.' },
  { version: '23.37.0', date: '2026-06-18', type: 'Feature', text: 'Wer Organizer werden möchte, stellt jetzt einen Antrag in der App; Admins sehen offene Anträge beim Öffnen der App und geben sie mit einem Klick frei. Die Antrags-Mail enthält einen Direkt-Link zum Bestätigen.' },
  { version: '23.36.0', date: '2026-06-18', type: 'Feature', text: 'Wochenbericht verständlicher: Entwürfe werden als solche gekennzeichnet, neu veröffentlichte und stattgefundene Events (mit Teilnehmerzahl) werden aufgeführt.' },
  { version: '23.35.0', date: '2026-06-18', type: 'Bugfix', text: 'Der automatische Wochenbericht wurde nie verschickt — behoben, er kommt jetzt zuverlässig bei allen Admins an.' },
  { version: '23.34.0', date: '2026-06-18', type: 'Feature', text: 'Teilnehmerlisten: kein überflüssiger Tooltip mehr beim Überfahren der Fotos.' },
  { version: '23.33.0', date: '2026-06-18', type: 'Feature', text: 'Teilnehmerliste: eingeklappte Personen-Spalte (Foto + Name + Position/Standort), Suchtreffer werden markiert, alle Spalten sortierbar.' },
  { version: '23.32.0', date: '2026-06-18', type: 'Feature', text: 'Assistenz-/People-Felder werden in der konsolidierten Teilnehmer-Matrix mit Foto und Name angezeigt; Suche über alle Spalten mit Markierung.' },
  { version: '23.31.0', date: '2026-06-18', type: 'Feature', text: 'In der QR-Code-Mail ist „DEX App" jetzt ein anklickbarer Link.' },
  { version: '23.30.0', date: '2026-06-18', type: 'Feature', text: 'Im Organizer Center steht „Warteliste" direkt neben „Angemeldet"; in der Event-Übersicht steht pro Event, wie viele auf der Warteliste sind.' },
  { version: '23.29.0', date: '2026-06-18', type: 'Bugfix', text: '„No-Show" wird nur noch für neu angelegte Events angeboten.' },
  { version: '23.28.0', date: '2026-06-18', type: 'Feature', text: 'Neuer Status „No-Show": das Check-in-Team kann Teilnehmer als „nicht erschienen" markieren.' },
  { version: '23.27.0', date: '2026-06-17', type: 'Bugfix', text: 'Der Teams-Chat-Link beim Organizer öffnet jetzt direkt die Teams-App.' },
  { version: '23.26.0', date: '2026-06-17', type: 'Feature', text: 'Mehrere Organizer in einer gemeinsamen Kachel; E-Mail- und Teams-Chat-Link je Person; freie Plätze neben dem Registrieren-Button.' },
  { version: '23.25.0', date: '2026-06-17', type: 'Feature', text: 'Bild-Einstellungen im Bild-Editier-Modal gebündelt; Organizer-Karte mit klickbarer Mail; Option „Organizer groß anzeigen".' },
  { version: '23.24.0', date: '2026-06-17', type: 'Bugfix', text: 'Event-Karte: angepasstes Foto überdeckt nicht mehr den Titel; freie Plätze nicht mehr doppelt.' },
];
