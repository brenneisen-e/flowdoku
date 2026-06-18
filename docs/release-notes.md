# DEX Event Experience Platform — Release Notes

Pflege-Regel: siehe `CLAUDE.md` → „Release Notes — Pflicht bei JEDEM Release".
Bei **jedem** Versions-Bump kommt hier eine Zeile dazu (neueste Version **oben**),
Spalten: **Version | Datum | Art | Beschreibung**. `Art` = `Feature` (neue
Funktionalität) oder `Bugfix` (Fehlerbehebung). Beschreibung auf Deutsch,
nutzerverständlich (was sich für Organizer/Teilnehmer ändert).

| Version | Datum | Art | Beschreibung |
|---------|-------|-----|--------------|
| 23.35.0 | 2026-06-18 | Bugfix | Der automatische Wochenbericht kam nie an: Er wurde mit leerem Event-Bezug in die Mail-Queue geschrieben, woran der Versand-Flow scheiterte. Behoben — der Bericht wird jetzt korrekt an alle Admins versendet. |
| 23.34.0 | 2026-06-18 | Feature | Teilnehmerlisten: Kein überflüssiger Tooltip mehr beim Überfahren der Teilnehmer-Fotos (Name + Position stehen ohnehin daneben); der Hover-Zoom bleibt. |
| 23.33.0 | 2026-06-18 | Feature | Normale Teilnehmerliste: standardmäßig eingeklappte „Teilnehmer"-Spalte (Foto + Name fett, darunter „Position • Standort"), Suchtreffer werden grün markiert, und alle Spalten inkl. eigener Abfragefelder sind sortierbar. |
| 23.32.0 | 2026-06-18 | Feature | Konsolidierte Teilnehmer-Matrix: People-Picker-Felder (z.B. „Assistenz") werden mit Foto + Name angezeigt; Personen-Spalte zweizeilig; Suche über alle Spalten mit grüner Markierung; event-spezifische Spalten sortierbar. |
| 23.31.0 | 2026-06-18 | Feature | In der QR-Code-Mail ist „DEX App" jetzt ein anklickbarer Link zur App. |
| 23.30.0 | 2026-06-18 | Feature | Organizer Center: Kachel „Warteliste" steht direkt neben „Angemeldet"; in der Event-Übersicht steht pro Event, wie viele auf der Warteliste sind. |
| 23.29.0 | 2026-06-18 | Bugfix | „No-Show" wird nur noch für neu angelegte Events angeboten; bestehende Events werden nicht mehr automatisch in ihrer Struktur verändert. |
| 23.28.0 | 2026-06-18 | Feature | Neuer Anmeldestatus „No-Show": Das Check-in-Team kann Teilnehmer als „nicht erschienen" markieren (eigener Zähler neben Eingecheckt). |
| 23.27.0 | 2026-06-17 | Bugfix | Der Teams-Chat-Link beim Organizer öffnet jetzt direkt die Teams-App statt der Browser-Zwischenseite. |
| 23.26.0 | 2026-06-17 | Feature | Mehrere Organizer werden in einer gemeinsamen Kachel nebeneinander gezeigt; E-Mail- und Teams-Chat-Link je Person; freie Plätze links neben dem Registrieren-Button; „Ich nehme nicht teil"-Button dezenter (grau + X). |
| 23.25.0 | 2026-06-17 | Feature | Bild-Darstellung pro Ansicht jetzt im „Bild editieren"-Modal gebündelt; Organizer-Karte mit klickbarer Mail + „Bei Fragen wende dich gerne an"; Option „Organizer groß anzeigen". |
| 23.24.0 | 2026-06-17 | Bugfix | Event-Karte: angepasstes Foto überdeckt nicht mehr den Event-Titel; freie Plätze nicht mehr doppelt; Banner „Anmeldung ab …" in Grün; Anmeldeseiten-Vorschau zeigt die echte Bildgröße. |
| 23.23.0 | 2026-06-17 | Feature | Event-Liste/Karte: Bildgröße einstellbar ohne dass ein Kreis-Zuschnitt abgeschnitten wird. |
| 23.22.0 | 2026-06-17 | Feature | Anmeldeseite: Größe (max. Höhe) des Event-Bildes pro Event einstellbar. |

> Ältere Versionen (vor 23.22) sind hier nicht rückwirkend erfasst; die Historie
> steht in der Git-Commit-Historie.
