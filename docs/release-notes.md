# DEX Event Experience Platform — Release Notes

Pflege-Regel: siehe `CLAUDE.md` → „Release Notes — Pflicht bei JEDEM Release".
Bei **jedem** Versions-Bump kommt hier eine Zeile dazu (neueste Version **oben**),
Spalten: **Version | Datum | Art | Beschreibung**. `Art` = `Feature` (neue
Funktionalität) oder `Bugfix` (Fehlerbehebung). Beschreibung auf Deutsch,
nutzerverständlich (was sich für Organizer/Teilnehmer ändert).

| Version | Datum | Art | Beschreibung |
|---------|-------|-----|--------------|
| 23.39.0 | 2026-06-18 | Feature | Die Archivierung räumt jetzt auch alte Mails und Protokoll-Einträge von Events weg, die bereits gelöscht wurden (nicht nur von abgelaufenen). So bleiben die Hintergrund-Listen dauerhaft schlank. |
| 23.38.0 | 2026-06-18 | Feature | Der Wochenbericht geht als EINE Mail an alle Admins gemeinsam; als Empfänger stehen jetzt die echten Namen der Admins statt eines generischen „Admin". |
| 23.38.0 | 2026-06-18 | Bugfix | Der automatische Wochenbericht wird jetzt zuverlässig selbst ausgelöst. Vorher konnte ein alter Merker im Browser die Auslösung dauerhaft verhindern, sodass nie ein Bericht kam — jetzt entscheidet allein der zentrale Versand-Eintrag, ob ein neuer Bericht fällig ist. |
| 23.38.0 | 2026-06-18 | Bugfix | Ergänzung zur Events-Kennzahl: Beim Anlegen eines Sub-Events wird die „Events"-Zahl auf dem Startbildschirm nicht mehr mit hochgezählt — so bleibt sie dauerhaft korrekt (zählt nur eigenständige Veranstaltungen). |
| 23.37.0 | 2026-06-18 | Feature | Wer Organizer werden möchte, stellt jetzt einen richtigen Antrag in der App. Admins sehen offene Anträge beim Öffnen der App als Hinweis und können sie mit einem Klick freigeben — die Person wird dann automatisch zum Organizer und per Mail informiert. Die Benachrichtigungs-Mail an die Admins enthält außerdem einen Direkt-Link, der den Antrag in der App zum Bestätigen öffnet (das Freigeben geht nur als Admin). |
| 23.37.0 | 2026-06-18 | Feature | Auf dem Start-/Ladebildschirm zählt die Kennzahl „Events" jetzt nur noch eigenständige Veranstaltungen (Haupt-/Klammer-Events) und nicht mehr jedes Sub-Event einzeln mit. Die Teilnehmerzahl zählt weiterhin alle. |
| 23.36.0 | 2026-06-18 | Feature | Der wöchentliche Admin-Bericht ist jetzt viel verständlicher geschrieben und zeigt mehr: Events, die noch Entwurf sind, werden klar als solche gekennzeichnet; es gibt einen Abschnitt für Events, die seit dem letzten Bericht veröffentlicht wurden, und einen für Events, die in der Woche stattgefunden haben — inklusive Teilnehmerzahl, wobei zusammengehörige Haupt- und Sub-Events übersichtlich zusammen aufgeführt werden. Die Mail heißt jetzt „Automatischer Wochenbericht". |
| 23.36.0 | 2026-06-18 | Feature | Im Demo-Modus („als Nutzer testen") sieht die Startseite jetzt genau so aus wie für einen echten Nutzer: Die Kachel „Organizer" ist ausgegraut (mit dem Hinweis „Organizer werden?") und die Check-In-Kachel wird ausgeblendet — vorher waren beide fälschlich verfügbar. |
| 23.35.0 | 2026-06-18 | Bugfix | Der automatische Wochenbericht wurde nie verschickt — durch einen technischen Fehler blieb er im Versand hängen, obwohl die App ihn als „erledigt" vermerkte. Behoben: Der Bericht kommt jetzt wieder zuverlässig bei allen Admins an. |
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
