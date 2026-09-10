# AI Use Case Platform — Agentic Banking Demo Hub

SPFx-Webpart für die Plattform, über die Deloitte-interne Agent-Demos für den
Banking-Sektor als Kacheln aufrufbar sind. Grundlage ist das Konzept-Deck
„Agentic Banking Demo Hub" (Use-Case-Analyse → interner Hackathon →
Demo-Plattform).

- **Site:** `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-AIUseCasePlatform`
- **Look and feel, Aufbau und Steuerung:** wie DEX (`../dex-event-app-spfx`)

## Was schon da ist

| Bereich | Stand |
|---|---|
| SPFx-Gerüst, Build, Paket | steht — `gulp bundle --ship` läuft ohne Warnung, `.sppkg` wird erzeugt |
| Design-System | `components/dexUi.ts`, `Modal.tsx`, `Icons.tsx` und das SCSS-Modul sind aus DEX übernommen |
| Rollen | Admin · Kurator · User, mit geprüfter Rechtevergabe und spiegelbildlichem Entzug |
| Listen | `AIUC_UseCases`, `AIUC_Roles`, `AIUC_Log` — werden beim ersten Start selbst angelegt |
| Kachelwand | Suche (bindestrich- und umlautunabhängig), Bereichsfilter, „nur aufrufbare" |
| Detailseite | Start-Knopf, die fünf Ressourcen, Bewertung, eingebettet oder neues Fenster |
| Pflege | Anlegen, ändern, löschen; Startbestand aus dem Konzept mit einem Klick |
| Rollenverwaltung | Vergeben, ändern, entziehen — mit Meldung, wenn ein Recht NICHT gesetzt wurde |

## Bauen

```bash
npm install --no-audit --no-fund      # einmalig, ~2 min
./node_modules/.bin/tsc --noEmit -p tsconfig.json
./node_modules/.bin/gulp bundle --ship
./node_modules/.bin/gulp package-solution --ship
# Ergebnis: sharepoint/solution/ai-usecase-platform.sppkg
```

Nie `npx tsc` — das zieht einen fremden, neueren Compiler und meldet Dinge,
die es im Projekt nicht gibt. Immer den Projekt-Compiler.

## Release

Version an **drei** Stellen: `package.json`,
`config/package-solution.json` (2×, mit `.0` am Ende) und
`src/webparts/aiUseCasePlatform/version.ts`.

## Die Entscheidungen, die man kennen muss

**Die Site-Adresse steht an EINER Stelle** — `constants.ts`. In DEX steht sie
an neun Stellen in fünf Dateien, und ein Umzug wäre dort eine Suche statt einer
Änderung. Wer eine zehnte Stelle braucht, importiert von dort.

**Ein Lesefehler ist keine Null.** `getRoles` und `getUseCases` liefern bei
einem Fehler `null`, nie `[]`. Die Oberfläche unterscheidet „lädt", „ok" und
„nicht lesbar" und benennt den dritten Fall — eine leere Kachelwand und ein
403 sehen sonst gleich aus, und die Kachelwand ist die ganze App.

**Eine Rechtevergabe gilt erst, wenn sie nachgelesen wurde.** `grantVerified`
wiederholt (1,5 s / 4 s) und liest danach `roledefinitionbindings`. Ein
`addroleassignment`-POST allein ist keine Vergabe — in DEX fehlten dadurch
18 von 126 Rollen-Einträgen mindestens ein Recht, und jede Zuweisung hatte
„erfolgreich" gemeldet.

**Wer Rechte vergibt, baut den Entzug im selben Commit.** `revokeAllAccess`
liest nach; SharePoint antwortet auf einen DELETE ohne vorhandene Zuweisung
mit 404 *oder* 500, der Status allein trägt also nicht.

**Eine Rolle wirkt nur mit Leserecht auf der Rollenliste.** Deshalb setzt jede
Zuweisung Read nach, und die Oberfläche sagt es, wenn das nicht geklappt hat.

**Demo-Aufruf: neues Fenster ist die Vorgabe, eingebettet die Ausnahme.**
Permissions Policy reicht Rechte nur abwärts durch — was das SharePoint-WebView
nicht hat, kann keine eingebettete Seite gewinnen. In DEX sind daran fünf
Anläufe für den Kamera-Scan gescheitert. Dazu verbieten viele Ziele das
Einbetten selbst per `X-Frame-Options`, und das merkt man erst im Kundentermin.
Je Use Case wählbar (Nutzer-Entscheidung 10.09.2026).

**„Live" ohne Deployment-Link wird beim Speichern abgefangen** — sonst steht
eine Kachel da, die nichts tut.

## Offen

- **Kachelbilder** liegen als URL vor, nicht als Anhang. Ein Upload wie in DEX
  (Item-Attachment plus Zuschnitt) fehlt noch.
- **Deep-Link auf einen Use Case** (`?uc=<id>`) ist im NavigationContext
  vorgesehen, aber noch nicht gelesen.
- **Betreuer je Use Case** werden angezeigt, geben aber keine Rechte. Ob ein
  Team seinen eigenen Eintrag pflegen darf, ist offen — das bräuchte
  Item-Level-Rechte, und dort gilt: Contribute reicht bei Zeilensicherheit NIE
  für fremde Zeilen.
- **Protokoll-Liste** wird geschrieben, aber noch nirgends angezeigt.
