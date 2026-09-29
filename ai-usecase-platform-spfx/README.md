# AI Use Case Platform — Agentic Banking Demo Hub

SPFx-Webpart für die Plattform, über die Deloitte-interne Agent-Demos für den
Banking-Sektor als Kacheln aufrufbar sind. Grundlage ist das Konzept-Deck
„Agentic Banking Demo Hub" (Use-Case-Analyse → interner Hackathon →
Demo-Plattform). Die Startliste stammt aus dem KI-Arbeitsplatz
(Repo `kiarbeitsplatz`, 18 Demos in sechs Clustern).

- **Site:** `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-AIUseCasePlatform`
- **Look and feel, Aufbau und Steuerung:** wie DEX (`../dex-event-app-spfx`)
- **Stand:** v1.3.0

## Wege durch die App

```
Landing  ──Start──▶  Start-Übersicht ─┬─▶ Use Cases        (Kachelwand, Suche, Filter)
                                      │        └─▶ Detailseite  (Aufruf, Ressourcen, Verlauf)
                                      ├─▶ Use Case Studio  (Organizer, Admin: anlegen, pflegen)
                                      │        └─▶ Protokoll
                                      └─▶ Rollenverwaltung (nur Admin)
```

| Seite | Was sie kann |
|---|---|
| **Landing** | Orb, „Guten Tag, <Vorname>.", Willkommenstext, Hinweiskarte, „Start", „Entwickelt von …" — wie DEX |
| **Kopfzeile** | DE/EN, Suche über alle Use Cases, „Hast du Fragen?", Menü, Profil |
| **Start-Übersicht** | Große Kacheln je Abschnitt; die Kachel „Use Case Studio" und die Rollenverwaltung erscheinen nur, wer sie nutzen darf |
| **Use Cases** | Kachelwand mit Suche (bindestrich- und umlautunabhängig), Bereichsfilter, „nur aufrufbare" |
| **Detailseite** | Start-Knopf (neues Fenster oder eingebettet), die fünf Ressourcen, Betreuer (mailto), Bewertung, „Link kopieren" (Deep-Link `&uc=<id>`), Verlauf |
| **Use Case Studio** | Anlegen, ändern, archivieren, löschen; Kachelbild-Upload mit 16:9-Zuschnitt; Betreuer-Editor; Startbestand mit einem Klick |
| **Protokoll** | Wer hat wann was geändert — bei Änderungen mit „Feld: alt → neu"; die letzten 500 Einträge, auf einen Use Case einschränkbar |
| **Rollenverwaltung** | Personensuche, Rolle vergeben/ändern/entziehen, „Rechte prüfen", „Fehlende Rechte nachsetzen", „Überzählige Rechte entziehen", Matrix „Rechte je Rolle" |

## Rollen

| Rolle in der Oberfläche | Gespeichert als | Darf |
|---|---|---|
| **Admin** | `Admin` | alles, dazu Rollenverwaltung; Full Control auf den drei Listen |
| **Use Case Organizer** | `Kurator` | Use Cases anlegen und pflegen, Protokoll lesen; Edit auf `AIUC_UseCases`, Contribute auf `AIUC_Log`, Read auf `AIUC_Roles` |
| **User** | `User` (oder keine Zeile) | Use Cases ansehen und aufrufen |

**Der Speicherwert heißt weiter `Kurator`.** Die Rolle hieß in v1.0–v1.2 so;
bestehende Zeilen in `AIUC_Roles` bleiben gültig, es gibt keine Datenmigration.
`utils/rollen.ts` übersetzt in beide Richtungen (`rolleAusSpeicher`,
`rolleFuerSpeicher`). Wer die Spalte in SharePoint von Hand füllt, schreibt
`Kurator`, nicht `Organizer`.

## Was schon da ist

| Bereich | Stand |
|---|---|
| SPFx-Gerüst, Build, Paket | steht — `gulp bundle --ship` läuft ohne Warnung, `.sppkg` wird erzeugt |
| Design-System | `components/dexUi.ts`, `Modal.tsx`, `Icons.tsx` und das SCSS-Modul sind aus DEX übernommen |
| Listen | `AIUC_UseCases`, `AIUC_Roles`, `AIUC_Log` — werden beim ersten Start selbst angelegt |
| Kachelbild | Upload als Item-Attachment mit Zuschnitt; erst hochladen, dann Zeile speichern, das alte Bild zuletzt entfernen |
| Deep-Link | `…AIUseCases.aspx?env=WebView&uc=<id>` öffnet direkt die Detailseite |
| Protokoll | Anzeige, Rollen-Ereignisse und Änderungsdiff |
| Rollenverwaltung | Rechte werden nachgelesen, Entzug spiegelbildlich, Selbstschutz, Prüfung aller Zeilen gegen die Rechte |

## Installation (Checkliste)

- [ ] `.sppkg` in den App-Katalog hochladen (`dist/ai-usecase-platform.sppkg`)
- [ ] Auf der Site die Seite `SitePages/AIUseCases.aspx` anlegen und das Webpart einfügen (Vollbreite-Abschnitt empfohlen). Der Deep-Link zeigt auf genau diese Adresse — `APP_URL` in `constants.ts`
- [ ] Als **Admin** zum ersten Mal öffnen: Die drei Listen werden angelegt und der Startbestand (18 Demos) wird einmalig eingespielt — nur bei leerem Katalog und nur, wenn er noch nie befüllt wurde. Ist er später leer, geht es im Use Case Studio mit einem Klick
- [ ] Rollenverwaltung öffnen, weitere Admins und Use Case Organizer eintragen
- [ ] **Bestehende Installation (v1.0–v1.2):** Rollenverwaltung → „Rechte prüfen" → „Fehlende Rechte nachsetzen". Admins bekommen ab v1.3 Full Control auf den drei Listen, und die Rollenliste kappt ihre Vererbung erst, wenn die Person Listen verwalten darf
- [ ] Mit einer Person testen, die **kein** Admin ist. Admins haben Full Control und bemerken fehlende Rechte nie

**Vererbung.** `AIUC_UseCases` und `AIUC_Log` erben zunächst von der Site.
Beim ersten Recht für Organizer bekommen sie eigene Rechte; dabei bleibt die
bisherige Zuweisung bestehen. Wer die Listen später von Hand wieder auf
„Vererbung" zurückstellt, verliert die Einzelrechte — „Rechte prüfen" zeigt das.

## Bauen

```bash
npm install --no-audit --no-fund      # einmalig, ~2 min
./node_modules/.bin/tsc --noEmit -p tsconfig.json
rm -rf dist release temp sharepoint/solution/debug lib
./node_modules/.bin/gulp bundle --ship        # Exit 1 nur bei Lint-Warnungen
./node_modules/.bin/gulp package-solution --ship
cp sharepoint/solution/ai-usecase-platform.sppkg ../dist/ai-usecase-platform.sppkg
```

Nie `npx tsc` — das zieht einen fremden, neueren Compiler und meldet Dinge,
die es im Projekt nicht gibt. Immer den Projekt-Compiler. Ziel ist ES5 mit
`lib: es5, dom, es2015.*` — also kein `Array.includes`, kein `finally`, kein
`Object.values`, keine `for…of`-Schleife über `Set`/`Map`.

`react-hooks/rules-of-hooks` steht auf `error`, und ESLint hält jede Funktion
mit Präfix `use` für einen Hook. Ein Link-Helfer heißt deshalb `linkZumUseCase`,
nicht `useCaseLink`.

## Release

Version an **drei** Stellen: `package.json`,
`config/package-solution.json` (2×, mit `.0` am Ende) und
`src/webparts/aiUseCasePlatform/version.ts`. Ein Release = ein Build: erst die
Version schreiben, dann bauen.

## Screenshot-Harness

`tools/harness/` rendert die echte App gegen ein SharePoint im Speicher und
lässt Chromium alle Seiten in allen Rollen, in beiden Sprachen und auf dem
Handy durchklicken — rund 300 Bilder, dazu ein Bericht mit Befunden
(Überlauf, abgeschnittene Texte, Konsolenfehler). Aufruf und Hinweise stehen
in `tools/harness/README.md`. Nach jedem Umbau an einer Seite einmal laufen
lassen und die Bilder ansehen; `tsc` sieht keine Layoutfehler.

## Die Entscheidungen, die man kennen muss

**Die Site-Adresse steht an EINER Stelle** — `constants.ts`. In DEX steht sie
an neun Stellen in fünf Dateien, und ein Umzug wäre dort eine Suche statt einer
Änderung. Wer eine zehnte Stelle braucht, importiert von dort.

**Ein Lesefehler ist keine Null.** `getRoles`, `getUseCases` und `getLog`
liefern bei einem Fehler `null`, nie `[]`. Die Oberfläche unterscheidet „lädt",
„ok" und „nicht lesbar" und benennt den dritten Fall — eine leere Kachelwand
und ein 403 sehen sonst gleich aus, und die Kachelwand ist die ganze App.
Dasselbe gilt für Aktionen: Wer die Rollen nicht lesen kann, bekommt keine
Schaltfläche „Neuer Use Case" und keine Erstbefüllung.

**Eine Rechtevergabe gilt erst, wenn sie nachgelesen wurde.** `grantVerified`
wiederholt (1,5 s / 4 s) und liest danach `roledefinitionbindings`. Ein
`addroleassignment`-POST allein ist keine Vergabe — in DEX fehlten dadurch
18 von 126 Rollen-Einträgen mindestens ein Recht, und jede Zuweisung hatte
„erfolgreich" gemeldet.

**Wer Rechte vergibt, baut den Entzug im selben Commit.** `revokeVerified`
liest nach; SharePoint antwortet auf einen DELETE ohne vorhandene Zuweisung
mit 404 *oder* 500, der Status allein trägt also nicht. Für das eigene Konto
läuft nie ein Entzug (Selbstschutz), und eine Rolle wird nie stillschweigend
herabgestuft: Der Entzug der Rechte folgt der Rollenänderung, nicht umgekehrt.

**Eine Rolle wirkt nur mit Leserecht auf der Rollenliste.** Deshalb setzt jede
Zuweisung Read nach, und die Oberfläche sagt es, wenn das nicht geklappt hat.

**Erst anlegen, dann löschen.** Beim Speichern eines Use Cases wird das neue
Bild zuerst hochgeladen, dann die Zeile geschrieben und erst danach das alte
Bild entfernt. Scheitert ein Schritt, bleibt der alte Stand. Löschen geht in den
Papierkorb (`/recycle`), nie per hartem DELETE.

**Demo-Aufruf: neues Fenster ist die Vorgabe, eingebettet die Ausnahme.**
Permissions Policy reicht Rechte nur abwärts durch — was das SharePoint-WebView
nicht hat, kann keine eingebettete Seite gewinnen. In DEX sind daran fünf
Anläufe für den Kamera-Scan gescheitert. Dazu verbieten viele Ziele das
Einbetten selbst per `X-Frame-Options`, und das merkt man erst im Kundentermin.
Je Use Case wählbar (Nutzer-Entscheidung 10.09.2026). Der KI-Arbeitsplatz steht
hinter einem Passwort; alle Start-Use-Cases sind deshalb `fenster`.

**„Live" ohne Deployment-Link wird beim Speichern abgefangen** — sonst steht
eine Kachel da, die nichts tut.

**Die Startliste läuft nur bei leerem Katalog — und nur für Organizer und Admins.**
Sie überschreibt nie einen gepflegten Bestand und läuft nicht beim ersten
Aufruf einer beliebigen Person (zwei gleichzeitige Erstaufrufe hätten sonst
doppelte Kacheln erzeugt). Wer eine bestehende Installation hat, bekommt die
Textkorrekturen der Startliste (v1.3: 43 belegte Änderungen gegen den aktuellen
Stand des Arbeitsplatzes) deshalb NICHT automatisch — die Einträge im Studio
öffnen und ändern, oder die Liste leeren und neu einspielen.

**Kein Microsoft Graph.** `webApiPermissionRequests` ist leer. Die Personensuche
nutzt den People Picker von SharePoint, das Profilfoto die SharePoint-eigene
Adresse. Eine Graph-Berechtigung müsste ein Tenant-Admin freigeben und würde
das Paket ohne Nutzen blockieren.

## Offen

- **Betreuer je Use Case** werden angezeigt (mit Mailto), geben aber keine
  Rechte. Ob ein Team seinen eigenen Eintrag pflegen darf, ist offen — das
  bräuchte Item-Level-Rechte, und dort gilt: Contribute reicht bei
  Zeilensicherheit NIE für fremde Zeilen.
- **Kein Ticketsystem.** „Organizer werden?" und „Hast du Fragen?" öffnen eine
  Mail an `KONTAKT_EMAIL`. Ein Gruppenpostfach wäre besser als eine Person —
  dann ist es eine Zeile in `constants.ts`.
- **Nicht im echten Tenant geprüft:** Rechtevergabe und Entzug auf den Listen,
  Verhalten unter Drosselung (429), `addroleassignment` auf Listen, die erst
  ihre Vererbung kappen müssen. Der Harness bildet SharePoint im Speicher nach
  und findet Layout- und Logikfehler, keine Tenant-Eigenheiten. Vor dem ersten
  Einsatz mit einer Testperson durchspielen.
