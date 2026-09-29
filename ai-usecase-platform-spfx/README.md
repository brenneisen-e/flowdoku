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
                                      ├─▶ Rollenverwaltung (nur Admin)
                                      └─▶ Fragen & Feedback (Dialog, für alle)
```

| Seite | Was sie kann |
|---|---|
| **Landing** | Orb, Begrüßung mit Vorname (je nach Tageszeit „Guten Morgen", „Guten Tag" oder „Guten Abend"), Willkommenstext, Hinweiskarte, „Start", „Entwickelt von …" — wie DEX |
| **Kopfzeile** | DE/EN, Suche über alle Use Cases, „Hast du Fragen?", Menü, Profil |
| **Start-Übersicht** | Große Kacheln je Abschnitt. „Use Cases" und „Fragen & Feedback" sieht jede Person. „Use Case Studio" steht für alle da, ist aber für Nicht-Organizer ausgegraut („Organizer werden?"); „Protokoll" gibt es nur für Organizer, „Rollenverwaltung" nur für Admins |
| **Use Cases** | Kachelwand mit Suche (bindestrich-, umlaut- und `ue`/`oe`/`ae`-unabhängig), Bereichsfilter, „nur aufrufbare" |
| **Detailseite** | Start-Knopf (neues Fenster oder eingebettet) für das Deployment, die vier weiteren Ressourcen (Source Code, Guide, Wiki, Video), Betreuer (mailto), Bewertung, „Link kopieren" (Deep-Link `&uc=<id>`), „Verlauf" (nur Organizer) |
| **Use Case Studio** | Anlegen, ändern, archivieren, löschen; Kachelbild-Upload mit 16:9-Zuschnitt; Betreuer-Editor; Startbestand mit einem Klick |
| **Protokoll** | Wer hat wann was geändert — bei kurzen Feldern als „Feld: alt → neu", bei Texten und Links nur mit dem Feldnamen; die letzten 500 Einträge, auf einen Use Case einschränkbar |
| **Rollenverwaltung** | Personensuche, Rolle vergeben/ändern/entziehen, „Rechte prüfen", „Fehlende Rechte nachsetzen", „Überzählige Rechte entziehen", Matrix „Rechte je Rolle" |

## Rollen

| Rolle in der Oberfläche | Gespeichert als | Darf |
|---|---|---|
| **Admin** | `Admin` | alles, dazu Rollenverwaltung; Full Control auf den drei Listen |
| **Use Case Organizer** | `Kurator` | Use Cases anlegen und pflegen, Protokoll lesen; **Contribute** auf `AIUC_UseCases` und `AIUC_Log`, Read auf `AIUC_Roles` |
| **User** | `User` (oder keine Zeile) | Use Cases ansehen und aufrufen |

**Der Speicherwert heißt weiter `Kurator`.** Die Rolle hieß in v1.0–v1.2 so;
bestehende Zeilen in `AIUC_Roles` bleiben gültig, es gibt keine Datenmigration.
`utils/rollen.ts` übersetzt in beide Richtungen (`rolleAusSpeicher`,
`rolleFuerSpeicher`). Wer die Spalte in SharePoint von Hand füllt, schreibt
`Kurator`, nicht `Organizer`.

**Organizer bekommen Contribute, nicht Edit.** „Edit" enthält „Manage Lists":
Ein Organizer könnte Spalten oder die ganze Liste löschen. Zeilen anlegen,
ändern, recyceln und Anhänge tragen nur Add/Edit/Delete Items. Bis v1.3 stand
in der Tabelle noch Edit — bei einer bestehenden Installation zeigt „Rechte
prüfen" das als „zu viele Rechte", „Überzählige Rechte entziehen" stuft sie
zurück.

## Was schon da ist

| Bereich | Stand |
|---|---|
| SPFx-Gerüst, Build, Paket | steht — `gulp bundle --ship` läuft ohne Warnung, `.sppkg` wird erzeugt |
| Design-System | `components/dexUi.ts`, `Modal.tsx`, `Icons.tsx` und das SCSS-Modul sind aus DEX übernommen; ungenutzte DEX-Regeln sind entfernt |
| Listen | `AIUC_UseCases`, `AIUC_Roles`, `AIUC_Log` — werden beim ersten Start selbst angelegt |
| Kachelbild | Upload als Item-Attachment mit Zuschnitt; beim Ändern erst hochladen, dann Zeile speichern, das alte Bild zuletzt entfernen |
| Deep-Link | `…AIUseCases.aspx?env=WebView&uc=<id>` öffnet direkt die Detailseite |
| Protokoll | Anzeige, Rollen-Ereignisse und Änderungsdiff |
| Rollenverwaltung | Rechte werden nachgelesen, Entzug spiegelbildlich, Selbstschutz, Prüfung aller Zeilen gegen die Rechte |
| Absicherung | Links, HTML und Adressen aus der Liste gelten als nicht vertrauenswürdig (`utils/sicher.ts`), Dialoge haben Escape-Stapel und Fokus-Falle, eine Fehlergrenze fängt Render-Fehler ab |

## Installation (Checkliste)

- [ ] `.sppkg` in den App-Katalog hochladen (`dist/ai-usecase-platform.sppkg`). Das Paket hat `skipFeatureDeployment: true` — es lässt sich tenant-weit bereitstellen oder auf der Site hinzufügen
- [ ] Auf der Site die Seite `SitePages/AIUseCases.aspx` anlegen und das Webpart einfügen (Vollbreite-Abschnitt empfohlen). Der Deep-Link zeigt auf genau diese Adresse — `APP_URL` in `constants.ts`
- [ ] **Ein Site Owner öffnet die Seite als Erster.** Die erste Person, die die leere Rollenliste öffnet, wird automatisch Admin (Erstinstallation). Listen anlegen kann nur, wer „Manage Lists" auf der Site hat, Rechte setzen nur mit „Manage Permissions" — ein Mitglied mit Edit-Recht würde Admin, könnte die Listen aber nicht absichern
- [ ] Beim ersten Öffnen werden die drei Listen angelegt und der Startbestand (18 Demos) einmalig eingespielt — nur bei leerem Katalog und nur, wenn er noch nie befüllt wurde. Ist er später leer, geht es im Use Case Studio mit einem Klick („Fehlende anlegen" ergänzt nur, was fehlt)
- [ ] Rollenverwaltung öffnen, weitere Admins und Use Case Organizer eintragen
- [ ] **Bestehende Installation (v1.0–v1.2):** Rollenverwaltung → „Rechte prüfen" → „Fehlende Rechte nachsetzen" (Admins bekommen ab v1.3 Full Control auf den drei Listen) und danach „Überzählige Rechte entziehen" (Organizer: Contribute statt Edit). Im SharePoint Admin Center unter „API access" die früher freigegebenen Graph-Berechtigungen (`User.Read`, `User.Read.All`) der App entfernen — v1.3 fordert sie nicht mehr an, die Freigabe bleibt aber stehen
- [ ] Mit einer Person testen, die **kein** Admin ist. Admins haben Full Control und bemerken fehlende Rechte nie

**Vererbung.** Bei der ersten Rechtevergabe überhaupt — schon für den ersten
Admin — kappt die App die Vererbung von `AIUC_UseCases` und `AIUC_Log` (mit
Kopie: was heute gilt, gilt weiter). Was danach am Web vergeben wird — eine neue
AD-Gruppe, eine neu geteilte Site —, erreicht diese zwei Listen nicht mehr; neue
Leser müssen auf der Liste selbst eingetragen werden. `AIUC_Roles` bekommt eigene
Rechte **ohne** Kopie (Owners und die anlegende Person, danach die vergebenen
Rollen). Wer die Listen später von Hand wieder auf „Vererbung" zurückstellt,
verliert die Einzelrechte — „Rechte prüfen" zeigt das.

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

**Das Stylesheet in `dexUi.ts` und `Modal.tsx` ist ein Template-Literal.** Ein
Backtick in einem CSS-Kommentar dort beendet den String, und `tsc` meldet
„',' expected" mitten in der Datei. Klassennamen in Kommentaren ohne Backticks
schreiben.

## Release

Version an **drei** Stellen: `package.json`,
`config/package-solution.json` (2×, mit `.0` am Ende) und
`src/webparts/aiUseCasePlatform/version.ts`. Ein Release = ein Build: erst die
Version schreiben, dann bauen.

## Screenshot-Harness

`tools/harness/` rendert die echte App gegen ein SharePoint im Speicher und
lässt Chromium alle Seiten in allen Rollen und auf dem Handy durchklicken —
rund 300 Bilder, dazu ein Bericht mit Befunden (Überlauf, abgeschnittene Texte,
Konsolenfehler). Deutsch ist die Vorgabe; `node shot.js lang=en` fährt die
Reisen auf Englisch. Aufruf und Hinweise stehen in `tools/harness/README.md`.
Nach jedem Umbau an einer Seite einmal laufen lassen und die Bilder ansehen;
`tsc` sieht keine Layoutfehler. In v1.3 hat genau das vier davon gefunden
(zusammenklebende Status-Pillen, ein Dialogfuß mit 22 px Formular darunter,
ein Hinweis über dem Symbol, rohe Speicherwerte im Protokoll).

## Die Entscheidungen, die man kennen muss

**Die Site-Adresse steht im Code an EINER Stelle** — `constants.ts` (nur
`config/serve.json` trägt sie für den lokalen Workbench zusätzlich). In DEX steht
sie an neun Stellen in fünf Dateien, und ein Umzug wäre dort eine Suche statt
einer Änderung. Wer eine zehnte Stelle im Code braucht, importiert von dort.

**Ein Lesefehler ist keine Null.** `getRoles`, `getUseCases` und `getLog`
liefern bei einem Fehler `null`, nie `[]` (auch bei mehr als 5 000 Zeilen, wo
der Lesepfad kappt). Die Oberfläche unterscheidet „lädt", „ok" und „nicht
lesbar" und benennt den dritten Fall — eine leere Kachelwand und ein 403 sehen
sonst gleich aus, und die Kachelwand ist die ganze App. Schlägt nur das
Nachladen fehl, bleibt der letzte Stand sichtbar, mit Hinweis. Dasselbe gilt für
die Rolle: Ein echter Lesefehler zeigt „Rolle: unbekannt", nicht „User". Wer die
Rollen nicht lesen kann, bekommt keine Schaltfläche „Neuer Use Case" und keine
Erstbefüllung.

**Eine Rechtevergabe gilt erst, wenn sie nachgelesen wurde.** `grantVerified`
wiederholt (1,5 s / 4 s) und liest danach `roledefinitionbindings`. Ein
`addroleassignment`-POST allein ist keine Vergabe — in DEX fehlten dadurch
18 von 126 Rollen-Einträgen mindestens ein Recht, und jede Zuweisung hatte
„erfolgreich" gemeldet. Das gilt auch dort, wo die App die Rollenliste sperrt:
Erst wird die anlegende Person mit Full Control eingetragen und nachgelesen,
sonst holt die App die Vererbung zurück, statt eine Liste ohne Zugriff zu
hinterlassen.

**Wer Rechte vergibt, baut den Entzug im selben Commit.** `entzieheListe` und
`reduziereListe` lesen nach; SharePoint antwortet auf einen DELETE ohne
vorhandene Zuweisung mit 404 *oder* 500, der Status allein trägt also nicht.
Für das eigene Konto läuft nie ein Entzug (Selbstschutz). Beim Herabstufen und
Entfernen kommen zuerst die Rechte (nachgelesen), dann die Zeile — scheitert der
Entzug, bleibt die Zeile; bei einer Erhöhung kommt die Zeile zuerst. Rollen
vergeben und ändern dürfen nur echte Admins (auch in der Vorschau „als User
ansehen" nicht).

**Eine Rolle wirkt nur mit Leserecht auf der Rollenliste.** Deshalb setzt jede
Zuweisung Read nach, und die Oberfläche sagt es, wenn das nicht geklappt hat.

**Erst anlegen, dann löschen.** Beim Ändern eines Use Cases wird das neue Bild
zuerst hochgeladen, dann die Zeile geschrieben und erst danach das alte Bild
entfernt; scheitert ein Schritt, bleibt der alte Stand. Beim Neuanlegen kommt
zuerst die Zeile (das Bild hängt an ihr) — scheitert dann nur das Bild, ist der
Use Case da und die Meldung sagt es. Geschrieben wird nur, was sich gegenüber dem
Stand beim Öffnen geändert hat: Wer nur den Titel ändert, überschreibt kein
Bild, das jemand inzwischen getauscht hat. Use Cases und ihre Bilder gehen beim
Löschen in den Papierkorb (`/recycle`), nie per hartem DELETE. Die einzige Ausnahme
sind Rollenzeilen (`deleteRole`); sie stehen im Protokoll.

**Listenwerte sind nicht vertrauenswürdig.** Die Liste ist in SharePoint auch
direkt beschreibbar (jede Person mit Schreibrecht), und die Beschreibung hat in
der Pflegeseite gar kein Feld. Deshalb gehen Links nur als absolute
`http(s)://`-Adressen weiter, die Beschreibung wird auf einfache Auszeichnungen
(Absatz, Fett, Liste, Link) bereinigt, Betreuer-Adressen müssen schlichte
Adressen sein (`?bcc=…` hängt in `mailto:` Empfänger an), und ein eingebetteter
Aufruf läuft in einem `sandbox`-iframe ohne Zugriff auf die Seite darüber — auf
derselben Herkunft wie SharePoint wird gar nicht eingebettet. Alles steht in
`utils/sicher.ts`; wer einen neuen Listenwert in `href`, `src`, `window.open`
oder HTML setzt, geht dort hindurch.

**Demo-Aufruf: neues Fenster ist die Vorgabe, eingebettet die Ausnahme.**
Permissions Policy reicht Rechte nur abwärts durch — was das SharePoint-WebView
nicht hat, kann keine eingebettete Seite gewinnen. In DEX sind daran fünf
Anläufe für den Kamera-Scan gescheitert. Dazu verbieten viele Ziele das
Einbetten selbst per `X-Frame-Options`, und das merkt man erst im Kundentermin.
Je Use Case wählbar (Nutzer-Entscheidung 10.09.2026). Der KI-Arbeitsplatz steht
hinter einem Passwort; alle Start-Use-Cases sind deshalb `fenster`.

**„Live" ohne Deployment-Link wird beim Speichern abgefangen** — sonst steht
eine Kachel da, die nichts tut. Dasselbe für Links, die keine gültige
`https://`-Adresse sind.

**Die Startliste hat genau EINEN Weg** (`seedStartUseCases`): automatisch beim
ersten Start eines Organizers oder Admins, per Knopf auf der Kachelwand und im
Studio. Sie läuft automatisch nur bei leerem Katalog und nur, wenn der Merker
„schon befüllt" im Protokoll fehlt; sie legt nur Titel an, die es noch nicht
gibt, und meldet einen Teilerfolg mit Zahl und Grund („Nur 7 von 18 angelegt"),
mit einem Knopf „Fehlende anlegen". Wer eine bestehende Installation hat,
bekommt die Textkorrekturen der Startliste (v1.3: 43 belegte Änderungen gegen den
aktuellen Stand des Arbeitsplatzes) deshalb NICHT automatisch — die Einträge im
Studio öffnen und ändern, oder die Liste leeren und neu einspielen.

**Start: erst die Rollen, dann die Listen.** Die Prüfung und Anlage der Listen
und Spalten (rund 24 GETs) läuft nur für Organizer und Admins; alle anderen
lesen sofort. Bis v1.3 wartete jede Person bei jedem Start darauf.

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
- **Das Protokoll ist ein Hinweis, kein Beweis.** Organizer haben Contribute auf
  `AIUC_Log` und könnten fremde Zeilen ändern oder löschen; `Wer` schreibt der
  Browser. Zeilensicherheit (`WriteSecurity=2`) und das Feld `Author` als
  Quelle wären der nächste Schritt. Die Protokolltexte stehen außerdem nur
  deutsch in der Liste.
- **Kachelbilder** dürfen auf beliebige Fremdhosts zeigen, wenn jemand eine
  Adresse von Hand einträgt (nur `http(s)` wird geprüft) — jeder Betrachter
  löst dann einen Aufruf dorthin aus. Bilder, die die App selbst hochlädt,
  liegen als Anhang auf der Site.
- **Gleichzeitige Erstaufrufe.** Öffnen zwei Organizer die leere Plattform in
  derselben Sekunde, können sie den Startbestand doppelt anlegen (es gibt keine
  serverseitige Sperre); zwei Admins, die sich im selben Moment gegenseitig
  herabstufen, lassen niemanden übrig, der es rückgängig macht — dann hilft ein
  Site Owner in SharePoint.
- **Nicht im echten Tenant geprüft:** Rechtevergabe und Entzug auf den Listen,
  Contribute (statt Edit) für Organizer beim Bildwechsel und Recyceln, das
  Kappen der Vererbung auf `AIUC_Roles`, Verhalten unter Drosselung (429),
  `addroleassignment` auf Listen, die erst ihre Vererbung kappen müssen. Der
  Harness bildet SharePoint im Speicher nach und findet Layout- und
  Logikfehler, keine Tenant-Eigenheiten. Vor dem ersten Einsatz mit einer
  Testperson durchspielen — je einmal als Organizer und als Person ohne Rolle.
