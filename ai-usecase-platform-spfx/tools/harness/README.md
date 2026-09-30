# Screenshot-Harness — die AI Use Case Platform ohne SharePoint ansehen

Vorbild ist `dex-event-app-spfx/tools/wizard-harness`. SPFx hat keinen lokalen
Workbench mehr, und die Plattform braucht den Tenant. Der Harness bündelt die
**echte** `AiUseCasePlatform` mit ihren echten Providern und einem gefälschten
`WebPartContext`, dessen `spHttpClient` ein **SharePoint im Speicher** ist.
Chromium klickt sich dann durch die Seiten wie ein Mensch und macht je Ansicht
ein PNG.

```bash
cd ai-usecase-platform-spfx/tools/harness
npm i --no-audit --no-fund     # einmalig — liest die package.json HIER
node build.js                  # Bundle + flaches CSS → out/
node shot.js                   # alle Bilder → out/shots/, Bericht → out/report.json
node serve.js                  # zum Selbstansehen: http://127.0.0.1:5173/?role=admin
```

**Immer `npm i` in diesem Ordner.** Ohne die eigene `package.json` sucht npm die
nächste weiter oben und trägt `esbuild` und `playwright` in
`ai-usecase-platform-spfx/package.json` ein — ins Produkt-Paket, aus dem SPFx
das `.sppkg` baut. Danach `git status` im Repo-Wurzelordner: `package.json` und
`package-lock.json` des Produkts dürfen nicht dabeistehen.

Kein `gulp`. Der Produkt-Build läuft im selben Verzeichnis und darf nicht
gestört werden; das Bundle des Harness geht nach `tools/harness/out/`.
`tsconfig.json` des Produkts nimmt mit `include: src/**` nichts aus `tools/`
auf. `build.js` prüft keine Typen (esbuild entfernt sie nur) — das macht der
Produkt-Build.

Chromium: `shot.js` sucht unter `/opt/pw-browsers/…` (Remote-Umgebung) und
nimmt sonst das von Playwright installierte; `PLAYWRIGHT_CHROMIUM=<Pfad>`
überschreibt.

## `sicher-test.js` — die Bereinigung von Listenwerten

`node sicher-test.js` bündelt `utils/sicher.ts` und wirft Angriffsbeispiele
(`<img onerror>`, `javascript:`-Links mit Tabulator und Großschreibung,
`mailto:`-Anhängsel, SVG-`onload`, `<form>`) durch `bereinigeHtml`, `sichereUrl`
und `sichereMail` — im echten Chromium, mit Ausführung im DOM —, dazu die
Suche (`ue`/`oe`/`ae` findet Umlaute, und die Faltung verliert nichts). Exit 1,
wenn ein Fall durchrutscht. Das Produkt hat keine Tests; diese Datei ist der Ersatz
für genau die Stelle, an der ein Fehler ein gespeichertes XSS wäre.

## Was `shot.js` tut

`node shot.js` läuft etwa 6 Minuten und macht rund 300 Bilder (davon rund 50 `-voll`), dazu `out/shots/index.html` — eine Galerie mit allen Bildern, ihrem Satz und den Befunden je Lauf — und `out/report.json` mit den Befunden nach Problem gruppiert (`gruppiert`).

| Aufruf | Wirkung |
| --- | --- |
| `node shot.js` | alles |
| `node shot.js role=user` | nur diese Rollen (`admin`, `organizer`, `user`; Komma-Liste) |
| `node shot.js device=mobile` | nur Handy (390 × 844, `isMobile`, 2-fach) bzw. `desktop` (1280 × 900) |
| `node shot.js lang=en` | die Reisen in englischer Oberfläche (Dateien enden auf `-en`) |
| `node shot.js states=0` | ohne Fehler- und Leerzustände |
| `node shot.js only=admin-desktop-reise` | nur Läufe, deren Name das enthält (räumt `out/shots` dann nicht auf) |

Je Rolle und Gerät laufen drei Läufe, jeder in einem eigenen Browser-Kontext
(leerer `localStorage`, eigene Konsole):

- **`-reise`** — die lange Reise per Klick: Landing → „Start" → Start-Übersicht →
  Kachel „Use Cases" → Filter → Karte → Detail (Live, In Arbeit, Live ohne Link,
  eingebettet, archiviert) → Zurück. Organizer und Admin weiter: Use Case
  Studio (Liste, Dialog „Neuer Use Case", **ausfüllen und speichern**, Bearbeiten,
  **löschen**), Protokoll (mit den eben geschriebenen Einträgen), „Verlauf",
  „User-Ansicht". Admin: Rollenverwaltung (Personensuche, **Rolle vergeben**,
  Entfernen-Dialog, „Rechte prüfen", „Fehlende Rechte nachsetzen",
  „Überzählige Rechte entziehen", Matrix „Rechte je Rolle"). User: „Organizer
  werden?".
- **`-kopf`** — Kopfzeile und Dialoge: Suche (Ergebnisfenster; auf dem Handy die
  Suche der Kachelwand), Burger-Menü, Profil-Popup, „Hast du Fragen?", „Über
  die App", Sprachwechsel.
- **`-deeplink`** — `?uc=3` (startet direkt auf der Detailseite; beide
  „Zurück"-Wege), unbekannte Nummer, archivierter Use Case, ungültige Id.
- **`-bild`** (Admin am Rechner, Organizer am Handy) — Kachelbild: Datei wählen,
  Dialog „Bild zuschneiden", „Übernehmen", Speichern (Anhang an der Zeile; die
  Attrappe prüft, dass Zeile UND Anhang stimmen und die Adresse ein Bild
  liefert), Kachel auf der Wand, „Bild entfernen".

Dazu die **Zustände** (`zustand-…`): 403 auf die Use-Case-Liste, HTTP 500,
leer (mit automatischer Erstbefüllung), leer nach absichtlichem Löschen, Liste
gefüllt aber Protokoll leer, Protokoll 500, **Erstinstallation ohne jede Liste**,
Rollenliste nicht lesbar (403), Ladezustand, Erstinstallation durch leere Rollenliste.

Bilder: `out/shots/<rolle>-<seite>-<desktop|mobile>.png`. Ein `…-voll.png` ist
dieselbe Seite mit so hohem Fenster, dass sie ganz draufpasst (die App-Hülle
folgt der Fensterhöhe; ein `fullPage`-Bild zeigte eine Seite, die es in
SharePoint nicht gibt). Scheitert ein Schritt, liegt ein `…-FEHLER-…png` vom
Stand dort und der Befund steht im Bericht.

**`out/report.json`** ist der Bericht: je Lauf die Bilder mit einem Satz dazu,
die Befunde (Seite, Problem, Evidenz), die Schritte, alle `console.error`,
`console.warn`, `pageerror`, gescheiterte Anfragen und REST-Aufrufe, die die
Attrappe nicht kannte. Daneben liegt je Lauf eine `<lauf>-console.log`.

**Geprüft wird nach jedem Bild:** weißer Bildschirm (`.app-layout` fehlt =
React-Baum abgestürzt), praktisch leerer Inhalt, horizontaler Überlauf (auf dem
Handy der häufigste Layoutfehler), Zeilentitel/-untertitel als `inline` (die
Texte laufen ineinander), Knöpfe im Browser-Standard-Aussehen (ohne Stil) und
— am Rechner, über alle Seiten — eine Inhaltsspalte, deren Breite vom Inhalt
abhängt. Ein `pageerror` ist immer ein Befund; ein `console.error` im
Normalzustand auch. In den Fehlerzuständen sind Konsolenzeilen der App
erwartbar und stehen nur im Bericht; dafür prüfen die Zustandsläufe, dass
eine Seite einen **Lesefehler nicht als „leer", „0" oder „nicht gefunden"**
darstellt (Kachelwand, Deep-Link, Studio, Protokoll). Erwartungen kommen aus
den Daten der Attrappe (Zahl der Kacheln je Rolle, welche Use Cases „In
Arbeit"/archiviert sind), nicht aus festverdrahteten Titeln.

## Adressparameter der gebauten Seite

Für einen gezielten Blick (im Browser über `serve.js` oder in Playwright):

| Parameter | Wirkung |
| --- | --- |
| `role=admin\|organizer\|user\|first` | Rolle der Person. Organizer steht in der Liste als **`Kurator`** (`utils/rollen.ts`). `first` = Rollenliste leer → die App macht die Person zum Admin |
| `lang=de\|en` | Sprache (setzt `aiuc_locale`) |
| `mobile=1` | Handy-Zweige auch in einem breiten Fenster (überschreibt `matchMedia` für die App-Abfrage); ein Handy-**Bild** braucht zusätzlich ein schmales Fenster |
| `state=ok` | Normalzustand |
| `state=forbidden` / `error` | Use-Case-Liste antwortet 403 / 500 (Lesen und Schreiben) |
| `state=empty` | Liste UND Protokoll leer → die App befüllt sich selbst (Produktverhalten); als User scheitert das am Recht |
| `state=leer` | Liste leer, Protokoll trägt den Erstbefüllungs-Merker → der Leerzustand „Noch keine Use Cases" |
| `state=nolog` | Liste gefüllt, Protokoll leer |
| `state=logerror` | nur das Protokoll antwortet 500 |
| `state=roles403` | die Rollenliste antwortet 403 (Person ist „User", obwohl sie eingetragen ist) |
| `state=fresh` | keine der drei Listen existiert — die App legt Listen, Spalten und Rechte an |
| `delay=<ms>` | künstliche Antwortzeit je Anfrage (Vorgabe 25) |
| `data=variety\|start` | Startdaten mit Abweichungen (Vorgabe) oder unverändert |
| `uc=<id>` | Deep-Link — liest die App selbst |

`data=variety` legt über die 18 Start-Use-Cases (`data/startUseCases.ts`) ein
paar Abweichungen, damit die Zustände sichtbar werden, die beim Umbau kaputtgehen:
Nr. 5 „In Arbeit", Nr. 9 „Geplant", Nr. 12 auf „Live" ohne Deployment-Link,
Nr. 14 archiviert, Nr. 2 „eingebettet", fünf Kacheln mit Bild.

## Was der Harness ersetzt

`fakeSharePoint.ts` beantwortet **genau** die REST-Aufrufe, die die App stellt:
Listen anlegen/prüfen, Spalten, `ListItemEntityTypeFullName`, Zeilen lesen,
anlegen, `MERGE`, `recycle`, Anhänge, `roleassignments`
(`breakroleinheritance`, `addroleassignment`, `removeroleassignment`,
`roledefinitionbindings`), `web/effectivebasepermissions` (Bit „Manage Lists"),
`sitegroups/getbyid(n)/users` (Owners), `ensureuser`, People-Picker,
`GetMyProperties`. Wächst die REST-Fläche der App, meldet der Bericht den neuen
Aufruf als „den die Attrappe nicht kennt" — dann in `fakeSharePoint.ts`
ergänzen (so kamen die beiden letzten dazu).
Es ist nicht nur Anzeige: Was die App schreibt, liest sie danach wieder, und
`shot.js` prüft am Speicher nach (Zeile da? Protokolleintrag? Rechte auf den
drei Listen?). Mit Absicht streng wie SharePoint:

- Schreiben in eine **Spalte, die es nicht gibt**, ist HTTP 400 — das hat die
  Plattform am 10.09.2026 leer gemacht.
- Ein falscher `__metadata.type` ist HTTP 400 (`_x005f_` im Listennamen).
- `$filter` auf eine fehlende Spalte ist HTTP 400, nicht „0 Treffer".
- `addroleassignment` auf einer Liste, die noch von der Site erbt, scheitert.
- Ein Entzug ohne vorhandene Zuweisung antwortet 404, das Nachlesen 404.
- **Was die Attrappe nicht kennt, antwortet 404 und steht im Bericht** unter
  „REST-Aufruf, den die Attrappe nicht kennt". Ein stilles 200 würde eine neue
  Abfrage der App als „läuft" durchwinken.

Die Rollenliste hat sieben Personen — darunter eine mit **fehlenden Rechten**,
eine **User mit Rest-Rechten**, eine unter einer **Alias-Adresse** (`j.klein@`
gegen das Konto `julia.klein@`) und eine **ausgeschiedene** (nicht auflösbar) —,
damit „Rechte prüfen" etwas zu berichten hat.

Ersetzt sind sonst nur: das SCSS-Modul (Proxy, der Klassennamen unverändert
zurückgibt; das SCSS wird flach zu `out/app.css`, `:global` danach entfernt),
`@microsoft/sp-http` (echte Attrappe: `SPHttpClient.configurations.v1` ist ein
**Wert**, den ein Proxy nicht trägt) und alle übrigen `@microsoft/*` (rekursiver
Proxy — die App importiert daraus nur Typen). Die Uhr steht fest auf
29.09.2026 13:00 Berliner Zeit („Guten Tag"), das Profilbild liefert
`shot.js` selbst (`userphoto.aspx`); `user-desktop-ohnefoto` zeigt die Initialen.

## Was der Harness NICHT prüft

- **Echte Rechte.** Es gibt eine grobe Rollen-Attrappe (User schreibt nichts,
  Organizer schreibt Use Cases und Protokoll, Admin alles → sonst 403) — nicht
  die Vererbung, Gruppen oder die Wirkung von Full Control/Edit in SharePoint.
  „Der Organizer sieht die Kachel, aber SharePoint verweigert das Speichern"
  bleibt unsichtbar.
- **Drosselung** (HTTP 429/503, `Retry-After`), Paging (`nextLink`), langsame
  Leitungen. Der Ladezustand kommt nur aus `delay`.
- **Microsoft Graph**, echtes Benutzerprofil, echte Personensuche
  (der Picker filtert auf `@deloitte.de`, das Verzeichnis hat neun Namen).
- **Die eingebettete Demo** (`aufrufArt: eingebettet`): Der Knopf „Demo hier
  öffnen" wird gezeigt, aber nicht geklickt — das iframe würde ins Netz laden.
  Die Härtung (`sandbox`, Schema-Prüfung, keine gleiche Herkunft) ist deshalb
  nur im Code geprüft, nicht im Bild.
  „Demo starten" öffnet einen neuen Tab; auch das wird nicht geklickt.
- **Kachelbild: echte Kamera- und Datei-Auswahl.** Der Lauf `-bild` wählt eine
  Beispieldatei, schneidet zu und speichert (Attrappe nimmt den Anhang; über
  `serve.js` liegt das Bild danach unter der Adresse, die die App merkt) — die
  Dateiauswahl des Betriebssystems und große Fotos vom Handy prüft er nicht.
- **Zwischenablage** („Link kopieren"), `mailto:`-Link, Tastaturbedienung.
- **Das Aussehen in SharePoint selbst**: Die Hülle wird nicht von SharePoint-
  Chrome umgeben (Kopfleiste, Seitentitel, Ränder) — `useShellHeight` rechnet
  mit `rect.top = 0`. Schriften: Aptos/Open Sans fehlen im Container, es gilt
  die Rückfallschrift (Breiten und Umbrüche können in SharePoint abweichen).
- **Der Produkt-Build.** Der Harness baut mit esbuild, ohne Typ-Prüfung und
  ohne ESLint; ein grüner Harness heißt nicht „gulp bundle läuft".

`out/`, `node_modules/` und `package-lock.json` sind per `.gitignore`
ausgeschlossen.
