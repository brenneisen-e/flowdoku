# Harness — den Wizard UND die Teilnehmer-Seiten ohne SharePoint ansehen

Seit v31.2. Nutzer-Ansage 07.09.2026: „am Ende bitte auch alle Schritte im
Wizard als PNG selber angucken." SPFx hat keinen lokalen Workbench mehr; der
Harness bündelt die echten Seiten mit gemockten Providern und macht mit
Chromium je Ansicht ein Bild.

Seit v31.7 rendert er ausser dem Wizard auch die fünf Teilnehmer-Seiten —
Landing, Start, Event-Übersicht, Anmeldeseite und „Meine Events", jeweils
einmal am Rechner und einmal am Handy. Wer eine dieser Seiten umbaut, sieht
sein Ergebnis, statt es zu behaupten.

```bash
cd dex-event-app-spfx/tools/wizard-harness
npm i --no-audit --no-fund               # einmalig, liest die package.json HIER
node build.js                            # Bundle + flaches CSS → out/
node shot.js edit                        # out/shots/edit-step-1.png … (Wizard, Beispiel-Event)
node shot.js create                      # out/shots/create-initial.png (Anlege-Modus)
node shot.js pages                       # out/shots/page-*.png (Teilnehmer-Seiten, Desktop + Handy)
```

**Immer `npm i` MIT der package.json aus diesem Ordner** (also von hier aus,
oder mit `--prefix`). Ohne sie sucht npm die nächste `package.json` weiter oben
und trägt esbuild und playwright in `dex-event-app-spfx/package.json` ein — ins
Produkt-Paket, aus dem SPFx das `.sppkg` baut. Deshalb liegt hier eine eigene,
private `package.json`.

Chromium: `shot.js` sucht unter `/opt/pw-browsers/…` (Remote-Umgebung); lokal
`npx playwright install chromium` oder den Pfad in `shot.js` anpassen.

## Eine einzelne Seite ansehen

`node shot.js pages` schießt alles. Für einen gezielten Blick reicht die
gebaute `out/index.html` mit Parametern — im Browser oder in Playwright:

| Parameter | Wirkung |
| --- | --- |
| `?mode=edit` / `?mode=create` | Event-Wizard (Vorgabe) |
| `?page=landing` | Landing Page (Orb, QR-Kasten, „Du bist angemeldet") |
| `?page=start` | Startmenü (Kacheln bzw. Zeilen) |
| `?page=list` | Event-Übersicht (Kachel-Ansicht) |
| `?page=list&view=list` | Event-Übersicht in der Listen-Ansicht |
| `?page=register` | Anmeldeseite (Beispiel-Event „Sommerfest Köln 2026") |
| `?page=register&event=ev-umb` | Anmeldeseite der Klammer — mit Termin-Auswahl |
| `?page=myevents` | Meine Events |
| `&event=<id>` | anderes Beispiel-Event auf der Anmeldeseite (`ev-open`, `ev-wait`, `ev-qr`, `ev-umb`, `ev-past`, `ev-can`) |
| `&mobile=1` | Handy-Zweige — setzt `window.__dexForceMobile` VOR dem ersten Render |

`mobile=1` allein reicht nicht für ein echtes Handy-Bild: Der Viewport muss
dazu passen, sonst laufen die CSS-Media-Queries weiter im Desktop-Zweig.
`shot.js pages` setzt beides (390 × 844, `isMobile`); `out/index.html` trägt
seit v31.7 ein `viewport`-Meta, ohne das Chromium im Handy-Modus ein 980 px
breites Layout anlegt und nur herunterskaliert.

## Was der Harness ersetzt

Siehe `entry.tsx`: `EventContext`, `RoleContext`, `NavigationContext`,
`UserContext` und (für die Startseite) `TicketContext` bekommen Proxy-Mocks;
`LanguageProvider` und `DialogProvider` laufen echt; `@microsoft/*` wird durch
einen rekursiven Proxy ersetzt; das SCSS-Modul liefert Klassennamen
unverändert. Der Dienst (`EventService`) wird nie gebaut —
`window.__dexSpfxContext` fehlt absichtlich.

Der Proxy-Fallback ist `async () => undefined`. Das trägt für alles, was nur
„gemacht" wird (Speichern, Mail, Refresh), aber **nicht** für Aufrufe, deren
Rückgabe die Seite gleich weiterverarbeitet: `undefined.length`,
`undefined.filter` und ein fehlender Aufräum-Rückruf sind der kürzeste Weg zur
weißen Seite. Alles, was ein Array, ein Objekt, ein `Set` oder eine Funktion
liefern MUSS, steht deshalb ausdrücklich in `entry.tsx` — wer eine Seite
ergänzt, ergänzt dort mit.

## Beispieldaten (`sampleData.ts`)

Bewusst nicht der Normalfall allein, sondern die Zustände, die beim Umbau
kaputtgehen:

| Event | Zustand |
| --- | --- |
| Sommerfest Köln 2026 | freie Plätze, keine eigene Anmeldung → Ziel von `?page=register`, Felder aller Typen inkl. Pflichtfeld und `showIf` |
| Kochkurs am Rhein | voll, eigene Zeile auf der **Warteliste** (Platz #2) |
| Team-Offsite Hamburg | eigene Anmeldung **„QR versendet"**, Start in zwei Tagen → QR-Kasten auf der Landing Page |
| DTP Basics Training | **Klammer-Event** (`subEventsOnlyMode`) mit drei Terminen: Tag 1 angemeldet, Tag 2 voll + Warteliste, Tag 3 frei |
| Neujahrsempfang 2026 | **vergangen**, Status „Eingecheckt" |
| Legal Update Q4 | **abgemeldet** — steht bewusst nicht in `getMyEventNumbers`, damit die Abgemeldet-Nachsuche wirklich läuft |

Die Belegung kommt aus dem gemockten Platzzähler (`getLiveCounterStats`), nicht
aus der Teilnehmerliste — genau wie live. Ohne ihn zeigten die Termin-Zeilen
„1/20", weil die zeilenweise gesicherte Liste nur die eigene Zeile hergibt.

## Die App-Hülle ansehen (Header + `.app-layout`)

`entry.tsx`/`shot.js` lassen den Rahmen bewusst weg (siehe unten). Wer die
**Hülle** selbst kartieren oder umbauen will, nimmt den zweiten Einstieg:

```bash
node shell-build.js     # → out/shell/{bundle.js,app.css,index.html}
node shell-shot.js      # → out/shots/shell-*-{1280,400}.png (+ -dom.json)
```

Gerendert wird derselbe Rahmen, den `DexEventPlatform.tsx` aufspannt:
`.dexApp > .app-layout > <Header/> + <main class="main-content"> > Seite`,
inklusive der JS-gesetzten Höhe auf `.app-layout` und der
`html, body { overflow: hidden; height: 100vh }`-Injektion. Deshalb wird hier
**nicht** `fullPage` geschossen — ein fullPage-Bild zeigte eine Seite, die es
in SharePoint nicht gibt.

| Parameter | Wirkung |
| --- | --- |
| `?view=landing` / `?view=start` | Seite unter dem Header (Vorgabe `landing`) |
| `?role=admin` / `?role=user` | echte Rolle; `user` zeigt die „Neu hier?"-Pille mittig im Header |
| `&mobile=1` | Handy-Zweige (`window.__dexForceMobile`) |

Neben jedem PNG liegt eine `*-dom.json`: die gemessene Schachtelung mit Kasten
und den tragenden CSS-Werten (`display`, `flex`, `overflow-y`, `padding`,
`grid-template-columns`, `scrollHeight` vs. `clientHeight`). Das Bild zeigt,
wie es aussieht — die JSON sagt, woran es liegt.

Die Hülle baut `shell-build.js` nach `out/shell/`, der Wizard-Harness nach
`out/`. Sie überschreiben sich nicht.

## Was der Harness NICHT zeigt

- **Kein Header, kein App-Rahmen** — im Wizard-/Seiten-Harness (`build.js` +
  `shot.js`). Gerendert wird dort nur die Seite in `.dexApp > .main-content`.
  Header, Banner und Boot-Loader fehlen; das `.app-layout` mit
  `overflow: hidden` fehlt absichtlich, sonst schnitte es die Bilder ab. Wer
  den Rahmen braucht, nimmt `shell-build.js`/`shell-shot.js` (oben).
- **Keine echten Bilder.** Event-Bilder sind erzeugte SVG-Platzhalter, das
  Mail-Logo fehlt ganz. Wer Bildzuschnitt oder Orb-Größen beurteilen will,
  braucht die echte App.
- **Keine Fluent-Icon-Font und keine Profilbilder.** Die
  `[request]`-Zeilen in den console.log nennen beides beim Namen:
  `res.cdn.office.net/…/fabric-icons-*.woff` (die Icon-Font des
  `@fluentui`-`<Icon>`, aus `file://` nicht erreichbar) und
  `/_layouts/15/userphoto.aspx` (SharePoint-Profilbilder — es gibt kein
  SharePoint). Fluent-Icons erscheinen deshalb als leere Kästchen,
  Profilbilder als Initialen-Kreis. Die eigenen SVG-Icons
  (`components/Icons`) stimmen.
- **Native Datums-/Zeitfelder zeigen die Browser-Sprache** (`mm/dd/yyyy`).
  Das ist Produktverhalten und der Grund für die Leitfaden-Regel „nie nativ" —
  kein Harness-Fehler.
- **Kein Speichern, keine Personensuche, keine Mail-Vorschau mit echten
  Vorlagen.** Es geht um Darstellung, Reihenfolge und Texte — genau das, was
  der UI-Leitfaden (`docs/ui-leitfaden.md`) regelt.

Jede Seite schreibt ihre `page-*-console.log` neben das PNG. Eine
`[pageerror]`-Zeile darin heißt: der React-Baum ist gestorben, das Bild ist
kein Beleg. `node shot.js pages` sagt es am Ende auch in einer Zeile.

`out/`, `node_modules/` und `package-lock.json` sind per `.gitignore`
ausgeschlossen.
