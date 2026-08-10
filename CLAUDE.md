# CLAUDE.md — DEX Event Experience Platform

Arbeitsanweisungen für Claude Code in diesem Repo. Ergänzt `ENTWICKLUNG.md`
(dort steht der ausführliche Build- und Bundle-Hintergrund).

## Projekt

**DEX** (Deloitte Event Experience Platform) — SPFx-Webpart, React + TypeScript.
Organizer legen Events an, Teilnehmer melden sich an; Mails und Outlook-Termine
laufen über Power-Automate-Flows.

- Code: `dex-event-app-spfx/src/webparts/dexEventPlatform/`
- SharePoint-Site: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- Event-Liste: `DEX_Events`, Teilnehmer je Event in einer eigenen Subsite-Liste

Die drei großen Dateien tragen fast alles: `components/EventCreationPage.tsx`
(~17k Zeilen, Wizard), `components/AdminPage.tsx` (~15k, Organizer Center),
`services/EventService.ts` (~12k, SharePoint-Zugriff).

**Branch:** wird pro Sitzung vorgegeben (zuletzt `claude/mach-claude-md-gax5yx`,
davor `claude/spfx-app-bugfixes-4kui16`) — Stand **v28.90.0**. Nur auf den
vorgegebenen Branch pushen. Keine PRs ohne ausdrückliche Aufforderung.

## Erst einrichten, dann bauen

Eine frische Sitzung hat **kein `node_modules`**. Ohne Install schlägt jeder
Befehl auf eine Art fehl, die nach einem Code-Fehler aussieht:

```bash
cd dex-event-app-spfx && npm install --no-audit --no-fund   # ~1 min, ~2300 Pakete
```

`npx tsc` zieht ohne lokale Installation ein **globales, neueres** TypeScript und
meldet Dinge, die es im Projekt nicht gibt (`webpack-env` nicht gefunden,
`target=ES5` deprecated). Deshalb immer den Projekt-Compiler nehmen:
`./node_modules/.bin/tsc --noEmit -p tsconfig.json`. Dasselbe gilt für `gulp`.

## Release-Ablauf — Reihenfolge einhalten

Ein Release = ein Build. Version und Release Notes **vor** dem Build schreiben,
sonst baut man zweimal (die Notes stecken im Bundle, sie werden in der App unter
„Was ist neu?" angezeigt).

```bash
cd dex-event-app-spfx
# 1) Version an DREI Stellen
#    package.json · config/package-solution.json (2×, mit .0 am Ende) · src/webparts/dexEventPlatform/version.ts
# 2) Release Notes an ZWEI Stellen
#    src/webparts/dexEventPlatform/data/releaseNotes.ts   (neuester Eintrag OBEN, nutzerverständlich)
#    docs/release-notes.md                                (neueste Zeile OBEN, technisch)
# 3) tsc (Projekt-Compiler, nicht npx — siehe oben)
./node_modules/.bin/tsc --noEmit -p tsconfig.json
# 4) Sauber bauen (stale Bundles blähen das .sppkg auf)
rm -rf dist release temp sharepoint/solution/debug
./node_modules/.bin/gulp bundle --ship; ./node_modules/.bin/gulp package-solution --ship
#    kein && — bundle endet wegen Lint-Warnungen mit Exit 1 (s.u.)
# 5) Paket an DREI Stellen
cp sharepoint/solution/dex-event-platform.sppkg ../dist/dex-event-platform.sppkg
#    ../docs/downloads/dex-event-platform-v<VERSION>.sppkg  (alte Datei per git mv umbenennen)
# 6) docs/index.html: Version, Download-Link, Tag, „enthält kumulativ"-Absatz
```

Erwartete Größe: ca. **2,2–2,3 MB**. Deutlich größer heißt: stale Bundles, Schritt 4
wiederholen. In `release/assets/` darf es nur **eine** `dex-event-platform-web-part_*.js` geben.

`gulp bundle` endet mit Exit-Code 1, sobald Lint-Warnungen auf stderr gehen —
das ist **kein** Fehlschlag. Es gibt Alt-Warnungen in `HotelSetupWizard.tsx`,
`HotelImportModal.tsx`, `HotelPlanningPanel.tsx`. Prüfen, ob eine Warnung aus der
eigenen Änderung stammt; sonst weiter. Häufigste eigene Warnung:
`react/no-unescaped-entities` bei `"` in JSX-Text → `&bdquo;` / `&ldquo;`.

Commits: klarer Betreff `vX.Y.Z: …`, im Body die Ursache erklären, nicht nur das
Was. Trailer `Co-Authored-By` und `Claude-Session` anhängen.

## Sprache

Oberfläche und Kommentare sind deutsch, zweisprachig über `isDe`. Neue
Kommentare auf Deutsch, im Ton der Umgebung: erklären **warum**, nicht was.
Versionsmarken wie `// v28.66: …` sind Konvention — beibehalten.

Echte Umlaute (ä/ö/ü), keine ue/oe/ae-Umschreibungen. **Ausnahme:** vier Werte
sind SharePoint-Daten und dürfen nie „korrigiert" werden —
`Nachruecken`, `OrgNachruecker`, `nachruecker`, `UeberUns`.

## Fallen, die in dieser Codebasis wehtun

**Große JSX-Blöcke nie per Textmarke verschieben.** Ein Versuch, den Innenteil
eines 500-Zeilen-Blocks zwischen zwei Textmarken zu schneiden, zerriss die
Tag-Balance (10 × TS17008/TS17015) und musste zurückgerollt werden. Entweder den
**ganzen** Block als Einheit bewegen (Anzeige-Bedingung umstellen statt
schneiden) oder die Verschachtelung vorher wirklich lesen. Vor solchen Eingriffen
eine Kopie ins Scratchpad legen.

**Wizard-Schritte hängen an festen Indizes.** `currentStep === N` in den
Anzeige-Bedingungen, dazu `steps`-Array, Schritt-Titel im Text, `STEP_HINTS_DE`/
`STEP_HINTS_EN`, `getStepErrors`-Cases, `SCOPE_AWARE_STEPS` und die Tour-Grenze
(`detail <= 7`). Wer einen Schritt hinzufügt oder entfernt, muss **alle** davon
nachziehen.

**Kommunikationsfelder der Sub-Events liegen nicht laufend im Draft.** Sie
werden erst bei `switchCommTab` in den Slot geschrieben. Wer sie liest
(Übertragen, Persistieren), muss vorher `flushActiveCommTabToState()` rufen und
aus `subEventsRef.current` lesen — sonst kopiert man den Stand vor der letzten
Bearbeitung. Objektwerte (`emailTemplateOverrides`) beim Kopieren klonen.

**Der Outlook-Flow verträgt kein leeres `EndDate`.** Er rechnet
`convertFromUtc(coalesce(OutlookEnd, EndDate))`; bei `null` bricht „Create event
(V4)" ab und der ganze Lauf scheitert — ohne Rückmeldung in der App. `EndDate`
fällt deshalb überall auf `StartDate` zurück, zentral in
`EventService.createEvent`/`updateEvent`.

**Der Flow triggert nur auf NEUE Listeneinträge** (`GetOnNewItems`). Ein MERGE
stößt ihn nie an. Fehlende Termine legt man über den nicht-destruktiven
Recreate-Pfad an: `deleteEventItemOnly` + `createEvent` mit `existingSubsiteUrl`
(Anmeldungen bleiben). Das **Hauptevent** nie so behandeln — seine Item-Id steht
in `ParentEventId` aller Kinder.

**Die Warteliste hat keine Positionsspalte.** Die Position ist der Rang nach
`TeilnehmerID asc`. Danach sortieren die App (`promoteFirstWaitlistItem`) **und**
der Flow `DEX_IDReorder_TeilnehmerIDs`. Eine eigene Prioritätsspalte würde der
Flow ignorieren — deshalb sortiert `setWaitlistPosition` die TeilnehmerIDs um.

**Piggyback-Konfiguration in `EmailTemplateOverrides`.** Flags wie
`_subEventsOnlyMode`, `_noDescription`, `_imageBanner` liegen im JSON dieser
Spalte. Wer ein neues Flag ergänzt, muss es **auch beim Laden strippen** —
sonst überschreibt der alte Wert beim Speichern den frisch berechneten.

**Inline-Styles können kein `:hover`.** Interaktive Elemente brauchen einen
Hover-State (`hoverIdx`, `evTabHover`), sonst lesen sie sich als Beschriftung.

**Der Scope-Umschalter gehört genau einmal auf die Seite.** Seit v28.78 rendert
`renderGlobalScopeBar` die Reiter global über dem Formular; die alten
`StickyTabStrip`-Instanzen je Schritt sind Altlast. In Schritt Kommunikation
stand sie bis v28.88 noch da — zwei identische Reiter-Reihen, die dasselbe
umschalten, liest der Organizer als zwei Navigationen und sucht die gültige.
Wer einen Schritt scope-fähig macht (`SCOPE_AWARE_STEPS`), hängt ihn an
`setScope` und baut **keine** eigene Leiste dazu.

**Guard-Meldungen müssen den tatsächlichen Grund nennen.** Die Anmeldeseite
sperrte mit „Bitte wähle mindestens das Haupt-Event oder ein Sub-Event aus" —
auch dort, wo es gar nichts zu wählen gibt: `willRegisterParent` ist bei
bestehender Anmeldung immer false, und ohne Sub-Events rendert die Auswahl-UI
gar nicht (sie hängt an `childEvents.length > 0`). Vor einer Sammel-Meldung
prüfen, welche der Bedingungen den Fall wirklich erzeugt hat
(`parentAlreadyRegistered`, `subEventsOnlyMode`, `parentRegBlocked`).
Und: leere Auswahl heißt nicht „nichts zu tun" — wer alle gebuchten Sub-Events
abwählt, meldet sie ab (`sessionsChanged` in `RegistrationPage`).

## Der Wizard, Stand v28.89

Neun Schritte. Über dem Formular steht die **Scope-Karte**
(`renderGlobalScopeBar`): Klammer/Haupt-Event und die Sub-Events als Reiter, ein
gemeinsamer Index (`activeScopeIdx` / `setScope`). In Schritten ohne Scope-Bezug
bleibt die Karte stehen und sagt, dass der Schritt für alles gilt.

**`setScope` ruft für die Kommunikation bewusst `switchCommTab(idx)`** statt
`setActiveCommTabIdx` — sonst gehen Sub-Event-Mailtexte verloren.

Seit v28.88 hat kein Schritt mehr eine eigene Reiter-Leiste; im
Kommunikations-Schritt steht an ihrer Stelle nur noch der Satz „Die
Einstellungen unten gelten für den oben gewählten Reiter" plus der bisherige
`InfoTooltip`.

**Schritt 1 ist seit v28.89 scope-fähig** (`SCOPE_AWARE_STEPS = [0,2,3,4,5]`).
Titel, Start, Ende, Beschreibung und Bild sind **dieselben** Eingaben und hängen
über `scopeSub`/`patchScopeSub` am Top-Level-State **oder** an
`subEvents[activeScopeIdx-1]`. Zwei Dinge, die dabei leicht kippen:

- **Zeitformate.** Top-Level ist Berliner Lokalzeit `YYYY-MM-DDTHH:MM`, ein
  Sub-Event UTC-ISO. Umgerechnet wird nur in `subIsoToDate`/`subDateToIso` —
  keine zweite Stelle aufmachen.
- **Was event-weit gilt, gehört auf die Klammer.** Opt-in, Bezeichnung,
  Anmelde-Modus, Vorlage, Entwurf/Aktivierung und die Sub-Event-Liste rendern
  nur bei `activeScopeIdx === 0`. Neue event-weite Felder in Schritt 1 müssen in
  diesen Block, sonst beantwortet man Grundsatzfragen „unter" einem Termin.
- **Pflichtfelder gehören dem Hauptevent.** `getStepErrors` prüft weiter den
  Top-Level; `proceedNext` wechselt deshalb bei Fehlern auf die Ebene, auf der
  der Fehler steht — sonst wirkt „Weiter" wie tot.

Die Sub-Event-Karten sind seither reine **Liste**: anlegen, „Bearbeiten"
(`setScope(idx+1)`), entfernen, Pflichtanmeldung. Keine Editor-Felder mehr.

## Offene Arbeit

Die vier Punkte aus v28.87 sind mit v28.88/v28.89 abgearbeitet (Feld-Ebene,
Reiter-Darstellung, Outlook-Dialog, Legacy-Rollen), v28.90 hat die Nachlese
dazu erledigt. Offen und **noch nicht begonnen**:

1. **Termin-Slots statt einzeln angelegter Sub-Events.** Der Wunsch: im
   Assistenten einen Zeitraum im Kalender markieren und daraus je Tag (oder je
   Stunde) ein Sub-Event erzeugen lassen, statt neun Karten von Hand
   anzulegen; auf der Anmeldeseite dann ein Kalender zur Auswahl statt einer
   Liste aus neun Funkbuttons. Vor dem Bauen zu klären, ob ein Slot ein
   **eigenes Sub-Event** bleibt (dann ist es reine Erzeugungs- und
   Darstellungs-Hilfe — Teilnehmerliste, Kapazität, Outlook-Termin und
   `ParentEventId` bleiben wie sie sind) oder ein **neuer Datentyp** unterhalb
   des Sub-Events wird (dann hängen Teilnehmerliste, Flows und Warteliste mit
   daran). Ersteres ist deutlich billiger und deckt den geschilderten Fall.

Bewusst **nicht** gebaut: ein Dropdown zum Springen zwischen Sub-Event-Reitern.
Es wäre eine zweite Bedienung für dieselbe Auswahl; die gescrollte Leiste hat
stattdessen Pfeile, Tastatur, Auto-Scroll zum aktiven Reiter und eine Zählung
„3 / 9". Falls die Reiter bei sehr vielen Sub-Events weiter stören, ist ein
Dropdown der nächste Kandidat — dann aber **statt** der Leiste, nicht daneben.

## Umgang

Nach dem Deploy den Wizard einmal von Schritt 1 bis 9 durchklicken — die
Umnummerierung aus v28.87 ist im Browser nicht verifiziert. Für v28.89
zusätzlich: in Schritt 1 zwischen Klammer und mehreren Sub-Events umschalten
(Titel/Zeiten/Beschreibung/Bild müssen dem Reiter folgen), ein Sub-Event über
die Liste anlegen und entfernen, und die Reiter-Leiste mit mehr als sechs
Sub-Events auf Pfeile, Zählung und Auto-Scroll ansehen.

**Bildschirmfotos zeigen den installierten Stand, nicht den Repo-Stand.** Ein
Screenshot mit zehn Wizard-Schritten kam aus einem Build vor v28.87; wer daraus
auf den Code schließt, sucht Fehler an Stellen, die es nicht mehr gibt. Erst die
Version im Bild (bzw. „Was ist neu?") mit `version.ts` abgleichen.

Bei Bildschirmfotos mit Fehlern: erst die Ursache im Code belegen, dann fixen.
Vermutungen als solche kennzeichnen. Wenn etwas nicht sauber fertig wird, lieber
den Teil sauber liefern und den Rest benennen, als halb umgebaut zu hinterlassen.
