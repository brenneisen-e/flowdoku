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

**Branch:** `claude/spfx-app-bugfixes-4kui16` — Stand **v28.87.0**.
Nur auf diesen Branch pushen. Keine PRs ohne ausdrückliche Aufforderung.

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
# 3) tsc
npx tsc --noEmit -p tsconfig.json
# 4) Sauber bauen (stale Bundles blähen das .sppkg auf)
rm -rf dist release temp sharepoint/solution/debug
npx gulp bundle --ship && npx gulp package-solution --ship
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

## Der Wizard, Stand v28.87

Neun Schritte. Über dem Formular steht die **Scope-Karte**
(`renderGlobalScopeBar`): Klammer/Haupt-Event und die Sub-Events als Reiter, ein
gemeinsamer Index (`activeScopeIdx` / `setScope`). In Schritten ohne Scope-Bezug
bleibt die Karte stehen und sagt, dass der Schritt für alles gilt.

Schritt 1 (Grundlagen) trägt seit v28.83–v28.87 auch: „Sub-Events nutzen?",
Bezeichnung, Anmelde-Modus und die Sub-Event-Karten.

**`setScope` ruft für die Kommunikation bewusst `switchCommTab(idx)`** statt
`setActiveCommTabIdx` — sonst gehen Sub-Event-Mailtexte verloren.

## Offene Arbeit, in dieser Reihenfolge

1. **Feld-Ebene vereinheitlichen** (der eigentlich gewünschte Endzustand):
   `title`, `startDate`, `endDate`, `description` und Bild in Schritt 1 je nach
   Scope an den Top-Level-State **oder** an `subEvents[activeScopeIdx-1]` binden:

   ```tsx
   value={activeScopeIdx > 0 ? sub.title : title}
   onChange={v => activeScopeIdx > 0 ? patchSub({ title: v }) : setTitle(v)}
   ```

   Wichtig: die **vorhandenen** Felder umhängen, keine zweite Box daneben bauen
   (das wurde in v28.81 versucht und in v28.82 zurückgenommen — „alles soll
   gleich aussehen", inklusive des „Bearbeiten & Vorschau"-Editors für die
   Beschreibung). Danach sind die Sub-Event-Karten überflüssig und können weg.

2. **Reiter-Darstellung** recherchieren und neu gestalten. Bei neun Sub-Events
   ist die Chip-Reihe grenzwertig; Kandidaten: Segmented Control, Dropdown mit
   Suche ab N Einträgen, Scroll-Leiste mit Pfeilen. Bisher nicht recherchiert.

3. **Outlook-Body-Dialog** übersichtlicher (Betreff/Termin/Ort, Überschriften,
   Header-Bild, Variablen und Editor stehen ungegliedert untereinander).

4. **Co-Organizer-Antrag:** `getRoleEmails('Organizer')` prüft nur den Wert
   `Organizer`. Falls in `DEX_Roles` noch Legacy-Einträge mit `EventAdmin`
   stehen, bekämen diese Personen unnötig einen Freigabe-Antrag.

## Umgang

Nach dem Deploy den Wizard einmal von Schritt 1 bis 9 durchklicken — die
Umnummerierung aus v28.87 ist im Browser nicht verifiziert.

Bei Bildschirmfotos mit Fehlern: erst die Ursache im Code belegen, dann fixen.
Vermutungen als solche kennzeichnen. Wenn etwas nicht sauber fertig wird, lieber
den Teil sauber liefern und den Rest benennen, als halb umgebaut zu hinterlassen.
