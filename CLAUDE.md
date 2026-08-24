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
davor `claude/spfx-app-bugfixes-4kui16`) — Stand **v29.43.0**. Nur auf den
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
# 5) Paket an ZWEI Stellen (seit v29.27 — vorher drei)
cp sharepoint/solution/dex-event-platform.sppkg ../dist/dex-event-platform.sppkg
```

**`docs/index.html` und `docs/downloads/` sind seit v29.27 EINGEFROREN** —
nicht mehr pflegen (User-Ansage 20.08.2026: „bitte bau nicht mehr die
index.html weiter … ich hole mir die aktuelle Version direkt aus dist im
GitHub-Repo"). Das Paket wird nur noch nach `dist/` kopiert; der
Kumulativ-Absatz, die Download-Links und die versionierte Kopie unter
`docs/downloads/` entfallen. `docs/release-notes.md` bleibt Pflicht.

Erwartete Größe: ca. **2,2–2,3 MB**. Deutlich größer heißt: stale Bundles, Schritt 4
wiederholen. In `release/assets/` darf es nur **eine** `dex-event-platform-web-part_*.js` geben.

**Release-Notes nie mit „…"-Literalen in Python-Heredocs schreiben.** Die
Notes brauchen deutsche Anführungszeichen, aber das schließende Zeichen der
Konvention hier ist das GERADE `"` — in einem Python-String beendet es den
String und das Skript stirbt mit `SyntaxError` (zweimal passiert: v29.16 und
v29.17; beim ersten Mal war die Version schon gebumpt, die Notes aber nicht
geschrieben — genau der halbe Zustand, den der Release-Ablauf verbietet).
Deshalb: typografische Paare über einen Helfer bauen (`q = lambda t: '„'+t+'“'`
mit U+201E/U+201C) oder gleich Write/Edit statt eines Skripts nehmen. Nach
jedem Skriptlauf `grep -c <version>` auf BEIDE Dateien — erst dann bauen.

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

**Immer duzen.** Oberfläche, Mails, Release Notes und auch Antworten an
Nutzer außerhalb der App (Ticket-Antworten, Support-Texte) sprechen die Person
mit „du" an — nie mit „Sie". Der Code hält das schon durch; die einzigen
Treffer auf „Sie" sind das Personalpronomen der 3. Person am Satzanfang
(„… Sie werden bereinigt"). Wer eine Formulierung prüft: `Sie haben/können/
sind` + `Ihre` sind die Kandidaten, aber nicht jeder Treffer ist eine Anrede.

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

**Die E-Mail-Adresse ist der einzige Schlüssel — und sie ist nicht eindeutig.**
`DEX_Participants`, die konsolidierte Matrix (`consolidatedRows`), die
Hotelplanung und die Doppel-Anmelde-Prüfung schlüsseln alle über
`ParticipantEmail.toLowerCase().trim()`. Dieselbe Person kann aber unter zwei
Schreibweisen in den Listen stehen (SMTP-Adresse vs. UPN/Alias) — der Code weiß
das an einer Stelle bereits: `canRegisterForOthers` sucht bewusst über
`pageContext.user.email` **und** die E-Mail aus dem `loginName`. Folge: zwei
Zeilen für eine Person, jede mit „—" beim Sub-Event der anderen. Wenn zwei
Ansichten sich über „angemeldet ja/nein" widersprechen, ist das der erste
Verdacht — **nicht** ein Anzeigefehler. Zum Nachrechnen: Matrix und
`HotelPlanningPanel` lesen dasselbe `subEventRegsByEventId`, und die
Status-Liste der Matrix (`+ Warteliste`) ist eine Obermenge von `ACTIVE_STATI`
im Panel. Bei gleicher E-Mail und gleichem Event **kann** es keine Abweichung
geben; bleibt nur: zwei Adressen oder zwei Klammer-Events. v29.2 weist das
über der Tabelle aus.

**Antworten stehen dort, wo angemeldet wurde — nicht auf der Klammer.** Bei
einem Klammer-Event ist die Zeile auf der Klammer eine Schattenzeile (v15.25):
kein Platz, keine Mail, oft **kein CustomData**. Die Formularantworten liegen
auf der Sub-Event-Zeile. Die konsolidierte Matrix löst Hauptevent-Felder
deshalb seit v15.3.1 zweistufig auf (Parent-Zeile zuerst, dann
Sub-Event-`CustomData`). Wer eine neue Auswertung über Antworten baut, muss
denselben Fallback nehmen — sonst widerspricht sie der Teilnehmerliste.
`HotelPlanningPanel.wishOf`/`formStayOf` waren bis v29.2 genau diese Ausnahme:
„Yes, I need accommodation" in der Matrix, „—" im Panel, und `autoDistribute`
übersprang die Person als `skippedNoWish`. Seit v29.3 gibt es dafür
`answerRowsOf(p)` = `[Klammer-Zeile, …Sub-Zeilen]`; die Feldsuche geht über
Parent- **und** Child-`eventSpecificFields`.

**`getAllRegistrations` wirft nicht.** Bei HTTP-Fehler bricht die Schleife ab
und liefert die bis dahin gelesenen Zeilen — bei 404 also `[]`. „Subsite
recycelt" und „Liste leer" sind damit ununterscheidbar, und genau darauf ist
`analyzeRegistryAgainstLists` in v29.0/29.1 hereingefallen (1045 „verwaiste"
Verweise = Rückstand des 3-Monats-Löschkonzepts). Seit v29.3 gibt es den
`onHttpError`-Rückruf; 404/410 heißt „Liste gelöscht" (eindeutig, Verweise
dürfen weg), alles andere heißt „übersprungen". Wer aus einem leeren Ergebnis
auf „nicht vorhanden" schließt, muss diesen Rückruf nutzen.

**Löschungen zuerst im Register, dann unwiderruflich.** `deleteParticipantData`
recycelte bis v29.2 erst die Subsite und räumte danach `DEX_Participants` auf —
mit nicht-striktem Lesen, einem `Promise.all` über alle Personen und
`.catch(() => null)`. Ein Throttling-429 hinterließ unbemerkt genau die
Verweise, die die Register-Prüfung später meldet, und nachrechnen ging nicht
mehr. Reihenfolge deshalb immer: prüfbare Nebenbuchhaltung zuerst, sequentiell,
mit Fehlerzähler — und bei Fehlern **abbrechen, bevor** etwas gelöscht wird.

**Bevor du zwei Ansichten „widersprüchlich" nennst: zähl die Spalten.** Der
Fall, der v29.2 und v29.3 ausgelöst hat („Hotel-Wunsch —" im Panel, „Yes, I
need accommodation" in der Matrix), war weder ein Datenproblem noch ein
Auflösungsproblem: In `HotelPlanningPanel` rendert der `<thead>`
`[wish, …childEvents, …]`, der `<tbody>` rendete `[…childEvents, wish, …]` —
v28.53 hat die Haken-Spalten nur im Kopf an der richtigen Stelle eingezogen.
Jede Zeile war um eine Spalte verschoben. Zwei Releases lang habe ich die
Daten untersucht, obwohl die Daten stimmten. Bei einer Tabelle mit
`.map`-erzeugten Spalten deshalb **zuerst** Kopf- und Zeilen-Reihenfolge
nebeneinanderlegen, dann erst die Werte.

**Vor einer neuen Ansicht prüfen, ob es sie schon gibt.** In v29.10 habe ich
einen Handbuch-Artikel „Systemarchitektur" gebaut — es gab längst
`ArchitecturePage` (v26.28, Kachel im Admin Hub, mit jsPDF-Export). Zwei
Darstellungen derselben Sache, die auseinanderlaufen; dieselbe Falle wie zwei
Bedienwege für dieselbe Auswahl. v29.12 hat konsolidiert: Einzelheiten stehen
auf der Architekturseite und wachsen dort automatisch ins PDF, der
Handbuch-Artikel liefert nur das Schaubild und verweist. **Neue Aufzählungen
gehören auf die Architekturseite**, nicht in den Artikel.

**Die Flow-Sammlung ist nicht vollständig.** `docs/flow-jsons.md` führt sieben
Flows, es gibt aber **acht**: `DEX_AssistantAccess_Grant` (v24.41) fehlt dort.
Wer die Zahl aus dieser Datei nimmt, dokumentiert sie falsch — mir passiert in
v29.10 bis v29.11, korrigiert in v29.12. Die vollständige Liste steht in
`ArchitecturePage` und ergibt sich aus den Queue-Listen im `EventService`
(`ensure*List`-Methoden nennen den zugehörigen Flow im Description-Feld).

**Es gibt ZWEI Bilder je Event, und sie haben nichts miteinander zu tun.**
Das Event-Bild (Schritt 1 → `EventImageUrl`, Item-Attachment) trägt Kachel und
Anmeldeseite; das Mail-Logo (Schritt Kommunikation → `EmailTemplateOverrides.
_eventLogo`, gespiegelt nach `EmailImageBase64`) trägt Mails und
Outlook-Termin — der Flow ersetzt damit `{{ORB_URL}}`. „Das Bild kommt in der
Mail an, aber nicht auf der Seite" ist deshalb kein Anzeigefehler, sondern
zwei Uploads, von denen einer leer blieb. Seit v29.13 fällt der Hero-Slot der
Anmeldeseite auf das Mail-Logo zurück (`heroImgUrl`, `usesMailImage`); die
Cover-Hintergründe bleiben bewusst am Event-Bild, weil ein Logo im Beschnitt
zerfällt.

**`subEventsOnlyMode` heißt: das Hauptevent ist keine Anmeldeeinheit.** Alles,
was die Kachel/Übersicht über die Klammer aussagt, ist dann eine Aussage über
etwas, das niemand buchen kann: `MaxParticipants` bleibt 0 und wurde als
„Unbegrenzt" gelesen, `RegistrationDeadline` ist ein Alt-Wert. Seit v29.13
entfallen beide auf `EventCard` und in `EventListView`. Und die Kinder heißen
für Teilnehmer nicht „Sub-Events", sondern „Events" — der Default von
`childTermSingular`/`childTermPlural` kippt in diesem Modus (`RegistrationPage`,
`MyEventSubEvents`). Wer neue teilnehmersichtbare Texte baut, nimmt diese
beiden Konstanten und **nie** ein fest verdrahtetes „Sub-Event"; für den
unbestimmten Artikel gibt es `childOneDe` (es heißt „ein Event", aber „eine
Session" — das war vorher überall falsch).

**Nachgerückt wird nur beim Abmelden — nicht bei einer Kapazitätsänderung.**
`promoteFirstWaitlistItem` hängt am Cancel-Pfad. Erhöht der Organizer eine
Kapazität, gibt es kein Ereignis, an dem etwas hinge; die Warteliste bleibt
stehen. Dafür ist die Aktion „Freie Plätze mit Warteliste füllen" da
(`runManualPromote`, seit v29.16 gruppen-bewusst und für alle freien Plätze).

**Bei geteilten Kapazitäten ist `maxParticipants` 0.** Die Kapazität steht in
`durchstarterCapacity`/`funstarterCapacity` (Beschriftungen `splitLabelA/B`).
Wer `maxParticipants` als Obergrenze prüft, prüft dort **gar nichts** — genau
daran ist das Nachrücken bis v29.16 vorbeigelaufen. Die Überbuchungs-Sperre in
`promoteFirstWaitlistItem` zählt über die GANZE Liste und kann eine Gruppe
nicht trennen; wer je Gruppe nachrückt, muss die Anzahl selbst ausrechnen
(`StarterType || PreferredStarterType`) und `onlyWithPreferredType` setzen.
Ausnahme: `splitSharedWaitlist` — dann ist es ein Topf mit der Summe beider
Kapazitäten.

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

## Modularisierung

Stand: **106k Zeilen in 132 Dateien**, davon die Hälfte in vier Dateien
(`EventCreationPage` 17,4k · `AdminPage` 15,3k · `EventService` 13,2k ·
`RegistrationPage` 6,3k). **Es gibt keine Tests.** Das einzige Netz sind
`tsc`, ESLint und der Build — was der Compiler nicht sieht, sieht niemand.

Deshalb in dieser Reihenfolge, nicht anders:

**Stufe 1 — Modul-Ebene (v28.94 erledigt).** Alles, was VOR der Komponente
steht, kennt den State nicht und lässt sich als Ganzes verschieben. Ergebnis:
`components/wizard/*` (StickyTabStrip, StepBadge, LocationMultiSelect,
FieldDescEditor, FieldTypeSuggestion, customFieldInput),
`components/admin/ActionsMenu.tsx` (ActionTile + Registry + Dropdown gehören
zusammen — eine Datei, nicht drei), `utils/{subEventTitle, eventStatus,
inviteGuards, fieldHeuristics}`, `data/{actionCategories,
descriptionTemplates}`. Rezept: Block ausschneiden, `export` davor, Import
zurück, `tsc`. **Danach immer `gulp bundle`** — ESLint findet die zu breit
gefassten Importe, die `tsc` durchgehen lässt.

**Stufe 2 — `EventService` (offen).** 13,2k Zeilen in einer Klasse, deren
Methoden über `this.context`/`this.siteUrl` laufen. Aufteilbar nach Thema
(Events, Teilnehmer, Hotels, Wartung), indem die Klasse die Methoden an
Modul-Funktionen delegiert, die den Kontext als Parameter bekommen. Mechanisch
und compiler-geprüft, aber viele Aufrufstellen — in einem Rutsch pro Thema,
nicht querbeet.

**Stufe 3 — die Render-Bäume (offen, teuer).** In `EventCreationPage` stecken
~16k Zeilen in EINER Funktion; die neun Schritte lesen aus rund 200
State-Variablen. Ein Schritt als eigene Komponente braucht einen echten
Props-Vertrag (oder einen Wizard-Context) — sonst schiebt man 200 Props durch.
Das ist genau der Umbau, der laut „Fallen" schon einmal die Tag-Balance
zerriss. **Nicht ohne Browser-Verifikation anfangen**, und immer nur EINEN
Schritt pro Release.

## Offene Arbeit

Die vier Punkte aus v28.87 sind mit v28.88/v28.89 abgearbeitet (Feld-Ebene,
Reiter-Darstellung, Outlook-Dialog, Legacy-Rollen), v28.90 hat die Nachlese
dazu erledigt. Offen und **noch nicht begonnen**:

1. **Stunden-Slots.** v28.91 kann Tage (`_subEventCalendar`): Im Assistenten
   Tage anklicken → je Tag ein normales Sub-Event; auf der Anmeldeseite ein
   Monatsraster statt der Liste. Ein Raster **innerhalb** eines Tages (09–11,
   11–13 …) ist bewusst nicht gebaut. Wenn es kommt, auf demselben Weg: ein
   Slot bleibt ein Sub-Event, der Kalender erzeugt nur mehrere pro Tag. Ein
   eigener Datentyp unterhalb des Sub-Events würde Teilnehmerliste, Flows und
   Warteliste mitziehen — das ist die teure Variante.

   **Wichtig bei erzeugten Terminen:** Start/Ende werden explizit auf
   00:00/23:59 gesetzt. Leer lassen wäre falsch — ein Sub-Event ohne Zeiten
   erbt seit v28.66 die Zeiten des Hauptevents, bei einer Reihe also den
   gesamten Zeitraum statt des einen Tages.

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
