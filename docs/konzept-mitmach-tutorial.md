# Konzept: Mitmach-Tutorial „Wir legen zusammen ein Test-Event an"

Stand 28.09.2026 (v32.0.15). Auslöser, Nutzer-Ansage: „das Tutorial muss viel
besser sein … ein echtes Test-Event … zusammen anlegen, Name eingeben, dann
Test-Datum … der User muss Dinge anklicken und selbst eingeben und dann auf
nächsten Schritt klicken … und im Event-Assistenten oben ein Button für das
Tutorial."

## 1. Ist-Stand

- `components/tutorial/TutorialGuide.tsx` (Provider, `openTutorial`,
  `startTour`, `TourChooser`, `TutorialOverlay`) und `tutorialTours.ts`
  (Touren `user` 9 Karten, `organizer` 17 Karten). Eingebunden in
  `DexEventPlatform.tsx` (~1350), `AppContent` liegt im Provider.
- Das Overlay hat einen vollflächigen Klick-Fänger — Mitmachen ist damit
  unmöglich. Spotlight per Box-Shadow-Loch, z-index 10800, Ziel wird gepollt.
- Der Wizard-Teil blättert nur (`dex-tutorial-wizard-step` →
  `setCurrentStep`), es wird nichts eingegeben oder gespeichert. Das
  Demo-Event lebt nur im Speicher (`tutorialDemoActive`,
  `buildDemoShowcaseEvents`).
- Einstiege: Header-Pille „Neu hier?" (nicht für Organizer/Admins),
  `LandingInfoModal`, Admin-Hub-Kachel, seit v32.0.14 Deep-Link
  `action=tutorial` aus der Onboarding-Mail.

## 2. Was ein echtes Test-Event heute auslöst

- Subsite, Teilnehmerliste und Rechte entstehen echt (30–60 s, 429 möglich).
- `DEX_CreateOutlookEvent` triggert auf JEDE neue `DEX_Events`-Zeile
  (`GetOnNewItems`) — auch bei Entwürfen. Empfänger: `OrganizerEmail`.
  Dass `DisableOutlook = true` den Termin verhindert, behauptet nur der
  App-Code; eine Trigger-Bedingung ist in `flow-jsons.md` nicht dokumentiert
  → vor Variante A im Tenant (Run history) prüfen.
- Mails beim Anlegen: „Event angelegt" an Organizer, Co-Organizer-Freigabe/
  Onboarding, Admin-Hinweis bei externer Zielgruppe, Check-in-Team-Freigabe.
- Entwürfe sind für KPI, Auto-Mails, Wochenbericht, Abrechnung und
  Teilnehmer ohnehin unsichtbar; sichtbar für Admins/Organizer im
  Organizer Center.
- Löschen: Register → Outlook-Absage (NUR wenn `CalendarLink` schon da ist)
  → Recycle. Ein eigenes Test-Flag gibt es nicht (`isFictive` = Entwurf).

## 3. Absicherung des Test-Events (zwei Netze)

1. **Im Formular** (Tutorial-Modus): Titel-Präfix „TEST – ", `isFictive`
   an, `activeFrom` leer, Organizer nur ich, Test-/Check-in-Team leer,
   Zielgruppe nur ich.
2. **Hart in `wizardSubmit`** (Kontextfeld `tutorialTest`): dieselben Werte
   erzwingen, Piggyback `_tutorialTest: { by, at }`, Co-Organizer-Freigabe,
   Admin-Hinweis und Scanner-Rechte überspringen. Flag in
   `useWizardVisibilityState` strippen und in `hotelCarryConfig` mittragen
   (CLAUDE.md-Regel), in `eventMapping` → `isTutorialTest`.

Risiken: eigener Entwurfsschlüssel `dex_event_creation_draft_tutorial_v1`
Pflicht (sonst überschreibt das Tutorial einen echten Entwurf); Vorlagen-
Fächer im Tutorial ausblenden; Aufräumen erst, wenn `CalendarLink` gesetzt
ist (sonst bleibt der Outlook-Termin stehen); Sub-Events nicht ansprechen.

## 4. Drehbuch (15 Stationen)

Grundregel: Eine Station ist ein ZUSTAND, kein Klick. Eingabe-Stationen
schalten „Weiter" auf der Karte erst frei, wenn der Zustand erfüllt ist;
Klick-Stationen rücken automatisch vor, wenn der echte Knopf gewirkt hat.

| # | Spotlight | Nutzer tut | Fertig, wenn |
|---|---|---|---|
| 0 | zentriert — Einleitung („nur du siehst es, ~5 Min.") | „Los geht's" | – |
| 1 | Nutzungsbedingungen (falls offen) | 2 Haken, Akzeptieren | `tcAccepted` |
| 2 | `wizard-title` | Namen tippen | `title.trim().length >= 3` |
| 3 | `wizard-dates` (+ Datepicker) | Beginn, Ende wählen | Start/Ende gesetzt, Ende > Start |
| 4 | `wizard-draft` | Haken bleibt an | `isFictive` |
| 5 | `wizard-next` | „Weiter" | `currentStep === 1` |
| 6 | `wizard-organizer` | lesen, „Weiter" | `currentStep === 2` |
| 7 | `wizard-location` | Ort tippen (oder überspringen) | `location.trim()` |
| 8 | `wizard-next` | „Weiter" | `currentStep === 3` |
| 9 | `wizard-capacity` / `wizard-deadline` | Plätze tippen, „Weiter" | `max > 0` oder unbegrenzt; `currentStep === 4` |
| 10 | `wizard-add-question` / `wizard-question-label` | Frage anlegen, benennen, „Weiter" | neues Feld mit Label; `currentStep === 5` |
| 11 | `wizard-channel` (+ `wizard-testmail`) | Kanal / Testmail | je nach Entscheidung 1 |
| 12 | `wizard-create-early` | „Event erstellen" | Countdown oder Submit läuft |
| 13 | `anlege-countdown` → Summary | warten | Organizer Center mit Test-Event offen |
| 14 | `admin-next-steps` | lesen, Anmeldeseite ansehen | – |
| 15 | zentriert — „Löschen oder behalten?" | Wahl | – |

Neue `data-tour`-Anker: `wizard-title`, `wizard-dates`, `wizard-draft`,
`wizard-organizer`, `wizard-location`, `wizard-capacity`, `wizard-deadline`,
`wizard-add-question`, `wizard-question-label`, `wizard-channel`,
`wizard-testmail`, `wizard-next` (beide Weiter-Knöpfe), `wizard-create-early`,
`anlege-countdown`, `admin-next-steps`.

## 5. Architektur

- Neu: `components/tutorial/WizardCoach.tsx` + `coachStations.ts`; Provider
  und Einstiege bleiben. Messlogik als gemeinsamer Hook
  `useSpotlightRect(selectors[])`.
- Overlay blockiert NICHT (`pointer-events: none`, nur Loch + pulsierender
  Rahmen); bei offenem App-Modal dockt die Karte unten rechts an; mobil
  Bottom-Sheet. Klassen `dex-ui-coach-*` in `dexUi.ts` UND Leitfaden.
- `EventCreationPage` meldet per `useEffect` einen flachen Snapshot
  (currentStep, title, Daten, isFictive, Organizer, Ort, Plätze,
  Fragen, Countdown …) an den Provider — vor keinem frühen Return.
- `testEventId` aus `dex-event-submit-success`.
- Wizard im Coach-Modus mit `key="coach"` neu starten, eigener
  Entwurfsschlüssel; Fortsetzen über `localStorage dex_wizard_coach_v1`.
- Einstiege: Chip „Tutorial: Test-Event gemeinsam anlegen" oben im
  Wizard (nur Neu-Anlage), dritte Option im `TourChooser`, Deep-Link
  `action=tutorial` startet für Organizer den Coach.

## 6. Stufenplan

- **Stufe 1** (~1.000 Zeilen): Coach, Stationen 0–13 + 15, Anker,
  Snapshot, eigener Entwurfsschlüssel, beide Sicherheitsnetze, Knopf im
  Wizard, Eintrag im Chooser, Löschen am Ende.
- **Stufe 2** (~500): Anmeldeseite + Test-Anmeldung, Fortsetzen nach
  Reload, `isTutorialTest`-Badge, „Live schalten"/„Einladungsmail" beim
  Test-Event ausblenden, Header-Pille für Organizer ohne Event.
- **Stufe 3** (~300): alte Organizer-Tour im Wizard-Teil ersetzen,
  Handbuch/Rollenmatrix/Architekturseite, Harness-Screenshots.

## 7. Entscheidungen (Eike, 28.09.2026)

1. Kommunikation: **B „nur an mich"** — Mails und Outlook-Termin laufen
   echt, Organizer und Zielgruppe sind erzwungen nur die Person selbst.
2. Ende: **Löschen als Hauptknopf**, „Als Entwurf behalten" daneben.
3. Alte Organizer-Tour: **Assistenten-Teil ersetzt** durch eine Karte mit
   Verweis; Organizer Center und Check-in bleiben (Label jetzt
   „Organizer-Rundgang").
4. Kennzeichnung: **Präfix „TEST – " plus Flag** `_tutorialTest`.

## 8. Stand der Umsetzung

**Stufe 1 ausgeliefert mit v32.1.0.** Dateien: `tutorial/WizardCoach.tsx`
(Stationen + Overlay), `tutorial/TutorialGuide.tsx` (Provider: `coachActive`,
`startCoach`, `stopCoach`, `reportWizard`, Löschen, Auswahl mit Empfehlung),
`EventCreationPage` (Schnappschuss-Effect VOR `if (submitted) return`,
`DRAFT_KEY` = `COACH_DRAFT_KEY`, erstes Netz), `wizardSubmit.tutorialTestSicher`
(zweites Netz), `WizardFormShell` (Chip oben, Hinweis-Kasten, Anker),
`DexEventPlatform` (key-Remount, Deep-Link startet den Coach).

Abweichungen vom Drehbuch, alle aus dem Harness-Durchlauf
(`tools/wizard-harness/coach-shot.js`, 20 Bilder, 0 Fehler):
- Nach den Nutzungsbedingungen kommt für Admins die Abrechnungsfrage — die
  Station wartet auf `tcAccepted && !billingPromptOpen`.
- Nicht jeder Dialog trägt `role="dialog"`: Der Coach prüft zusätzlich per
  `elementFromPoint`, ob sein Ziel verdeckt ist, und dockt dann unten rechts an.
- Plätze stehen per Vorgabe auf „Unbegrenzt" (kein Zahlenfeld): erst die
  Kachel „Begrenzt" (`wizard-capacity-mode`), dann das Feld.
- `dexUiPulse` animiert box-shadow und überschrieb das Abdunkeln — Loch und
  Puls-Rahmen sind zwei Elemente.
- ESC beendet den Coach NICHT (schließt im Formular Datepicker und Dialoge).

Offen (Stufe 2/3): Fortsetzen nach Reload, Anmeldeseite + Test-Anmeldung,
`isTutorialTest`-Badge und „Live schalten"/„Einladungsmail" beim Test-Event
ausblenden, Header-Pille für Organizer ohne Event, Handbuch-Artikel.
