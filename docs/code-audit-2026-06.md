# Code-Audit v20.0 (2026-06-09)

Vollständiges Audit der SPFx-App (66.084 Zeilen Quellcode). Methodik: Die
gesamte Codebasis wurde in 7 Bereiche aufgeteilt und jede Quelldatei von genau
einem Auditor-Durchlauf Zeile für Zeile gelesen (EventService, übrige
Services/Utils/Types, EventCreationPage, AdminPage, RegistrationPage+MyEventsPage,
Contexts+App-Root, restliche Komponenten+Config). Jeder gemeldete Befund wurde
anschließend **einzeln im Code verifiziert**, bevor etwas geändert wurde —
mehrere Meldungen stellten sich dabei als False Positives heraus (siehe unten).
Leitplanke: **keine einzige Funktion darf verloren gehen** — alle Änderungen
sind verhaltenserhaltend (Bug-Fixes stellen dokumentiertes Soll-Verhalten her).

## 1. Behobene Bugs (v20.0)

### 1.1 Sub-Event-Edit verliert Felder (gleiche Klasse wie v19.32) — `EventCreationPage.tsx`

Der UPDATE-Payload für bestehende Sub-Events (`subUpdates` in
`persistSubEventsForParent`) schrieb nur einen Teil der Felder, die der
Create-Pfad (`childPayload`) kennt. Folge: Änderungen an diesen Feldern auf
einem **bestehenden** Sub-Event gingen beim Speichern still verloren:

- `Audience` / `LocationFilter` / `FilterMode` (v19.27-AudiencePicker pro Sub-Event!)
- `LocationAddress` (strukturierte Adresse, v15.3)
- `Agenda` / `Transfers`
- `LastDeregisterDate`, `WaitlistEnabled`
- `AskSalutation` (pro-Sub-Event-Anrede-Checkbox)

Alle neun Felder werden jetzt im `subUpdates`-Payload mitgeschrieben
(Werte konsistent aus `childPayload`). `AudienceResolvedEmails` wird bewusst
NICHT geschrieben — der Create-Pfad für Sub-Events schreibt sie ebenfalls
nicht (nur Top-Level löst DLs vor-auf), Konsistenz zum Create-Verhalten.

### 1.2 `WaitlistEnabled` fehlte auch im Top-Level-Edit — `EventCreationPage.tsx`

Das Umschalten der Warteliste auf einem **bestehenden** Event wurde nie
persistiert (nur der Create-Pfad schrieb das Feld). Jetzt im `updates`-Payload
ergänzt.

### 1.3 QR-Code-Erzeugung scheiterte still — `EventContext.tsx`

Beim Auto-Send-QR fiel die App bei einem `QRCode.toDataURL`-Fehler still auf
den Text-Fallback zurück, ohne den Fehler zu loggen. Jetzt `console.warn` im
Catch (Beobachtbarkeit; Fallback-Verhalten unverändert).

## 2. Bundle-Architektur (größter Hebel)

**Hauptbundle: 5.087.658 → 891.351 Bytes (−82 %)** — jede App-Öffnung lädt
und parst jetzt ~0,9 MB statt ~5,1 MB. Zwei Mechanismen:

### 2.1 Route-Level-Code-Splitting (`DexEventPlatform.tsx`)

Neun Sekundär-Seiten werden per `React.lazy` + `Suspense` erst beim ersten
Aufruf als eigener Webpack-Chunk geladen: `EventCreationPage` (Wizard, 13k
Zeilen + react-datepicker), `AdminPage` (9,5k Zeilen), `CheckInPage`,
`SettingsPage`, `RoleMatrixPage`, `ParticipantsPage`, `FlowchartPage`,
`ManualPage`, `SelfCheckInDisplayPage`. Eager bleiben die Teilnehmer-Pfade
(Landing, Start, EventList, Registration, MyEvents, Profile) und
`SelfCheckInPage` (Deep-Link wird VOR dem Boot-Loader gerendert).

### 2.2 Call-Site-Lazy-Loading schwerer Bibliotheken

| Bibliothek | Vorher | Jetzt geladen bei |
|---|---|---|
| `xlsx` (~1,3 MB) | statisch in AdminPage | Klick auf Excel-Export (`import('xlsx').then(...)`) |
| `react-pdf`/pdfjs | statisch via PdfViewer in MyEventsPage | Öffnen einer PDF-Vorschau (`React.lazy`) |
| `qr-scanner` | statisch in CheckInPage | Kamera-Start (Typ-Import bleibt statisch) |
| `jspdf` + `qrcode` | statisch in selfCheckInPdf | Self-Check-in-PDF-Download |
| `qrcode` | statisch in EventContext (Boot-Pfad!) + AdminPage | QR-Mail-Erzeugung / QR-Vorschau |

**Konvention ab v20.0:** Neue schwere Dependencies (> ~50 KB) gehören nicht in
statische Imports von Boot-Pfad-Dateien (Contexts, eager Pages) — stattdessen
`await import('lib')` an der Verwendungsstelle oder `React.lazy` für ganze
Komponenten.

## 3. Performance-Fixes

- **Context-Provider-Values memoizt** (`LanguageContext`, `UserContext`,
  `NavigationContext`, `RoleContext`): vorher erzeugte jeder Provider-Re-Render
  (z.B. durch die asynchronen Boot-Nachlade-Schritte Profil/Foto/Gruppen/Rollen)
  ein neues Value-Objekt → App-weite Re-Renders aller Consumer.
  **`EventContext` wurde bewusst NICHT memoizt** — sein Value enthält ~30
  Funktionen mit Closures über diverse State-Atome; ohne Laufzeit-Tests ist das
  Stale-Closure-Risiko höher als der Nutzen. Sauberer Folge-Schritt: Funktionen
  einzeln auf `useCallback` umstellen, dann Value memoizen.
- **`EventListPage`:** komplette Filter-/Sortier-Kette (Status → Entwurf →
  Sichtbarkeit → Sortierung) in `useMemo`; `isOwnOrganizer` als `useCallback`.
  Lief vorher bei jedem Render über alle Events.
- **`SharePointService.searchUsers`:** Profil-Anreicherung (Location/JobTitle)
  der Picker-Treffer parallel statt sequentiell — vorher bis zu N serielle
  Roundtrips pro Tipp-Suche.
- **`MyEventsPage`:** `allMyNumbers.indexOf(...)` in vier Filter-Schleifen
  durch ein `Set` ersetzt (vorher O(Events × Anmeldungen) pro Ladevorgang).

## 4. Totcode entfernt

- `AdminPage.exportCsv`: CSV-Escaper `esc` + `void esc` (seit XLSX-Umstieg
  v8.4 ungenutzt).
- `AdminPage.getRegListUrl`: nie aufgerufen, lieferte ohnehin nur `<base>/Lists`.

**Bewusst NICHT entfernt:** `CheckInPage.getBrowserUrl` /
`openExternalScanner` — als „für zukünftige Nutzung" markiert (geparkter
Code des Maintainers); `fillDemo` (dokumentiert als bewusst behalten);
`ActionsCollapsibleCard` (Kompat-Shim).

## 5. Verifizierte False Positives (NICHT ändern!)

Diese Befunde aus dem Audit wurden geprüft und sind **korrekt wie sie sind**:

1. **OData-`$filter` mit `encodeURIComponent`** (`EventService` ~4505/4546/
   5214/5234/7373/7394, `SharePointService` ~598): gemeldet als „Filter matcht
   nie". Falsch — der Server dekodiert Percent-Encoding im Query-String VOR
   dem OData-Parsen; `encodeURIComponent` im Filter-WERT ist funktional
   korrekt (und schützt `&`/`#`). Die Pfade laufen seit v19.6 produktiv.
2. **`selfCheckIn.ts` „Timezone-Bug" (lokale Mitternacht):** Das Check-in-
   Fenster „nur am Event-Tag" ist bewusst in **lokaler** Zeit gerechnet —
   korrekt für den deutschen Tenant. Der vorgeschlagene UTC-Fix würde für
   UTC+2 einen ECHTEN Off-by-one-Bug einführen.
3. **`RegistrationPage.t()`-Stale-Closure:** deps `[locale]` sind korrekt —
   `locale` wird pro Render aus `forcedRegLang || appLocale` abgeleitet.
4. **`MyEventsPage` sessionStorage-Parse:** `JSON.parse` ist bereits in
   try/catch mit sauberem Fallback.
5. **`EventService.loadFileAsBase64` „tot":** wird an 2 Stellen aufgerufen
   (Logo-Cache).
6. **AdminPage `adminAddMemberQueryTimer`-„Leak":** die Set-Stelle ruft
   `clearTimeout` vor jedem neuen Timer.
7. **EventListPage „Organizer sieht unsichtbare eigene Events":** ist
   Spec-Verhalten laut Rollenmatrix.

## 6. Offene Punkte (bewusst zurückgestellt, nach Priorität)

1. **`RegistrationPage` Modal-Re-Trigger via `setTimeout(50ms)`** (Extern-
   Mail-/Confirm-/CC-Modal → `handleSubmit()`-Neustart): theoretisches
   Doppel-Submit-Fenster. Redesign auf reine Ref-Flags wäre der saubere Fix,
   ist aber ein Eingriff in den fragilsten Flow der App → nicht ohne
   Tenant-Verifikation umbauen.
2. **`AdminPage` konsolidierte Ansicht:** `consolidatedFiltered`-Pipeline +
   `roommateChoice`-Map werden in `renderConsolidatedView()` (Plain-Function
   im Render) bei jedem Render neu gebaut — Memoization erfordert Hoisting an
   den Komponenten-Top-Level (Hook-Regeln). Lohnt bei großen Klammer-Events.
3. **`EventContext`-Value-Memoization** (siehe 3.) — erst Funktionen auf
   `useCallback` heben.
4. **`EventService.fixRegistrationListColumns`:** Duplikat-Erkennung feuert
   pro Kandidat einen einzelnen REST-Call (bei vielen Duplikat-Feldern
   Throttling-Gefahr beim „Spalten fixen") — batchen oder in-memory filtern.
5. **`reserveSeat` Drift-Mitigation:** `realActive` wird einmal vor der
   CAS-Retry-Schleife gelesen (dokumentierte Best-Effort-Grenze).
6. **Encoding-Altlasten:** ae/oe/ue in ALTEN Kommentaren diverser Dateien
   (kein UI-Text betroffen). Per CLAUDE.md-Regel nur bei Berührung fixen.
7. **`BulkUserImportModal`:** Import-Schleife ohne Abbruch-Flag beim
   Modal-Schließen (nur verschwendete Hintergrund-Calls, kein Datenfehler).
8. **`RegistrationPage` Dokument-Upload nach Anmeldung:** schlägt ein Upload
   fehl, zeigt der Erfolgs-Screen trotzdem Erfolg (Datei fehlt still) —
   Per-Datei-Status + Fehler-Toast wäre der Fix.
9. **CheckInPage Scanner-Lifecycle:** unter schnellem Mount/Unmount könnte
   `stop()` mit neuem `startCamera()` kollidieren (Mount-Flag wäre die
   Härtung).

## 7. Gesamteinschätzung

Der Code ist in deutlich besserem Zustand als die Dateigrößen vermuten lassen:
konsequente Fehlerbehandlung, durchdachte SP-REST-Patterns (CAS/ETag,
nometadata-Disziplin), keine XSS-Funde (alle `dangerouslySetInnerHTML`-Stellen
rendern intern erzeugtes HTML), keine Permission-Lücken in den
Admin-Mutationspfaden gefunden. Die strukturellen Schwächen sind bekannt und
dokumentiert (Monolith-Dateien, Comm-Tab-Spiegelung, zweiter
CustomFields-Write) — die wiederkehrende Bug-Klasse „Feld im Create- aber
nicht im Update-Payload" wurde mit 1.1/1.2 erneut bestätigt und sollte bei
jedem neuen Event-Feld als Checklisten-Punkt gelten (siehe CLAUDE.md).
