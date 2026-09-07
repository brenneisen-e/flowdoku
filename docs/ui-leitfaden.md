# UI-Leitfaden — Wizard, Organizer Center, Modale (Stand v31.3)

Dieser Leitfaden ist die **verbindliche Arbeitsgrundlage** für jede Änderung an
der Oberfläche des Event-Wizards, des Organizer Centers und aller Modale. Er
entstand mit v31.2, als rund 60 Dateien parallel modernisiert wurden
(Nutzer-Auftrag 07.09.2026: „jeden einzelnen Schritt im Event-Wizard moderner,
intuitiver und aufgeräumter darstellen, gute Mouseover-Effekte, alle Modale
optimieren"). Wer eine Datei anfasst, hält sich an diesen Leitfaden — sonst
laufen 60 Dateien wieder auseinander.

Die Klassen liegen in `src/webparts/dexEventPlatform/components/dexUi.ts`
(`ensureDexUiStyles()` injiziert sie einmal ins Dokument; `WizardFormShell`
und `Modal` rufen das bereits). **Neue Klassen: dort ergänzen UND hier
eintragen.**

---

## 1. Grundsätze

1. **Der Inhalt führt, nicht der Rahmen.** Weiße Karten auf hellem Grund,
   dünne Linien (1 px, `--dex-gray-200`), Radius 12–14 px, wenig Schatten. Grün
   (`--dex-green` #86bc25) nur dort, wo es etwas bedeutet: aktive Auswahl,
   Primär-Knopf, Fortschritt, Erfolg. Nie als Flächenfarbe für Kästen.
2. **Eine Frage je Block.** Ein Abschnitt (`dex-ui-section`) stellt genau eine
   Frage und beantwortet sie mit einem Bedienelement. Erklärtext maximal zwei
   Zeilen sichtbar; alles Weitere in `InfoTooltip` oder einen Aufklapper
   (`dex-ui-disclosure`, standardmäßig zu).
3. **Jedes klickbare Element hat einen Hover — und nur klickbare.** Inline-
   Styles können kein `:hover` — deshalb die Klassen unten. Ein Element ohne
   Hover liest sich als Beschriftung; umgekehrt verspricht ein Hover eine
   Aktion. Nutzer-Ansage 07.09.2026: „Mouseover natürlich nur dort, wo es Sinn
   ergibt, und kein Mouseover um des Mouseovers willen." Also: Karte mit
   Hover nur, wenn Klick sie öffnet/bearbeitet; Zeile mit Hover nur, wenn
   Klick etwas tut; reine Anzeigen (Kennzahlen, Statuspillen, Hinweiskästen,
   Vorschau) bleiben ruhig — `dex-ui-card` ohne `--hover`, `dex-ui-pill`
   statt `dex-ui-chip`. Kein `onMouseEnter`-State mehr für reine Optik.
4. **Der Knopf ist der Schritt.** Wo ein Ablauf erklärt wird („1. Vorschau,
   2. Test, 3. Senden"), sitzt der Knopf in derselben Zeile wie die Nummer
   (`dex-ui-step`). Nummern und Knöpfe getrennt zwingen zum Zuordnen.
5. **Genau ein Primär-Knopf je Ansicht.** `btn btn-primary` einmal, alles
   andere `btn-secondary`, `btn-outline` oder `dex-ui-textbtn`. Destruktives
   ist `btn-danger` und braucht eine Rückfrage (`confirmDialog`, `danger: true`).
6. **Selten Gebrauchtes ist zu.** Feineinstellungen, Alt-Optionen,
   Expertenschalter: hinter `dex-ui-disclosure` („Weitere Einstellungen"),
   Standard eingeklappt. Was leer ist, zeigt einen leeren Zustand
   (`dex-ui-empty`) statt einer leeren Fläche.
7. **Deutsch, du, echte Umlaute.** Beide Sprachen über `isDe`. Typografische
   Anführungszeichen in JSX-Text als `&bdquo;…&ldquo;`. Ausnahmen aus
   CLAUDE.md gelten (`Nachruecken`, `OrgNachruecker`, `nachruecker`, `UeberUns`).
8. **Keine nativen Datums-/Zeitfelder.** Chrome/Edge zeigen sie in der
   Browser-Sprache. Datum über `react-datepicker` (`dd.MM.yyyy`, locale de),
   Zeit als Text mit `normalizeTimeInput` (siehe `AgendaEditor`).
9. **Mobil bleibt bedienbar.** Raster über `dex-ui-grid-2/3` (brechen bei
   768 px auf eine Spalte), Knopfreihen mit `flexWrap`, bestehende
   `isMobile`-Zweige erhalten.

---

## 2. Was verändert werden SOLL — und was nicht

Die Modernisierung ist mehr als Optik. Nutzer-Ansage 07.09.2026: „Bei den
Agenten schaust du dir aber auch die richtige Logik der Reihenfolge, die
Formulierungen und wie eine Frage überhaupt abgefragt wird an." Jede Datei
wird deshalb auf drei Ebenen überarbeitet:

### 2a. Reihenfolge — in welcher Folge fragt der Schritt?

- **Pflicht zuerst, dann Optionales, dann Feineinstellungen.** Was ohne
  Antwort nicht weitergeht (Titel, Datum, Organizer), steht oben; was das
  Event nur verfeinert, darunter; was selten jemand anfasst, in einem
  Aufklapper „Weitere Einstellungen" ganz unten.
- **Abhängiges folgt seinem Schalter.** Team-Größe steht direkt unter „Team-
  Anmeldung erlauben", nicht drei Kästen weiter. Was ausgeschaltet ist,
  bleibt sichtbar, aber gedämpft (`dex-ui-card--muted`) — der Organizer
  sieht, was er bekäme.
- **Zusammengehöriges steht zusammen.** Frist und Selbst-Abmeldung, Bild und
  Zuschnitt, Kopfbild und Vollbild-Schalter. Wer zwei Blöcke zum selben Thema
  findet, legt sie in einen Abschnitt.
- **Ein Schritt erzählt eine Geschichte:** „Was ist das Event? → Wer darf
  hin? → Was passiert nach der Anmeldung?" Wenn die Reihenfolge im Code
  historisch gewachsen ist (v9 unten, v29 oben), wird sie umgestellt. Das
  ist ein JSX-Umzug INNERHALB der Datei — erlaubt, sofern jeder Block dieselben
  State-Bindungen behält und die Hook-Reihenfolge unangetastet bleibt (Hooks
  stehen ohnehin am Komponentenanfang).

### 2a′. Ausrichtung — wo steht der Knopf?

Nutzer-Ansage 07.09.2026 (nach den ersten Screenshots): „Es hängt immer sehr
viel rechts, z.B. der Button für Bearbeiten, obwohl das gar nicht notwendig
ist — dann ist nur der Button rechts. Sinnvoll auflösen, dass der Button
links ausgerichtet ist."

- **Aktionen stehen links beim Inhalt, nicht rechts außen.** Ein Knopf, der
  zu einem Text oder einer Zeile gehört („Bearbeiten", „Bearbeiten &
  Vorschau", „Massenimport", „für alle übernehmen", „+ Programmpunkt"), folgt
  dem Inhalt unmittelbar (in derselben Zeile mit `gap`, oder darunter
  linksbündig) — nicht mit `justify-content: space-between` oder
  `margin-left: auto` an den rechten Rand geschoben. Ein Knopf allein rechts
  in einer sonst leeren Zeilenhälfte ist der Fehler.
- **Rechts außen bleibt nur, was WEG vom Inhalt gehört:** Entfernen/Löschen
  (`dex-ui-iconbtn--danger`), Schließen (×), der Zähler eines Aufklappers
  („3 angepasst") und die Knöpfe einer Modal-Fußzeile.
- **Eine Zeile, die eine Aktion hat, ist selbst klickbar** — mit Hover
  (`dex-ui-row`/`dex-ui-card--hover`) und `cursor: pointer`; der Klick auf die
  Zeile tut dasselbe wie ihr Hauptknopf („Bearbeiten"). Die Sub-Event-Zeilen
  in Schritt 1 sind das Beispiel des Nutzers: Zeile überfahren → hebt sich,
  Zeile anklicken → Bearbeiten. Nebenknöpfe in der Zeile stoppen die
  Weitergabe (`e.stopPropagation()`), damit „Entfernen" nicht öffnet.

### 2b. Frageform — wie wird eine Frage gestellt?

| Die Frage ist … | Dann ist das Bedienelement … |
|---|---|
| Ein/Aus für einen ganzen Bereich („Team-Anmeldung erlauben") | `dex-ui-switch` mit kurzem Satz, was dann passiert |
| Ja/Nein mit Erklärung („Wartelisten-Plätze automatisch nachrücken") | `dex-ui-toggle-row` (Titel + eine Zeile Folge) |
| Eine von 2–4 Alternativen mit Konsequenz („Nur Sub-Events buchbar" vs. „Haupt-Event + Sub-Events") | `dex-ui-choice`-Kacheln nebeneinander (`dex-ui-grid-2`), jede mit Titel + einer Zeile „was das heißt" |
| Eine von vielen kurzen Werten (Sprache, Bezeichnung, Gruppe) | `dex-ui-chip`-Reihe oder `dex-ui-tabs` |
| Mehrere aus vielen (Standorte, Tage) | `dex-ui-chip` mehrfach aktiv |
| Eine Zahl mit Grenzen (Team-Größe 2–20) | kompaktes Zahlenfeld mit Grenzen im Hilfetext, nie ein Dropdown mit 19 Einträgen |
| Freitext, der irgendwo erscheint (Hinweis unter der ID) | Eingabe mit Beispiel im Platzhalter und einem Satz, WO der Text erscheint |
| Eine Entscheidung mit Folgen (Löschen, Versand an alle) | `btn-danger`/`btn-primary` plus Rückfrage, die die Folge nennt („An 47 Personen senden — der gespeicherte Text geht raus.") |

Eine Checkbox ohne Erklärung, ein Dropdown mit zwei Einträgen, ein Radio-
Paar ohne Konsequenz-Zeile: alles Kandidaten für die Tabelle oben.

### 2c. Formulierung — versteht es jemand beim ersten Lesen?

- **Die Beschriftung ist die Frage oder die Aussage, nicht der Feldname.**
  „Wann endet die Anmeldung?" statt „Anmeldefrist"; „Teilnehmer dürfen sich
  selbst abmelden" statt „Selbstabmeldung aktiv".
- **Eine Zeile Folge unter jedem Schalter:** „Dann bekommt jede neue
  Anmeldung ihren QR-Code automatisch." Der Organizer soll wissen, was
  passiert — nicht, wie das Feld heißt.
- **Kurz, konkret, du.** Kein „Bitte beachten Sie", kein „ggf.", kein
  Passiv, wo ein Subjekt geht. Fachbegriffe der App (Klammer, Sub-Event,
  Warteliste, Nachrücken) bleiben, erhalten aber beim ersten Auftreten im
  Schritt eine Halbzeile Erklärung oder ein `InfoTooltip`.
- **Erklärtext ist nachrangig:** maximal zwei sichtbare Zeilen je Block; der
  Rest geht in `InfoTooltip` (Struktur „Was du hier einstellst / Anzeige in
  der App / Auswirkung für Teilnehmer" beibehalten) oder einen Aufklapper.
- **Beide Sprachen gleichwertig.** Jede Änderung am deutschen Text hat ihr
  englisches Gegenstück. Englisch ebenso kurz und direkt („you").
- **Keine Aussage geht verloren.** Kürzen und umstellen ja; eine Folge,
  eine Grenze („max. 20"), eine Warnung weglassen nein. Wer unsicher ist, ob
  ein Satz eine Information trägt: drinlassen, in den Tooltip.

### 2d. Was NICHT verändert werden darf

Es gibt keine Tests; das einzige Netz sind `tsc`, ESLint und der Diff.
Deshalb ohne Ausnahme:

- **Kein Verhalten anfassen**: Handler-Logik, Bedingungen, State-Semantik,
  Effects, Service-Aufrufe, Validierungen, Defaults, was gespeichert wird.
  Ein Block darf wandern und anders aussehen — er bindet danach dieselben
  Werte an dieselben Setter.
- **Exportierte Schnittstellen bleiben**: `Props`-Interfaces, Export-Namen,
  Prop-Namen — nichts hinzufügen, nichts entfernen, nichts umbenennen. Andere
  Dateien werden parallel bearbeitet und kennen deine neue Prop nicht.
- **`visible`-Muster bleibt**: Wizard-Schritte rendern
  `<div style={{ display: visible ? 'block' : 'none' }}>` als Wurzel —
  `display:none` statt Unmount, damit Eingaben beim Schrittwechsel bleiben.
- **`data-tour`, `id`, `aria-*`, `role`, `htmlFor`/`id`-Paare** bleiben
  erhalten (Tour und Tests verweisen darauf).
- **`StepBadge`-Nummern bleiben** vor den Beschriftungen (Support verweist
  darauf), auch wenn Blöcke wandern.
- **Hook-Reihenfolge bleibt** (`react-hooks/rules-of-hooks` ist `error`):
  kein Hook hinter einem frühen Return, keine Hooks in Bedingungen, keine
  Hooks in verschobene IIFEs ziehen.
- **`react-hooks/rules-of-hooks`** ist `error`: kein Hook hinter einem
  frühen Return, keine Hooks in Bedingungen.
- **Kommentare**: Deutsch, `// v31.2: …`, erklären WARUM. Bestehende
  Versionskommentare (`// v28.66: …`) bleiben stehen.
- **Nur die eigene Datei.** Geteilte Dateien (`dexUi.ts`, `Modal.tsx`,
  `WizardFormShell.tsx`, `adminStyles.ts`, `Icons.tsx`, `InfoTooltip.tsx`,
  SCSS-Modul, `EventCreationPage.tsx`, `AdminPage.tsx`) werden zentral
  gepflegt. Fehlt eine Klasse: mit den vorhandenen auskommen und den Wunsch
  im Bericht nennen.

---

## 3. Farben, Abstände, Schrift

| Zweck | Wert |
|---|---|
| Marke / aktiv / primär | `var(--dex-green, #86bc25)` · dunkler `var(--dex-green-dark, #6b9a1e)` · Text auf Weiß `#4a7c1f` |
| Text | `--dex-gray-800` (#333) · sekundär `--dex-gray-600` · Hinweis `--dex-gray-500` · Platzhalter `--dex-gray-400` |
| Linien | `--dex-gray-200` (#e8e8e8) · sanfter Grund `var(--dex-gray-50, #fafafa)` · Flächen `--dex-gray-100` |
| Warnung | `--dex-orange` #ed8b00 · Text `#b35a00` · Fläche `#fff7e6` |
| Fehler | `--dex-red` #da291c · Fläche `--dex-red-light` #fce8e6 |
| Info | Blau `#3860b2` · Fläche `#f4f9fd` |
| Radius | Karten 14 px · Kacheln/Zeilen 12 px · Eingaben 10 px · Chips/Pills 999 px |
| Schatten | Ruhe: keiner oder `0 1px 4px rgba(0,0,0,.04)` · Hover: `0 6px 20px rgba(0,0,0,.07)` · Modal: `0 24px 64px rgba(0,0,0,.22)` |
| Schrift | Titel 1.45 rem/800 · Abschnitt 0.72 rem/700 uppercase · Label 0.88 rem/600 · Text 0.85–0.9 rem · Hinweis 0.78 rem |
| Abstände | Karteninnen 16–18 px · zwischen Karten 12 px · zwischen Abschnitten 22 px · Feld zu Feld 16 px |
| Bewegung | `transition: … 0.18s ease` · Hover-Lift `translateY(-1px)` · Einblenden `dex-ui-fade-in` |

---

## 4. Die Klassen (`dex-ui-*`)

Zustände als zusätzliche Klasse: `is-active`, `is-open`, `is-disabled`,
`is-done`, `is-pending`. Zusammensetzen mit `cx()` aus `dexUi.ts`:
`className={cx('dex-ui-chip', on && 'is-active')}`.

### Karten und Kacheln

| Klasse | Wofür |
|---|---|
| `dex-ui-card` | Weiße Karte mit Linie. Modifier: `--soft` (grauer Grund), `--accent` (grüne Kante links), `--hover` (hebt sich beim Überfahren), `--muted` (60 % Deckkraft, 100 % bei Hover — für ausgeschaltete Bereiche) |
| `dex-ui-choice` | **Eine von mehreren Optionen** (Anmelde-Modus, Kopfbild-Wahl, Hub-Kacheln). `<button type="button">` mit `is-active`. Innen optional `dex-ui-choice-icon`, `dex-ui-choice-body` (`-title`, `-desc`), `dex-ui-choice-check` (Häkchen-Kreis rechts) |
| `dex-ui-toggle-row` | **Ein/Aus mit Text** als `<label>` um eine Checkbox: `<label className={cx('dex-ui-toggle-row', on && 'is-active', off && 'is-disabled')}><input type="checkbox"/><span className="dex-ui-toggle-row-body"><span className="dex-ui-toggle-row-title">…</span><span className="dex-ui-toggle-row-desc">…</span></span></label>` |
| `dex-ui-switch` | Schalter statt Checkbox für **Haupt-Ein/Aus** eines Bereichs: `<label className="dex-ui-switch"><input type="checkbox"/><span className="dex-ui-switch-track"/><span className="dex-ui-switch-label">…</span></label>` |
| `dex-ui-step` | Nummerierte Ablauf-Zeile: `dex-ui-step-num`, `dex-ui-step-body` (`-title`, `-hint`), `dex-ui-step-action` (der Knopf). `is-done` / `is-pending` färben die Nummer |
| `dex-ui-kpi` | Kennzahl-Kachel: `dex-ui-kpi-value`, `dex-ui-kpi-label`. Modifier `--green`, `--orange` |
| `dex-ui-empty` | Leerer Zustand (gestrichelter Rahmen): `dex-ui-empty-icon`, `dex-ui-empty-title` |

### Chips, Pills, Reiter

| Klasse | Wofür |
|---|---|
| `dex-ui-chip` | Umschaltbarer Chip (Sprache, Filter, Punkt-Auswahl) — `<button>` mit `is-active` |
| `dex-ui-pill` | **Nur Anzeige** (Status, Zähler): `--green`, `--gray`, `--orange`, `--red`, `--blue` |
| `dex-ui-tabs` + `dex-ui-tab` | Segment-Reiter (2–5 Ansichten): Container `dex-ui-tabs`, Knöpfe `dex-ui-tab` mit `is-active` |

### Knöpfe

| Klasse | Wofür |
|---|---|
| `btn btn-primary` | DER Hauptknopf (einmal je Ansicht) |
| `btn btn-secondary` / `btn-outline` | Nebenaktionen |
| `btn btn-danger` | Löschen/Verwerfen — mit Rückfrage |
| `dex-ui-btn-sm` | Zusatzklasse für kompakte `.btn` in Zeilen und Modalen |
| `dex-ui-iconbtn` | Runder Symbol-Knopf ohne Fläche (X, Stift, Kopieren). `--danger` färbt rot bei Hover, `--green` grün |
| `dex-ui-textbtn` | Textknopf („Alle anzeigen", „Zurücksetzen"). `--muted` grau, `--danger` rot |

### Listen und Zeilen

| Klasse | Wofür |
|---|---|
| `dex-ui-row` | Zeile mit Hover-Grund: `dex-ui-row-main` (`-title`, `-sub`), `dex-ui-row-actions` (erscheinen kräftiger bei Hover). `--bordered` mit Trennlinie |
| `dex-ui-drag-handle` | Griff zum Sortieren (≡), `cursor: grab` |
| `dex-ui-table` | Tabelle mit ruhigem Kopf und Zeilen-Hover; in `dex-ui-table-wrap` für Rand und Scroll |
| `dex-ui-avatar` | Runde Foto-Kachel 32 px (`--lg` 44 px); `dex-ui-avatar-stack` überlappt mehrere |

### Struktur und Formular

| Klasse | Wofür |
|---|---|
| `dex-ui-section` + `dex-ui-section-title` | Abschnitt mit kleiner Versal-Überschrift und Linie rechts; optional `dex-ui-section-desc` |
| `dex-ui-field` + `dex-ui-label` + `dex-ui-help` | Feldgruppe: Beschriftung (mit `InfoTooltip` daneben), Eingabe, Hilfetext darunter. `dex-ui-label-optional` für „(optional)" |
| `dex-ui-grid-2` / `-3` / `-auto` | Raster, mobil eine Spalte |
| `dex-ui-input` / `dex-ui-select` / `dex-ui-textarea` | Kompakte Eingaben (Modale, Zeilen). `--sm` noch kleiner. Im Wizard-Formular bleiben `form-input`/`form-select` (48 px) für Hauptfelder |
| `dex-ui-disclosure` | Aufklapper-Knopf: `<button className={cx('dex-ui-disclosure', open && 'is-open')}><span className="dex-ui-disclosure-chevron"><ChevronDown size={16}/></span>Weitere Einstellungen<span className="dex-ui-disclosure-count">3</span></button>` und darunter `dex-ui-disclosure-body` (Chevron zeigt zu nach unten, offen nach oben) |
| `dex-ui-callout` | Hinweiskasten: `--info`, `--success`, `--warn`, `--danger`, `--neutral`; Symbol in `dex-ui-callout-icon` |
| `dex-ui-divider`, `dex-ui-muted`, `dex-ui-stack`, `dex-ui-inline` | Trennlinie, Kleintext, vertikaler Stapel (gap 10), horizontale Reihe (gap 8, wrap) |

### Nachzug aus der Runde (Wünsche der Umbau-Agenten, seit v31.2)

| Klasse | Wofür |
|---|---|
| `dex-ui-row.is-active` / `.is-done` | Aktive Zeile (grüne Kante) · erledigte Zeile (gedämpft, bei Hover voll) |
| `dex-ui-rowbtn` | `<button>` als ganze Zeile (Reset: kein Rand, volle Breite, linksbündig) — zusammen mit `dex-ui-row` |
| `dex-ui-btn-reset` | Nackter `<button>` ohne Browser-Optik, z.B. um eine Karte klickbar zu machen |
| `dex-ui-card--list` | Karte als Container für `dex-ui-row`-Zeilen (Innenabstand 6 px) |
| `dex-ui-select--sm` | Kompaktes Select passend zu `dex-ui-input--sm` (Pfeil bleibt frei) |
| `dex-ui-input--error` (auch `-select-`, `-textarea-`) | Roter Rand und roter Fokus-Ring bei ungültiger Eingabe |
| `dex-ui-label-required` | Roter Stern hinter der Beschriftung eines Pflichtfelds |
| `dex-ui-range` | Schieberegler in Markenfarbe |
| `dex-ui-dropzone` (+ `is-over`) | Ablagefläche für Dateien (gestrichelt, hebt sich beim Ziehen) |
| `dex-ui-callout--sm` / `--flush` | Kompakter Hinweis neben einem Knopf · randloser Hinweis innerhalb einer Karte (nur Linie unten) |
| `dex-ui-tab.is-off` | Gedämpfter Reiter für einen abgeschalteten Bereich |
| `dex-ui-iconbtn.is-active` | Gedrückter Symbol-Knopf (z.B. Beschreibung sichtbar) |
| `dex-ui-grid-3-1` / `dex-ui-grid-1-3` | Zwei Spalten 3:1 bzw. 1:3 (Straße/Hausnummer, PLZ/Ort), mobil eine Spalte |
| `dex-ui-disclosure-chevron.is-open` | Chevron-Drehung auch ohne umgebenden `dex-ui-disclosure`-Knopf |
| `dex-ui-chip-remove` | Kleines rundes „×" im Chip (kein `<button>` im `<button>`) |

Der Chevron eines Aufklappers ist im Ruhezustand `ChevronDown` und dreht
geöffnet um 180° nach oben (die Klasse macht das). Wer einen anderen Pfeil
nimmt, setzt die Drehung selbst. Neue Symbole in
Icons.tsx seit v31.2: `ChevronRight`, `ImageIcon`, `Crop`. `.dex-ui-field`
verliert innerhalb eines `dex-ui-grid-*` seinen unteren Abstand (das Raster
hat den `gap`). Der Untertitel des Modals ist ein `<div>` — auch Blöcke sind
erlaubt.

### Modale

`Modal` (components/Modal.tsx) hat seit v31.2 `title`, `subtitle`, `icon`,
`footer`, `hideClose`. **Jedes Modal nutzt sie** — kein eigener `<h3>` mehr,
keine eigene Knopfzeile am Ende:

```tsx
<Modal open={open} onClose={close} maxWidth={560}
  title={isDe ? 'QR-Codes an Teilnehmer' : 'QR codes to attendees'}
  subtitle={isDe ? 'Jede Person bekommt ihren persönlichen Code per Mail.' : '…'}
  icon={<QrCode size={20} />}
  footer={<>
    <button type="button" className="btn btn-secondary" onClick={close}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
    <button type="button" className="btn btn-primary" onClick={send}>{isDe ? 'Senden' : 'Send'}</button>
  </>}>
  …Inhalt…
</Modal>
```

Regeln: Primär-Knopf rechts außen, Abbrechen links davon; ein Knopf, der
etwas Unwiderrufliches auslöst, steht nie direkt neben „Abbrechen" ohne
Abstand (Fuß mit `dex-ui-modal-foot--split` nutzen, wenn links eine
Nebenaktion steht). Inhalt gliedern mit `dex-ui-section`; Kennzahlen als
`dex-ui-pill` oder `dex-ui-kpi`; Abläufe als `dex-ui-step`; Erklärtext als
`dex-ui-callout--neutral` oder Aufklapper. Ein Modal, das beim Öffnen mehr als
einen Bildschirm Text zeigt, ist zu voll.

### Wizard-Schritte

Kopf jedes Schritts — `WizardFormShell` definiert die Klassen:

```tsx
<h2 className="dex-step-head-title">
  <span className="dex-step-eyebrow">{isDe ? 'Schritt 7 von 9' : 'Step 7 of 9'}</span>
  {isDe ? 'Team-Anmeldung' : 'Team registration'}
</h2>
<p className="dex-step-head-lead">…eine bis zwei Sätze…</p>
```

Darunter: `dex-ui-section`s in der Reihenfolge, in der ein Organizer die
Fragen beantwortet — Pflicht zuerst, dann Optionales, dann ein Aufklapper
„Weitere Einstellungen". Große Ein/Aus-Entscheidungen als `dex-ui-switch`
oder `dex-ui-toggle-row`; Ja/Nein-Alternativen mit Erklärung als zwei
`dex-ui-choice`-Kacheln nebeneinander (`dex-ui-grid-2`). Bereiche, die durch
einen Schalter ausgeschaltet sind, bleiben sichtbar, aber `dex-ui-card--muted`.
`StepBadge`-Nummern bleiben vor den Beschriftungen stehen.

### Organizer Center (seit v31.3)

| Klasse | Wofür |
|---|---|
| `dex-ui-page-head` | Seitenkopf: `dex-ui-page-head-title` (h2), `dex-ui-page-head-meta` (Datum · Ort · Pillen), `dex-ui-page-head-actions` (Zurück, Bearbeiten — rechts, weil sie WEG von der Seite führen) |
| `dex-ui-kpi-row` | Raster für Kennzahl-Kacheln (auto-fit, min 140 px). Neue KPI-Modifier `--red`, `--blue`, `--gray`; `dex-ui-kpi-sub` für eine Zeile unter der Beschriftung; `is-clickable` + `is-active`, wenn die Kachel die Liste filtert |
| `dex-ui-toolbar` | Werkzeugleiste über einer Tabelle: `dex-ui-searchbar` (Eingabe mit `dex-ui-searchbar-icon`) links, Filter-`dex-ui-chip`s daneben, `dex-ui-toolbar-spacer`, Export/Spalten rechts |
| `dex-ui-table--compact` | Engere Tabelle für lange Listen. Kopfzellen `is-sortable` (Zeiger, Hover) und `is-sorted` (grün) mit `dex-ui-table-sort` (▲/▼); Zeilen `is-clickable` (nur wenn Klick öffnet), `is-selected`, `is-muted` (abgemeldet/inaktiv); Zellen `is-num` (rechtsbündig, Tabellenziffern), `is-actions` (Symbol-Knöpfe, erscheinen bei Zeilen-Hover kräftiger). `dex-ui-table-wrap--sticky` hält den Kopf beim Scrollen; `dex-ui-table-foot` für Zähler/Seitenwechsel |
| `dex-ui-person` | Personen-Zelle: `dex-ui-avatar` + `dex-ui-person-name` + `dex-ui-person-sub` (E-Mail/Standort), abgeschnitten mit Ellipse |
| `dex-ui-dot` | Statuspunkt vor einem Text: `--green`, `--orange`, `--red`, `--blue` |
| `dex-ui-action-group` | Gruppe von Aktionen mit `dex-ui-action-group-title` (Versal) und `dex-ui-action-grid` (auto-fill, min 240 px) |
| `dex-ui-action` | EINE Aktion als `<button>`: `dex-ui-action-icon`, `dex-ui-action-body` (`-title`, `-desc` = eine Zeile Folge), optional `dex-ui-action-badge` (Pill rechts). `--danger` für Unwiderrufliches; `:disabled` gedämpft — der Grund steht dann in `-desc` |
| `dex-ui-card-head` | Kopfzeile einer Karte/Sektion: `dex-ui-card-head-title` (h3 mit Symbol), `dex-ui-card-head-meta` (Zähler), `dex-ui-card-head-actions` (rechts: Aufklapper-Chevron, Export) |
| `dex-ui-tile` | Navigations-Kachel (Admin Hub): `dex-ui-tile-icon`, `dex-ui-tile-title`, `dex-ui-tile-desc`; ganze Kachel ist der Knopf |
| `dex-ui-progress` | Auslastungsbalken: `dex-ui-progress-bar` mit `style={{ width: pct+'%' }}`, `--orange` ab 90 %, `--red` bei Überbuchung |

---

## 5. Organizer Center

Das Organizer Center (`AdminPage`, `AdminHubPage`, `components/admin/**`) ist
keine Formular-Seite, sondern ein **Arbeitsplatz**: Der Organizer kommt mit
einer Absicht („Wer ist angemeldet?", „QR-Codes raus", „Warum steht da eine
Warnung?") und soll in einem Blick sehen, wo er klickt. Die Grundsätze aus
Abschnitt 1 und die Regeln aus 2a′–2d gelten unverändert; dazu kommt:

### 5a. Reihenfolge der Event-Seite

Von oben nach unten, nach der Frage „was muss ich JETZT wissen oder tun?":

1. **Seitenkopf** (`dex-ui-page-head`): Titel, Status-Pille (Entwurf/Aktiv/
   Vergangen/Archiviert), Datum · Ort · Sub-Event-Zahl als Meta-Zeile; rechts
   „Zurück zur Übersicht" und „Event bearbeiten" (die einzigen Knöpfe, die
   rechts stehen dürfen — sie führen weg von der Seite).
2. **Hinweise, die Handeln verlangen** (`dex-ui-callout--warn/--danger`):
   nicht lesbare Listen, Dubletten, ID-Lücken, Überbuchung, fehlende
   Rechte. Jeder Hinweis nennt die Folge und trägt seinen Knopf IN der
   Zeile („3 Dubletten — Bereinigen"). Reihenfolge nach Dringlichkeit:
   Datenverlust > Rechte > Dubletten > Kosmetik.
3. **Kennzahlen** (`dex-ui-kpi-row`): Angemeldet, Warteliste, Eingecheckt,
   Abgemeldet, freie Plätze — mit Auslastungsbalken. Eine Kachel, die die
   Liste filtert, ist `is-clickable` und zeigt `is-active`. Nicht lesbar
   heißt „–" plus Hinweis, **nie 0** (CLAUDE.md: ein Lesefehler ist keine
   Null).
4. **Nächste Schritte** (`dex-ui-step`, `is-done`): der Ablauf eines Events
   (Einladen → QR versenden → Check-in → Abrechnung) mit dem Knopf in der
   Zeile. Erledigtes bleibt sichtbar, aber gedämpft.
5. **Aktionen** (`dex-ui-action-group`): gruppiert nach Absicht —
   *Kommunikation* (Einladung, Massenmail, QR-Codes), *Teilnehmer*
   (Hinzufügen, Nachrücken, Teams, Import), *Daten & Export* (Excel, Liste
   an F&A, Bescheinigungen), *Einstellungen* (Spalten, Sichtbarkeit,
   Rechte reparieren), *Gefahrenzone* zuletzt (`--danger`: Alle abmelden,
   Event löschen). Jede Aktion: Titel + eine Zeile, was passiert. Eine
   Aktion, die gerade nicht geht, bleibt sichtbar als `:disabled` — mit dem
   Grund in der Folgezeile („noch keine Anmeldungen").
6. **Teilnehmer** — Werkzeugleiste (`dex-ui-toolbar`: Suche links, Status-
   Chips, Spaltenwahl und Export rechts), darunter die Tabelle (5b).
7. **Warteliste** (mit Position und „Nachrücken" in der Zeile), dann
   **Abmeldungen** — beide als Karten mit `dex-ui-card-head` und Aufklapper,
   Abmeldungen standardmäßig zu.
8. **Auswertungen** zuletzt (Anwesenheit je Programmpunkt, Quiz, Hotel,
   Teams): eingeklappt, Kopf mit Zähler.

Die Übersichtsseite (kein Event gewählt) folgt demselben Muster: Kopf mit
„+ Neues Event", Suche und Sortierung als Werkzeugleiste, Events als
klickbare Zeilen oder Kacheln (`dex-ui-row`/`dex-ui-tile`) mit Status-Pille,
Datum, Anmeldezahl; Vergangenes und Archiviertes hinter Aufklappern.

### 5b. Tabellen

- **Ruhiger Kopf, sichtbare Sortierung.** `dex-ui-table` in
  `dex-ui-table-wrap`; sortierbare Spalten `is-sortable` mit Pfeil nur an
  der sortierten Spalte. Lange Listen `--compact` und `--sticky`.
- **Die Person ist eine Zelle.** Avatar, Name, darunter E-Mail oder Standort
  (`dex-ui-person`) — nicht drei Spalten. Status als `dex-ui-pill`
  (grün angemeldet, blau eingecheckt, orange Warteliste, grau abgemeldet,
  rot No-Show). Zahlen rechtsbündig (`is-num`). Aktionen in der letzten
  Spalte als `dex-ui-iconbtn` (`is-actions`), Löschen/Abmelden rot.
- **Zeilen-Hover nur mit Klick.** Öffnet ein Klick auf die Zeile das Detail
  (`ParticipantDetailModal`), ist sie `is-clickable`; Nebenknöpfe stoppen die
  Weitergabe. Reine Anzeige-Tabellen (Auswertungen) behalten den leichten
  Zeilen-Hover der Klasse, aber keinen Zeiger.
- **Leer ist ein Zustand, unbekannt ein anderer.** Keine Zeilen →
  `dex-ui-empty` mit einem Satz, was zu tun ist („Noch niemand angemeldet —
  Einladung verschicken?"). Nicht lesbar (`null`, `regStaleHint`, gesperrte
  Sub-Event-Listen) → `dex-ui-callout--warn` mit den Namen der gesperrten
  Listen; die Zahlen darüber zeigen „–".
- **Fußzeile statt Kopfzähler:** „47 von 120 · 3 gefiltert" in
  `dex-ui-table-foot`. Massenauswahl (Checkbox-Spalte) zeigt ihre Aktionen
  in der Werkzeugleiste, nicht unter der Tabelle.
- **Mobil**: Tabelle scrollt in `dex-ui-table-wrap` horizontal; bestehende
  `isMobile`-Kartenansichten bleiben.

### 5c. Was im Organizer Center NICHT verändert werden darf

Zusätzlich zu 2d — hier steckt die Logik, die CLAUDE.md in zwanzig Fallen
beschreibt:

- **`reloadRegistrations()` ist der einzige Nachlade-Pfad**, `onHttpError`-
  Rückrufe, `regStaleHint`, `deniedSubEventLists`, `null` = unbekannt — kein
  Umbau macht daraus eine leere Liste oder eine 0.
- **Anker und IDs bleiben:** `#admin-waitlist-anchor`, `data-tour`,
  Element-IDs, auf die `JumpButtons`, Deep-Links (`?action=admin&event=…
  &ticket=…`) und die Tour zeigen.
- **`ActionsRegistryProvider`/`ActionTile`-Registrierung** (ActionsMenu):
  Kacheln melden sich beim Mount an; wer eine Kachel umbaut, behält
  `ActionTile` als Wurzel und dessen Props.
- **Bestätigungen und Folgen-Texte** vor destruktiven Aktionen (Alle
  abmelden, Event löschen mit Tipp-Bestätigung, Subsite recyceln) bleiben
  wortgleich in beiden Sprachen.
- **Spaltenkonfiguration (`useColumnConfig`), Sortier-/Filter-State,
  Excel-Export-Spalten** — Darstellung ja, Reihenfolge der Datenfelder im
  Export nein.
- **Hotel-Dateien** (`HotelPlanningPanel`, `HotelSetupWizard`,
  `HotelImportModal`) tragen Alt-Warnungen (`no-explicit-any`): die dürfen
  bleiben, neue kommen nicht dazu. Der `<thead>`/`<tbody>`-Spaltenabgleich
  aus CLAUDE.md (v29.2) ist Pflichtprüfung nach jedem Umbau einer Tabelle
  mit `.map`-Spalten.

---

## 6. Symbole

`components/Icons.tsx` (Inline-SVG, Props `size`, `strokeWidth`):
`ChevronLeft ChevronUp ChevronDown Settings GraduationCap BarChart3 CaptainHat
Mail Star Book QrCode Info Calendar Pin MessageSquare Trash2 Send Plus X
FileText Search Users Download Copy Check Pencil ExternalLink AlertCircle Hash
Shirt Columns Wrench RefreshCw Video Link2`. Dazu Fluent-Icons über
`<Icon iconName="…" />` aus `@fluentui/react/lib/Icon`, wo schon im Einsatz.
Keine neuen Icon-Bibliotheken.

---

## 7. Prüfen vor dem Abschluss

```bash
cd dex-event-app-spfx
./node_modules/.bin/tsc --noEmit -p tsconfig.json 2>&1 | grep -F "<eigene Datei>"   # muss leer sein
./node_modules/.bin/eslint <eigene Datei>                                             # keine neuen Warnungen
```

Häufigste eigene Warnung: `react/no-unescaped-entities` bei `"` in JSX-Text →
`&bdquo;` / `&ldquo;`. Zweithäufigste: ungenutzte Importe nach dem Umbau
(`@typescript-eslint/no-unused-vars`) — entfernen.

Dann den Diff lesen wie ein Prüfer: Ist jeder Handler noch da? Jede
Bedingung? Jeder Text in beiden Sprachen? Jedes `data-tour`? Und einmal die
Seite in Gedanken von oben nach unten lesen: Stimmt die Reihenfolge der
Fragen, weiß der Organizer bei jedem Schalter, was danach passiert, und ist
jede Frage in der Form gestellt, die die Tabelle in 2b vorsieht?
