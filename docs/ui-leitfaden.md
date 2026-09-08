# UI-Leitfaden — Wizard, Organizer Center, Modale, Teilnehmer-Seiten (Stand v31.7)

Dieser Leitfaden ist die **verbindliche Arbeitsgrundlage** für jede Änderung an
der Oberfläche des Event-Wizards, des Organizer Centers, aller Modale und der
teilnehmersichtbaren Seiten. Er
entstand mit v31.2, als rund 60 Dateien parallel modernisiert wurden
(Nutzer-Auftrag 07.09.2026: „jeden einzelnen Schritt im Event-Wizard moderner,
intuitiver und aufgeräumter darstellen, gute Mouseover-Effekte, alle Modale
optimieren"). Wer eine Datei anfasst, hält sich an diesen Leitfaden — sonst
laufen 60 Dateien wieder auseinander.

Die Klassen liegen in `src/webparts/dexEventPlatform/components/dexUi.ts`
(`ensureDexUiStyles()` injiziert sie einmal ins Dokument; `WizardFormShell`
und `Modal` rufen das bereits — auf den Teilnehmer-Seiten ruft es die
Seiten-Komponente selbst, siehe 6). **Neue Klassen: dort ergänzen UND hier
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

## 6. Teilnehmer-Seiten

Startseite (`LandingPage`), Einstieg (`StartPage`), Event-Liste und Kachel
(`EventListPage`, `components/EventCard.tsx`), Anmeldeseite
(`RegistrationPage` + `components/registration/**`) und Meine Events
(`MyEventsPage` + `components/myEvents/**`).

Das Organizer Center ist ein Arbeitsplatz: dieselben Leute, freiwillig, immer
wieder — und wer etwas nicht versteht, fragt. Die Anmeldeseite sieht **jeder,
einmal, meist auf dem Handy, meist zwischen zwei Terminen** und danach nie
wieder. Was hier unklar ist, erzeugt keine Rückfrage, sondern eine
Nicht-Anmeldung, von der niemand erfährt. Die Grundsätze aus Abschnitt 1 und
die Regeln aus 2a–2d gelten unverändert; dazu kommen 6a–6f.

**Geltungsbereich.** Nur die oben genannten Dateien. Ebenfalls
teilnehmersichtbar, aber **nicht** von Abschnitt 6 erfasst und deshalb nicht
„mitmodernisiert": `SelfCheckInPage`/`SelfCheckInDisplayPage` (v30.95, Optik
aus v31.1), `ProfilePage`, Header/Navigation, `InquiryModal`, `TicketsPage`
und das Handbuch (`components/manual/**`).

**Stand.** Diese Dateien haben v31.2 und v31.3 nicht mitgemacht: `dex-ui-`
kommt in `RegistrationPage.tsx`, `registration/EventCard.tsx`,
`PersonalDataSection.tsx`, `EventSpecificSection.tsx`, `TeamSection.tsx`,
`RegistrationActionBar.tsx`, `RegistrationBanners.tsx`, `MyEventsPage.tsx`,
`MyEventCard.tsx`, `MyEventSubEvents.tsx`, `components/EventCard.tsx` und
`StartPage.tsx` **0×** vor. Umgestellt sind nur die Modale
(`ProxyWizardModal`, `MassImportModal`, `SmallModals`, `SubmitConfirmModal`,
`SubEventFieldsModal`, `MyEventsModals`) sowie zwei Einzelkästen
(`EventListPage.tsx:396-435` v31.6, `LandingPage.tsx:1131-1137` v31.4). **Die
Modale sind die Vorlage, die Formular- und Kartenflächen sind die Arbeit.**

**Pflicht vor der ersten `dex-ui-`-Klasse: `ensureDexUiStyles()`.** Das
Stylesheet wird nicht aus dem SCSS geladen, sondern von `dexUi.ts` einmal ins
`document.head` injiziert. Gerufen wird es heute u.a. von `Modal.tsx:52`,
`WizardFormShell`, `AdminPage.tsx:237`, `EventListPage.tsx:368` und
`LandingPage.tsx:28` — **nicht** von `RegistrationPage`, `MyEventsPage`,
`StartPage` und keiner Datei in `registration/` oder `myEvents/`. Wer dort
`dex-ui-callout` setzt, ohne den Aufruf zu ergänzen, liefert unformatiertes
Markup aus, das im Test wie ein Layoutfehler aussieht. Also: **die vier
Seiten-Komponenten rufen `ensureDexUiStyles()` genau einmal, Unterkomponenten
nie** — wie `EventListPage.tsx:366-369` NACH dem frühen Return, mit dem
Kommentar „idempotent, kein Hook".

**Zeilennummern altern, Symbolnamen nicht.** Die Angaben unten stammen aus dem
Stand vor der Runde; nach dem ersten Schnitt stimmen sie nicht mehr. Vor jedem
Eingriff selbst nachzählen (dieselbe Regel wie im Modularisierungs-Rezept in
CLAUDE.md). Und: **es gibt zwei `EventCard.tsx`** —
`components/EventCard.tsx` (425 Zeilen, die Kachel der Event-Liste) und
`components/registration/EventCard.tsx` (474 Zeilen, Station 1 der
Anmeldeseite). Jede Fundstelle wird mit Pfad genannt.

---

### 6a. Reihenfolge — je Seite eine Sollfolge

#### Landing Page (`components/LandingPage.tsx`)

Sortiert nach Dringlichkeit, nicht nach Alter des Codes:

1. **Was heute zu tun ist:** Check-in-Kasten mit QR und Einlassnummer
   (:1138-1183) und „Du bist angemeldet" mit Countdown (:1186-1223). Heute
   stehen sie an vierter Stelle — unter Orb (:1099), tageszeitabhängiger
   Begrüßung (:1103) und zwei Zeilen Willkommenstext (:1106). Am Eventmorgen
   liegt der QR-Code damit unter drei Deko-Blöcken.
2. **Begrüßung, Orb, Willkommenstext.** Deko steht nie zwischen zwei
   Handlungen. Sie bleiben — nur nicht vor der Handlung.
3. **„Start"** (:1227) ist die Handlung der Seite und damit der einzige
   Primär-Knopf (Grundsatz 1.5). Heute ist er `btn-outline`, während der
   Werbekasten „DEX für dein Event nutzen" direkt darunter (:1232-1317,
   Desktop-Zweig) vollflächig grün gefüllt ist — das kehrt die Rangfolge um
   und verstößt gegen Grundsatz 1.1 („Grün nie als Flächenfarbe für Kästen").
   Werbung wird eine `dex-ui-card--hover`, der Start-Knopf bekommt das Gewicht.
4. **Verwaltungs-Hinweise** (inaktive Konten, Entwürfe, Archivierung, Archiv
   aufräumen, ablaufende und zu löschende Teilnehmerlisten, :692-961): bis zu
   sechs gestapelte Kästen mit vier destruktiven Knöpfen. Sie sind
   Organizer-Arbeit und gehören gebündelt in EINEN `dex-ui-callout--warn` mit
   `dex-ui-disclosure` — und auf dem Handy (`position:'static'`, :695) nie vor
   die eigene Anmeldung.

Die Versionsmarke (:682-687) ist absolut positioniert und hat keine Stelle in
der Lesereihenfolge — sie bleibt, wo sie ist.

#### Einstieg (`components/StartPage.tsx`)

Der Kachel-Zweig (`clusters`, :190-195) und der Mobil-Zweig (`rowClusters`,
:209-238) definieren dasselbe Menü zweimal und laufen bereits auseinander
(andere Untertitel, andere `strokeWidth`, **kein einziges `data-tour` im
Mobil-Zweig**). Solange beide bestehen: **jede Änderung an einem Zweig wird im
selben Commit im anderen nachgezogen.** Wer sie zusammenlegt, muss die
`data-tour`-Attribute mitnehmen (6e).

#### Event-Liste (`components/EventListPage.tsx`)

Zugriffs-/Fehlerkasten mit Grund und „Erneut versuchen" (:396-435) → Suche,
Filter, Ansichtsumschalter → „Deine Events" → „Weitere Events" →
Leerzustand. Der Leerzustand ist bereits richtig gebaut (:583-589: nur wenn
`eventsReadStatus !== 'forbidden' && !== 'error'`) — er bekommt `dex-ui-empty`
und einen Satz, was zu tun ist; heute ist er ein grauer Absatz und **hart
deutsch** („Keine Events für dich gefunden.", :587).

#### Event-Kachel (`components/EventCard.tsx`)

Sie beantwortet vier Fragen: *Was? Wann/wo? Kann ich noch buchen? Wen frage
ich?* Rangfolge: Bild/Titel/Ort → Zeitraum → Zustand (angemeldet · N frei ·
voll · Warteliste) → Ansprechpartner → Knopf.

- **Die Geometrie liegt im geteilten SCSS.** `.event-card__image` (feste
  `height: 320px`), `__overlay`, `__title`, `__meta`, `__body`, `__dates`,
  `__deadline` stehen im `:global`-Block von `DexEventPlatform.module.scss`
  (:517-540) — nach 2d zentral gepflegt. Umsortieren heißt deshalb: **innerhalb
  von `__body` verschieben**; alles, was SCSS bräuchte, wird als Wunsch
  gemeldet, nicht lokal gebaut.
- **Der Ansprechpartner darf auf die Kachel — aber nur unter seinen
  Bedingungen.** Heute sind die Organizer nur IM Frist-Overlay als
  Hover-Portal erreichbar (:246-260), also erst, wenn die Anmeldung zu ist, und
  auf dem Handy gar nicht. Wer sie nach oben holt, nimmt `hiddenOrgEmails` und
  `allOrgsHidden` (:144-146, aus `hideOrganizer` + `hideOrganizerIndividualOnly`)
  **mit** — eine unbedingte Zeile „Organizer: …" macht genau die Namen sichtbar,
  die der Wizard verbergen sollte. Und: die Kontaktkarte hängt am Regex
  `/(organizer)/i` über `t('events.deadlinepassed.hint')` (:226); wer sie aus
  diesem Satz löst, ersetzt den Anker im selben Commit durch eine eigene Prop
  (6e).
- **„Frist abgelaufen" steht bewusst zweimal.** Der rote Kasten im Body (:394-408)
  rendert bei `isDeadlinePassed` für ALLE — auch für Organizer und bereits
  Angemeldete; der Overlay (:212) nur für reguläre User
  (`showDeadlineOverlay`, :127). Der Kasten ist keine Dublette und bleibt.
- **Eine Kachel darf nicht dort schweigen, wo die Entscheidung liegt.** Bei
  `subEventsOnlyMode` entfallen Plätze (:380) und Frist (:390) zu Recht
  (v29.13) — an ihre Stelle treten die Zahl der buchbaren Termine und der Satz,
  dass die Anmeldung je Termin läuft. Bezeichnung über `childTermPlural` (6d).
- **Datum:** `formatDate(startDate) … formatDate(endDate)` (:374-379) lässt bei
  leerem `endDate` „… 09:00 bis" plus Leerzeile stehen. Ein Eintages-Event
  bekommt eine Zeile, kein Fragment.

#### Anmeldeseite (`components/RegistrationPage.tsx`)

Die drei nummerierten Stationen (Dein Event → Deine Daten → Anmeldung
abschließen, :2567-2616) erzählen die Geschichte aus 2a richtig und bleiben
unverändert (6e). Die Arbeit liegt **innerhalb** der Stationen:

1. **Station 1** ist reine Anzeige (Bild, Titel, Datum, Ort, Ansprechpartner,
   Beschreibung, Programm). Nichts hinzufügen, was Eingabe ist.
2. **Station 2:** zuerst die eigenen Daten, dann der Team-Schalter. Der
   Stellvertreter-Umschalter und „Massenimport"
   (`registration/PersonalDataSection.tsx:75-133`) sind Organizer-Funktionen,
   stehen heute aber GANZ OBEN im Kartenkopf und per `marginLeft:'auto'` (:106)
   rechts angedockt — beides falsch nach 2a und 2a′. Sie gehören linksbündig in
   einen `dex-ui-disclosure` am Ende der Sektion.
3. **Station 3:** Gruppen-Auswahl → Auswahl von Haupt-Event und Terminen →
   Fragen. Die **zwei Render-Orte** des Hauptfeld-Blocks
   (`EventSpecificSection.tsx:425-427` und :913) bleiben — sie sind Absicht
   (v29.28).
4. **Fehlerkasten, Aktionsleiste, Datenschutz-Fußnote** zuletzt. Der
   Fehlerkasten (:2632-2636) bekommt `role="alert"` und scrollt ins Bild; auf
   dem Handy ist ein Kasten unter dem Formular sonst unsichtbar.

**Die Reihenfolge der eventspezifischen Felder wird NICHT nach `required`
sortiert.** Sie ist Organizer-Absicht (Drag-Sortierung und „Nach oben"/„Nach
unten" in `wizard/steps/FieldsStep.tsx`), und der `showIf`-Filter (:1533-1539)
blendet Felder abhängig von der Antwort eines **anderen** Feldes ein — eine
Umsortierung stellt das Folgefeld über seine Auslöserfrage. Verbessert wird die
Darstellung (Raster, Abstände, Pflicht-Markierung), nicht die Sortierung.

**Die zwei Team-Wege sind sachlich eine Frage** — „eigenes Team anmelden"
(`PersonalDataSection.tsx:382-405`) und „offenem Team beitreten"
(`RegistrationPage.tsx:2600`) schließen sich gegenseitig aus
(`togglePendingJoinTeam`). Zusammenlegen heißt aber: zwei Dateien plus eine
neue gemeinsame Sichtbarkeitsbedingung (die heutigen sind nicht deckungsgleich)
— das verletzt 2d („nur die eigene Datei", „kein Verhalten anfassen") und
gehört in ein eigenes Ticket. In dieser Runde: beide Wege bekommen dieselbe
Form und stehen direkt untereinander, mit einem Satz, dass man sich für einen
entscheidet.

#### Meine Events (`MyEventsPage.tsx`, `myEvents/MyEventCard.tsx`)

Seite: Ladezustand → Ladefehler → seitenweite Kästen → „Aktive Events" →
„Vergangene Events" → Aufklapper „Abgemeldete Events".

- **Seitenweite Kästen hängen nie an der Zahl der eigenen Anmeldungen.**
  „Offene Anforderungen an mich" (:1150-1183) und „Von deiner Assistenz
  verwaltet" (:1184-1222) liegen innerhalb von `activeEntries.length > 0`
  (:1126): Wer selbst nirgends angemeldet ist, aber für jemanden etwas
  erledigen soll, sieht statt seiner Aufgabe den Leerzustand.
- **„Vergangene Events" sagt heute nur, was NICHT mehr geht** („eine Abmeldung
  ist hier nicht mehr möglich", :1238-1242). Wonach man nach einem Event sucht
  — Teilnahmebescheinigung, Dokumente — steht nicht dort, obwohl die
  Bescheinigung genau dort existiert (`MyEventCard.tsx:759-780`).
- **Die Sortierung bleibt vorerst.** `getEvents` liefert `$orderby=StartDate
  desc` (`services/events/eventsCrud.ts:91`), die Seite reicht das ungefiltert
  durch (:903-908) — das morgige Event steht unten, die „sessionsOnly"-Karten
  aus der zweiten Schleife (:530-575) ohne Datumsbezug ganz hinten. Das ist
  eine **Verhaltensänderung** (2d) und hängt zusätzlich am Tour-Selektor
  `.my-event-card` (6e): eigener Commit, eigenes Review, nicht in dieser Runde.

Karte (`myEvents/MyEventCard.tsx`, 1.286 Zeilen) — Sollfolge, heutiger Ist
daneben:

1. **Kopfzone:** Bild, Titel, Status-Pille, **Wann · Wo · Teams-Knopf · „Mein
   QR-Code"**, „Angemeldet am …" als Metazeile. Heute stehen Ort und Datum erst
   nach der Team-Box (:311-547, mit Fotozeilen mehrere hundert Pixel), nach
   Unterkunft (:243-281) und zwei Hinweiskästen (:228-238, :286-305); der
   QR-Knopf ist der vierte Chip einer Chip-Zeile weit unten (:740-754). Die
   Kopfzone muss den Fall „noch kein QR" aushalten — der Knopf hängt an
   `Status ∈ {QR versendet, Eingecheckt}` (v28.7) und bleibt daran.
2. **Bei einem Klammer-Event: die Terminliste** (:1066). In `subEventsOnlyMode`
   IST sie die Anmeldung; heute rendert sie nach Programm (:866),
   Transferzeiten (:951), Dokumenten (:987) und Quiz (:992).
3. **Hinweise, die eine Einschränkung erklären** („für dich angelegt",
   „nur für Termine angemeldet") als `dex-ui-callout`, höchstens zwei sichtbare
   Zeilen. Der `hiddenRow`-Hinweis (:228-238) ist heute ein Fünfzeiler mit vier
   Fettungen.
4. **Team, Unterkunft, Ansprechpartner** — sie ergänzen die Anmeldung, sie sind
   nicht die Anmeldung. „Organizer" (:597-611) und „Ansprechpartner"
   (:617-639) beantworten dieselbe Frage in zwei Optiken: ein Block. Beide
   Blöcke behalten ihre Bedingungen (`hideOrganizer`,
   `hotelVisibleToAttendees`, siehe 6e).
5. **Inhalt:** Beschreibung, Programm, Transferzeiten, Dokumente, Uploads,
   Quiz. Die Gewichtung der Aufklapper ist heute umgekehrt: Beschreibung zu
   (:648-667), Team-Liste und Agenda immer offen.
6. **Aktionszeile zuletzt, links beim Inhalt.** „Abmelden" (:1237-1268) steht
   heute per `justify-content: space-between` rechts neben „Angemeldet am …"
   (:1107-1112) — genau das Muster, das 2a′ verbietet.

**Der Umbau dieser Karte ist eine eigene Aufgabe, nicht ein Auftrag neben zehn
anderen.** Es gilt die CLAUDE.md-Regel für große JSX-Blöcke: den **ganzen**
Block als Einheit bewegen, vorher eine Kopie ins Scratchpad, nach jedem Schnitt
`tsc`.

---

### 6b. Handy — was nur beim Überfahren erscheint, existiert nicht

Das ist die Regel, die diese Seiten von den Organizer-Seiten unterscheidet, und
sie wird lautlos verletzt. Heutige Fälle:

- Die Kalender-Tage melden an und ab; das Wort „abmelden" erscheint nur im
  Hover (`myEvents/MyEventSubEvents.tsx:617`), dazu ein `title` (:547-575). Auf
  dem Handy steht dort ein „✓", und ein Fingertipp meldet ab.
- Der Grund einer gesperrten Anmeldung steht nur im `title` des Knopfs
  (`registration/RegistrationActionBar.tsx:124-140`); der Knopf selbst nennt
  ihn in zwei von sechs Fällen.
- Die Folge von „Ich nehme nicht teil" steht nur im `title` (:214-220) — sie
  sagt für das GANZE Event inklusive aller Termine ab
  (`RegistrationPage.tsx:1614-1619`).
- „Belegung nicht ermittelbar" steht nur im `title`
  (`MyEventSubEvents.tsx:769`, `registration/EventSpecificSection.tsx:586-589`).

**Ein `title`-Attribut ist eine Zugabe, nie der einzige Träger einer
Information.** Was den Zustand oder die Folge erklärt, steht als Text in der
Zeile darunter oder in einem `dex-ui-callout`.

Weiter:

- **Bedienelemente nach der 2b-Tabelle.** Die Anmeldeseite besteht aus rohen
  Checkboxen mit Inline-Styles (also ohne Hover): Team-Anmeldung
  (`PersonalDataSection.tsx:385`) → `dex-ui-switch`; Haupt-Event-Haken
  (`EventSpecificSection.tsx:392`), Termin-Haken (:720), Leistungsnachweis
  (:235), Team-Zustimmung (`TeamSection.tsx:153`) und Checkbox-Custom-Fields
  (`RegistrationPage.tsx:2324`) → `dex-ui-toggle-row`; die Gruppen-Kacheln
  (`EventSpecificSection.tsx:180-227`) → `dex-ui-choice` in `dex-ui-grid-2`.
  Mehrfachauswahl als `dex-ui-chip`-Reihe — `SubEventFieldsModal.tsx:141-163`
  zeigt für genau diese Feldtypen schon, wie es aussieht.
- **Eine gesperrte Aktion bleibt ein sichtbarer `:disabled`-Knopf mit
  unveränderter Beschriftung**, der Grund steht in der Zeile darunter. Kein
  Knopf, dessen Beschriftung die Fehlermeldung ist („Abmeldung gesperrt",
  `MyEventSubEvents.tsx:820`), und kein `<span>` in Knopf-Optik an der Stelle
  des Knopfs (`MyEventCard.tsx:1190-1234`, drei Stück mit Rahmen, Radius und
  Knopf-Innenabstand direkt in der Knopfzeile). Deren Texte nennen jeweils den
  echten Grund und bleiben inhaltlich wortgleich — nur die Form ändert sich.
- **Entscheidungen mit Folgen brauchen die Rückfrage aus 2b.** „Ich nehme nicht
  teil" (`RegistrationActionBar.tsx:209-227`) ist heute ein `btn-secondary`
  ohne jede Rückfrage. Also `confirmDialog` mit `danger: true`, dessen Text die
  Folge mit `childTermPlural` benennt. Hinweis: `.btn-danger` ist in dieser App
  bewusst **grau** (`DexEventPlatform.module.scss:324-325`, im Modal-Overlay
  mit `!important`, v24.63) — die Warnung trägt der Dialogtext, nicht die
  Farbe. Kein lokal gebauter roter Knopf und kein roter Text auf
  `btn-secondary`.
- **Keine nativen Datums-/Zeitfelder** (Grundsatz 8) — mit Formatauflage:
  `RegistrationPage.tsx:2425-2435` rendert `<input type="date">` bzw.
  `datetime-local` für Custom-Felder. Der Wert ist die **Antwort**, die als
  String in `CustomData` landet und im Organizer Center, im Excel-Export und in
  der Bearbeiten-Ansicht wieder gelesen wird. Ein Ersatz zeigt deutsch
  (`dd.MM.yyyy`) und **speichert weiterhin zeichengleich** (`YYYY-MM-DD` bzw.
  `YYYY-MM-DDTHH:mm`) — sonst sind Alt- und Neu-Antworten unvergleichbar
  (dieselbe Roundtrip-Falle wie `{{Organizer}}`, v30.74). Vorbild für die
  Anzeige: `StayRangePicker` (:2416).
- **Die frühen Return-Seiten sind vollwertige Seiten** — für viele Menschen die
  einzige Seite der App, die sie je sehen. Vorlage sind
  `RegistrationPage.tsx:1280-1303` (Bildband, Symbol, Titel, Grund, Knopf
  zurück) und der Erfolgsschirm :1908-2061. Nachzuziehen: **„Event nicht
  gefunden"** (:1242-1249) — bloßer Text ohne Karte, ohne Symbol und ohne einen
  Satz, warum (gelöscht? kein Zugriff? alter Link?); die **Absage-Seite**
  (:1736-1754) — Karte ohne Bild und Symbol; die **Team-Beitritts-Seite**
  (:1777-1791) — gestaltet, aber ohne jeden Knopf, also eine Sackgasse.
- **Jede dieser Seiten rendert den Zustand, mit dem sie erreicht wurde.**
  `handleDecline` (:1587-1634) setzt bei Teilfehlschlag `setDeclined(true)`
  **und** `setError("… bei N Terminen hat es nicht geklappt")` — die
  `declined`-Seite (:1735-1754) gibt `error` nirgends aus. Die Person liest
  „Absage erfasst", obwohl Termine offen blieben. Dasselbe für
  `submittedAsWaitlist`, `submittedSessionsRef`, `submittedWaitlistRef`.
- **Der Mobil-Zweig ist eine eigene Ansicht.** `StartPage.tsx:209-238`,
  der `isMobile`-Zweig der Zeilenansicht in `EventListPage` und
  `registration/regHelpers.tsx:154-214` (Sektionen klappen NUR auf dem Handy
  ein) werden im selben Commit nachgezogen wie der Desktop-Zweig — sonst
  driften sie weiter auseinander.

---

### 6c. Zustände — leer, unbekannt, Warteliste, vorbei

- **Statusnamen kommen nie roh aus SharePoint.** `getStatusLabel`
  (`myEvents/myEventsHelpers.tsx:130-138`) hat keinen Fall für „QR versendet"
  und „No-Show" und gibt über `default: return status` den rohen Wert aus — im
  englischen UI steht dann deutsch „QR versendet". Die Schlüssel existieren
  längst in beiden Sprachen (`LanguageContext.tsx:207-212` / :660-665:
  `status.registered`, `.qrsent`, `.waitlist`, `.checkedin`, `.noshow`,
  `.cancelled`): **Das ist ein Ein-Datei-Fix ohne neue Übersetzungsschlüssel.**
  Ebenso `getStatusBadgeClass` (:119-127) — heute bekommt „QR versendet"
  dieselbe grüne Klasse wie „Angemeldet" und No-Show fällt auf grau durch,
  ausgerechnet der Zustand mit dem größten Erklärbedarf.
  **Der `default`-Zweig bleibt** und gibt weiterhin den Rohwert zurück: `Status`
  ist eine SharePoint-Choice, die erweitert werden kann — ein unbekannter Wert
  muss sichtbar bleiben, nicht zu „" oder „Unbekannt" werden.
- **Warteliste ist kein Zustand der Farbe allein.** Heute ist nur das Badge
  orange (`MyEventCard.tsx:217-221`), die Fußzeile sagt weiter „Angemeldet am
  …", alle Aktionen sind dieselben. Text und Fußzeile benennen es. Die Position
  bleibt an `maxParticipants > 0` gebunden (:218) — bei geteilten Kapazitäten
  ist der Wert 0 (CLAUDE.md), dort wäre die Rechnung Unsinn; lieber keine
  Position als eine falsche.
- **Ein Lesefehler ist keine Null — auf diesen Seiten lädt er zur
  Doppelanmeldung ein.** `getMyEventNumbers`/`getEventNumbersForEmail`
  (`context/actions/participantFiles.ts:35-64`) fangen jeden Fehler und liefern
  leere Arrays; `getRegistrationCount`
  (`services/events/registrationStatus.ts:715-739`) bricht bei `!response.ok`
  mit `break` ab und liefert `{registered:0, waitlist:0}`. Beide werfen nicht —
  ein 403 auf `DEX_Participants` liest sich deshalb als „du bist nirgends
  angemeldet" (Kachel ohne Overlay, Knopf „Registrierung starten"), ein
  gedrosseltes Lesen als „140 frei" bei vollem Event
  (`components/EventCard.tsx:113`, :386). **Die Reparatur der Verträge liegt in
  geteilten Dateien mit vielen Aufrufern (inklusive des Handbuch-Stubs) und
  gehört in ein eigenes Ticket.** Für die Oberfläche gilt bis dahin: **kein
  neuer Schluss auf ein leeres Ergebnis** — keine neue Meldung, keine neue
  Sperre, keine neue Zahl, die aus `[]` oder `0` „niemand" bzw. „frei" macht.
- **Rollen und Rechte nie aus einer leeren Datenliste schließen.** `StartPage`
  prüft `rolesReadStatus === 'forbidden'` (:137-143, v30.81), leitet aber
  `isOrganizerOfAnyEvent` (:35-39) und `isCheckInTeamOfActive` (:45-50) aus
  `events` ab: Bei einem 403 auf DEX_Events ist `events` leer, die
  Organizer-Kachel wird ausgegraut und bewirbt ausgerechnet dem Organizer
  gegenüber „Organizer werden?", die Check-in-Kachel verschwindet.
  `eventsReadStatus` gehört genauso geprüft wie `rolesReadStatus`.
- **Leer wird nur behauptet, wenn der Lesevorgang nachweislich geglückt ist.**
  Vorbilder im eigenen Bestand: `EventListPage.tsx:583-589` und der
  Grund-Kasten :396-435 mit „Erneut versuchen"; `LandingPage.tsx:303-320`
  (behält bei Lesefehler den gecachten Stand) und der Check-in-Kasten mit
  `checkInBoxError` (:414, :1131-1137). Leere Listen bekommen `dex-ui-empty`
  mit einem Satz, was zu tun ist; Ladefehler `dex-ui-callout--danger` mit
  Wiederholen-Knopf — heute ist der Ladefehler in „Meine Events" eine rote
  Textzeile ohne Symbol, ohne Knopf und ohne englische Fassung (:634,
  :1113-1117).
- **„Unbekannt" ist ein eigener Zustand und wird benannt.** `count === null`
  heißt weder „voll" noch „N frei" — `MyEventSubEvents.tsx:501-505/721-724`
  rechnet das richtig, zeigt es aber nur als „—" plus `title` (:769); ebenso
  `EventSpecificSection.tsx:210/663/790-791`. Daneben gehört ein sichtbarer
  Satz bzw. ein `dex-ui-callout--warn`. Und **nie 0**.
- **Jeder Zustand außer dem Normalfall wird an drei Stellen sichtbar: Pille,
  Karten-Ton, Aktionszeile.** Vergangenes und Abgemeldetes ist
  `dex-ui-card--muted`, **behält aber Datum, Ort und Bild** — sonst ist „wann
  war das noch mal?" nicht mehr zu beantworten. Heute sehen vergangene Karten
  aus wie aktive (nur Sektionsüberschrift und ein grauer Satz unterscheiden
  sie), während die abgemeldete Zeile alles außer Titel und Abmeldedatum
  verliert (`myEvents/CancelledEventsCollapsible.tsx:60-79`) und „Zur
  Anmeldung" in die Liste statt zum Event führt (`MyEventsPage.tsx:1263-1264`).
- **Der Zustand „angemeldet" wird markiert, nicht verdeckt — aber der Overlay
  ist mehr als Optik.** `components/EventCard.tsx:280-321` enthält ZWEI Knöpfe
  („Meine Events" und, für `canCreateEvents || isOwnOrganizer`, „Für andere
  Person registrieren" — die Quelle des navIntent `register-other`, :313), und
  die Kachel-Wurzel navigiert bei `alreadySignedUp` bewusst nirgendwohin
  (:157). Wer auf grüne Kante plus Pille umstellt (wie
  `EventListPage.tsx:630-636`, :688-697), muss beide Knöpfe und das Klickziel
  `my-events` (`EventListPage.tsx:613-615`) im selben Zug übernehmen — sonst
  bleibt eine Karte, die auf Klick nichts tut. Sperren (Frist, „Anmeldung ab")
  dürfen weiter verdecken; sie blockieren den Klick auch optisch.
- **Kachel und Zeile sagen dasselbe.** Beide lesen dieselbe Liste, weichen aber
  ab: freie Plätze abzüglich Warteliste und mit Split-Kapazität
  (`components/EventCard.tsx:95-113`) gegen `event.maxParticipants || '∞'`
  (`EventListPage.tsx:685` — bei geteilten Gruppen ist der Wert 0, die Liste
  behauptet „unbegrenzt"); Frist und „Anmeldung ab" sperren die Kachel und
  fehlen in der Liste ganz; das Klickziel ist einmal die Anmeldeseite, einmal
  `my-events`. Begründet sind nur Bildgröße und Zeilendichte. Wer eine Angabe
  ändert, zieht die andere Ansicht im selben Commit nach; das Ziel ist eine
  gemeinsame Quelle (`eventCardFacts(event)`) — das ist eine Code-Änderung und
  gehört in ein eigenes Ticket, nicht in eine Optik-Runde.
- **Hinweiskästen nur als `dex-ui-callout`, Farbe gleich Bedeutung**, höchstens
  zwei sichtbare Zeilen (orange = Warnung, blau = Info, grün = Erfolg/aktiv,
  rot = Gefahr). Heute stehen an derselben Stelle vier handgebaute Kästen
  (`MyEventCard.tsx:228`, :257, :296, :338) plus zwei auf Seitenebene
  (`MyEventsPage.tsx:1151`, :1185), und Grün heißt einmal Hotel, einmal Team,
  einmal Assistenz.
- **Anzeige ist `dex-ui-pill`, Schaltbares ist `dex-ui-chip`** (Grundsatz 3).
  Die Antwort-Tags (`myEventsHelpers.tsx:22-29`, heute ein handgebautes Tag mit
  `borderRadius: 4` und ohne Klasse) und der Chip „Gruppe: Durchstarter"
  (`MyEventCard.tsx:586-592`) sind reine Anzeige — der echte Umschalter „Gruppe
  wechseln" steht 550 Zeilen tiefer (:1155-1176). Umgekehrt kein Hover ohne
  Klick: der 240-%-Foto-Zoom auf Team-Mitgliedern (:398-412, ebenso
  `MyEventsModals.tsx:282-296`) und der 260-%-Zoom im Personen-Antwort-Tag
  (`myEventsHelpers.tsx:39-40`) versprechen eine Aktion, die es nicht gibt.
- **Aufklapper sind `dex-ui-disclosure` mit `ChevronDown`, klickbare Zeilen
  `<button>` bzw. `dex-ui-row` mit Hover, Zeiger und Tastaturfokus.** Ein ▶ oder
  ▲▼ im Text ist kein Bedienhinweis, ein `<div onClick>` ist keine Zeile:
  `CancelledEventsCollapsible.tsx:26-43` (ganze Sektion),
  `MyEventCard.tsx:652-667` (Beschreibung), `myEvents/DocumentsViewer.tsx:128-146`
  (Dokumentzeile öffnet die Vorschau, der Pfeil ist ein `<span>`),
  `MyEventSubEvents.tsx:838-858` (Gruppenköpfe), `myEvents/QuizPlayer.tsx:340-359`
  (Antwortknöpfe ohne Hover).

---

### 6d. Formulierung und Sprache

- **Guard-Meldungen nennen den tatsächlichen Grund — und das ist hier bereits
  erreicht.** `registration/submitFlow.ts:192-244` fächert „nichts
  abzuschicken" in fünf benannte Fälle auf, `regFailMessage` (:771-811) nennt
  `not-allowed`, `deadline`, `dup-check-failed`, `insert-failed`,
  `already-registered` und `full` einzeln; der alte Sammelsatz
  (`LanguageContext.tsx:402`) steht bewusst nur noch als letzter Zweig (:242).
  **Nichts davon zu einer Meldung zusammenfassen** — das haben v28.88, v29.9,
  v29.13 und v30.67 nacheinander repariert. Offen ist genau eine Meldung: „Bitte
  alle Pflichtfelder ausfüllen." (`LanguageContext.tsx:130`, gerufen
  `submitFlow.ts:290` und :395) nennt kein Feld, obwohl dieselbe Datei es für
  Custom-Felder namentlich tut (:379, :438).
- **Erst was du tun kannst, dann die Mechanik.** „Ein Event kannst du jederzeit
  nachträglich an- oder abmelden" (`MyEventSubEvents.tsx:410-416`) verspricht
  „jederzeit", wo Fristen gelten, und erklärt Mail und Outlook, bevor es sagt,
  was möglich ist; die Einschränkung kommt als Anhängsel und nur, wenn gerade
  etwas gesperrt IST (:400-409).
- **Ganze Sätze statt Kurzcodes.** „Abmeldefrist war am 12.10." statt „war bis
  12.10." (`MyEventSubEvents.tsx:622`), „Selbst-Abmeldung gesperrt" statt „fix"
  (:617).
- **Etiketten werden Fragen (2c) — aber nur die App-Texte.** „Persönliche
  Informationen" und „Event-spezifische Informationen"
  (`LanguageContext.tsx:83-84`) sind Feldnamen; der zweite ist zusätzlich
  irreführend, weil dort die Termin-AUSWAHL steht. „Aktive Events"
  (`MyEventsPage.tsx:1226`) meint `upcomingEntries`, also kommende — ein
  Wartelisten-Eintrag steht dort unter „aktiv". „Registrierungen"/„registriert"
  (`LanguageContext.tsx:157-158`) heißen in dieser App Anmeldungen. **Nicht
  angefasst werden Organizer-Daten**, die als Überschrift dienen:
  `event.splitSectionTitle` (`EventSpecificSection.tsx:102`), `splitHelpText`
  (:107-110) und `childTermPlural` (:344-350) gewinnen vor dem App-Text — nur
  der Fallback-Zweig wird umformuliert.
- **Bezeichnungen kommen aus den Term-Konstanten**, inklusive Artikel:
  `childTermSingular`/`childTermPlural`/`childOneDe`
  (`RegistrationPage.tsx:468-497`, `MyEventSubEvents.tsx:392-394`,
  `MyEventCard.tsx:200/287`) und `agendaTermSingular`/`agendaTermPlural`
  (:775/881). Fest verdrahtet sind heute „Sub-Event"
  (`MyEventSubEvents.tsx:227/234/920/923/1035`, Badge-Rückfall
  `MyEventCard.tsx:203`, `LanguageContext.tsx:406-407`), „Session ohne Titel"
  als Fallback-Titel (`MyEventSubEvents.tsx:826`) und „Event-Section"
  (`MyEventsModals.tsx:412-414`) — ein Begriff, den es sonst nirgends in der App
  gibt.
- **Sprachquelle ist ausschließlich `isDe`/`t()` aus dem `LanguageContext`** —
  nie `event.emailLanguage` (das ist die Mailsprache; heute lesen
  `MyEventUpload.tsx:23`, `MyEventDocField.tsx:22` und `MyEventCard.tsx:623`
  daraus die UI-Sprache, sodass der Upload-Kasten englisch sein kann, während
  die Seite deutsch ist; `MyEventSubEvents.tsx:40-46` erklärt im Kommentar,
  warum das falsch ist).
- **Kein Text ohne Gegenstück.** Hart deutsch: „Registrierungen konnten nicht
  geladen werden." (`MyEventsPage.tsx:634`, dazu ohne einen Hinweis, was zu tun
  ist), „Keine Events für dich gefunden." (`EventListPage.tsx:587`) sowie
  „Entwurf" (:665), „Teilnehmer" (:685), „Organizer:" (:673), „Angemeldet"
  (:690), „Warteliste" (:695) in der Zeilenansicht. Hart englisch: „So far used
  for…" im Boot-Loader (`DexEventPlatform.tsx:863-870`, dazu `locale="en"` fest
  verdrahtet).
- **Wer `LanguageContext.tsx` anfasst, ist EINER.** Fast jede Textregel oben
  endet in derselben Datei mit zwei Sprachblöcken — bei parallelen Agenten ist
  das die Datei, in der sie sich gegenseitig überschreiben. Also: Änderungen an
  `t()`-Schlüsseln sammelt ein Vorab-Commit oder ein einzelner Agent; die
  Seiten-Agenten ändern nur den JSX-Text ihrer eigenen Datei.
- **Typografie:** `&bdquo;…&ldquo;` statt gerader Anführungszeichen (heute u.a.
  `MyEventCard.tsx:347-348/1160/1169-1170`, `MyEventsPage.tsx:1166`,
  `MyEventUpload.tsx:66`, `MyEventDocField.tsx:47`), echte Ellipse statt „...",
  und kein Ladezustand, der einen Knopf zu „…" macht
  (`MyEventCard.tsx:1264`, `LanguageContext.tsx:157/161/166`).

---

### 6e. Was auf den Teilnehmer-Seiten NICHT verändert werden darf

Zusätzlich zu 2d. Diese Seiten hängen an mehr Außenwelt als jede andere Fläche
der App: Tour, Deep-Links aus Mails, Handbuch, geparste Texte,
Datenschutz-Schalter.

- **Kein Hook hinter einem frühen Return — und die Returns kippen zur
  Laufzeit.** `RegistrationPage.tsx`: letzter Hook `searchUsersInAudience`
  (heute :1214-1218), erster Return `if (!event)` (:1220), dann :1278
  `notYetActive`, :1306 `isFullyClosed`, :1735 `declined`, :1759
  `submitted && submittedJoinKind`, :1794 `submitted`; Haupt-Return :2542.
  Zwischen dem ersten Return und dem Dateiende steht heute kein einziges
  `React.use*` — auch nicht versteckt: `createSubmitFlow` (:1562) und die
  Props-Bündel (:2447-2541) dürfen **niemals** zu `useMemo`/`useCallback`
  werden. Der Warnkommentar :1158-1165 bleibt stehen. Dasselbe Verbot:
  `StartPage.tsx:240` (`if (isMobile)` — kippt bei Resize UND über
  `window.__dexForceMobile` der Handbuch-Vorschau), `MyEventsPage.tsx:910`
  und :951, `EventListPage.tsx:346`, `regHelpers.tsx:167` vor :171,
  `MyEventUpload.tsx:39`, `QuizPlayer.tsx:189`, `RegisterPreviewModal.tsx:262`.
  `ensureDexUiStyles()` ist **kein** Hook und darf nach dem Return stehen.
- **Die Deklarationskette zwischen den Returns ist Abhängigkeit, kein Zufall:**
  `errorBorder` → `parentAlreadyRegistered` → `parentRegBlocked` →
  `sessionsChanged` → `nothingToSubmit` → `renderMainFieldsSection` →
  `createSubmitFlow` → `handleDecline` → `renderRegField` → Props-Bündel.
- **Die drei `beforeunload`-Wächter bleiben** mit State-Flag und Cleanup:
  `RegistrationPage.tsx:318` (`isSubmitting`), `MyEventsPage.tsx:329`
  (`cancelProgress`), `MyEventSubEvents.tsx:70`. Ohne sie bricht ein Tabwechsel
  eine halbe An-/Abmeldung ab und hinterlässt Termine ohne Klammer. Ebenso der
  Live-Counter-Effect `RegistrationPage.tsx:520-542` — sein Rückgabewert ist das
  `subscribeEventRealtime`-Unsubscribe plus das Entfernen des `focus`-Listeners.
- **Tour-Anker — drei davon sind Struktur, kein `data-tour`.**
  `data-tour="landing-start"` (`LandingPage.tsx:1227`) und
  `tile-register`/`tile-myevents`/`tile-admin`/`tile-checkin`
  (`StartPage.tsx:97/104/111/148`) bleiben am selben Element. Dazu zeigen
  `components/tutorial/tutorialTours.ts:74/82/98` auf `.event-grid`,
  `.event-grid > *:first-child` und `.my-event-card`: Die Kartenansicht muss
  `.event-grid` heißen (`EventListPage.tsx:535`), ihr erstes Kind muss die
  `EventCard`-Wurzel bleiben (kein Wrapper-`div` dazwischen), und die AKTIVEN
  Karten (`MyEventCard.tsx:157`) müssen im DOM VOR den abgemeldeten stehen
  (`CancelledEventsCollapsible.tsx:63`) — sonst hebt die Tour eine abgemeldete
  Anmeldung hervor. Ein Selektor ins Leere wirft nicht, er zeigt nichts
  (`TutorialGuide.tsx:190/203`).
- **`id={\`dex-myevent-${event.id}\`}` (`MyEventCard.tsx:157`) ist ein
  Deep-Link-Ziel, keine Dekoration.** `MyEventsPage.tsx:897-898` scrollt darauf,
  angestoßen von `?action=cancel&event=<Nr>` aus der Outlook-Absage-Mail
  (`DexEventPlatform.tsx:435-446`); die abgemeldete Karte trägt bewusst keine
  id. `.main-content` bleibt der Scroll-Container (`RegistrationPage.tsx:57`,
  definiert `DexEventPlatform.tsx:1065`).
- **Die globalen Layout-Klassen bleiben** (`:global`-Block des SCSS-Moduls,
  zentral gepflegt): `.page-container`, `.event-card` samt
  `__image/__overlay/__title/__meta/__body/__dates/__deadline/__register-btn`,
  `.event-grid`, `.my-event-card` samt
  `__thumb/__header/__details/__specific/__actions`, `.start-card` samt
  `__icon/--admin/--checkin`, `.landing` samt
  `__hero/__card/__orb/__text/__actions`, `.section-header`, `.badge` +
  `-green/-orange/-red/-gray`, `.my-events-list`, `.dex-cancel-btn`. Neue Optik
  über zusätzliche `dex-ui-`-Klassen, nie durch Umbenennen; fehlt eine Klasse,
  wird sie im Bericht gewünscht (offen: ein dunkler Karten-Overlay, dreimal
  wortgleich in `components/EventCard.tsx:213/264/281`).
- **Die drei nummerierten Stationen der Anmeldeseite bleiben, wo und wie sie
  sind.** `.reg-step-num` ist absolut auf `left:-44px` positioniert und
  funktioniert NUR als Kind von `.registration-layout`
  (`RegistrationPage.tsx:2567-2616`, SCSS :548-591); ein Stationskopf, der aus
  dem Container wandert, hat seine Nummer außerhalb des Bildschirms. Nummern
  und Reihenfolge bleiben — wie `StepBadge` im Wizard.
- **Zwei Zeichenketten werden geparst, nicht gelesen.** (1)
  `t('events.deadlinepassed.hint')` (`LanguageContext.tsx:79/537`) MUSS das Wort
  „organizer" enthalten — `components/EventCard.tsx:226` zerlegt ihn mit
  `/(organizer)/i`, um die Organizer-Kontaktkarte anzuhängen; ohne Treffer
  verschwindet sie lautlos. (2) `t('myevents.agenda')` MUSS auf Deutsch exakt
  „Programm" bleiben — `QuizPlayer.tsx:62`, `DocumentsViewer.tsx:154/175` und
  `MyEventCard.tsx:881/933` benutzen `=== 'Programm'` als Sprach-Detektor.
  Beides ist ein Konstruktionsfehler; **neue Stellen dieser Art entstehen
  nicht**, und wer eine der beiden Zeichenketten anfasst, ersetzt den Mechanismus
  im SELBEN Commit an allen Stellen (Sprache über `isDe`, Kontakt über eine
  eigene Prop) — sonst ändert er Verhalten, nicht Wortlaut.
- **Gespeicherte Formate sind kein Anzeigetext:** die Regex
  `^(.+?)\s*<([^>]+@[^>]+)>\s*$` für People-Picker-Antworten
  (`myEventsHelpers.tsx:14`, `RegistrationPage.tsx:667/1373`,
  `submitFlow.ts:742/1313`), die Platzhalter `{{EventTitle}}`, `{{Organizer}}`,
  `{{Name}}`, `{{AppUrl}}`, `{{ContactEmail}}` in der Event-Beschreibung
  (`registration/EventCard.tsx:427-442`) und die Antwort-Strings der
  Custom-Felder (6b, Datum).
- **SharePoint-Werte werden nie „korrigiert":** die Status-Literale
  `'Angemeldet'`, `'Warteliste'`, `'Abgemeldet'`, `'Eingecheckt'`,
  `'QR versendet'` (mit Leerzeichen), `'No-Show'`, der EmailType `'Info'`
  (`MyEventsPage.tsx:828`; die Choice-Liste steht in
  `services/events/emailQueue.ts:94` mit `Nachruecken` OHNE Umlaut) und die
  Spaltennamen aus `SPRegistration` (`TeilnehmerID`, `CustomData`,
  `AgendaCheckIns`, `Hotel*`, `CancellationDate`). Die vier ue/oe-Werte aus
  CLAUDE.md kommen auf diesen Seiten **nicht** vor — was hier mit Umlaut steht,
  ist Fließtext und darf umformuliert werden.
- **Datenschutz-Schalter wandern mit ihrem Block.** `hideOrganizer` +
  `hideOrganizerIndividualOnly` (`components/EventCard.tsx:144-146`,
  `MyEventCard.tsx:597-607`, `RegistrationPage.tsx:2035/2047`),
  `hotelVisibleToAttendees` (`MyEventCard.tsx:243`), `teamOpenSlotsVisible`
  (`RegistrationPage.tsx:720/2600`) und die Zeilen-Sicherheit der
  Teilnehmerlisten entscheiden, WER WESSEN Daten sieht. Kein Umbau blendet
  Personen-, Hotel- oder Teamdaten an einer Stelle ein, an der sie heute an
  einer Bedingung hängen.
- **Anzeige-Bedingungen und Rechenwege sind Aussagen, keine Formatierung:**
  `!subOnly` vor Plätze-Badge (`components/EventCard.tsx:380`) und Frist (:390)
  sowie `if (event.subEventsOnlyMode) return null` (`MyEventCard.tsx:199/286`);
  `freePlaces = effectiveMax - currentParticipants - waitlistCount` (:113 — die
  Wartelisten-Subtraktion IST die Aussage „frei", v24.72); „Angemeldet gewinnt"
  (:290, v30.2); die Rechte-Kette
  `!canCreateEvents && !isOwnOrganizer && !alreadySignedUp` (:127/134/135);
  die Wartelisten-Position nur bei `maxParticipants > 0`
  (`MyEventCard.tsx:218`); der QR-Knopf erst ab Status „QR versendet"/
  „Eingecheckt" (:740, v28.7); die Bescheinigung nur bei `agendaCheckIn` plus
  mindestens einer Marke (:759-762); `isFullyClosed` rechnet über ALLE
  buchbaren Sub-Events statt über die sichtbarkeitsgefilterte Liste
  (`RegistrationPage.tsx:1296-1299`, v30.20); die Organizer-Ausnahme in
  `notYetActive`/`isFullyClosed` (:1278/:1306); `sessionsChanged` (leere
  Auswahl IST die Änderung, :1398-1402); `kidsFirst`
  (`MyEventsPage.tsx:774-788`, v30.68: erst die Termine abmelden, die Klammer
  nur, wenn keiner fehlschlug).
- **Das Handbuch rendert diese Seiten LIVE.**
  `components/manual/sections/{findEvent,myEvents,subEvents,intro,checkIn,registerForOther}.tsx`
  mounten `EventListPage`, `RegistrationPage`, `MyEventsPage` und `LandingPage`
  gegen die **handgepflegten** Stubs in
  `components/manual/previews/PreviewProviders.tsx:148-216`. Ruft eine dieser
  Seiten künftig eine dort fehlende Context-Funktion in einem Mount-Effect,
  stirbt der ganze React-Baum mit „… is not a function" — **nur im Handbuch,
  nicht in der App**, im normalen Test also unsichtbar (so geschehen in v28.11,
  Kommentar :199-203). Zweiter Host: `RegisterPreviewModal.tsx:379`. Und
  `window.__dexForceMobile` (`utils/useIsMobile.ts:17`) heißt: die Vorschau
  rendert die Mobil-Zweige mit. Das Handbuch **zitiert außerdem Beschriftungen
  wörtlich** („Aktuelle Events", „Deine Events", „Registrierung starten",
  „Zusätzliche Sessions", Badge „Nur Sessions" —
  `manual/sections/findEvent.tsx:33`, `myEvents.tsx:95-100`): wer sie ändert,
  zieht `manual/sections/*` nach.
- **Exportierte Symbole und Storage-Schlüssel bleiben.** Exporte:
  `isEventVisibleForUser` (`EventListPage.tsx:132` → `RegistrationPage`,
  `MyEventCard`, `MyEventSubEvents`) und `KpiRow` (`LandingPage.tsx:1387` →
  `DexEventPlatform.tsx:38`). Schlüssel (Umbenennen = stiller Cache-Verlust,
  jeder Zugriff in try/catch): `dex:myevents:cache`, `dex-eventlist-view`,
  `dex_assist_<email>`, `dex_landing_regboxes_v1:<email>`,
  `dex_demo_impersonation` und `INACTIVE_SUMMARY_CACHE_KEY`
  (`LandingPage.tsx:343` — **geteilt** mit dem Organizer Center).

---

### 6f. Prüfen — die sieben Zustandsfälle

`tsc` und ESLint fangen hier fast nichts, und für die Teilnehmer-Seiten gibt es
**keinen Harness**: `tools/wizard-harness/entry.tsx` bündelt ausschließlich
`EventCreationPage`. Die einzige Vorschau sind die Handbuch-Seiten (oben) —
also ist der Stub-Provider Pflichtprüfung, und die Fälle unten werden am Diff
durchgegangen. Zusätzlich zu Abschnitt 8:

- [ ] Kachel als normaler User **nach Fristablauf** (Overlay + roter Kasten,
      Organizer-Kontakt erreichbar, Klick blockiert).
- [ ] Kachel eines **`subEventsOnlyMode`**-Events (keine Plätze, keine Frist —
      dafür Terminzahl und der Satz zur Anmeldung je Termin).
- [ ] Karte mit **Warteliste** und mit **geteilten Kapazitäten**
      (`maxParticipants === 0`: keine erfundene Position, Text und Fußzeile
      sagen „Warteliste").
- [ ] **`hiddenRow`**-Karte (fremd angelegt: sichtbar, nicht abmeldbar, Grund
      als Text, kein Knopf-Imitat).
- [ ] Anmeldeseite ohne freigegebene Termine (**`hiddenChildCount > 0`**) und
      als **Organizer nach Frist** (Banner statt Sperre).
- [ ] Liste und Kachel bei **`eventsReadStatus === 'forbidden'`** (Grund-Kasten
      statt „keine Events", Kacheln nicht als „0 angemeldet"/„frei").
- [ ] **Meine Events ohne eigene Anmeldung, aber mit Assistenz-Aufgabe** (die
      seitenweiten Kästen rendern, nicht der Leerzustand).

---

## 7. Symbole

`components/Icons.tsx` (Inline-SVG, Props `size`, `strokeWidth`):
`ChevronLeft ChevronUp ChevronDown Settings GraduationCap BarChart3 CaptainHat
Mail Star Book QrCode Info Calendar Pin MessageSquare Trash2 Send Plus X
FileText Search Users Download Copy Check Pencil ExternalLink AlertCircle Hash
Shirt Columns Wrench RefreshCw Video Link2`. Dazu Fluent-Icons über
`<Icon iconName="…" />` aus `@fluentui/react/lib/Icon`, wo schon im Einsatz.
Keine neuen Icon-Bibliotheken.

---

## 8. Prüfen vor dem Abschluss

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
