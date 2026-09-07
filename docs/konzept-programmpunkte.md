# Konzept: Programmpunkte (Sub-Events ohne eigene Anmeldung)

Stand 07.09.2026, Entwurf zur Abstimmung. Nutzer-Anforderung:

> Ich brauch die Möglichkeit Sub-Events ohne eigene Teilnehmerliste und
> Sub-SharePoints einzurichten. Ich brauch das nur für einen separaten
> Check-in für die Anwesenheitskontrolle. Damit werden die Sub-Events eher zu
> Agenda-Bausteinen. … Ich muss aber die Möglichkeit haben, dass ein
> bestehendes Event (ich habe eins mit 20 Sub-Events) migriert werden kann.

## 1. Was ein Programmpunkt ist — und was nicht

Ein **Programmpunkt** ist ein Zeitblock innerhalb eines Events: Titel, Start,
Ende, optional Ort und Beschreibung. Er hat **keine** eigene Anmeldung, keinen
Platz, keine Warteliste, keine Frist, keine eigene Mail, keinen eigenen
Outlook-Termin und keine eigene Teilnehmerliste. Wer zum Event angemeldet ist,
darf zu jedem Programmpunkt kommen. Das Einzige, was je Programmpunkt
festgehalten wird, ist die **Anwesenheit**: wer wann eingecheckt wurde.

Damit ist die Trennung zu den bestehenden Sub-Events scharf:

| | Sub-Event (bisher) | Programmpunkt (neu) |
|---|---|---|
| Anmeldung | eigene, mit Platz, Frist, Warteliste | keine — die Anmeldung zum Event genügt |
| Daten | eigene DEX_Events-Zeile, eigene Subsite, eigene Teilnehmerliste | JSON-Spalte am Hauptevent |
| Kommunikation | Mail und Outlook je Termin möglich | nichts eigenes; das Programm kann in Mail und Termin des Events stehen |
| Check-in | je Sub-Event auf dessen Liste | je Programmpunkt in der Teilnehmerzeile des Events |
| Flows | Outlook-Flow, Reorder-Flow, Zähler-Kette je Liste | **keiner** |
| Kosten je Stück | Subsite anlegen (Sekunden, Drossel-Risiko), 7 Listen-Schreibvorgänge je Anmeldung | ein JSON-Eintrag |

Die zweite Zeile ist der Kern: Ein Event mit 20 Sub-Events bedeutet heute 21
Subsites, 21 Teilnehmerlisten, 21 Zähler und bei jeder An- oder Abmeldung 20
Schreibvorgänge mehr — das ist der Ursprung eines Großteils der Drossel-Fälle
der letzten Wochen (Abmelde-Welle, Speichern bei 82 %, Reorder-Aufträge).
Programmpunkte kosten davon nichts.

## 2. Datenmodell

### 2.1 Programmpunkte am Hauptevent

Neue Spalte **`AgendaJson`** (mehrzeiliger Text) auf `DEX_Events`, angelegt
idempotent über `ensureEventsList` (eventsListSchema.ts, wie
`SelfCheckInEnabled` in v18.33). Bewusst eine eigene Spalte und **kein**
Piggyback in `EmailTemplateOverrides`: Das Piggyback-JSON ist schon der Ort für
zwölf Flags, und jedes davon muss beim Laden gestrippt werden (CLAUDE.md-
Falle). Ein Programm mit 20 Einträgen ist Fachdatum, kein Flag.

```json
[
  { "id": "a1", "title": "Begrüßung", "start": "2026-10-12T08:00:00Z", "end": "2026-10-12T08:30:00Z", "location": "Plenum", "description": "", "order": 1 },
  { "id": "a2", "title": "Keynote", "start": "…", "end": "…", "location": "Plenum", "order": 2 }
]
```

- `id` ist stabil (kurze Zufalls-Id), damit Check-ins auch nach Umbenennen
  oder Umsortieren am richtigen Punkt hängen.
- Zeiten als UTC-ISO wie bei Sub-Events; Umrechnung nur über `subIsoToDate`/
  `subDateToIso` (dieselbe Regel wie im Wizard).
- Größe: 20 Einträge ≈ 4 KB. Keine Grenze in Sicht.

Im `DeloitteEvent`-Typ: `agenda?: AgendaItem[]`. Ein Event hat **entweder**
Sub-Events **oder** Programmpunkte, nie beides (Abschnitt 4 begründet das).

### 2.2 Anwesenheit in der Teilnehmerzeile

Neue Spalte **`AgendaCheckIns`** (mehrzeiliger Text) auf jeder Teilnehmerliste,
angelegt über `ensureRegistrationList` **und** `fixRegistrationListColumns`
(„Spalten fixen" für Bestandslisten — dieselbe Pflicht wie bei jeder neuen
Spalte, s. CLAUDE.md zu `StarterType`).

```json
{ "a1": { "at": "2026-10-12T08:03:12Z", "by": "scanner@deloitte.de" },
  "a2": { "at": "2026-10-12T09:01:40Z", "by": "self" } }
```

Warum in der Teilnehmerzeile und nicht in einer neuen Liste `DEX_Checkins`:

1. Alles, was der Check-in braucht (Name, Foto, Teilnehmer-ID, Status,
   QR-Code), steht schon in dieser Zeile. Der bestehende QR-Code bleibt
   derselbe — er identifiziert die Person, der gewählte Programmpunkt kommt
   vom Scanner-Gerät.
2. Die Rechte stimmen schon: Wer heute einchecken darf (Organizer, Check-in-
   Team), schreibt heute schon `Status = Eingecheckt` per MERGE auf diese
   Zeile. Kein neues Berechtigungskonzept, kein neuer Entzug.
3. Die Auswertung (Matrix Person × Programmpunkt) liest die Liste, die das
   Organizer Center ohnehin lädt. Kein zweiter Ladepfad.
4. Ein Check-in ist ein MERGE auf eine Zeile — genau wie heute. Zwei Scanner,
   die verschiedene Personen einchecken, kollidieren nicht; derselbe Mensch
   zweimal ist idempotent (gleicher Schlüssel, spätere Zeit gewinnt nicht —
   der erste Zeitstempel bleibt).

Der bisherige Event-Check-in (`Status = Eingecheckt`, `CheckedInDate`) bleibt
unverändert und bedeutet weiterhin „ist heute da". Vorschlag: Der **erste**
Check-in an irgendeinem Programmpunkt setzt ihn automatisch mit — wer um 9 Uhr
in der Keynote sitzt, ist angekommen. (Entscheidung 1, s. Abschnitt 8.)

## 3. Wo Programmpunkte in der App auftauchen

### 3.1 Wizard, Schritt 1

Im Abschnitt, in dem heute die Sub-Event-Liste steht, kommt **vor** die Liste
eine Wahl mit drei Karten (einmalig, danach umschaltbar nur solange keine
Anmeldung existiert):

- **Nur das Event** — wie heute ohne Sub-Events.
- **Sub-Events mit eigener Anmeldung** — wie heute; Erklärsatz: „Jeder Termin
  hat eigene Plätze, Fristen und Kommunikation. Teilnehmer wählen aus."
- **Programmpunkte (Agenda)** — neu; Erklärsatz: „Ein Ablauf mit Zeiten. Die
  Anmeldung gilt fürs ganze Event; je Programmpunkt wird nur die Anwesenheit
  per Check-in erfasst. Keine eigenen Listen, keine eigenen Mails."

Die Programmpunkt-Liste ist bewusst schlank: Titel, Start, Ende, Ort — eine
Zeile je Punkt, Drag-Reihenfolge, „Punkt hinzufügen", „aus Tag kopieren" (den
Vortag klonen mit +1 Tag), Sammelaktion „Alle Zeiten um X Minuten schieben".
Keine Kommunikations-Reiter, keine Kapazität, keine Frist, keine Pflicht —
genau das, was der Nutzer will: Agenda-Bausteine.

Die Scope-Karte (Reiter Klammer/Sub-Events) erscheint bei Programmpunkten
**nicht**; sie hängt an `subEvents`, und die bleiben leer.

### 3.2 Anmeldeseite

Unter der Beschreibung ein Block **„Programm"**: die Punkte chronologisch mit
Tag-Zwischenüberschriften (dieselbe Gruppierung nach Datum, die
`groupSubEventTabs` heute nach Präfix macht). Nur lesen, nichts auswählen. Wer
sich anmeldet, meldet sich fürs Event an.

### 3.3 Meine Events

Auf der Karte derselbe Block „Programm", je Punkt mit dem eigenen
Anwesenheits-Haken („anwesend 08:03"), sobald einer gesetzt ist. Kein
Abmelden je Punkt — es gibt nichts abzumelden.

### 3.4 Mail und Outlook

Ein neuer Platzhalter **`{{Programm}}`** (Tabelle Zeit · Titel · Ort) für
Bestätigungs-, QR- und Outlook-Text. Ohne Platzhalter ändert sich nichts. Das
ist reine App-Logik beim Zusammenbauen der Mail — **kein Flow ändert sich**.

### 3.5 Check-in-Seite

Nach der Event-Wahl erscheint bei Events mit Programmpunkten eine zweite
Wahl **„Programmpunkt"**:

- Vorbelegt mit dem Punkt, der **jetzt läuft** (oder als Nächstes beginnt) —
  am Einlass will niemand suchen. Darüber die Option „Nur Ankunft (Event)",
  das ist der heutige Check-in.
- Live-Scanner, Foto-Weg und Teilnehmer-ID-Feld arbeiten unverändert; der
  gewählte Punkt entscheidet nur, **wohin** geschrieben wird.
- Die Bestätigungskarte (Foto, Name, Status) zeigt zusätzlich, an welchen
  Punkten die Person heute schon war — ein zweiter Scan derselben Person am
  selben Punkt meldet „bereits um 08:03 eingecheckt" statt still zu
  überschreiben.
- Die drei Kacheln (Angemeldet / Eingecheckt / No-Show) beziehen sich auf den
  gewählten Punkt; „Nur Ankunft" zeigt die heutigen Zahlen.
- **Self-Check-in (Live-QR):** Der rotierende Code trägt zusätzlich die
  Punkt-Id; am Eingang jedes Raums hängt ein anderer Bildschirm oder das
  Tablet wird umgestellt. Der Ausdruck (PDF) bekommt je Punkt eine Seite.

### 3.6 Organizer Center

Bei Events mit Programmpunkten:

- Reiter **„Anwesenheit"** (statt der konsolidierten Sub-Event-Matrix): Zeilen
  Personen, Spalten Programmpunkte, Haken mit Uhrzeit im Tooltip, Kopfzeile
  mit Zahl je Punkt; gruppiert nach Tag wie die Abmelde-Sicht „nach Tag"
  (v30.83). Ab etwa 15 Punkten die Chip-Darstellung je Person und die
  Sicht „nach Programmpunkt" mit Balken — dieselben zwei Sichten wie bei den
  Abmeldungen, dieselbe Komponente.
- Export: die Teilnehmer-Excel bekommt je Punkt eine Spalte (Uhrzeit oder
  leer). Eine Teilnahmebescheinigung je Person (PDF, „war anwesend bei …")
  ist mit denselben Daten ein Folgeschritt.
- Manuelles Nachtragen: in der Anwesenheits-Matrix je Zelle „setzen"/
  „entfernen" mit Audit (ChangeLog `AgendaCheckIn`/`AgendaCheckInRemoved`),
  für „hat vergessen zu scannen".

## 4. Entweder Sub-Events oder Programmpunkte — nicht beides

Beides gleichzeitig wäre technisch möglich (zwei Listen im Wizard, zwei Blöcke
auf der Anmeldeseite), aber es verdoppelt jede Frage: Wo steht der Termin?
Zählt der Check-in für den Punkt oder das Sub-Event? Was zeigt die Matrix? Die
Erfahrung mit dem Scope-Umschalter (CLAUDE.md: „Der Scope-Umschalter gehört
genau einmal auf die Seite") und den zwei Reiter-Leisten sagt: zwei Wege für
ähnliche Dinge nebeneinander liest der Organizer als zwei Navigationen. Ein
Event entscheidet sich einmal. Wer später doch Sub-Events braucht, kann
Programmpunkte löschen und umschalten, solange keine Anwesenheit erfasst ist.

Kalender-Tage (`_subEventCalendar`) bleiben eine Spielart der Sub-Events —
jeder Tag ist dort eine Anmeldeeinheit. Ein Tages-Programm innerhalb eines
Office-Tags wäre dagegen ein klassischer Fall für Programmpunkte, aber erst,
wenn ein Kalender-Tag selbst Programmpunkte tragen dürfte. Das ist Stufe 2 und
gehört nicht in den ersten Wurf.

## 5. Migration eines bestehenden Events mit Sub-Events

Als Aktion im Organizer Center: **„Sub-Events in Programmpunkte umwandeln"**.
Sie läuft in drei Phasen und ist bis zur letzten reversibel — nach dem Muster,
das v30.67 für alle Löschpfade festgelegt hat: erst anlegen, dann prüfen, der
unumkehrbare Schritt kommt zuletzt.

### Phase 1 — Prüfen (nichts wird geschrieben)

Die Aktion lädt alle Sub-Event-Listen (mit `onHttpError`; eine nicht lesbare
Liste bricht die Prüfung ab, sie wird nicht als leer gerechnet) und zeigt:

- N Sub-Events → N Programmpunkte (Titel, Zeiten, Ort übernommen).
- M Personen, die auf mindestens einem Sub-Event angemeldet sind.
- Davon **K ohne Anmeldung am Hauptevent** (im Modus „Nur Sub-Events" sind das
  die, deren Klammer-Schattenzeile fehlt; im Normalmodus alle, die nur Termine
  gebucht haben). Für sie legt Phase 2 eine echte Anmeldung am Hauptevent an.
- Check-ins auf Sub-Events → werden zu Anwesenheiten.
- **Blocker**, bei denen die Migration nicht startet:
  - Personen auf einer **Warteliste** eines Sub-Events (sie haben keinen
    Platz; am Hauptevent gäbe es plötzlich einen). Der Organizer entscheidet
    vorher: nachrücken lassen oder abmelden.
  - **Geteilte Kapazitäten** (Durchstarter/Funstarter) auf einem Sub-Event.
  - Sub-Events mit eigenen **Formularfeldern und Antworten**: Die Antworten
    hängen an der Sub-Event-Zeile. Phase 2 kopiert sie in `CustomData` der
    Hauptzeile unter dem Präfix des Programmpunkts, aber der Organizer muss
    das sehen und bestätigen, denn die Felder gibt es danach als Fragen nicht
    mehr.
- **Hinweise**: eigene Kommunikationstexte je Sub-Event (gehen verloren, das
  Event hat künftig einen Text), Hotelplanung (bezieht sich auf Sub-Events,
  wird auf das Event umgehängt).

### Phase 2 — Anlegen (alles additiv, jederzeit abbrechbar)

1. `AgendaJson` am Hauptevent schreiben (Punkte aus den Sub-Events, Ids neu).
2. Für jede Person: fehlende Hauptevent-Anmeldung anlegen (echte Zeile, Status
   `Angemeldet`, ohne Mail, ohne Outlook — sie ist ja schon eingeladen).
   Hat das Hauptevent `MaxParticipants`, wird die Zahl vorher auf mindestens
   M gesetzt; im Modus „Nur Sub-Events" wird der Modus abgeschaltet, die
   Schattenzeilen werden zu echten Anmeldungen.
3. `AgendaCheckIns` je Person aus den Sub-Event-Check-ins füllen.
4. Formularantworten kopieren (s. Blocker oben).
5. Sequentiell, mit Fehlerzähler, mit Fortschritt (dasselbe Overlay wie bei
   der Abmeldung). **Ein einziger Fehler stoppt vor Phase 3** — der alte Stand
   ist dann noch vollständig da, die neuen Daten sind zusätzlich da und
   stören nicht (Programmpunkte werden erst mit Phase 3 sichtbar, über ein
   Flag `agendaMigrationPending`).

### Phase 3 — Abschalten (erst hier wird etwas weggenommen)

1. Outlook: für jede Person und jedes Sub-Event `Ausladen` in die Queue —
   der Termin des Hauptevents bleibt. Das sind M × N Queue-Zeilen; bei 20
   Sub-Events und 80 Personen 1.600 Ausladungen. Die Aktion pace-t das und
   sagt vorher, wie viele es sind. Alternative, die der Organizer wählen
   kann: **keine** Ausladung, wenn die Sub-Event-Termine als Kalender-Erinnerung
   erwünscht bleiben (sie schaden nicht, sie sind nur nicht mehr mit DEX
   verbunden).
2. `DEX_Participants`: die Sub-Event-Nummern je Person entfernen.
3. Die Sub-Event-Zeilen in `DEX_Events` auf `Status = Migrated` setzen — sie
   verschwinden aus allen Ansichten, **Subsites und Listen bleiben 30 Tage
   stehen** (Archiv-Konzept, „Altes Archiv löschen" räumt sie nach der Frist
   weg). Das ist die Rücktür: Solange die Listen stehen, kann ein Admin die
   Migration mit einem Klick zurücknehmen (Status zurück, `AgendaJson`
   leeren, Hauptevent-Anmeldungen, die Phase 2 angelegt hat, wieder
   entfernen — sie sind markiert).
4. Audit: ChangeLog `SubEventsMigratedToAgenda` mit Zahlen.

### Für das konkrete Event mit 20 Sub-Events

Phase 1 sagt dir vorab, ob Wartelisten oder Formularantworten im Weg stehen.
Wenn nicht, ist der Lauf eine Sache von Minuten; die Ausladungen sind der
teuerste Teil und der einzige, den Teilnehmer merken — deshalb die Wahl.

## 6. Was sich NICHT ändert

- Keine Änderung an einem Power-Automate-Flow. Programmpunkte erzeugen keine
  DEX_Events-Zeilen, keine Queue-Aufträge, keine Zähler. Der Reorder-Flow, der
  Outlook-Flow und die Mail-Flows sehen ein Event mit Programmpunkten wie ein
  Event ohne Sub-Events.
- Bestehende Events mit Sub-Events laufen unverändert weiter. Niemand muss
  migrieren.
- Der QR-Code je Person bleibt derselbe; Mails müssen nicht neu verschickt
  werden.

## 7. Umsetzung in Stufen

| Stufe | Inhalt | Ergebnis für dich |
|---|---|---|
| 1 | Spalten, Typ, Wizard-Wahl mit Programmpunkt-Liste, Anzeige auf Anmeldeseite und Meine Events, `{{Programm}}` | Neue Events können Programmpunkte haben |
| 2 | Check-in je Programmpunkt (Scanner, ID, Foto-Weg), Live-QR je Punkt, Kacheln, Doppel-Scan-Meldung | Anwesenheit wird erfasst |
| 3 | Anwesenheits-Reiter im Organizer Center (Matrix + „nach Punkt"), Excel-Spalten, manuelles Nachtragen | Auswertung |
| 4 | Migration (drei Phasen, Rücknahme) | Dein 20er-Event umziehen |
| 5 (später) | Teilnahmebescheinigung, Programmpunkte unter Kalender-Tagen | |

Stufen 1 bis 3 sind je ein Release; Stufe 4 braucht einen Testlauf an einer
Kopie deines Events, bevor sie am echten läuft.

## 8. Entscheidungen, die ich von dir brauche

1. **Ankunft automatisch:** Setzt der erste Check-in an einem Programmpunkt
   auch den Event-Check-in („ist da")? Mein Vorschlag: ja.
2. **Entweder/oder:** Ein Event hat Sub-Events oder Programmpunkte, nie beides.
   Mein Vorschlag: ja, s. Abschnitt 4.
3. **Migration und Outlook:** Sollen die Sub-Event-Termine bei der Migration
   aus den Kalendern ausgeladen werden, oder bleiben sie als Erinnerung
   stehen? Ich würde die Wahl im Dialog anbieten, Vorgabe „ausladen".
4. **Name:** „Programmpunkte" (Vorschlag) oder „Agenda"? Der Begriff steht
   dann in Wizard, Anmeldeseite, Check-in und Mails; wie bei
   `childTermSingular` sollte er je Event umbenennbar sein („Session",
   „Workshop", „Slot").
