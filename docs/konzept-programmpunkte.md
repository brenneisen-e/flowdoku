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

**Stand v30.86: kein neues Feld — die bestehende Agenda IST die Liste.**
Beim Umsetzen zeigte sich, dass es die Struktur längst gibt: Die Spalte
`Agenda` auf `DEX_Events` trägt seit v22 ein `AgendaItem[]` (`id`, `date`,
`time`, `endTime`, `title`, `description`, `icon`), gepflegt in Schritt 3
„Ort & Programm", angezeigt in „Meine Events" als „Programm". Eine zweite
Spalte `AgendaJson` daneben wäre genau die Doppelung, vor der CLAUDE.md warnt
(„Vor einer neuen Ansicht prüfen, ob es sie schon gibt"). Also:

- `AgendaItem` bekommt `location?` (Raum). Sonst unverändert.
- Neu ist nur ein **Modus-Flag** am Event: `agendaCheckIn` (Piggyback
  `_agendaCheckIn` in `EmailTemplateOverrides`) — „die Agenda ist nicht nur
  Anzeige, sondern die Liste der Check-in-Stationen". Dazu die Bezeichnung
  `_agendaTerm = { singular, plural }` (leer = „Programmpunkt(e)").
- Beide werden beim Laden gestrippt (`useWizardVisibilityState`) und in
  beiden Save-Pfaden über `agendaCheckInPiggyback()` frisch gebaut; das Flag
  wird nie zusammen mit `subEventsOptIn` geschrieben.
- `id` ist stabil (`ag-<ts>-…`), damit Check-ins (Stufe 2) auch nach
  Umbenennen oder Umsortieren am richtigen Punkt hängen. Zeiten bleiben
  Lokal-Strings (`YYYY-MM-DD` + `HH:mm`) wie bisher in der Agenda.

Ein Event hat **entweder** Sub-Events **oder** Programmpunkte, nie beides
(Entscheidung 2, Abschnitt 4).

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
unverändert und bedeutet weiterhin „ist heute da". **Entscheidung 1
(07.09.2026): Ein Check-in an einem Programmpunkt setzt NUR diesen Punkt** —
„ich möchte wissen, ob die Person bei 1, 3, 5 da war". Kein automatischer
Event-Check-in; wer den will, checkt zusätzlich am Event ein (wie heute).
Die Auswertung (Stufe 3) zeigt beides getrennt.

## 3. Wo Programmpunkte in der App auftauchen

### 3.1 Wizard, Schritt 1

**Umgesetzt (v30.86):** Unter „Nutzung von Sub-Events" steht neben der
Kachel „Sub-Events aktivieren" eine zweite Kachel **„Programmpunkte mit
Anwesenheits-Check-in nutzen"** — mit Erklärkasten „Was ist ein
Programmpunkt?" nach dem Muster der Sub-Event-Kachel. Ist sie an, ist der
Sub-Event-Haken gesperrt („nicht kombinierbar"), und umgekehrt verschwindet
die Programmpunkt-Kachel, sobald Sub-Events aktiv sind. In der Kachel:
Bezeichnung (Presets Programmpunkt, Session, Vortrag, Workshop, Slot oder
eigener Begriff), die Zahl der angelegten Punkte und ein Knopf, der zu
Schritt 3 springt.

Die Punkte selbst werden **nicht** in Schritt 1 gepflegt, sondern dort, wo
die Agenda schon immer war: Schritt 3 „Ort & Programm". Der Editor trägt im
Programmpunkte-Modus die gewählte Bezeichnung, ein Badge „Check-in je
Punkt", einen Erklärkasten, neu ein Feld **Raum** je Punkt und den Knopf
**„Letzten Tag kopieren (+1 Tag)"** für mehrtägige Abläufe. Keine
Kommunikations-Reiter, keine Kapazität, keine Frist, keine Pflicht.

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

**Entscheidung 3 (07.09.2026): Migration = ein NEUES Event.** „Migration
würde ich gerne als neues Event machen, damit das alte nicht verloren geht."
Das alte Event bleibt vollständig stehen (Sub-Events, Listen, Check-ins,
Mails, Historie); der In-Place-Umbau mit drei Phasen aus der ersten Fassung
entfällt. Stattdessen eine Aktion im Organizer Center:

**„Als neues Event mit Programmpunkten kopieren"** (Stufe 4)

1. **Prüfen (nichts wird geschrieben).** Alle Sub-Event-Listen mit
   `onHttpError` laden — eine nicht lesbare Liste bricht ab, sie zählt nicht
   als leer. Anzeige: N Sub-Events → N Programmpunkte (Titel, Datum, Zeiten,
   Ort), M Personen, die auf mindestens einem Sub-Event oder der Klammer
   aktiv angemeldet sind, K davon mit Check-in, Wartelisten-Personen
   (werden NICHT kopiert — sie hatten keinen Platz; namentlich aufgeführt),
   Formularantworten je Sub-Event (werden in `CustomData` der neuen
   Hauptzeile unter dem Präfix des Punkts abgelegt).
2. **Neues Event anlegen** über den normalen `createEvent`-Pfad: Titel
   (Vorgabe „<alt> (Programmpunkte)"), Bild, Beschreibung, Ort, Organizer,
   Kommunikation der Klammer; `_agendaCheckIn: true`; `Agenda` aus den
   Sub-Events (neue Ids). `MaxParticipants` mindestens M. Das neue Event
   startet als **Test-Event/Entwurf**, damit vor der Freigabe alles geprüft
   werden kann.
3. **Anmeldungen still kopieren:** je Person eine echte Zeile am neuen Event
   (Status `Angemeldet`, `TeilnehmerID` neu, **ohne Mail, ohne Outlook** —
   die Person ist schon eingeladen; das alte Event trägt Mail und Termin
   weiter). `AgendaCheckIns` aus den Sub-Event-Check-ins füllen (Punkt-Id
   ↔ Sub-Event-Id aus Schritt 2). Sequentiell, mit Fehlerzähler und
   Overlay; jeder Fehler wird gezählt, am Ende Bericht. Nichts am alten
   Event wird verändert.
4. **Optional, per Haken:** Anmeldung am alten Event schließen
   (`_klammerDeadline` = jetzt, Sub-Event-Fristen = jetzt), damit sich
   niemand mehr doppelt anmeldet. Kein Löschen, keine Ausladung — die
   Outlook-Termine der Sub-Events bleiben als Erinnerung; das ist der
   Preis dafür, dass nichts verloren geht, und er ist bewusst gewählt.
5. Audit: ChangeLog `CopiedToAgendaEvent` (alt) / `CreatedFromSubEvents`
   (neu) mit Zahlen.

Für das konkrete Event mit 20 Sub-Events: Schritt 1 sagt vorab, ob
Wartelisten oder Formularantworten betroffen sind; der Lauf ist eine
Sache von Minuten, und weil das alte Event unangetastet bleibt, kann die
Kopie jederzeit verworfen und wiederholt werden.

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

| Stufe | Inhalt | Stand |
|---|---|---|
| 1 | Modus-Flag + Bezeichnung, Wahl in Schritt 1, Agenda-Editor mit Raum und „Tag kopieren", Programm-Block auf Anmeldeseite und in Meine Events | **v30.86 — ausgeliefert** |
| 2 | Check-in je Programmpunkt (Scanner, Teilnehmer-ID, Name), Punkt-Wahl mit Vorschlag „läuft gerade", Zähler je Punkt, Doppel-Scan-Meldung; Spalte `AgendaCheckIns` (neue Listen + „Spalten fixen"); Haken in Meine Events | **v30.91 — ausgeliefert** (offen: Self-Check-in/Live-QR je Punkt) |
| 3 | Abschnitt „Anwesenheit je Programmpunkt" im Organizer Center (Matrix Person × Punkt, „nach Punkt"), Excel mit zwei Blättern, manuelles Nachtragen/Zurücknehmen mit Audit | **v30.92 — ausgeliefert** |
| 4 | „Als neues Event mit Programmpunkten kopieren" (Abschnitt 5) | offen |
| 5 (später) | `{{Programm}}`-Platzhalter für Mail/Outlook, Teilnahmebescheinigung, Programmpunkte unter Kalender-Tagen | offen |

Stufen 2 bis 4 sind je ein Release; Stufe 4 braucht einen Testlauf an einer
Kopie des 20er-Events, bevor sie am echten läuft — was mit dem Kopier-Ansatz
ohnehin der normale Weg ist.

## 8. Entscheidungen (getroffen am 07.09.2026)

1. **Ankunft automatisch: NEIN.** Ein Check-in an einem Programmpunkt setzt
   nur diesen Punkt. Der Event-Check-in bleibt eine eigene Handlung.
2. **Entweder/oder: JA.** Ein Event hat Sub-Events oder Programmpunkte, nie
   beides. Der Assistent sperrt die jeweils andere Wahl.
3. **Migration: als NEUES Event.** Das alte Event bleibt unverändert
   erhalten; keine Ausladungen, kein Umbau in place (Abschnitt 5).
4. **Name: „Programmpunkte" als Vorgabe, je Event umbenennbar** (Session,
   Vortrag, Workshop, Slot, eigener Begriff) — Piggyback `_agendaTerm`.

Ergänzt bei der Umsetzung: **Kein neues Datenfeld** — die vorhandene Agenda
wird zur Programmpunkt-Liste (Abschnitt 2.1). Und die Anmeldeseite zeigt das
Programm jetzt für ALLE Events mit Agenda; das hatte der Assistent seit v22
versprochen und nie eingelöst.
