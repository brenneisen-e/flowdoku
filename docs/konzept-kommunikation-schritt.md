# Konzept: Kommunikations-Schritt im Wizard entlasten

Stand 07.09.2026 (v30.88). Auftrag: „Informier dich im Netz mit Best Practices,
wie wir den Wizard-Schritt Kommunikation noch intuitiver und besser aufbauen.
Der ist für die meisten Organizer in der jetzigen Form überfordernd."

## 1. Was der Schritt heute zeigt

`components/wizard/steps/CommunicationStep.tsx` (1.100 Zeilen) rendert auf
EINER Seite, je Reiter (Klammer / jedes Sub-Event):

| # | Kasten | Inhalt | Steuerelemente |
|---|---|---|---|
| — | Übersichtsbox „automatische Kommunikation" | 8 Zeilen Status (Mails an/aus, Anmelde-/Abmelde-Bestätigung, Outlook, Absage → Auto-Abmeldung, inaktives Konto, Sprache, BCC) | 0 (nur Text) |
| 25 | Mail-Sprache | DE / EN | 2 |
| 26 | Benachrichtigungen (Aufklapper) | `disableEmails`, `disableRegistrationEmail`, `disableCancellationEmail`, `disableOutlook`, `autoDeregisterOnDecline`, `inactiveHandling` | 6 + Erklärtexte |
| 27 | Organizer mitlesen (Aufklapper) | BCC bei Anmeldung (nie/immer/ab Datum), bei Abmeldung (nie/immer/nach Frist), Datum | 7 |
| 28 | Mail-Logo (Aufklapper) | Upload, aus Event-Foto, Zuschnitt, Breite/Abstand | 5+ |
| 29 | Outlook-Logo (Aufklapper) | dasselbe noch einmal, „wie Mail-Logo" | 5+ |
| 30 | Outlook-Text (Aufklapper) | Betreff, Überschrift, Unterüberschrift, Body (HTML-Editor), Zeit-Override | 6+ |
| 31 | Vorlagen (Aufklapper) | je Mail-Typ (Anmeldung, Abmeldung, Warteliste, Nachrücken, QR …) Betreff/Überschrift/Text im HTML-Editor | 6 Typen × 4 |

Dazu bei Sub-Events: der Schalter „gemeinsam / einzeln" (v30.71), die
Themen-Chips „wie Haupt-Event / abweichend" und der Knopf „für alle
übernehmen". Zusammen rund **40 Steuerelemente**, alle auf einer Seite, jeder
Aufklapper mit dem Zusatz „Standard – empfohlen, klick zum Anpassen".

Das Kernproblem ist nicht ein einzelner Kasten, sondern dass **Entscheiden
und Feinjustieren auf derselben Ebene** stehen. Wer nur „Mail und Termin wie
üblich" will (die meisten), muss durch sieben Kästen lesen, um festzustellen,
dass er nichts tun muss. Und die Übersichtsbox oben sagt ihm das zwar — aber
in acht Zeilen Prosa, bevor er die Frage überhaupt gestellt hat.

## 2. Was die Literatur dazu sagt

Gelesen: NN/G (Nielsen Norman Group) zu Cognitive Load in Formularen und zu
Tabs vs. Accordions, Baymard zu Inline-Accordions in Formularen, UXPin und
LogRocket zu Progressive Disclosure, Toptal zu Settings-Seiten. Was daraus
für DIESEN Schritt trägt:

1. **Kern-Ebene und Erweiterungs-Ebene trennen (Progressive Disclosure).**
   Die Kern-Ebene ist „die kleinste Menge Steuerelemente, mit der die
   Mehrheit die Aufgabe erledigt"; alles Seltene, Riskante oder
   Erklärungsbedürftige kommt auf Anforderung. Heute ist alles Kern-Ebene,
   nur eingeklappt — das ist keine Trennung, sondern eine Verkleidung.
2. **Nielsens Einschränkung:** Stufenweises Aufteilen funktioniert nur, wenn
   die Stufen wenig voneinander abhängen. Mail-Sprache, Vorlagen und Logo
   hängen eng zusammen (Sprache bestimmt die Vorlage, Logo steht in jeder
   Vorlage). Deshalb **keinen zweiten Wizard-Schritt** aufmachen, sondern
   innerhalb des Schritts Ebenen bilden.
3. **Smart Defaults, sichtbar und änderbar.** Ein Default, den man erst durch
   Aufklappen sieht, ist für den Nutzer kein Default, sondern eine Unbekannte.
   Der Default muss als getroffene Entscheidung dastehen („Mail + Outlook,
   Deutsch, Organizer nicht in Kopie") — mit einem Klick zum Ändern.
4. **NN/G, vier Prinzipien gegen Cognitive Load:** Struktur (gruppieren,
   Reihenfolge nach Wichtigkeit), Transparenz (vorab sagen, was passiert),
   Klarheit (eine Frage je Steuerelement, positive Formulierung — „Mails
   senden" statt „Mails deaktivieren"), Unterstützung (Fehler und Folgen
   direkt am Element).
5. **Tabs vs. Accordions (NN/G):** Accordions für viele kurze Inhalte, Tabs
   für wenige lange. Die vier Anpassungs-Themen (Vorlagen, Mail-Logo,
   Outlook-Logo, Outlook-Text) sind wenige und lang → Reiter, nicht sieben
   Aufklapper untereinander.
6. **Baymard zu Inline-Accordions in Formularen:** Nutzer wissen nicht, ob
   eingeklappte Bereiche „mitgespeichert" werden. Genau die Frage stellt
   sich hier bei jedem `<details>`: Gilt der Standard, wenn ich ihn nie
   aufgeklappt habe? Die Antwort muss am Element stehen, nicht im Tooltip.
7. **Settings-Seiten (Toptal, Microsoft):** zwei Ebenen — „Basic" für die
   Masse, „Advanced" bewusst getrennt; Änderungen sofort sichtbar
   (Vorschau), nicht erst nach dem Speichern.

## 3. Vorschlag: „Drei Entscheidungen, eine Vorschau, ein Anpassen-Bereich"

### Ebene 1 — Entscheiden (immer sichtbar, drei Fragen)

Eine Karte pro Frage, als Radio-Karten wie in Schritt 1 („Anzeige-Bezeichnung"):

1. **Wie erreicht DEX die Teilnehmer?**
   - Mail + Outlook-Termin (Standard) · nur Mail · nur Outlook · keine
     Kommunikation
   - setzt `disableEmails`/`disableOutlook`; „keine Kommunikation" zeigt die
     Folge in EINER roten Zeile darunter („niemand erfährt von seiner
     Anmeldung — nur sinnvoll, wenn du selbst einlädst").
2. **In welcher Sprache?** DE · EN (wie heute, Kasten 25).
3. **Wer liest mit?** Organizer nie · immer in Kopie · ab Stichtag — EIN
   Radio für Anmeldung UND Abmeldung, mit einem kleinen „abweichend für
   Abmeldungen"-Link, der die zweite Zeile erst dann zeigt.

Alles andere aus Kasten 26 (`disableRegistrationEmail`,
`disableCancellationEmail`, `autoDeregisterOnDecline`, `inactiveHandling`)
wandert nach Ebene 3. Diese vier Schalter braucht ein Event in zwanzig.

### Ebene 2 — Sehen (eine Vorschau statt acht Zeilen Text)

Die Übersichtsbox wird ersetzt durch **eine gerenderte Vorschau der
Anmeldebestätigung** (Kopfbild, Überschrift, Text, darunter die
Outlook-Karte „so sieht der Termin aus") — die Vorschau-Logik gibt es in
`HtmlEditorModal` schon (Live-Preview + `buildOutlookBody`), sie muss nur als
Karte ohne Editor gerendert werden. Daneben ein Knopf **„Testmail an mich"**
(Mechanik: `qrTestSendAction` / Queue mit Empfänger = aktuelle Person).
Eine Vorschau beantwortet „was kommt raus?" schneller als jede Erklärung,
und sie zeigt sofort, ob Logo, Sprache und Text zusammenpassen.

Unter der Vorschau eine **Zusammenfassungs-Zeile** mit Chips: „Anmeldung:
Standardtext · Abmeldung: angepasst · Mail-Logo: Event-Foto · Outlook-Text:
Standard". Das ist die Transparenz aus NN/G — und die Antwort auf Baymards
Frage („gilt der Standard?"): ja, und hier steht welcher.

### Ebene 3 — Anpassen (ein Link, dann Reiter)

Ein Link **„Texte und Bilder anpassen"** klappt EINEN Bereich mit vier
Reitern auf: **Mail-Texte** (heutiger Kasten 31) · **Mail-Logo** (28) ·
**Outlook-Termin** (29 + 30 zusammengelegt: Logo, Betreff, Text, Zeiten) ·
**Feineinstellungen** (die vier seltenen Schalter aus 26). Reiter statt
Aufklapper, weil es wenige, lange Inhalte sind (NN/G). Jeder Reiter trägt
oben den Chip „Standard aktiv" oder „angepasst — zurücksetzen".

### Sub-Events

`commShared` bleibt der Schalter, aber er rückt an den Anfang von Ebene 1
als vierte Frage nur bei Sub-Events: **„Gelten diese Einstellungen für alle
Termine?"** Ja (Standard, neue Events) → keine Reiter, kein Chip-Vergleich,
ein Formular. Nein → die Scope-Karte oben wird aktiv wie heute, und Ebene 3
zeigt je Thema den Chip „wie Haupt-Event / abweichend" (das ist
`topicDiffersFromParent`, existiert). Damit sieht ein Organizer mit 20
Terminen im Normalfall dieselbe Seite wie einer ohne Sub-Events.

## 4. Was sich NICHT ändert

- Kein neues Datenfeld, kein Flow. Alle Werte sind dieselben States;
  geändert wird nur, welche davon auf welcher Ebene stehen.
- Der HTML-Editor (Modal) bleibt — nur die Einstiegsstelle wandert in den
  Reiter „Mail-Texte".
- `SCOPE_AWARE_STEPS`, `setScope`/`switchCommTab`, `persistSubEvents` und
  `dc` (v30.71) bleiben unangetastet; die Fallen aus CLAUDE.md zum
  Kommunikationsfeld-Slot gelten weiter.

## 5. Umsetzung in Stufen

| Stufe | Inhalt | Aufwand |
|---|---|---|
| A | Ebene 1 (drei Radio-Karten) + Ebene 3 als Reiter-Bereich hinter einem Link; Kasten 26 aufteilen; Übersichtsbox auf die Chip-Zeile eindampfen | 1 Release, nur `CommunicationStep.tsx` + ein Chip-Helfer |
| B | Vorschau-Karte (Mail + Outlook) aus der `HtmlEditorModal`-Preview herausziehen | 1 Release |
| C | „Testmail an mich" | klein, mit B |
| D | Sub-Event-Frage als vierte Karte; Reiter nur im Einzel-Modus | mit A |

Stufe A ist der eigentliche Hebel: Sie macht aus „sieben Kästen, in denen
alles gleich wichtig aussieht" eine Seite mit drei Fragen, deren Antworten
man liest, ohne etwas anzuklicken.

## 6. Offene Entscheidung

- Soll „keine Kommunikation" überhaupt in Ebene 1 stehen? Alternative: nur
  „Mail + Outlook / nur Mail / nur Outlook", und „gar nichts" liegt in
  Ebene 3 hinter einer Rückfrage. Dafür spricht, dass der Modus fast nur bei
  Klammer-Events mit `_subEventsOnlyMode` vorkommt — und dort setzt die App
  ihn ohnehin selbst.

## Quellen

- NN/G, „Few Guesses, More Success: 4 Principles to Reduce Cognitive Load in Forms" — https://www.nngroup.com/articles/4-principles-reduce-cognitive-load/
- NN/G, „Tabs vs. Accordions: When to Use Each" — https://www.nngroup.com/videos/tabs-vs-accordions/
- NN/G, „Tabs, Used Right" — https://www.nngroup.com/articles/tabs-used-right/
- Baymard, „Accordion UX: The Pitfalls of Inline Accordion and Tab Designs" — https://baymard.com/blog/accordion-and-tab-design
- UXPin, „What Is Progressive Disclosure in UX?" — https://www.uxpin.com/studio/blog/what-is-progressive-disclosure/
- LogRocket, „Progressive disclosure in UX design: Types and use cases" — https://blog.logrocket.com/ux-design/progressive-disclosure-ux-types-use-cases/
- Toptal, „How to Improve App Settings UX" — https://www.toptal.com/designers/ux/settings-ux
- Microsoft Learn, „Guidelines for app settings" — https://learn.microsoft.com/en-us/windows/apps/design/app-settings/guidelines-for-app-settings
