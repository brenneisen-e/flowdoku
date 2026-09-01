# Audit v30.66 — offene Befunde

Ergebnis des Fehler-Audits nach der Modularisierung (103 Prueferwaehlen, jeder Fund von drei
unabhaengigen Skeptikern gegengelesen). **Alle 29 Befunde sind Logikfehler, die es vor dem Umbau
schon gab** — kein einziger Umbaufehler. Die Zuordnung „schon im Ausgangsstand" ist je Eintrag
gegen den unveraenderten Vergleichsbaum geprueft.

Sieben Befunde sind in v30.66 behoben (siehe `docs/release-notes.md`). Die uebrigen stehen hier,
damit sie nicht verlorengehen — bewusst NICHT im selben Release behoben: ein Release, das eine
Modularisierung dieser Groesse mit zwei Dutzend Verhaltensaenderungen mischt, ist im Fehlerfall
nicht mehr aufzutrennen.

Reihenfolge = Schwere. Jeder Eintrag nennt Datei:Zeile, Fehlszenario, Ursache und einen Fix-Vorschlag.

## 1. [hoch] Anmeldeseite zählt eine ABGEMELDETE Klammer-Zeile als vorhanden — die Klammer wird nicht reaktiviert, die Person fällt aus der Hotelplanung heraus

**Wo:** `components/registration/submitFlow.ts:795`  
**Im Ausgangsstand:** Identisch im Ausgangsstand: components/RegistrationPage.tsx:2354 (`const parentAlreadyHasRow = !!myParentReg;`). Der Umbau hat den handleSubmit-Körper 1:1 nach submitFlow.ts verschoben.

**Szenario:** Klammer-Event (subEventsOnlyMode) mit Terminen und Hotelfrage auf Hauptevent-Ebene. Eine Person hat vorher abgesagt bzw. der Organizer hat im Abmelde-Dialog gezielt nur die Klammer-Zeile abgemeldet (createKlammerActions.ts:485-499 stellt die Klammer als eigene, abwählbare Zeile in den Dialog) — ihre Klammer-Zeile steht auf 'Abgemeldet'. Jetzt bucht sie über die Anmeldeseite einen Termin und füllt dabei die übergreifenden Hauptevent-Felder aus (Hotel, Verpflegung, Anreise). `myParentReg` ist die abgemeldete Zeile, also ist `parentAlreadyHasRow` true → `shouldShadowRegisterParent` false → der Schritt-3-Block (Zeile 1106) läuft nicht, und die Sub-Event-Anmeldungen laufen mit `skipShadowParent: true` (Zeile 1054), also greift auch das zentrale Netz in EventContext.tsx:1314 nicht, das die Zeile über den Reaktivierungs-Pfad wieder scharf gestellt hätte. Stattdessen schreibt `updateMyRegistration` die frischen Antworten in die ABGEMELDETE Zeile und lässt den Status stehen. Ergebnis: Der Termin ist gebucht, die Klammer-Zeile bleibt 'Abgemeldet'. Konkrete Folge in der Hotelplanung: HotelPlanningPanel bekommt die Klammer-`registrations` (admin/sections/HotelPlanningSection.tsx:61) und filtert `people` auf ACTIVE_STATI (HotelPlanningPanel.tsx:575-580) — die Person ist im Hotel-Roster gar nicht vorhanden, obwohl sie „Yes, I need accommodation“ beantwortet hat und an dem Termin teilnimmt. `autoDistribute` kann sie nicht einmal als `skippedNoWish` melden. Zusätzlich taucht sie in der „Abmeldungen“-Liste des Gesamt-Events auf (AdminPage.tsx:1519-1524), während die Matrix sie als Teilnehmerin führt.

**Ursache:** Zwei Ableitungen aus demselben `myParentReg`, aber mit unterschiedlicher Semantik: `parentAlreadyRegistered` (RegistrationPage.tsx:1290) prüft den Status, `parentAlreadyHasRow` (submitFlow.ts:795) prüft nur die Existenz des Objekts. Für den Zweck „muss die Klammer-Zeile geschrieben/reaktiviert werden?“ ist die Existenz die falsche Frage — eine abgemeldete Zeile ist keine Klammer-Anmeldung. Verstärkt wird es durch `skipShadowParent: true`: die Anmeldeseite schaltet das zentrale Netz aus und muss den Fall deshalb selbst abdecken.

**Fix-Vorschlag:** `parentAlreadyHasRow` auf den aktiven Stand einschränken: `const parentAlreadyHasRow = !!(myParentReg && myParentReg.Status && myParentReg.Status !== 'Abgemeldet');` — dann ist es identisch zu `parentAlreadyRegistered` und `shouldShadowRegisterParent` wird true. Der Schritt-3-Block ruft `doParentRegistration(true)` → registerForEvent findet über getMyRegistration die abgemeldete Zeile und geht in den Reaktivierungs-Pfad (EventContext.tsx:804-806), der Status UND CustomData korrekt setzt. Der `else if`-Zweig auf Zeile 967 (nur updateMyRegistration) bleibt dann dem Fall vorbehalten, für den er gedacht war: eine BESTEHENDE aktive Schatten-Zeile beim Nachbuchen eines weiteren Termins.

## 2. [hoch] Check-in per QR-Code liest die ÄLTESTE Zeile zur Adresse — eine abgemeldete Alt-Zeile weist eine angemeldete Person am Einlass ab

**Wo:** `services/events/registrationStatus.ts:415`  
**Im Ausgangsstand:** Identisch im Ausgangsstand: services/EventService.ts:8020 (getRegistrationByEmail). Der korrekte Nachbar steht dort auf Zeile 5806 (getMyRegistration, mit $orderby=Id desc). Der Umbau hat die Zeile nur verschoben.

**Szenario:** Voraussetzung sind zwei Zeilen zur selben Adresse — genau der Zustand, den v27.11 in getMyRegistration und v30.54 in utils/b2runBibImport.ts (`Bei zwei Zeilen zur selben Adresse gewinnt die AKTIVE — sonst gilt eine alte Abmeldung als aktueller Stand`) ausdrücklich als real behandeln. Entstehung: Person meldet sich ab (Zeile #50 → 'Abgemeldet'); später meldet eine Assistenz sie erneut an. Wegen ReadSecurity=2 sieht die Assistenz Zeile #50 nicht, getMyRegistration liefert ihr nichts, es entsteht Zeile #200 mit Status 'Angemeldet'. Am Event-Tag scannt das Check-in-Team ihren QR-Code. `processCode` (components/CheckInPage.tsx:640) ruft `getRegistrationByEmail`; SharePoint liefert ohne $orderby die Zeilen in Id-Reihenfolge, $top=1 nimmt also #50. Die Seite antwortet `<Name> — Anmeldung storniert` (CheckInPage.tsx:656-659) und bricht ab. Die Person steht mit gültigem QR-Code vor der Tür und gilt als abgemeldet; nur wer weiß, dass er stattdessen über das Namens-/Teilnehmer-ID-Suchfeld gehen muss, kommt rein. Derselbe Griff auf eine 'Eingecheckt'-Alt-Zeile meldet fälschlich „bereits eingecheckt“.

**Ursache:** `getRegistrationByEmail` schließt aus EINER Adresse auf GENAU EINE Zeile: `$top=1` ohne `$orderby` und ohne Status-Filter. Welche der mehreren Zeilen zurückkommt, entscheidet SharePoint (praktisch die kleinste Id = die älteste). Die Funktion hat genau EINEN Aufrufer — den QR-Scan am Check-in-Tisch — und ist damit die einzige Stelle der Registrierungs-Lookups, die den v27.11-Fix nicht mitbekommen hat.

**Fix-Vorschlag:** Query wie in getMyRegistration/markConsentPendingByEmail sortieren und den aktiven Stand bevorzugen: `&$orderby=Id desc` ergänzen und `email.trim()` statt `email` escapen. Sauberer noch (analog zu buildBibReport in utils/b2runBibImport.ts): `$top=20` holen und clientseitig die erste Zeile mit Status in ['Angemeldet','QR versendet','Eingecheckt'] nehmen, sonst die neueste — dann gewinnt die aktive Zeile auch dann, wenn sie NICHT die neueste ist.

## 3. [hoch] E-Mail-Änderung im Organizer Center pflegt DEX_Participants nicht nach — Event verschwindet aus „Meine Events“ und die Doppel-Anmelde-Prüfung ist für die neue Adresse blind

**Wo:** `components/admin/logic/useEditModalHandlers.ts:150`  
**Im Ausgangsstand:** Identisch im Ausgangsstand: components/AdminPage.tsx:1399 (`if (newEmail !== oldEmail) {` im saveEdit-Körper, Zeile 1348 ff.). Der Umbau hat den Körper 1:1 nach useEditModalHandlers.ts verschoben.

**Szenario:** Organizer öffnet in der Teilnehmerliste „Bearbeiten“ und korrigiert die E-Mail einer Person (Tippfehler bei manueller Anlage, Namensänderung nach Heirat) von alt.name@deloitte.de auf neu.name@deloitte.de. Speichern meldet Erfolg. Danach: (1) `DEX_Participants` trägt die EventNumber weiterhin unter alt.name@deloitte.de, für neu.name@deloitte.de existiert kein Eintrag. (2) Die Person öffnet „Meine Events“ — MyEventsPage.tsx:421 lädt `getMyEventNumbers()`; hat sie noch irgendein anderes Event im Register, greift der Altdaten-Fallback (MyEventsPage.tsx:600 `} else {`) NICHT, und die Nachlese-Schleife ab Zeile 566 nimmt nur Zeilen mit Status 'Abgemeldet' auf. Das Event fehlt komplett, sie hält sich für nicht angemeldet. (3) Sie meldet sich erneut an: die v28.22-Prüfung auf unsichtbare Doppel-Anmeldungen (components/registration/submitFlow.ts:851 `getEventNumbersForEmail(participantEmail)`) fragt das Register unter der NEUEN Adresse, bekommt nichts zurück, warnt nicht — und es entsteht eine zweite Zeile, die einen zweiten Platz belegt. Genau diese Kette beschreibt die App selbst in AdminActionsCard.tsx:610 als Symptom eines fehlenden Register-Eintrags. (4) Bei der späteren Abmeldung ruft der Cancel-Pfad `removeParticipantEvent(neueAdresse, eventNumber)` — findet keinen Datensatz, die EventNumber bleibt für immer auf der alten Zeile stehen und taucht in `analyzeRegistryAgainstLists` als verwaister Verweis auf.

**Ursache:** Die E-Mail ist laut CLAUDE.md der einzige Schlüssel des Datenmodells; DEX_Participants schlüsselt ausschliesslich darüber (`getParticipantByEmail`, `$filter=Email eq '...'`). Alle Schreibpfade, die diesen Schlüssel VERGEBEN, machen einen Dual-Write (registerForEvent, registerTeam, addTeamMember, cancelRegistration). Der einzige Pfad, der den Schlüssel ÄNDERT, macht ihn nicht: `saveEdit` behandelt ParticipantEmail wie ein gewöhnliches Textfeld und schiebt es in dasselbe `patch`-Objekt wie Vorname/JobTitle. Dass der Wert die Identität der Person in einer zweiten Liste ist, kommt im Code nicht vor.

**Fix-Vorschlag:** In `saveEdit` nach erfolgreichem `adminUpdateRegistration` den Registerwechsel nachziehen, wenn `newEmail !== oldEmail` und `selectedEvent.eventNumber` gesetzt ist — in dieser Reihenfolge (prüfbare Nebenbuchhaltung zuerst, CLAUDE.md): `await eventServiceRef.upsertParticipant(newVorname, newNachname, newEmail, selectedEvent.eventNumber, editingReg.Status === 'Warteliste' ? 'Warteliste' : 'Angemeldet')`, danach `await eventServiceRef.removeParticipantEvent(oldEmail, selectedEvent.eventNumber)`. Bei einem Klammer-Event zusätzlich über alle aktiven Sub-Event-Zeilen der Person laufen (deren EventNumbers hängen an derselben Adresse). Schlägt einer der beiden Schritte fehl, den Organizer darauf hinweisen statt still Erfolg zu melden. Zweitens: Da die Zeile nach dem Wechsel der alten Person gehört (Zeilen-Autor), zusätzlich `trySetItemAuthor(subsiteUrl, REG_LIST_NAME, editingReg.Id, newEmail)` aufrufen — sonst sieht die neue Adresse ihre eigene Zeile wegen ReadSecurity=2 nicht.

## 4. [hoch] Outlook-Update des Hauptevents wird am Sub-Event-Schalter vorbei verworfen — OutlookDirty wird trotzdem gelöscht

**Wo:** `components/wizard/logic/wizardSubmit.ts:1117`  
**Im Ausgangsstand:** components/EventCreationPage.tsx:5630 im Ausgangsstand — identische Bedingung.

**Szenario:** Event mit Sub-Events. Der Organizer ändert in Schritt 1 den Titel, wechselt in Schritt „Kommunikation" auf den Reiter eines Sub-Events, für das „Outlook-Termin deaktivieren" gesetzt ist, und speichert von dort. `switchCommTab` (components/wizard/logic/commTabs.ts:133) hat `setDisableOutlook(true)` aus dem Sub-Slot gesetzt — der rohe State hält jetzt den SUB-Wert. Das Update-Modal listet das Hauptevent trotzdem (die Erkennung nutzt korrekt `resolveTopLevelCommState().disableOutlook`, siehe outlookChanges.ts:142), der Organizer hakt es an → `confirmOutlookSave` setzt `pendingOutlookDirtyWriteRef.current = !topChecked` = false (outlookActions.ts:263), das wird in Zeile 650 als `OutlookDirty:false` geschrieben. Danach greift Zeile 1117 nicht, weil `disableOutlook` (Sub-Wert) true ist: KEIN `UpdateEvent` in der Queue. Ergebnis: Der Kalender aller Teilnehmer behält den alten Titel, der Dirty-Marker ist gelöscht, der Wizard bietet die Aktualisierung nie wieder an — und die App meldet „gespeichert".

**Ursache:** Der Save-Pfad prüft an dieser einen Stelle den rohen `disableOutlook`-State statt `effDisableOutlook` (= `topComm.disableOutlook`), obwohl derselbe Wert zwei Zeilen weiter oben schon aufgelöst vorliegt und in Zeile 640 in die Spalte geschrieben wird. Genau diese Falle ist an zwei anderen Stellen bereits behoben (v18.45/v18.50 in `detectOutlookRelevantChangesImpl`); der Queue-Zweig wurde damals nicht mitgezogen.

**Fix-Vorschlag:** In Zeile 1117 (und in der Zählung in Zeile 1102) `disableOutlook` durch `effDisableOutlook` ersetzen. Zusätzlich absichern: `OutlookDirty:false` erst schreiben, wenn `queueOutlookEvent` `true` geliefert hat — sonst geht der Marker auch bei einem 429 verloren.

## 5. [hoch] Shadow-Heal-Timer wird vom `events`-Update abgeraeumt und nie neu gestellt — die Nachheilung der Klammer-Zeilen laeuft nie

**Wo:** `context/EventContext.tsx:2809`  
**Im Ausgangsstand:** context/EventContext.tsx:5522-5543 — zeichengleich, der Fehler ist NICHT durch den Umbau entstanden.

**Szenario:** Eine Anmeldung fuer ein Sub-Event scheitert beim Schreiben der Klammer-Zeile (Throttling 429); der Merker landet in `utils/shadowHeal`. Beim naechsten App-Start soll er 20 s nach dem Boot abgearbeitet werden. Ablauf real: (1) `loadEvents` ruft `setEvents(mapped)` (EventContext.tsx:307) — der Effekt laeuft, setzt `shadowHealRanRef.current = true` und stellt den 20-s-Timer. (2) Wenige Sekunden spaeter meldet der Nachlauf die Teilnehmerzahlen und ruft `setEvents(withCounts)` (EventContext.tsx:316) — `events` ist ein NEUES Array, React fuehrt die Cleanup-Funktion aus und ruft `window.clearTimeout(t)`. (3) Der Effekt laeuft erneut, faellt sofort auf `if (shadowHealRanRef.current) return` — es wird KEIN neuer Timer gestellt. Ergebnis: der Timer feuert nie, in keiner Sitzung. Dieselbe Kette loest auch `ensureEventDocuments` aus (EventContext.tsx:347, `setEvents(prev => …)`), sobald der Nutzer „Meine Events" oeffnet. Folge fuer den Nutzer: Die fehlende Klammer-Zeile wird nie nachgetragen — die Person steht im Sub-Event, fehlt aber in der Teilnehmerliste/konsolidierten Matrix des Hauptevents, und der Merker verfaellt nach 14 Tagen ungenutzt. Genau vor diesem Muster warnt der Code an anderer Stelle ausdruecklich: `components/DexEventPlatform.tsx:615` — „v23.12 BUG-FIX: KEIN clearTimeout-Cleanup zurueckgeben … sonst bricht ein Re-Render (events-Aenderung) den Timer ab und er feuert nie." Dort wurde die Lehre gezogen, hier nicht.

**Ursache:** Der Effekt kombiniert einen Einmal-pro-Sitzung-Merker (`shadowHealRanRef`) mit einer Cleanup-Funktion, die den Timer beim Dependency-Wechsel abbricht. Der Merker verhindert das Neu-Armen, die Cleanup-Funktion verhindert das Feuern — zusammen ergibt das „nie". Die Dependency `events` aendert beim Boot garantiert mehrfach ihre Identitaet (mindestens zweimal durch `setEvents(mapped)` + `setEvents(withCounts)`), also lange vor Ablauf der 20 s.

**Fix-Vorschlag:** Den Timer nicht an den Dependency-Zyklus haengen. Entweder wie in `DexEventPlatform` bewusst OHNE Cleanup arbeiten (Kommentar dazu, damit es nicht als Versehen zurueckgebaut wird), oder den Timer in einem Ref halten und nur beim Unmount abraeumen:

```ts
const shadowHealTimerRef = React.useRef<number | null>(null);
// … im Effekt:
shadowHealTimerRef.current = window.setTimeout(() => { /* … */ }, 20000);
// KEIN return () => clearTimeout(t) mehr hier.
// Stattdessen einmalig:
React.useEffect(() => () => {
  if (shadowHealTimerRef.current !== null) window.clearTimeout(shadowHealTimerRef.current);
}, []);
```
Alternativ den Merker erst IM Timeout-Callback setzen, dann armt ein `events`-Update den Timer sauber neu (kostet aber je Update einen neuen Timer — die Ref-Variante ist sauberer).

## 6. [hoch] Team-Anmeldung: Promise.all über die Insert-Schreibvorgänge ohne Fehlerzähler — Teilerfolg wird als voller Erfolg gemeldet

**Wo:** `context/EventContext.tsx:1539`  
**Im Ausgangsstand:** context/EventContext.tsx:2881 — wortgleich; pre-existing.

**Szenario:** Ein Team-Lead meldet am Freigabetag um 09:00 ein 5-köpfiges B2Run-Team an. `reserveSeat(subsiteUrl, group, cap, 5)` reserviert atomar 5 Plätze (context/EventContext.tsx:1445/1451). Danach starten alle 5 `registerTeamMember`-POSTs GLEICHZEITIG (`Promise.all`) — bei einer Anmelde-Welle genau die Konstellation, die SharePoint mit 429 drosselt. Zwei POSTs scheitern; `registerTeamMember` fängt das ab und liefert `{ ok: false }` (services/events/registration.ts:371 ff., try/catch um den ganzen Body). `anyOk` ist true, also gibt `registerTeam` `{ ok: true }` zurück. In components/registration/submitFlow.ts:604-613 greift `if (!result.ok)` nicht → `setSubmitted(true)`, der Lead sieht die Erfolgsseite. Die Bestätigungs-Mail wird nur an die 3 erfolgreichen Zeilen geschickt und ihr Team-Block listet auch nur 3 Namen — die beiden fehlenden Personen erfahren nichts, stehen in keiner Teilnehmerliste und tauchen am Eventtag ohne QR-Code auf. Gleichzeitig steht `SeatsTaken` bei 5, obwohl nur 3 Zeilen existieren: zwei Phantom-Plätze blockieren zwei andere Anmeldungen (die auf die Warteliste rutschen), bis jemand „Zähler abgleichen" klickt. Nirgends wird ein Fehler protokolliert oder angezeigt.

**Ursache:** Schreiboperationen parallel per `Promise.all` ohne Fehlerzähler und ohne Kompensation — exakt das Muster, das CLAUDE.md für `deleteParticipantData` als Ursache des Register-Rückstands beschreibt („alle Schreibvorgänge liefen als EIN `Promise.all` gleichzeitig los … jeder Fehler wurde weggeworfen") und das dort auf sequentiell + Fehlerzähler umgestellt wurde. Hier steht es unverändert. `some(r => r.ok)` ist zudem die falsche Erfolgsdefinition: für eine Team-Anmeldung ist Teilerfolg schlimmer als Totalausfall, weil der Totalausfall wenigstens gemeldet würde. Und die reservierten Plätze werden bei Teilerfolg nicht auf die tatsächliche Anzahl zurückgeführt.

**Fix-Vorschlag:** Fehlerzähler einführen, Teilerfolg als solchen melden und die überzähligen Reservierungen zurückgeben:

```ts
const results = await Promise.all(insertPromises);
const failed = results.filter(r => !r.ok);
if (failed.length === results.length) return { ok: false, reason: 'insert-failed' };
if (failed.length > 0) {
  // Reservierte, aber nicht belegte Plaetze zurueckgeben — sonst blockieren
  // sie andere Anmeldungen (adjustSeatCounterField ist additiv + ETag-CAS).
  await eventService.adjustSeatCounterField(
    subsiteUrl, eventService.seatFieldFor(effectiveStarterType || ''), -failed.length,
  ).catch(() => { /* */ });
  return { ok: false, reason: `partial-insert:${failed.map(f => f.email).join(',')}` };
}
```

In components/registration/submitFlow.ts:604 einen Zweig für `partial-insert:` ergänzen, der die betroffenen Adressen nennt („Für folgende Personen ist die Anmeldung nicht durchgekommen: … — bitte einzeln nachmelden"), statt die Erfolgsseite zu zeigen. Alternativ (robuster, weniger 429): die Inserts sequentiell in der Schleife ausführen — bei N ≤ 20 ist das kein Performance-Thema, und dieselbe Begründung steht schon 150 Zeilen weiter oben für die Doppel-Anmelde-Prüfung.

## 7. [hoch] reconcileCounters schließt genau die Split-Kapazitäts-Events aus — Filter über maxParticipants, wo die Kapazität in den Split-Feldern steht

**Wo:** `context/EventContext.tsx:488`  
**Im Ausgangsstand:** context/EventContext.tsx:1268-1279 — identischer Filter, identisches totes `isSplit`. Beim Umbau unverändert übernommen.

**Szenario:** B2Run-Event mit geteilten Kapazitäten (durchstarterCapacity 200, funstarterCapacity 150, `maxParticipants` = 0 — genau so legt der Wizard es an). Ein Admin startet die App: `reconcileCounters()` läuft, aber `(e.maxParticipants || 0) > 0` ist für dieses Event `false` → das Event ist NIE Ziel des Abgleichs. Damit läuft der einzige App-seitige Selbstheilungs-Lauf für `SeatsTaken`/`SeatsTakenDurch`/`SeatsTakenFun` an den Events vorbei, die als einzige ZWEI Zähler synchron halten müssen. Konkrete Folge: Ein Teilnehmer meldet sich selbst ab; `releaseSeatAfterCancel` findet `WaitlistTaken` = -1 („unbekannt", Bestands-Event ohne je gepflegtes Feld) und zählt fail-closed NICHTS herunter (seats.ts:511-515). Der Gruppenzähler bleibt dauerhaft zu hoch; die nächste Anmeldung für diese Gruppe bekommt von `reserveSeat` 'full' und landet auf der Warteliste, obwohl ein Platz frei ist — und die App repariert das von sich aus nie, weil das Event aus der Zielmenge gefiltert ist. Dass `isSplit` drei Zeilen tiefer berechnet wird, obwohl kein Split-Event den Filter passieren kann, zeigt die Absicht: der Lauf SOLLTE Split-Events erfassen.

**Ursache:** Der Filter benutzt `maxParticipants` als „hat dieses Event überhaupt eine Kapazität?"-Prüfung. Bei geteilten Kapazitäten ist `maxParticipants` per Konvention 0 (CLAUDE.md: „Bei geteilten Kapazitäten ist `maxParticipants` 0 … Wer `maxParticipants` als Obergrenze prüft, prüft dort **gar nichts**"). Die Split-Felder werden im Filter nicht berücksichtigt.

**Fix-Vorschlag:** Den Filter auf eine effektive Kapazität umstellen, z.B. `const effCap = (e.maxParticipants || 0) > 0 ? (e.maxParticipants || 0) : ((e.durchstarterCapacity || 0) + (e.funstarterCapacity || 0));` und dann `effCap > 0` prüfen. Das bereits vorhandene `isSplit` unterhalb bleibt unverändert und wird dadurch endlich wirksam.

## 8. [hoch] setWaitlistPosition kennt die getrennten Gruppen-Wartelisten nicht — „Platz ändern" verschiebt in die falsche Liste

**Wo:** `services/events/seats.ts:78`  
**Im Ausgangsstand:** services/EventService.ts:6096 (`const waitlist = allItems.filter(i => i.Status === 'Warteliste').sort(byTidThenId);`) bzw. components/AdminPage.tsx:12927 (`setWlPosModal({ reg, currentPos: truePos != null ? truePos : (i + 1), total: regs.length });`) — zeichengleich, der Fehler ist beim Umbau 1:1 mitgewandert.

**Szenario:** Split-Kapazitäts-Event (B2Run) mit getrennten Wartelisten. Durchstarter-Warteliste: 10 Personen (TeilnehmerID 101–110), Funstarter-Warteliste: 5 Personen (TID 111–115). Der Organizer öffnet im Organizer Center die Tabelle „Warteliste Funstarter", sieht dort Person X auf **Platz 4**, klickt „Platz ändern" und trägt **2** ein (Wunsch: X soll Zweite der Funstarter-Warteliste werden). setWaitlistPosition baut aber EINE globale Warteliste [101…115], findet X bei fromIdx=13 und setzt toIdx=1 — X landet auf globaler Position 2, also zwischen dem ersten und zweiten DURCHSTARTER-Wartenden. Ergebnis: X ist innerhalb der Funstarter-Warteliste auf Platz 1 (nicht 2), alle zehn Durchstarter-Wartenden bekommen neue TeilnehmerIDs und rücken je eine Position nach hinten, und die Erfolgsmeldung nennt „jetzt Platz 2 (vorher 14)" — Zahlen, die der Organizer nirgends gesehen hat. Weil das Modal `max = wlPosModal.total = regs.length` (= 5) erzwingt, sind für die Funstarter-Liste überhaupt nur globale Positionen 1–5 erreichbar: JEDE Eingabe schiebt die Person an den Anfang ihrer Gruppe und mischt gleichzeitig die fremde Gruppe um. Die Funktion ist für die zweite Gruppe damit vollständig unbrauchbar.

**Ursache:** `setWaitlistPosition` liest nur `$select=Id,Status,TeilnehmerID` und filtert ausschließlich auf `Status === 'Warteliste'` — `PreferredStarterType` wird weder abgefragt noch berücksichtigt. Die UI (waitlistDurch/waitlistFun/waitlistUnassigned in AdminPage.tsx:1731-1739, Rang aus waitlistTruePos, das je Gruppe rankt) und `promoteFirstWaitlistItem` (filtert `PreferredStarterType eq …`, waitlist.ts:78-81) rechnen dagegen PRO GRUPPE. `targetPosition` wird als Index in die globale Liste interpretiert, kommt aber als Rang innerhalb der Gruppe herein.

**Fix-Vorschlag:** In `setWaitlistPosition` die Gruppe mitlesen und mitfiltern: `$select=Id,Status,TeilnehmerID,PreferredStarterType`, einen optionalen Parameter `group?: string` ergänzen und die Zielliste als `allItems.filter(i => i.Status === 'Warteliste' && (!group || (i.PreferredStarterType || '') === group))` bilden. Die Umnummerierung muss dann die ANDEREN Wartelisten-Zeilen in ihrer bisherigen TeilnehmerID-Reihenfolge unangetastet mitschreiben (globale Sequenz weiterhin 1..N Aktive, dann alle Wartenden — nur die Reihenfolge INNERHALB der Zielgruppe ändert sich). Aufrufer: `WaitlistPositionModal` reicht `reg.PreferredStarterType` durch, wenn `selectedEvent.durchstarterCapacity/funstarterCapacity` gesetzt sind und `splitSharedWaitlist` NICHT aktiv ist.

## 9. [mittel] Demo-Vorlagen schreiben eine 10-Zeichen-Frist („YYYY-MM-DD") in einen State, den der DatePicker als UTC-Mitternacht liest — die Frist springt von 23:59 auf 02:00

**Wo:** `components/wizard/hooks/useWizardOptionState.tsx:497`  
**Im Ausgangsstand:** Ausgangsstand: components/EventCreationPage.tsx:3567, :3591, :3615, :3656 (jeweils `setRegistrationDeadline(fmtDate(deadline))`) mit fmtDate/fmtDatetime bei :3480-3487 — identisch.

**Szenario:** Nachgerechnet mit TZ=Europe/Berlin: `new Date("2026-09-05")` ist laut ECMAScript ein Date-Only-Form und wird als UTC-Mitternacht geparst → `Sat Sep 05 2026 02:00:00 GMT+0200`. Ein Organizer klickt im Wizard auf „Demo" (loadDemoStandard) — im Feld „Anmeldung bis" steht daraufhin „05.09.2026, 02:00", obwohl die Vorlage 23:59 berechnet hatte. Zwei Folgen: (a) Speichert er ohne Anfassen, schreibt `deadlineToEndOfDayIso` (EventCreationPage.tsx:392-398) wegen der 10 Zeichen 23:59 — Anzeige und gespeicherter Wert gehen um 22 Stunden auseinander. (b) Korrigiert er die offensichtlich falsche „02:00" durch erneutes Anklicken desselben Tages im Kalender, liefert `onChange` `date.getHours() === 2` und schreibt „2026-09-05T02:00"; jetzt sind es 16 Zeichen, `deadlineToEndOfDayIso` lässt sie unverändert und die Anmeldung schließt am 05.09. um 02:00 statt 23:59. Teilnehmer, die am 05.09. tagsüber buchen wollen, bekommen „Anmeldefrist abgelaufen" — einen ganzen Tag zu früh gesperrt. Im Winter ist es 01:00 statt 02:00, der Effekt bleibt.

**Ursache:** Der State `registrationDeadline`/`lastDeregisterDate` trägt überall sonst das 16-Zeichen-Format „YYYY-MM-DDTHH:MM" (Berliner Lokalzeit) — so laden ihn `useWizardEventFieldState.ts:66/75` über `isoToLocal`, so schreiben ihn alle DatePicker. Nur die vier Demo-Loader legen die 10-Zeichen-Form hinein. `new Date()` behandelt beide Formen unterschiedlich (Date-Only = UTC, Date-Time ohne Offset = Browser-Lokalzeit), und `deadlineToEndOfDayIso` verzweigt zusätzlich auf `dateStr.length === 10`. Damit hängt an einem State-Wert zweierlei Bedeutung.

**Fix-Vorschlag:** In allen vier Demo-Loadern `fmtDate(deadline)` durch `fmtDatetime(deadline)` ersetzen (useWizardOptionState.tsx:497, :498, :521, :522 und wizardTemplates.ts:348, :349, :425, :426). Das Date trägt die 23:59 bereits (`beforeNextSaturday(n, 23, 59)`), es geht nur die Uhrzeit-Komponente verloren. Danach ist der State durchgängig 16-stellig; `deadlineToEndOfDayIso` behält seinen 10-Zeichen-Zweig als Alt-Pfad, wird aber nicht mehr getroffen.

## 10. [mittel] Die in v30.63 dokumentierte Organizer-Selbstheilung des Platzzählers ist nicht verdrahtet — `onlyMine` wird nirgends übergeben

**Wo:** `components/DexEventPlatform.tsx:584`  
**Im Ausgangsstand:** components/DexEventPlatform.tsx:581-599 — zeichengleich, ebenfalls ohne `onlyMine`. Vorbestehend, nicht durch den Umbau entstanden.

**Szenario:** Ein Organizer (kein Admin) betreut ein volles Event. Der Zähler ist nach oben gedriftet (Selbst-Abmeldung bei unbekanntem Wartelisten-Stand zählt fail-closed nicht herunter, seats.ts:511-515). Laut Release Notes v30.63 („Jetzt läuft er auch für Organizer, per neuem `onlyMine`-Filter auf die selbst betreuten Events beschränkt") soll der Abgleich beim Start des Organizers laufen und das heilen. Er läuft nicht: Der Effect bricht bei `if (!isAdmin) return;` ab, und der einzige Aufruf im ganzen Repo (`grep -rn reconcileCounters`) übergibt kein `opts`. Bis sich ein echter Admin anmeldet UND die 6-h-Drossel (`dex_counter_reconcile_lastrun`) abgelaufen ist, bleibt der Zähler falsch — die Anmeldeseite zeigt seit v30.62 genau diesen Zähler an, meldet also „ausgebucht", während Plätze frei sind. `canCreateEvents` steht bereits im Deps-Array, wird im Rumpf aber nie benutzt: der Organizer-Zweig war vorgesehen und ist nie eingebaut worden.

**Ursache:** `reconcileCounters(opts?: { onlyMine?: boolean })` hat den Parameter bekommen (EventContext.tsx:462, samt ausführlicher Begründung im Kommentar ab Zeile 464), aber die einzige Aufrufstelle in DexEventPlatform.tsx wurde nicht nachgezogen — weder das `isAdmin`-Gate noch das Argument.

**Fix-Vorschlag:** Gate und Aufruf zusammenführen: `if (!isAdmin && !canCreateEvents) return;` und `reconcileCounters(isAdmin ? undefined : { onlyMine: true })`. Die Drossel sollte dabei pro Rolle getrennt geschlüsselt werden (z.B. `dex_counter_reconcile_lastrun_org`), sonst blockiert ein Organizer-Lauf den nächsten Admin-Lauf für sechs Stunden.

## 11. [mittel] Endgültige Löschungen zählen Versuche statt Erfolge — deleteRegistration liefert `false`, das nirgends geprüft wird

**Wo:** `components/admin/participants/CancelledList.tsx:182`  
**Im Ausgangsstand:** components/AdminPage.tsx:13150 — wortgleich; pre-existing. Dieselbe Stelle noch einmal in components/admin/logic/useCancelPipeline.ts:258 (`cleanupShadowDuplicates`, Basis AdminPage.tsx:524) und useCancelPipeline.ts:192 (`performSilentDuplicateDelete`).

**Szenario:** Organizer klickt in der konsolidierten Abmelde-Matrix auf „Person überall löschen". Der Bestätigungsdialog verspricht „Alle 12 Einträge dieser Person (Gesamt-Event + Sub-Events) werden endgültig entfernt". Bei 3 der 12 Zeilen liegt die Zeile auf einer Sub-Event-Subsite, auf der der Organizer kein Delete-Recht hat (v30.37-Fall) oder SharePoint drosselt — `deleteRegistration` fängt das ab und gibt `false` zurück (services/events/registrationEdit.ts:692-699). Da der Rückgabewert nicht ausgewertet wird und die Funktion nicht wirft, greift auch das `catch` nicht. Anschließend schreibt der Audit-Log `count: 12, everywhere: true` — die revisionssichere Historie behauptet eine vollständige Löschung, die nicht stattgefunden hat. Es gibt keine Fehlermeldung; die betroffene Person bleibt in 3 Terminen angemeldet, bekommt weiter Mails und Outlook-Termine und zählt weiter in die Kapazität. Gleiches Bild bei „Doppelte Klammer-Zeilen bereinigen" (useCancelPipeline.ts:258): dort wird `removed += 1` unmittelbar nach dem ungeprüften Aufruf hochgezählt und am Ende als grüne Erfolgsmeldung „N doppelte Klammer-Zeile(n) entfernt." angezeigt — N ist die Zahl der VERSUCHE. Auffällig: der Einzel-Löschpfad 100 Zeilen weiter oben in derselben Datei (CancelledList.tsx:88) macht es richtig (`const ok = await …; if (ok) { … } else { showAlert('Löschen fehlgeschlagen.') }`).

**Ursache:** `deleteRegistration` ist bewusst nicht-werfend und meldet Misserfolg über den Rückgabewert (`return resp.ok` bzw. `catch → return false`). Alle drei Sammel-Löschpfade rufen sie im `await` ohne Zuweisung auf und verlassen sich stattdessen auf ein `try/catch`, das per Konstruktion nie auslöst. Damit ist ein fehlgeschlagenes DELETE von einem erfolgreichen ununterscheidbar — dieselbe Klasse wie das leere Leseergebnis, nur auf der Schreibseite: Erfolgsmeldung und Audit-Eintrag sind Aussagen über Aufrufe, nicht über Daten. CLAUDE.md verlangt für Löschpfade ausdrücklich „sequentiell, mit Fehlerzähler".

**Fix-Vorschlag:** In allen drei Pfaden den Rückgabewert auswerten und die Meldung ehrlich machen. In CancelledList.tsx:181-186:

```ts
let delOk = 0; let delFailed = 0;
for (const t of targets) {
  const ok = await eventServiceRef.deleteRegistration(t.sub, t.id).catch(() => false);
  if (ok) delOk += 1; else delFailed += 1;
}
await eventServiceRef.writeChangeLog({ ..., details: { deletedStatus: 'Abgemeldet', count: delOk, failed: delFailed, everywhere: delFailed === 0 } }).catch(() => { /* */ });
if (delFailed > 0) {
  showAlert(isDe
    ? `${delOk} von ${targets.length} Einträgen gelöscht — ${delFailed} konnten nicht entfernt werden (fehlende Rechte auf dem Termin oder Drosselung). Bitte „Organizer-Berechtigungen reparieren" ausführen und erneut versuchen.`
    : `${delOk} of ${targets.length} entries deleted — ${delFailed} could not be removed.`, { variant: 'error' });
}
```

Analog in useCancelPipeline.ts:258 (`removed` nur bei `ok` erhöhen, `failed` zählen und in der Abschlussmeldung nennen, `variant` bei `failed > 0` auf `'error'`) und in useCancelPipeline.ts:192 (`performSilentDuplicateDelete`: bei `false` einen Toast statt stiller Rückkehr).

## 12. [mittel] Hotelplanung: der Warnkasten „Unterkunftsbedarf ohne Hotel“ liest nur die Klammer-Zeile — bei einem Klammer-Event bleibt er leer, obwohl Personen ohne Zimmer sind

**Wo:** `components/HotelPlanningPanel.tsx:991`  
**Im Ausgangsstand:** Identisch im Ausgangsstand: components/HotelPlanningPanel.tsx:991-999 — zeichengleich, die Datei wurde vom Umbau nicht angefasst.

**Szenario:** Klammer-Event, bei dem die Hotel-Frage im Formular des Termins steht (der Normalfall laut CLAUDE.md: „Wer sich über ein Sub-Event angemeldet hat, hat auch die Hotel-Frage in dessen Formular beantwortet, und die Klammer-Zeile ist bei einer Klammer nur eine Schattenzeile ohne Antworten“). Verschärft für jede Person, die NICHT über die Anmeldeseite kam: das zentrale Netz in EventContext.tsx:1331-1336 legt die Klammer-Zeile mit `registerForEvent(parentEv.id, {}, ...)` an — CustomData ist buchstäblich `{}`. Der Organizer öffnet „Hotels & Übernachtungen“, verteilt manuell und verlässt sich auf den orangen Kasten „N Person(en) mit Unterkunftsbedarf ohne Hotel“ (Zeile 1468-1483, Text: „genau der Fall, der in einer Excel-Liste durchrutscht“). Der Kasten rendert gar nicht, weil `JSON.stringify` der leeren Schattenzeile weder 'hotel' noch 'yes' enthält — obwohl `autoDistribute` dieselben Personen über `wishOf` sehr wohl als Kandidaten führt. Die Ansicht sagt „alles versorgt“, und genau die Person ohne Zimmer reist an. Zweite Lücke im selben Ausdruck: wer den Zeitraum über ein `daterange`-Feld angegeben hat (v28.63), wird nicht erkannt, weil `formStayOf` nicht abgefragt wird.

**Ursache:** v29.3 hat die Antwort-Auflösung `answerRowsOf(p)` (Klammer-Zeile + alle Sub-Event-Zeilen derselben Adresse) eingeführt und `wishOf`/`formStayOf` darauf umgestellt — dieser dritte Leser derselben Frage wurde übersehen und arbeitet weiter auf `p` allein. Damit ist er die von CLAUDE.md beschriebene Ausnahme („Antworten stehen dort, wo angemeldet wurde — nicht auf der Klammer“), nur an einer anderen Stelle als 2029.2.

**Fix-Vorschlag:** Den Filter über dieselbe Quelle laufen lassen wie `autoDistribute`, dann verschwindet auch die Heuristik: `const wantsHotelWithout = React.useMemo(() => people.filter(p => !(p.Hotel || '').trim() && wishOf(p) === true), [people, wishOf]);`. `wishOf` deckt beide Ebenen ab, kennt die Zusatznächte-Ausnahme (v28.59) und wertet `formStayOf` mit aus; ausserdem entfällt der Blob-Scan, der heute jede Zeile mit irgendeinem 'yes' irgendwo im JSON als Hotel-Wunsch liest. Der useMemo braucht `wishOf` in den Deps (ist bereits ein useCallback).

## 13. [mittel] Konsolidierte Matrix: `hasParentReg` und die Parent-Feld-Auflösung ignorieren den Status und nehmen die älteste Zeile — abgemeldete Klammer-Zeile schluckt die Fehlermeldung und fängt die Bearbeitung ab

**Wo:** `components/admin/sections/ConsolidatedView.tsx:138`  
**Im Ausgangsstand:** Identisch im Ausgangsstand: components/AdminPage.tsx:5491 (`const hasParentReg = (emailKey: string): boolean =>`), die `registrations.find`-Varianten im selben Block. Der Umbau hat den JSX-Block 1:1 nach ConsolidatedView.tsx bzw. die Handler nach createKlammerActions.ts verschoben.

**Szenario:** Ausgangslage wie im vorigen Fund: eine Person hat aktive Termin-Zeilen, ihre Klammer-Zeile steht auf 'Abgemeldet' (gezielte Klammer-Abmeldung über den Dereg-Dialog, oder der Anmeldeseiten-Fehler oben). In der konsolidierten Matrix: (1) `hasParentReg` liefert true, weil kein Status geprüft wird → der rote Kasten „Fehlende Klammer-Anmeldung“ (Zeile 361) listet sie NICHT, und die Aktion „Zur Klammer hinzufügen“ ist ausgeblendet (Zeile 1068, `!hasParentReg`). Der Organizer hat keinen Weg, den Zustand zu erkennen oder zu reparieren. (2) Die Aktion „Felder“ (Zeile 1067/1080 → `openMainFieldsEdit`) ist dagegen sichtbar, und createKlammerActions.ts:313 nimmt wieder dieselbe Zeile: der Organizer trägt Hotel/Verpflegung in eine ABGEMELDETE Zeile ein. HotelPlanningPanel liest nur ACTIVE_STATI, sieht die Werte nie; die Person bleibt ohne Zimmer. (3) Analog schreibt „Assistenz zuordnen“ (createKlammerActions.ts:266) den RegisteredBy/Autor auf die abgemeldete Zeile. Zweite Variante desselben Fehlers: Bei ZWEI Klammer-Zeilen zur selben Adresse — laut DuplicateRegHintBox.tsx:58 ein ausdrücklich toleriertes „Aufräumen ist optional“-Szenario — liefert `getAllRegistrations` mit `$orderby=Id asc` (registrationEdit.ts:665) immer die ÄLTESTE; steht die aktuelle Antwort auf der neueren Zeile, zeigt die Matrix die veraltete, weil der Sub-Event-Fallback (Zeile 959) nur bei LEEREM Wert greift.

**Ursache:** Die Matrix schliesst aus dem E-Mail-Schlüssel auf genau eine Klammer-Zeile — mit `some`/`find` ohne Status-Filter und ohne definierte Reihenfolge. `registrations` kommt aus `getAllRegistrations` und enthält bewusst ALLE Zeilen inkl. 'Abgemeldet' (die Abmeldungen-Liste braucht sie). Dass genau dieselbe Suche eine Zeile weiter unten im Dereg-Dialog korrekt auf ACTIVE gefiltert wird, zeigt, dass es Vergessen und keine Absicht ist.

**Fix-Vorschlag:** Eine gemeinsame Auflösung einführen und an allen fünf Stellen benutzen — `const activeParentRegOf = (emailKey: string) => registrations.filter(r => (r.ParticipantEmail || '').toLowerCase().trim() === emailKey).sort((a,b) => (b.Id||0)-(a.Id||0)).find(r => ACTIVE.indexOf(r.Status||'') >= 0)`, also: nur aktive Zeilen, davon die neueste. `hasParentReg` = `!!activeParentRegOf(key)` — dann meldet der rote Kasten den Fall wieder und „Zur Klammer hinzufügen“ repariert ihn (registerForEvent geht über den Reaktivierungs-Pfad, es entsteht keine Dublette). `openMainFieldsEdit`/`submitAssignAssistant` in createKlammerActions.ts auf dieselbe Funktion umstellen.

## 14. [mittel] QR-Mail-Konfiguration eines Sub-Events geht beim nächsten Wizard-Speichern verloren (headerImage, blockLang, blockNote)

**Wo:** `components/wizard/hooks/useWizardVisibilityState.tsx:367`  
**Im Ausgangsstand:** components/EventCreationPage.tsx:1784-1804 im Ausgangsstand — zeichengleicher Filter.

**Szenario:** Organizer Center → QR-Mail eines Sub-Events gestalten: Kopf-Bild auf „Event-Foto" + volle Breite stellen, Block-Sprache auf DE, eigene Notiz unter dem Code. `saveQrMailOverride` (components/admin/logic/createQrMailActions.tsx:181-186) schreibt das als `QRCode.headerImage` / `.blockLang` / `.blockNote` in `EmailTemplateOverrides` der Sub-Event-Zeile. Danach öffnet derselbe Organizer das Event im Wizard und speichert (aus irgendeinem Grund — Titel, Frist, egal). Die Sub-Event-Hydration kopiert aus dem `QRCode`-Eintrag nur die Text- und Formatfelder, `persistSubEvents.ts:198-205` schreibt genau diesen beschnittenen Stand zurück. Ergebnis: Die QR-Mails dieses Termins zeigen wieder das DEX-Orb statt des Event-Fotos, die Block-Sprache kippt auf die Mail-Sprache und die Notiz ist weg — ohne Meldung, ohne dass jemand die QR-Mail angefasst hat. Beim HAUPTEVENT passiert das nicht: dort landet der komplette Eintrag ungefiltert in `rest` (Zeile 116) und wird unverändert zurückgeschrieben. Genau diese Asymmetrie macht den Fall schwer auffindbar.

**Ursache:** `EmailOverrideEntry` (components/wizard/emailOverrideEntry.ts) kennt nur Text- und Formatfelder; die Sub-Event-Hydration baut den Eintrag Feld für Feld daraus neu auf, statt ihn durchzureichen. Die v30.52/v30.60-Erweiterungen des QR-Overrides (`headerImage`, `blockLang`, `blockNote`) wurden im Organizer-Center-Pfad ergänzt, aber nicht im Wizard-Rundlauf. Zusätzlich fehlt `headerImage` in der Aufnahme-Bedingung: Ein Override, der NUR das Kopf-Bild ändert, fiele komplett durch.

**Fix-Vorschlag:** Den Eintrag durchreichen statt neu zusammensetzen — `filtered[key] = { ...val, subject: val.subject || '', heading: val.heading || '', bodyHtml: val.bodyHtml || '' }` — und `val.headerImage || val.blockLang || val.blockNote` in die Aufnahme-Bedingung mit aufnehmen. `EmailOverrideEntry` um die drei optionalen Felder ergänzen, damit der Typ die Erweiterung künftig erzwingt.

## 15. [mittel] Rollierende Fristen rechnen mit 24-h-Tagen statt Kalendertagen — nach der Sommerzeit-Umstellung fällt die Frist auf den Vortag

**Wo:** `components/EventCreationPage.tsx:256`  
**Im Ausgangsstand:** Ausgangsstand: components/EventCreationPage.tsx:519-523 (rollingDeadlineIso) und utils/eventFormat.ts:64-75 (subEventRegDeadline) — zeichengleich, der Fehler ist nicht durch den Umbau entstanden.

**Szenario:** Office-Tage-Reihe (Kalender-Termine, Ganztags, Start 00:00 Berlin) über Ende März, Regel „Anmeldung bis 1 Tag vor dem jeweiligen Termin". Nachgerechnet mit TZ=Europe/Berlin: Termin 27.03.2026 → Frist 26.03. 00:00 (Muster: Vortag 00:00). Termin 30.03.2026 (der Montag nach der Zeitumstellung) → Start 2026-03-29T22:00Z, minus 86400000 ms = 2026-03-28T22:00Z = **28.03. 23:00 Berlin** statt 29.03. 00:00. Für diesen einen Termin endet das Buchungsfenster damit 25 Stunden früher als bei allen Geschwistern; wer am 29.03. den 30.03. buchen will, bekommt in EventSpecificSection.tsx:517-518 (`deadlinePassed`) einen gesperrten Kalendertag mit „Anmeldefrist abgelaufen". Die Liste im Wizard (SubEventsSection.tsx:934) und die Teilnehmeransicht zeigen dazu „Frist 28.03." — ein Datum, das aus der Reihe fällt, während CapacityStep.tsx:463-467 weiterhin „Entspricht der rollierenden Regel" meldet, weil der Abgleich dieselbe fehlerhafte Rechnung benutzt. Im Herbst kippt es andersherum: Termin 26.10.2026 → Frist 25.10. **01:00** statt 00:00.

**Ursache:** `amount * 86400000` verschiebt einen absoluten UTC-Zeitpunkt um exakt 24 h. Über eine DST-Grenze hinweg entspricht das keinem Kalendertag mehr: der Berliner Offset wechselt zwischen +1 und +2, das Ergebnis liegt eine Stunde daneben — und weil die erzeugten Termine bewusst auf 00:00 stehen (toggleDaySubEventImpl in wizard/logic/wizardMisc.ts:288-289), rutscht diese Stunde über Mitternacht auf den Vortag. Der Kommentar über der Funktion beansprucht ausdrücklich „glatte Tagesgrenzen", was nur ohne DST-Wechsel stimmt.

**Fix-Vorschlag:** Bei `unit === 'days'` im Berliner Kalender rechnen statt in Millisekunden: Start über `isoToLocal(startIso)` in „YYYY-MM-DDTHH:MM" zerlegen, nur den Datumsteil um `amount` Tage verschieben (z.B. über `Date.UTC(y, m-1, d)` + `setUTCDate`) und das Ergebnis mit derselben Uhrzeit wieder durch `berlinLocalToUtcIso` schicken. `unit === 'hours'` darf weiter mit 3600000 rechnen (Stunden sind absolut gemeint). Die Rechnung gehört danach EINMAL in `utils/eventFormat.ts` und wird von EventCreationPage.tsx importiert — die Kopie in `subEventRegDeadline` (utils/eventFormat.ts:74) muss zwingend mitgezogen werden, sonst widersprechen sich der Wizard-Wert und der Anmeldeseiten-Fallback zusätzlich.

## 16. [mittel] Sub-Event-Outlook-Update: Queue-Ergebnis wird ignoriert, OutlookDirty wird trotzdem auf false gesetzt

**Wo:** `components/wizard/logic/wizardSubmit.ts:1144`  
**Im Ausgangsstand:** components/EventCreationPage.tsx:5657-5658 im Ausgangsstand.

**Szenario:** Terminreihe mit vielen Sub-Events. Der Organizer ändert Zeiten und hakt im Update-Modal alle Termine an. Nach rund 20 Schreibzugriffen drosselt SharePoint (429) — `queueOutlookEvent` fängt das intern ab und liefert `false` (services/events/outlookQueue.ts:169-172), es WIRFT nicht. Der `catch` greift also nie, `OutlookDirty:false` wird für jeden dieser Termine geschrieben. Ergebnis: Für die betroffenen Termine steht kein Eintrag in DEX_Outlook, der Kalender bleibt auf dem alten Stand, und weil der Dirty-Marker weg ist, meldet der Wizard beim nächsten Öffnen keine offene Änderung mehr. Derselbe Fehler in der Aktion „Alle Termine aktualisieren" (components/wizard/logic/outlookActions.ts:189-190): dort zählt `done += 1` jeden fehlgeschlagenen POST als Erfolg und die Meldung sagt „N Termine angestoßen".

**Ursache:** `queueOutlookEvent` meldet Fehlschlag als Rückgabewert, nicht als Ausnahme — dieselbe Klasse, die in v29.21 (`updateEvent`) und v29.48 (`deleteEvent`) schon zweimal zugeschlagen hat. Hier wurde der Rückgabewert nicht ausgewertet und der Dirty-Marker vorbehaltlos gelöscht.

**Fix-Vorschlag:** Rückgabewert prüfen: `const ok = await svc.queueOutlookEvent(...); if (ok) await updateEvent(subId, { OutlookDirty: false }, { skipReload: true }); else failedOutlookTitles.push(subTitle);` — und die gescheiterten Termine am Ende namentlich melden (wie `failedSubTitles` in persistSubEvents.ts). In `triggerOutlookUpdateAllImpl` `done`/`failed` am Rückgabewert statt am `catch` festmachen.

## 17. [mittel] Wartelisten-Platz-Dialog rechnet mit der gefilterten Trefferzahl statt mit der echten Wartelistenlänge

**Wo:** `components/admin/participants/WaitlistTables.tsx:132`  
**Im Ausgangsstand:** components/AdminPage.tsx:12927 (setWlPosModal-Aufruf im Wartelisten-Block) und components/AdminPage.tsx:16655 (valid-Prüfung im Wartelisten-Platz-Modal) — identischer Code, der Fehler ist nicht durch den Umbau entstanden.

**Szenario:** Event ohne Split-Kapazität, 12 Personen auf der Warteliste. Der Organizer tippt im Suchfeld über der Teilnehmerliste „meier“. Die Wartelisten-Tabelle zeigt jetzt genau eine Zeile, korrekt beschriftet mit „#1 (Platz 7)“ — die Platz-Anzeige nutzt bewusst waitlistTruePos aus der UNGEFILTERTEN Liste. Klick auf „Platz ändern“ in dieser Zeile: Der Dialog meldet „… steht aktuell auf Platz 7 von 1.“, das Zahlenfeld bekommt max=1 und `valid` verlangt parsed <= 1. Jeder Wunschwert außer 1 lässt „Platz setzen“ deaktiviert; die Person auf Platz 3 zu setzen ist bei aktiver Suche unmöglich. Der einzige akzeptierte Wert ist die Vorbelegung '1' — wer den Knopf trotzdem drückt, schiebt die Person ganz nach oben, obwohl er etwas anderes wollte. Ohne Suchbegriff funktioniert derselbe Klick korrekt (total = 12).

**Ursache:** `renderWaitlistTable(title, regs, accent)` bekommt als `regs` immer eine such-gefilterte Liste: `waitlistRegs = registrations.filter(r => r.Status === 'Warteliste').filter(matchesSearch)` (components/AdminPage.tsx:1495), und `waitlistDurch`/`waitlistFun`/`waitlistUnassigned` sind Teilmengen davon (AdminPage.tsx:1731-1739). Für die ANZEIGE des Platzes wurde das Problem in v26.31 bereits gelöst — `waitlistTruePos` (AdminPage.tsx:1745-1760) rankt ausdrücklich über `registrations.filter(r => r.Status === 'Warteliste')`, also ungefiltert, und `posDisplay` unterscheidet zusätzlich den Filterfall. Das mitgereichte `total` hat diese Behandlung nie bekommen und blieb `regs.length`, also die Anzahl der Suchtreffer. Im Modal ist `total` gleich drei Dinge: Obergrenze der Validierung, `max` des Eingabefelds und der Nenner im Erklärtext — alle drei sind bei aktiver Suche falsch. Der Fehler steckt unverändert schon im Ausgangsstand (AdminPage.tsx:12927 bzw. :16655); der Umbau hat ihn nur 1:1 mitgenommen.

**Fix-Vorschlag:** Die wahre Gruppengröße genauso mitliefern wie die wahre Position. In components/AdminPage.tsx die IIFE von `waitlistTruePos` (ab Zeile 1745) um eine zweite Map erweitern — `rank()` kennt die Länge ja bereits:

  const waitlistTrueTotal: Record<number, number> = {};
  const rank = (arr: SPRegistration[]): void => {
    const s = arr.slice().sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
    s.forEach((r, idx) => { if (typeof r.Id === 'number') { map[r.Id] = idx + 1; waitlistTrueTotal[r.Id] = s.length; } });
  };

(beide Maps aus der IIFE zurückgeben), `waitlistTrueTotal` in `waitlistTablesProps` (AdminPage.tsx:2040) und in `WaitlistTablesProps` aufnehmen und in WaitlistTables.tsx:132 verwenden:

  setWlPosModal({ reg, currentPos: truePos != null ? truePos : (i + 1), total: waitlistTrueTotal[reg.Id] != null ? waitlistTrueTotal[reg.Id] : regs.length });

Damit stimmen Nenner, `max` und Validierung auch bei aktivem Suchfilter, und die Split-Gruppen bleiben getrennt gezählt (rank() wird je Gruppe aufgerufen). WaitlistPositionModal.tsx selbst muss nicht angefasst werden.

## 18. [mittel] `URL.revokeObjectURL`-Cleanup mit `[shots]`-Dependency zerstoert die Vorschau-URLs der bereits angehaengten Screenshots

**Wo:** `components/QuestionButton.tsx:122`  
**Im Ausgangsstand:** components/QuestionButton.tsx:122 — zeichengleich, nicht durch den Umbau entstanden.

**Szenario:** Nutzer oeffnet „Hast du Fragen?", klickt „Screenshot aufnehmen" → `setShots(s => [...s, { file: f, url }])` (Zeile 205). Danach haengt er ein zweites Bild an (Zeile 205 oder `addFiles`, Zeile 222). Damit aendert sich `shots`; React fuehrt vor dem neuen Effekt die Cleanup-Funktion des VORIGEN Laufs aus — und die widerruft alle URLs des alten Arrays, also auch die des ersten Screenshots, der im neuen Array unveraendert weiterlebt (dasselbe `{file, url}`-Objekt wird per Spread uebernommen). Ergebnis: Die Miniatur des ersten Screenshots (Zeile 552, `<img src={s.url} …>`) bricht zu einem kaputten Bild, und ein Klick darauf oeffnet das ImageAnnotateModal (Zeile 712, `src={shots[annotateIdx].url}`) mit toter Quelle — Markieren ist fuer alle ausser dem zuletzt angehaengten Bild nicht mehr moeglich. Beim dritten Bild trifft es die ersten beiden usw. Der Versand selbst bleibt heil, weil `submitQuestion` die `File`-Objekte nutzt (Zeile 256) — es ist rein die Vorschau/Markierung, die kaputtgeht, und genau die ist der Grund, warum der Screenshot ueberhaupt angehaengt wird.

**Ursache:** Der Kommentar sagt „beim Unmount", die Dependency-Liste sagt „bei jeder Aenderung von `shots`". Eine Cleanup-Funktion mit nicht-leerer Dependency laeuft bei JEDEM Dependency-Wechsel, nicht nur beim Unmount. Da die Eintraege beim Anhaengen per Spread uebernommen werden, sind die widerrufenen URLs noch in Benutzung. (Zum Vergleich: `components/tickets/TicketCard.tsx:74` nutzt dasselbe Muster korrekt — dort ist `imgUrl` ein Einzelwert, der beim Wechsel wirklich ungueltig wird.)

**Fix-Vorschlag:** Den aktuellen Stand ueber ein Ref halten und nur beim Unmount abraeumen:

```tsx
const shotsRef = React.useRef(shots);
shotsRef.current = shots;
React.useEffect(() => () => {
  shotsRef.current.forEach((s) => { try { URL.revokeObjectURL(s.url); } catch { /* */ } });
}, []);
```
Das Widerrufen einzeln entfernter Bilder passiert bereits an den richtigen Stellen (`removeShot`, Zeile 230; `onAnnotateSave`, Zeile 237; `resetForm`, Zeile 135) und bleibt unangetastet.

## 19. [mittel] reloadSubEventRegs zeigt nach einer Abmeldung nicht lesbare Sub-Event-Listen als leer — ohne den v30.37-Warnhinweis

**Wo:** `components/admin/logic/createKlammerActions.ts:470`  
**Im Ausgangsstand:** components/AdminPage.tsx:1882-1891 — wortgleich; pre-existing.

**Szenario:** Organizer meldet in der konsolidierten Klammer-Matrix eine Person über das Abmelde-Modal ab (`runDeregModal`, createKlammerActions.ts, ruft am Ende `reloadSubEventRegs`). Beim Nachladen antwortet die Teilnehmerliste eines der 19 Termine mit 429 (Drosselung direkt nach den Abmelde-Schreibvorgängen) oder 403. `getAllRegistrations` wirft nicht — es liefert `[]`. `map[ch.id] = []` überschreibt den vorher korrekt geladenen Stand, und `setSubEventRegsByEventId(map)` ersetzt die GANZE Map. Die Matrix zeigt für diesen Tag ab sofort bei allen Personen „—", die Spalten-Summe 0 und die KPI-Kachel entsprechend weniger — ohne dass der v30.37-Warnbanner erscheint, denn `setDeniedSubEventLists` wird hier nicht gesetzt (der Banner in components/AdminPage.tsx:2443 hängt allein am Lade-Effect ab Zeile 400). Der Organizer sieht also genau das Bild, das v30.37 abschaffen sollte — ein volles Event, das wie ein leeres aussieht — und schließt daraus, seine Abmeldung habe zu viel gelöscht. Dieselbe Map ist außerdem Grundlage für `deletePerson` (components/admin/participants/CancelledList.tsx:161-181) und das HotelPlanningPanel.

**Ursache:** Der Lade-Effect in components/AdminPage.tsx:400-411 wurde in v30.37 auf `onHttpError` umgestellt und sammelt nicht lesbare Termine in `denied`. Der zweite Pfad, der dieselbe Map baut — `reloadSubEventRegs` nach einer Abmeldung — wurde dabei nicht mitgezogen und steht noch auf dem v30.36-Verhalten. Das `catch { map[ch.id] = []; }` ist zudem toter Code, weil `getAllRegistrations` nie wirft; der eigentliche Fehlerfall läuft am catch vorbei direkt in den `try`-Zweig mit leerem Array. CLAUDE.md-Merksatz: „Ein leeres Ergebnis ohne geprüften Status ist keine Aussage über die Daten, sondern über gar nichts."

**Fix-Vorschlag:** Denselben Rückruf nutzen wie der Lade-Effect und `deniedSubEventLists` mitpflegen — die Setter-Funktion muss dafür in den ctx von `createKlammerActions` aufgenommen werden:

```ts
const reloadSubEventRegs = async (): Promise<void> => {
  if (!selectedEvent || !selectedEvent.subEventsOnlyMode) return;
  const children = childEventsOf(selectedEvent.id);
  const map: Record<string, SPRegistration[]> = {};
  const denied: string[] = [];
  for (const ch of children) {
    const regs = await getAllRegistrations(ch.id, st => {
      if (st === 401 || st === 403 || st === 404 || st === 0) denied.push(ch.title || ch.id);
    });
    map[ch.id] = regs;
  }
  setSubEventRegsByEventId(map);
  setDeniedSubEventLists(denied);
};
```

Einfachere Alternative mit identischer Wirkung: den Body durch `setSubRegReloadTick(t => t + 1)` ersetzen und das Nachladen komplett dem bereits abgesicherten Effect überlassen — dann gibt es nur noch EINEN Pfad, der diese Map baut.

## 20. [mittel] „Freie Plätze mit Warteliste füllen" überbucht bei gemeinsamer Warteliste eine einzelne Gruppe

**Wo:** `components/admin/logic/useWaitlistActions.ts:288`  
**Im Ausgangsstand:** components/AdminPage.tsx:2308-2322 (groups/cap) und 2365-2376 (Promote-Schleife) — der Rumpf ist beim Umbau zeichengleich nach useWaitlistActions.ts gewandert.

**Szenario:** Split-Event mit `splitSharedWaitlist = true`. Gruppe A (Durchstarter) cap 30, 30 aktiv (voll). Gruppe B (Funstarter) cap 30, 25 aktiv. Auf der gemeinsamen Warteliste stehen 8 Personen, die ersten 6 davon mit `PreferredStarterType = 'Durchstarter'`. Der Organizer klickt „Freie Plätze mit Warteliste füllen": `cap = 60`, `active = 55`, `free = 5`, `count = 5`. Fünfmal wird `promoteFirstWaitlistItem(sub, undefined, undefined, undefined)` gerufen — ohne Typfilter, also die ersten fünf nach TeilnehmerID. Jede bekommt `StarterType = firstWaiting.PreferredStarterType = 'Durchstarter'` (waitlist.ts:112-115). Ergebnis: Gruppe A steht auf 35/30, Gruppe B unverändert auf 25/30. Die Anmeldeseite zeigt für A „ausgebucht", die Überbuchungs-Box im Organizer Center flaggt fünf Personen als „über Kapazität", und `detectOverbooking` (overbooking.ts:54) markiert sie mit `OverbookReview='Pending'` — durch eine Aktion, die der Organizer gerade selbst ausgelöst hat. Die Gruppengrenze wird bei der regulären Anmeldung sehr wohl erzwungen (EventContext.tsx:731-734 reserviert immer gegen die EINZELNE Gruppenkapazität, unabhängig von `splitSharedWaitlist`) — nur dieser Nachrück-Pfad kennt sie nicht.

**Ursache:** Bei `splitSharedWaitlist` wird die Warteliste zu einem Topf zusammengefasst und als Obergrenze nur die SUMME beider Kapazitäten geprüft. Die Gruppenzuordnung der nachrückenden Person bleibt aber erhalten (`StarterType` wird aus ihrem `PreferredStarterType` gesetzt) und die Gruppenkapazitäten gelten weiter — an jeder anderen Stelle (`reserveSeat`, `detectOverbooking`, Gruppen-Karten der Anmeldeseite). Gemeinsame Warteliste heißt „gemeinsame Reihenfolge", nicht „gemeinsamer Kapazitätstopf".

**Fix-Vorschlag:** Im `!perGroup`-Zweig bei `isSplitCapacity` nicht über die Summe rechnen, sondern die freien Plätze je Gruppe ermitteln und die Reihenfolge aus der gemeinsamen Warteliste nehmen: freie Plätze `freeA = capA - activeA`, `freeB = capB - activeB` bestimmen und je Nachrück-Schritt `promoteFirstWaitlistItem(sub, undefined, undefined, <Gruppe, die noch frei ist>)` aufrufen — d.h. den Typfilter auch bei gemeinsamer Warteliste setzen, solange nur eine der beiden Gruppen noch Platz hat. Alternativ (falls fachlich gewünscht) müsste `reserveSeat` bei `splitSharedWaitlist` ebenfalls gegen die Summe prüfen — dann aber konsistent an allen Stellen, inklusive der Gruppen-Karten der Anmeldeseite.

## 21. [niedrig] StayRangePicker: automatische Abreise per +86400000 ms erzeugt in der Nacht der Zeitumstellung einen 0-Nächte-Zeitraum

**Wo:** `components/StayRangePicker.tsx:109`  
**Im Ausgangsstand:** Ausgangsstand: components/StayRangePicker.tsx:109 — die Datei ist zwischen beiden Ständen byte-identisch.

**Szenario:** Nachgerechnet mit TZ=Europe/Berlin: Ein Teilnehmer öffnet auf der Anmeldeseite (bzw. der Organizer in der Hotelplanung) den Hotel-Zeitraum und wählt als Anreise den 25.10.2026 — die Nacht, in der die Uhr von 3 auf 2 zurückgestellt wird. `dayToDate('2026-10-25')` = 25.10. 00:00 MESZ; `+ 86400000 ms` landet auf 25.10. **23:00** MEZ, also demselben Kalendertag. `toLocalDay(cand)` liefert „2026-10-25", `to === from`, `stayNights` = 0. Das Feld zeigt „0 Nächte", und gespeichert wird der Wert „2026-10-25 – 2026-10-25" — genau der Zustand, den der Kommentar zwei Zeilen darüber verhindern soll. Für jeden anderen Anreisetag (auch für die Frühjahrs-Umstellung 29.03.) funktioniert es korrekt.

**Ursache:** Die Tages-Arithmetik läuft über eine feste Millisekunden-Konstante auf einem Date, das lokale Mitternacht darstellt. In der Nacht der Rückstellung hat der Kalendertag 25 Stunden, +24 h bleibt deshalb im selben Tag.

**Fix-Vorschlag:** Den Folgetag über die Kalender-Komponenten bilden statt über Millisekunden: `const cand = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);` — `Date` normalisiert den Überlauf und der Konstruktor löst DST korrekt auf. `stayNights`/`nightsBetween` selbst bleiben unverändert richtig, weil sie über `T00:00:00Z` verankert rechnen. Dieselbe Konstante steht auch in HotelPlanningPanel.tsx:86 und HotelSetupWizard.tsx:87 (`addDays`), dort aber auf `T00:00:00Z` verankert und damit unkritisch.

## 22. [niedrig] Unmount-Cleanup in DocumentsViewer haelt `blobUrl` aus dem ERSTEN Render fest ('') und gibt deshalb nie etwas frei

**Wo:** `components/myEvents/DocumentsViewer.tsx:109`  
**Im Ausgangsstand:** components/MyEventsPage.tsx:634-637 — zeichengleich uebernommen, nicht durch den Umbau entstanden.

**Szenario:** Nutzer oeffnet in „Meine Events" ein PDF-Dokument. `toggleDoc` laedt die Datei per XHR und setzt `setBlobUrl(URL.createObjectURL(correctBlob))` (Zeile 100). Er wechselt die Seite (Navigation, Tab-Wechsel) statt das Dokument per zweitem Klick zuzuklappen — die Komponente wird unmountet. Die Cleanup-Funktion wurde beim ersten Render mit `[]` erzeugt und hat `blobUrl` als `''` eingeschlossen; `if (blobUrl)` ist damit immer falsch, `URL.revokeObjectURL` wird nie aufgerufen. Der Blob (bei Handbuechern/Programmen schnell mehrere MB) bleibt fuer die gesamte Lebensdauer der Seite im Speicher. Wer in einer Sitzung mehrere Dokumente ueber mehrere Events oeffnet und dazwischen navigiert, sammelt sie alle an — auf Mobilgeraeten der Weg in den Tab-Neustart. Der explizite Weg (Dokument zuklappen, Zeile 48) raeumt korrekt auf; nur der Unmount-Weg ist wirkungslos, obwohl der Kommentar genau das verspricht.

**Ursache:** Klassischer stale-closure: leere Dependency-Liste + Zugriff auf einen State-Wert in der Cleanup-Funktion. Die Funktion wird genau einmal erzeugt und sieht den Initialwert des States, nicht den zum Unmount-Zeitpunkt aktuellen.

**Fix-Vorschlag:** Den aktuellen Wert ueber ein Ref mitfuehren, das die Cleanup-Funktion zur Unmount-Zeit ausliest:

```tsx
const blobUrlRef = React.useRef('');
React.useEffect(() => { blobUrlRef.current = blobUrl; }, [blobUrl]);
React.useEffect(() => () => {
  if (blobUrlRef.current) URL.revokeObjectURL(blobUrlRef.current);
}, []);
```
(Die Dependency schlicht auf `[blobUrl]` zu setzen waere falsch — dann wuerde beim Wechsel des Dokuments die alte URL zwar korrekt widerrufen, aber die Absicht „beim Unmount" bliebe unerfuellt, sobald der letzte Wert gesetzt wird.)

