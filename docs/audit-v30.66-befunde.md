# Audit v30.66 — vollstaendige Befundliste

Zwei unabhaengige Audit-Laeufe ueber die DEX-Codebasis:

| Lauf | Pruefer | Funde | widerlegt |
|---|---:|---:|---:|
| A — gesamter Code, 17 Dimensionen | 334 | 95 | 10 |
| B — nach der Modularisierung, Umbau + Logik | 103 | 29 | 0 |

Jeder Fund wurde von drei unabhaengigen Skeptikern gegengelesen (Linsen: Code-Lesart,
faengt eine andere Stelle den Fall ab, ist der Pfad ueberhaupt erreichbar); ein Fund gilt nur
als bestaetigt, wenn hoechstens einer ihn widerlegen konnte.

**Zentrales Ergebnis von Lauf B: kein einziger Fund stammt aus der Modularisierung.**
Alle Befunde sind gegen den unveraenderten Ausgangsbaum geprueft und dort nachweisbar.

Nach Zusammenfuehrung beider Laeufe (Dubletten ueber Datei+Titel entfernt): **124 Befunde**.

## In v30.66 behoben

- [kritisch] Festes „Anmeldung ab"-Datum geht beim Öffnen des Wizards still verloren und wird beim nächsten Speichern gelöscht (isoToLocal vor seiner Deklaration aufgerufen) — `components/EventCreationPage.tsx`
- [kritisch] Lösch-Sperre countExternalRegistrations öffnet sich bei jedem Lesefehler — Cascade-Delete ohne Aufbewahrungsschutz — `context/EventContext.tsx`
- [kritisch] Sammelmail nach Termin-Änderung: wrapTemplate-Argumente vertauscht — der Mail-Text ist wörtlich „#0076a8" — `context/actions/billing.ts`
- [kritisch] archiveEventStats schreibt 0-KPIs, wenn die Teilnehmerliste nicht lesbar ist — und danach wird sie unwiderruflich gelöscht — `services/events/eventStats.ts`
- [hoch] Sammelmail „Anmeldung aktualisiert": wrapTemplate-Argumente vertauscht — Mailtext landet in der Überschrift, Body enthält nur „#0076a8" — `context/actions/billing.ts`
- [hoch] `{{NewLeadBlock}}` wird HTML-escaped — Team-Lead-Hinweis erscheint als roher Markup-Text — `context/actions/cancellation.ts`
- [hoch] `{{WaitlistPositionBlock}}` wird HTML-escaped — in der Überbuchungs-Entschuldigung stehen die Tags als Text — `services/events/overbooking.ts`
- [hoch] archiveEventStats schreibt lauter Nullen ins Statistik-Archiv, wenn die Teilnehmerliste nicht lesbar ist — `services/events/eventStats.ts`
- [hoch] reorderParticipantIDs renummeriert eine teilweise gelesene Liste und meldet „0 Fehler" — `services/events/seats.ts`
- [hoch] „Aktiv ab" (ActiveFrom) ist das einzige Datumsfeld, das die Berlin-Konvertierung umgeht — der Freischalt-Zeitpunkt hängt an der Browser-Zeitzone — `components/wizard/logic/wizardSubmit.ts`

## Offen

Bewusst nicht in diesem Release behoben: Ein Update, das eine Modularisierung dieser
Groesse mit dutzenden Verhaltensaenderungen mischt, ist im Fehlerfall nicht mehr
aufzutrennen. Reihenfolge = Schwere.

### 1. [kritisch] Check-in: fehlgeschlagener Listen-Abruf wird als „keine Teilnehmer" gecacht und nie wiederholt

**Wo:** `components/CheckInPage.tsx:172` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Helfer am Einlass öffnet Check-in und wählt das Event; der Effekt Z.182-186 ruft `loadRegsForSearch`. Der Lese-Request scheitert (403, weil die Rechte auf der Sub-Event-Subsite noch nicht propagiert sind, oder 429 durch Throttling, wenn mehrere Helfer gleichzeitig laden). `getAllRegistrations` wirft dabei NICHT — `services/events/registrationEdit.ts:670` macht `if (!response.ok) { onHttpError?.(status); break; }` und liefert die bis dahin gelesenen (also null) Zeilen. `[]` landet als gültiges Ergebnis im Cache. Folge: der KPI-Block rendert gar nicht (Z.1299 prüft `.length > 0`), die Liste meldet „Keine Teilnehmer für dieses Event.", `searchLoadError` bleibt leer, und `checkInByParticipantId` (Z.279) sieht `regs` als truthy an und antwortet „Keine Anmeldung mit der Teilnehmer-ID X bei diesem Event" — genau die Behauptung, die der Kommentar darüber ausdrücklich verhindern will. Weil der Cache-Key jetzt gesetzt ist, blockiert `if (searchRegsCache[eventId]) return;` (Z.168) jeden weiteren Versuch; nur ein kompletter Seiten-Reload hilft.

**Ursache:** Das `catch` fängt nur geworfene Fehler. `getAllRegistrations` meldet HTTP-Fehler ausschließlich über den optionalen `onHttpError`-Rückruf, der hier nicht übergeben wird — die in CLAUDE.md zweimal beschriebene Falle (v29.0/29.1, v30.37) in einem dritten Pfad. `components/admin/ShirtSizeModal.tsx:57` macht es im selben Repo korrekt vor.

**Fix-Vorschlag:** Status prüfen und den Cache-Eintrag nur bei geprüftem Erfolg schreiben: `let ok = true; const regs = await getAllRegistrations(eventId, () => { ok = false; }); if (!ok) { setSearchLoadError('Teilnehmerliste konnte nicht gelesen werden (keine Berechtigung oder Drosselung) — bitte erneut versuchen.'); setIsLoadingSearchRegs(false); return; } setSearchRegsCache(prev => ({ ...prev, [eventId]: regs }));`. Ohne Cache-Eintrag versucht der Effekt es beim nächsten Rendern von selbst wieder.

### 2. [kritisch] F&A-Teilnehmerliste: leeres Leseergebnis wird als „0 Teilnehmer" gemeldet und dauerhaft gestempelt

**Wo:** `context/actions/billing.ts:91` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Admin (nicht Organizer dieses Events, also ohne Leserecht auf der Event-Subsite) bootet die App; 11 s später läuft `maybeSendBillingAutoMails` und ruft für ein Event mit 120 Anmeldungen `sendFAMail(ev,'list',{auto:true})`. `getAllRegistrations` bekommt 403, bricht ab und liefert `[]` — ohne dass jemand es merkt, weil KEIN `onHttpError` übergeben wird. Ergebnis: F&A erhält die offizielle Mail „anbei die Teilnehmerliste … **0** Personen", `listSentAt`/`autoListSentAt` werden gesetzt und `listSnapshot: []` gespeichert. Der Status springt auf „Teilnehmerliste versendet — Abschluss durch F&A offen", der Auto-Versand wiederholt sich nie (`!b.autoListSentAt && !b.listSentAt` ist ab jetzt falsch), und das F&A Center zeigt gleichzeitig den Widerspruch „Teilnehmerliste versendet" (Status-Pill) und „Es wurde bisher keine Teilnehmerliste an F&A versendet" (Bereich 3, weil `listSnapshot.length > 0` scheitert). Dasselbe passiert bei 429-Drosselung oder wenn `subsiteMap` den Eintrag nicht kennt (dort feuert sogar explizit `onHttpError(0)` — nur hört niemand zu).

**Ursache:** Genau die in CLAUDE.md dokumentierte Falle: `getAllRegistrations` wirft nicht, sondern bricht die Schleife ab und gibt die bis dahin gelesenen Zeilen zurück (services/events/registrationEdit.ts:648 ff.). Seit v29.3 gibt es dafür den `onHttpError`-Rückruf, und `EventContext.getAllRegistrations` meldet auch die fehlende Subsite darüber — `sendFAMail` ist die einzige Versandstelle, die ihn nicht nutzt. Der Schwesterpfad im selben Feature macht es richtig: `BillingActionPanel.downloadXlsx` (Zeile 125) bricht bei `rows.length === 0` mit einer Meldung ab; der Versandpfad hat diesen Guard nicht. Zweite, unabhängige Quelle für ein zu kurzes Ergebnis: gelesen wird nur die Liste des Hauptevents. Bei Klammer-Events hängen die echten Anmeldungen an den Sub-Event-Listen; die Klammer-Schattenzeile schreibt `registerForEvent` erst seit v30.42 zuverlässig, Altbestand aus AddParticipantsModal/MyEventsPage/AssistantPage hat sie nicht.

**Fix-Vorschlag:** `onHttpError` durchreichen und bei Fehler ODER leerem Ergebnis NICHT versenden: `let httpErr = 0; const regs = await getAllRegistrations(ev.id, s => { httpErr = s || -1; }); if (httpErr) return { ok:false, reason:'read-failed' }; const rows = faRowsFromRegistrations(regs); if (rows.length === 0) return { ok:false, reason:'no-participants' };` — beide Gründe in `BillingActionPanel.doSend` als eigene Meldung ausgeben („Teilnehmerliste konnte nicht gelesen werden — Versand abgebrochen"). Im Auto-Pfad denselben Fall überspringen statt zu stempeln, damit der nächste Lauf es erneut versucht. Zusätzlich für Klammer-Events die Zeilen der Sub-Events mit einbeziehen (per E-Mail dedupliziert), sonst meldet die Datei weniger Personen als die Teilnehmerliste im Organizer Center zeigt.

### 3. [kritisch] Recreate-Pfad prüft deleteEventItemOnly nicht — bei Fehlschlag löscht der Aufräum-Lauf die geteilte Subsite mit allen Anmeldungen

**Wo:** `components/wizard/logic/persistSubEvents.ts:372` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer klickt bei einem Klammer-Event mit 19 Terminen „Fehlende Termine jetzt anlegen" (forceOutlookRecreateRef) oder schaltet an einem bestehenden Sub-Event mit Anmeldungen „Outlook-Termin erstellen" nachträglich ein. Ab ca. dem 20. Schreibvorgang drosselt SharePoint (HTTP 429 — in dieser Codebasis dokumentiert, s. v29.74). deleteEventItemOnly liefert dann false. Der Code läuft trotzdem weiter: createEvent legt eine ZWEITE DEX_Events-Zeile an, die auf DIESELBE Subsite zeigt; die alte Zeile existiert weiter. Weil draft.dbId nicht in keptDbIds landet, ruft die Aufräum-Schleife am Ende (Zeile 583) deleteEvent(alteId) — und das ist der KASKADIERENDE Löschpfad: er recycelt event.SubsiteUrl (die geteilte Subsite mit allen Teilnehmern), queued einen Outlook-DeleteEvent und räumt die EventNumber aus DEX_Participants. Ergebnis: alle Anmeldungen des Termins weg (93 Tage Papierkorb), der frisch angelegte Termin zeigt auf eine recycelte Subsite, Meldung an den Organizer: „Änderungen gespeichert".

**Ursache:** deleteEventItemOnly wirft nie — es liefert bei HTTP-Fehler `false` (services/EventService.ts:3384: `const response = await this._delete(...); return response.ok;`). Das try/catch hier ist toter Code, der Rückgabewert wird verworfen. Genau diese Lücke wurde in v29.21 für den ERSTEN Recreate-Pfad (Zeile 310, `const itemDeleted = await deleteEventItemOnly(...)`) geschlossen; der zweite, ältere Pfad (DisableOutlook-Toggle / v28.69 „Fehlende Termine") wurde dabei übersehen.

**Fix-Vorschlag:** Denselben Riegel wie in Zeile 310 einziehen: `const itemDeleted = await deleteEventItemOnly(draft.dbId).catch(() => false); if (!itemDeleted) { showAlert(...); keptDbIds.add(draft.dbId); }` und nur im Erfolgsfall createEvent+continue ausführen. Zusätzlich absichern, dass eine Id, deren Item nachweislich noch existiert, NIE in den kaskadierenden deleteEvent-Zweig fällt.

### 4. [kritisch] Rollen-Entzug lässt Full Control auf dem gesamten Web stehen

**Wo:** `context/RoleContext.tsx:264` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Admin setzt in der Rollenverwaltung einen Organizer (oder Admin) auf „User" zurück, weil die Person das Team verlassen hat — oder löscht die Zeile ganz (`removeRole`, Zeile 281-292). Die App meldet Erfolg, die Person verschwindet aus DEX_Roles und sieht in der App keine Organizer-Funktionen mehr. In SharePoint behält sie aber weiterhin **Full Control auf dem Web-Root** der Site-Collection: sie kann jede Teilnehmer-Subsite und damit alle personenbezogenen Anmeldedaten lesen und ändern, und weil Full Control „Manage Permissions" einschließt, kann sie sich die gerade entzogenen Rechte auf DEX_Roles/DEX_Events selbst zurückgeben. Zweite Variante desselben Fehlers: Beim Downgrade Admin → Organizer (Zeile 261-263) wird nur `grantReadOnRolesList` gerufen; `addroleassignment` ist ADDITIV, die alte Full-Control-Bindung auf DEX_Roles bleibt daneben bestehen — der Ex-Admin kann die Rollenliste weiter bearbeiten und sich selbst wieder hochstufen.

**Ursache:** `grantOrganizerPermissions` (services/SharePointService.ts:376-395) vergibt zwei Rechte: Full Control auf DEX_Events UND Full Control auf `${siteUrl}/_api/web/roleassignments`. Für das zweite gibt es im gesamten Code kein Gegenstück — `grep 'roleassignments/getbyprincipalid'` findet nur die beiden Listen-Revokes in SharePointService.ts:431/454 sowie den generischen Helfer im Berechtigungs-Audit. Die Entzugs-Pfade sind also nicht spiegelbildlich zu den Vergabe-Pfaden aufgebaut. Zusätzlich unterstellt der Downgrade-Pfad, `addroleassignment` würde bestehende Bindungen ersetzen — SharePoint fügt sie aber zur bestehenden Rollenzuweisung hinzu.

**Fix-Vorschlag:** In SharePointService eine `revokeSiteAccess(userEmail)` ergänzen, die analog zu `revokeAccessOnEventsList` ein `X-HTTP-Method: DELETE` auf `${siteUrl}/_api/web/roleassignments/getbyprincipalid(${userId})` schickt, und sie in `updateRole` (newRole === 'User') und in `removeRole` aufrufen. Für das Downgrade Admin → Organizer/F&A zusätzlich vor `grantReadOnRolesList` die bestehende Zuweisung auf DEX_Roles löschen (`revokeAccessOnRolesList`), damit Read wirklich Read ist und nicht Read + Full Control. Der Erfolg jedes Revoke gehört geprüft und bei Fehlschlag sichtbar gemeldet — ein still fehlgeschlagener Entzug ist derselbe Zustand wie gar kein Entzug.

### 5. [kritisch] deleteParticipantData meldet Erfolg, obwohl die Teilnehmer-Subsite nicht gelöscht wurde

**Wo:** `services/events/eventStats.ts:264` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Admin startet im Admin Hub „Teilnehmerlisten löschen" (3-Monats-Löschkonzept). Für Event 4711 antwortet der recycle-POST mit 403 (Rechte auf der Subsite fehlen) oder 429 (Drosselung beim Lauf über viele Events). Die App zeigt „1 gelöscht, 0 fehlgeschlagen". Tatsächlich steht die Subsite mit allen personenbezogenen Anmeldedaten weiter da — dauerhaft, denn die DEX_EventStats-Zeile wurde geschrieben und `getParticipantDeletionDue` filtert das Event über `archived.has(e.eventNumber)` künftig aus. Es gibt keinen zweiten Versuch und keine Meldung.

**Ursache:** `svc._post` wirft bei HTTP 4xx/5xx NICHT (siehe den eigenen Kommentar in emailTemplatesList.ts: „v18.66: _post wirft NICHT bei HTTP 4xx/5xx") — es liefert die Response zurück. Der `catch` greift also nur bei Netzwerk-Exceptions, `response.ok` wird nie geprüft, und die Funktion läuft in jedem Fall auf `return true`. Der Aufrufer (EventContext.runParticipantDeletion, Zeile ~2625) verlässt sich laut dem Kommentar direkt darüber ausdrücklich auf `false` als „später erneut versuchen" und rollt nur dann die Statistik-Zeile per `deleteEventStatsRow` zurück.

**Fix-Vorschlag:** Rückgabewert prüfen und den Fehlschlag durchreichen: `const rec = await svc._post(`${event.SubsiteUrl}/_api/web/recycle`, {}); if (!rec.ok) { console.warn(…); return false; }` (Exception ebenfalls zu `return false` machen). Analog für den `RegistrationListName`-Recycle darunter. Nur wenn beide Löschschritte bestätigt sind, darf `true` zurückkommen.

### 6. [kritisch] removeParticipantEvent meldet Erfolg, ohne die MERGE-Antwort zu prüfen — der v29.3-Abbruchschutz vor der Löschung kann nie greifen

**Wo:** `services/EventService.ts:1242` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Admin startet das 3-Monats-Löschkonzept (runParticipantDeletion → deleteParticipantData, jetzt services/events/eventStats.ts:243-262). Für ein Event mit 200 Teilnehmern laufen 200 sequentielle GET+MERGE auf DEX_Participants. SharePoint drosselt ab Person ~80 mit HTTP 429 (withThrottleRetry gibt nach 2 Versuchen bzw. bei offenem Schutzschalter die 429-Antwort ZURÜCK, es wird nicht geworfen). removeParticipantEvent bekommt die 429-Response, ignoriert sie und liefert `true`. In deleteParticipantData bleibt `registryFailed === 0`, der Abbruch greift nicht, die Subsite wird recycelt. Ergebnis: 120 Personen tragen die EventNumber weiter im Register, die Teilnehmerliste ist weg — genau die „Verweis ohne Zeile"-Rückstände, die analyzeRegistryAgainstLists später meldet und die sich nicht mehr nachrechnen lassen.

**Ursache:** `_merge` wirft bei HTTP-Fehlern nicht, sondern gibt die SPHttpClientResponse zurück (EventService.ts:7486ff, 406 wird bewusst als Erfolg gemappt). `removeParticipantEvent` awaitet den Aufruf, verwirft aber `response.ok` und gibt unbedingt `true` zurück. Nur ein geworfener Netzwerkfehler oder ein fehlender Registereintrag liefern `false`. Die v29.3-Umstellung (erst Register, sequentiell, mit Fehlerzähler, bei Fehlern gar nicht löschen) hat die Reihenfolge korrigiert, verlässt sich aber auf einen Rückgabewert, der Fehler nicht transportieren kann.

**Fix-Vorschlag:** Rückgabewert an die Antwort koppeln: `const resp = await this._merge(...); return !!resp.ok;` (406 ist in `_merge` bereits auf ok=true normalisiert). Damit zählt deleteParticipantData den Fehlschlag und bricht VOR dem recycle ab. Zusätzlich ein `console.warn` mit `resp.status` und der E-Mail, damit der Admin sieht, welche Register-Zeilen offen blieben.

### 7. [kritisch] „Unvollständige Anmeldungen"-Box löscht echte Anmeldungen, wenn eine Termin-Liste nicht lesbar war

**Wo:** `components/AdminPage.tsx:5538` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Klammer-Event (subEventsOnlyMode) mit 19 Terminen. Eine Co-Organizerin wurde nachträglich benannt und hat deshalb nur auf der Klammer Rechte, nicht auf 5 der 19 Sub-Event-Subsites (genau der in CLAUDE.md/v30.37 dokumentierte Fall). Der Lade-Effekt (Zeile 668-678) schreibt für diese 5 Termine `map[ch.id] = []` und meldet sie in `deniedSubEventLists`. `consolidatedRows` ist durch die 14 lesbaren Termine NICHT leer, der frühe Return in `renderConsolidatedView` (Zeile 5422) greift also nicht. Jede Person, die nur einen der 5 gesperrten Termine gebucht hat, fehlt in `anySubEmails`, hat aber (seit v30.42 garantiert) eine aktive Klammer-Zeile → sie landet in `orphanShadowRegs` und erscheint in der roten Box „Unvollständige Anmeldungen (N)" mit dem Knopf „Rest-Anmeldung entfernen". Der Knopf ruft `performSilentDuplicateDelete` → `deleteRegistration` und löscht die Klammer-Zeile inklusive aller Hauptevent-Antworten HART (kein „Abgemeldet"-Status, keine Mail, kein Nachrücken). Der Begleittext („die Person kann sich danach neu anmelden") bestätigt den Organizer noch darin, dass er Datenmüll wegräumt.

**Ursache:** Die Geist-Erkennung liest ein leeres Ergebnis als Aussage über die Daten („keine Sub-Event-Zeile") statt als Aussage über die Lesbarkeit. Der dafür seit v30.37 vorhandene Zustand `deniedSubEventLists` wird an dieser Stelle nicht ausgewertet — er steuert ausschließlich das Banner in Zeile 11611. Exakt der Merksatz aus CLAUDE.md: „Ein leeres Ergebnis ohne geprüften Status ist keine Aussage über die Daten, sondern über gar nichts."

**Fix-Vorschlag:** Die Box an die Lesbarkeit koppeln: `if (deniedSubEventLists.length > 0) return null;` bzw. — falls sie sichtbar bleiben soll — mindestens den „Rest-Anmeldung entfernen"-Knopf deaktivieren und im Kasten benennen, dass N Terminlisten nicht gelesen werden konnten und die Liste deshalb nicht belastbar ist. Sauberer: `orphanShadowRegs` nur über die Termine bilden, deren Liste nachweislich gelesen wurde (Menge der erfolgreich geladenen `ch.id` mitführen) und die Box unterdrücken, sobald diese Menge nicht alle `consolidatedChildren` abdeckt.

### 8. [hoch] Anmeldeseite zählt eine ABGEMELDETE Klammer-Zeile als vorhanden — die Klammer wird nicht reaktiviert, die Person fällt aus der Hotelplanung heraus

**Wo:** `components/registration/submitFlow.ts:795` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Klammer-Event (subEventsOnlyMode) mit Terminen und Hotelfrage auf Hauptevent-Ebene. Eine Person hat vorher abgesagt bzw. der Organizer hat im Abmelde-Dialog gezielt nur die Klammer-Zeile abgemeldet (createKlammerActions.ts:485-499 stellt die Klammer als eigene, abwählbare Zeile in den Dialog) — ihre Klammer-Zeile steht auf 'Abgemeldet'. Jetzt bucht sie über die Anmeldeseite einen Termin und füllt dabei die übergreifenden Hauptevent-Felder aus (Hotel, Verpflegung, Anreise). `myParentReg` ist die abgemeldete Zeile, also ist `parentAlreadyHasRow` true → `shouldShadowRegisterParent` false → der Schritt-3-Block (Zeile 1106) läuft nicht, und die Sub-Event-Anmeldungen laufen mit `skipShadowParent: true` (Zeile 1054), also greift auch das zentrale Netz in EventContext.tsx:1314 nicht, das die Zeile über den Reaktivierungs-Pfad wieder scharf gestellt hätte. Stattdessen schreibt `updateMyRegistration` die frischen Antworten in die ABGEMELDETE Zeile und lässt den Status stehen. Ergebnis: Der Termin ist gebucht, die Klammer-Zeile bleibt 'Abgemeldet'. Konkrete Folge in der Hotelplanung: HotelPlanningPanel bekommt die Klammer-`registrations` (admin/sections/HotelPlanningSection.tsx:61) und filtert `people` auf ACTIVE_STATI (HotelPlanningPanel.tsx:575-580) — die Person ist im Hotel-Roster gar nicht vorhanden, obwohl sie „Yes, I need accommodation“ beantwortet hat und an dem Termin teilnimmt. `autoDistribute` kann sie nicht einmal als `skippedNoWish` melden. Zusätzlich taucht sie in der „Abmeldungen“-Liste des Gesamt-Events auf (AdminPage.tsx:1519-1524), während die Matrix sie als Teilnehmerin führt.

**Ursache:** Zwei Ableitungen aus demselben `myParentReg`, aber mit unterschiedlicher Semantik: `parentAlreadyRegistered` (RegistrationPage.tsx:1290) prüft den Status, `parentAlreadyHasRow` (submitFlow.ts:795) prüft nur die Existenz des Objekts. Für den Zweck „muss die Klammer-Zeile geschrieben/reaktiviert werden?“ ist die Existenz die falsche Frage — eine abgemeldete Zeile ist keine Klammer-Anmeldung. Verstärkt wird es durch `skipShadowParent: true`: die Anmeldeseite schaltet das zentrale Netz aus und muss den Fall deshalb selbst abdecken.

**Fix-Vorschlag:** `parentAlreadyHasRow` auf den aktiven Stand einschränken: `const parentAlreadyHasRow = !!(myParentReg && myParentReg.Status && myParentReg.Status !== 'Abgemeldet');` — dann ist es identisch zu `parentAlreadyRegistered` und `shouldShadowRegisterParent` wird true. Der Schritt-3-Block ruft `doParentRegistration(true)` → registerForEvent findet über getMyRegistration die abgemeldete Zeile und geht in den Reaktivierungs-Pfad (EventContext.tsx:804-806), der Status UND CustomData korrekt setzt. Der `else if`-Zweig auf Zeile 967 (nur updateMyRegistration) bleibt dann dem Fall vorbehalten, für den er gedacht war: eine BESTEHENDE aktive Schatten-Zeile beim Nachbuchen eines weiteren Termins.

### 9. [hoch] Auto-versendete QR-Mail enthält keine Teilnehmer-ID

**Wo:** `context/EventContext.tsx:1181` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer startet den QR-Massenversand (setzt AutoSendQRCode=true). Danach meldet sich ein Nachzügler an. Er bekommt automatisch die QR-Mail — aber ohne die große ID-Zeile unter dem Code und ohne den Hinweis „Falls der Scan nicht klappt: einfach diese Nummer am Einlass nennen.“ Am Einlass scheitert (auf Android/SharePoint-App-WebView) der Kamera-Scan, und die Person kann die Ersatz-Nummer nicht nennen, weil sie in ihrer Mail nicht steht. Beim manuellen Massenversand (createQrMailActions.tsx:328, `reg.TeilnehmerID`) und beim Test (SAMPLE_QR_ID) steht sie drin — nur genau bei den Spätanmeldern nicht.

**Ursache:** Der 7. Parameter `teilnehmerId` von `qrCodeEmail` wird hier hart als `undefined` übergeben. In `buildQrBlockHtml` (services/EmailTemplates.ts:953) gilt dann `hasId = false`, wodurch der komplette ID-Block plus Hinweiszeile weggelassen wird. Die ID wäre verfügbar: derselbe Block lädt wenige Zeilen später `getMyRegistration`, dessen `SPRegistration.TeilnehmerID` an allen anderen Aufrufstellen genutzt wird — der Aufruf steht nur hinter statt vor dem Mailaufbau.

**Fix-Vorschlag:** Die Registrierung VOR dem Mailaufbau laden und die ID durchreichen: `const myReg = event.subsiteUrl ? await eventService.getMyRegistration(event.subsiteUrl, emailToUse) : null;` dann `qrCodeEmail(..., qrOverride, myReg?.TeilnehmerID, qrHeroPhoto)` und `myReg` unten für `setQRSentStatus` wiederverwenden (spart zusätzlich einen Request).

### 10. [hoch] Bei freigegebener Hotel-Anzeige fragt jede Massenzuordnung EINMAL PRO PERSON nach — Abbrüche werden als Erfolg gemeldet

**Wo:** `components/HotelPlanningPanel.tsx:797` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** „Teilnehmer sehen ihr Hotel" ist eingeschaltet (visible = true). Der Organizer klickt „Automatisch verteilen", bestätigt den Verteilplan für 80 Personen — und bekommt danach 80 weitere Modals „Die Hotel-Anzeige ist für dieses Event freigegeben … Fortfahren?". Klickt er bei Person 17 auf „Abbrechen", kehrt `writeAssignment` sofort zurück, ohne zu schreiben; `done++` in Zeile 414 zählt trotzdem weiter und die Schlussmeldung in Zeile 420 sagt „80 Person(en) verteilt". Die Person 17 bleibt ohne Hotel, ohne dass irgendwo ein Fehler steht. Gleiches Muster in `assignWholeSub` (Zeile 284), `assignSelectedTo` (913) und `applyStayToSelected` (925).

**Ursache:** `writeAssignment` ist gleichzeitig der Einzel-Handler der Tabellenzelle UND das Schreib-Primitiv aller Massenläufe. Die Sichtbarkeits-Rückfrage steckt im Primitiv, hat kein „schon bestätigt"-Flag, und die Funktion liefert `Promise<void>` — die Schleifen können einen Abbruch gar nicht bemerken. `applyWizard` (Zeile 473 ff.) umgeht das bewusst und ruft `svc.setHotelAssignment` direkt; die vier anderen Massenpfade nicht.

**Fix-Vorschlag:** `writeAssignment(rows, hotel, from, to, opts?: { skipConfirm?: boolean }): Promise<boolean>` — Rückfrage nur wenn `!opts.skipConfirm`, Rückgabe `false` bei Abbruch/Fehler. In autoDistribute/assignWholeSub/assignSelectedTo/applyStayToSelected einmal vorab fragen (die Verteil-Rückfrage gibt es dort ohnehin schon) und mit `skipConfirm: true` aufrufen; `done` nur bei `true` hochzählen und die Schlussmeldung aus den tatsächlich geschriebenen Zeilen bilden.

### 11. [hoch] Belegungszahl der Sub-Events in „Meine Events" kommt aus der item-level-gefilterten Liste — zeigt 0/20 bei vollem Termin

**Wo:** `components/myEvents/MyEventSubEvents.tsx:128` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Klammer-Event mit Tages-Terminen à 20 Plätzen, Tag 3 ist mit 20 Personen voll. Ein Teilnehmer öffnet „Meine Events" → Sub-Event-Liste/Kalender. `getAllRegistrations(ce.id)` liefert unter ReadSecurity=2 nur die eigene(n) Zeile(n) → `counts[ce.id]` = 0 (nicht angemeldet) bzw. 1. Gerendert wird `{count}/{ce.maxParticipants}` (Zeile 671) also „0/20"; `isFull = count >= max` ist false, deshalb fehlt das rote „(voll)" und der Anmelde-Button bleibt aktiv (`disabled` Zeile 635 wertet `isFull` aus). Der Teilnehmer klickt an, `reserveSeat` setzt ihn serverseitig auf die Warteliste — nach einer Anzeige, die freie Plätze versprochen hat. Dieselbe Person sieht auf der Anmeldeseite für denselben Termin die richtige Zahl, weil `RegistrationPage.tsx:802` zuerst `getLiveCounterStats` (für alle lesbar) fragt und die Liste nur als Rückfall nutzt.

**Ursache:** Zwei Ansichten derselben Daten mit zwei verschiedenen Quellen: `RegistrationPage` wurde in v24.73/v30.62 auf den Counter (`getCounterStats`, unterscheidet sogar „nie gesetzt" von 0) umgestellt, `MyEventSubEvents` blieb beim direkten Listen-Zählen. Ein HTTP-Status-Check hilft hier nicht — die Antwort ist 200, nur inhaltlich beschnitten.

**Fix-Vorschlag:** In `refresh()` denselben Weg nehmen wie `RegistrationPage.occupancyOf`: zuerst `getLiveCounterStats(ce.id)` und nur bei `seatsKnown === true` die Zahl verwenden; sonst `null` speichern und in der UI statt „0/20" gar keine Zahl (bzw. „Belegung unbekannt") rendern, `isFull` in dem Fall auf false lassen, aber auch kein „frei" behaupten. `getLiveCounterStats` liegt bereits im `useEvents()`-Kontext und muss nur bis `MyEventSubEvents` durchgereicht werden.

### 12. [hoch] Bestätigungs-Dialog: die dort geänderte Auswahl wird beim Absenden ignoriert (veraltete Closure)

**Wo:** `components/registration/SubmitConfirmModal.tsx:175` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Event mit aktiviertem Sicherheitshinweis (event.confirmDialogEnabled, Modus „Auswahl", nicht freetext). Teilnehmer hakt auf der Seite Sub-Event A und B an, klickt „Anmelden" → der Dialog öffnet mit A und B. Er entfernt im Dialog den Haken bei B und klickt „Anmeldung bestätigen". Ergebnis: er wird für A UND B angemeldet. Umgekehrt: hakt er im Dialog zusätzlich C an, wird C NICHT angemeldet. Und bei einem normalen Event mit Sub-Events: entfernt er im Dialog den Haken beim Haupt-Event, wird das Haupt-Event trotzdem angemeldet (inkl. Mail und Outlook-Termin).

**Ursache:** `handleSubmit` wird bei jedem Render von `createSubmitFlow(...)` neu gebaut (RegistrationPage.tsx:1509) und schließt über `selectedSessions`, `registerForParent`/`willRegisterParent` GENAU DIESES Renders. Der onClick-Handler destrukturiert `handleSubmit` aus den Props (Zeile 36) und ruft im `setTimeout` diese bereits gebundene Funktion auf — also die aus dem Render VOR den setState-Aufrufen. Die 60 ms erzeugen zwar einen Re-Render mit einem neuen handleSubmit, aber die Timeout-Closure zeigt weiter auf das alte. Die neue Auswahl (`confirmDraftSessions`) landet nur im State, nie in der ausgeführten Funktion. Der Dialog erfüllt damit genau die Aufgabe nicht, für die er gebaut wurde („die (ggf. angepasste) Auswahl in den echten State übernehmen, dann Submit erneut anstoßen").

**Fix-Vorschlag:** Die Auswahl nicht über den State zurückspielen, sondern am Submit vorbeireichen: entweder `handleSubmit` in RegistrationPage in einem Ref spiegeln (`handleSubmitRef.current = handleSubmit;` bei jedem Render) und hier `handleSubmitRef.current()` im Timeout aufrufen, oder — sauberer — ein Override-Ref setzen (`submitOverrideRef.current = { parent: confirmDraftParent, sessions: new Set(confirmDraftSessions) }`), das `handleSubmit`/`performRegistration` zu Beginn ausliest und statt `selectedSessions`/`willRegisterParent` verwendet. Dieselbe Falle steckt in ExternalEmailWarningModal/CcSelfModal/AssistantModal — dort wird aber nur ein Ref gesetzt, kein State, deshalb dort unkritisch.

### 13. [hoch] Check-in meldet Erfolg, obwohl das Schreiben fehlschlug (Rückgabewert ignoriert)

**Wo:** `components/CheckInPage.tsx:705` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Am Einlass wird eine Person über QR oder Teilnehmer-ID gesucht, die Bestätigungskarte erscheint, der Helfer klickt „Einchecken". Der MERGE auf die Teilnehmerliste scheitert (403, weil das Check-in-Team keine Schreibrechte auf der Sub-Event-Subsite hat, oder 429 bei einer Einlasswelle, oder 400 weil die Spalte fehlt). Die Seite zeigt trotzdem die grüne Meldung „… eingecheckt", der Sitzungszähler steigt, die Person geht rein. In der Teilnehmerliste steht sie weiter als „Angemeldet" und taucht später in der No-Show-Auswertung auf. Niemand merkt es, weil kein einziger Fehlversuch sichtbar wird.

**Ursache:** `EventService.checkInParticipant` delegiert an `services/events/registrationStatus.ts:315-335`. Diese Funktion fängt jeden Fehler selbst ab und liefert `response.ok` bzw. `false` zurück — sie wirft NIE. Der `catch`-Zweig in `confirmCheckIn` ist damit toter Code, und der Rückgabewert wird gar nicht erst ausgewertet. Zum Vergleich: `markNoShowFromSearch` (Zeile 355-357 derselben Datei) macht es richtig (`const success = await …; if (success) …`).

**Fix-Vorschlag:** Rückgabewert auswerten statt auf einen Throw zu hoffen: `const ok = await eventService.checkInParticipant(...); if (!ok) { setResultMessage(`${pendingCheckIn.name} — Check-in fehlgeschlagen (bitte erneut versuchen oder Organizer informieren).`); setResultType('error'); setPendingCheckIn(null); processingRef.current = false; return; }` — erst danach Zähler hochzählen und Erfolg melden. Den `catch`-Block als zusätzliche Absicherung stehen lassen.

### 14. [hoch] Check-in per QR-Code liest die ÄLTESTE Zeile zur Adresse — eine abgemeldete Alt-Zeile weist eine angemeldete Person am Einlass ab

**Wo:** `services/events/registrationStatus.ts:415` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Voraussetzung sind zwei Zeilen zur selben Adresse — genau der Zustand, den v27.11 in getMyRegistration und v30.54 in utils/b2runBibImport.ts (`Bei zwei Zeilen zur selben Adresse gewinnt die AKTIVE — sonst gilt eine alte Abmeldung als aktueller Stand`) ausdrücklich als real behandeln. Entstehung: Person meldet sich ab (Zeile #50 → 'Abgemeldet'); später meldet eine Assistenz sie erneut an. Wegen ReadSecurity=2 sieht die Assistenz Zeile #50 nicht, getMyRegistration liefert ihr nichts, es entsteht Zeile #200 mit Status 'Angemeldet'. Am Event-Tag scannt das Check-in-Team ihren QR-Code. `processCode` (components/CheckInPage.tsx:640) ruft `getRegistrationByEmail`; SharePoint liefert ohne $orderby die Zeilen in Id-Reihenfolge, $top=1 nimmt also #50. Die Seite antwortet `<Name> — Anmeldung storniert` (CheckInPage.tsx:656-659) und bricht ab. Die Person steht mit gültigem QR-Code vor der Tür und gilt als abgemeldet; nur wer weiß, dass er stattdessen über das Namens-/Teilnehmer-ID-Suchfeld gehen muss, kommt rein. Derselbe Griff auf eine 'Eingecheckt'-Alt-Zeile meldet fälschlich „bereits eingecheckt“.

**Ursache:** `getRegistrationByEmail` schließt aus EINER Adresse auf GENAU EINE Zeile: `$top=1` ohne `$orderby` und ohne Status-Filter. Welche der mehreren Zeilen zurückkommt, entscheidet SharePoint (praktisch die kleinste Id = die älteste). Die Funktion hat genau EINEN Aufrufer — den QR-Scan am Check-in-Tisch — und ist damit die einzige Stelle der Registrierungs-Lookups, die den v27.11-Fix nicht mitbekommen hat.

**Fix-Vorschlag:** Query wie in getMyRegistration/markConsentPendingByEmail sortieren und den aktiven Stand bevorzugen: `&$orderby=Id desc` ergänzen und `email.trim()` statt `email` escapen. Sauberer noch (analog zu buildBibReport in utils/b2runBibImport.ts): `$top=20` holen und clientseitig die erste Zeile mit Status in ['Angemeldet','QR versendet','Eingecheckt'] nehmen, sonst die neueste — dann gewinnt die aktive Zeile auch dann, wenn sie NICHT die neueste ist.

### 15. [hoch] Co-Organizer dürfen nach Ablauf der Frist nicht anmelden, obwohl die Rollenmatrix es zusagt

**Wo:** `services/events/registration.ts:117` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Co-Organizer (steht in `EmailTemplateOverrides._coOrganizers`, nicht in `OrganizerEmail`) öffnet nach Ablauf der Anmeldefrist im Organizer Center „Teilnehmer hinzufügen" — der Dialog ist ausdrücklich für „nachträgliche Zusagen" gedacht und wird über `isOrganizerFor(selectedEvent)` auch für Co-Organizer eingeblendet. Jede Zeile im Ergebnis-Report zeigt `deadline`, keine einzige Person wird angemeldet. Dieselbe Sperre trifft ihn auf der Anmeldeseite. Die Rollenmatrix führt „Nach Anmeldefrist registrieren" für coorganizer aber ausdrücklich als „Eigene Events" (components/RoleMatrixPage.tsx:200-202). Zweiter Auslöser derselben Zeile: `sessionEmail` stammt hier allein aus `pageContext.user.email` — steht im Event die SMTP-Adresse und liefert der pageContext den UPN/Alias, wird auch der HAUPT-Organizer abgewiesen. Genau dafür sammelt `canRegisterForOthers` (services/events/profileData.ts:35-42) seit v19.6 zusätzlich die Adresse aus `loginName`.

**Ursache:** Check B liest nur die SharePoint-Spalte `OrganizerEmail` und splittet ausschließlich an `;`. Der Co-Organizer-Kreis liegt seit v9.18 im Piggyback `_coOrganizers` und wird hier nicht ausgewertet — obwohl Check A wenige Zeilen darüber (Zeile 104) über den Parameter `actorIsEventOrganizer` genau diesen erweiterten Kreis akzeptiert und `canRegisterForOthers` ihn serverseitig ebenfalls auflöst. Der Bypass-Parameter wurde bewusst nur auf Check A angewandt (weil eine Assistenz die Frist nicht umgehen darf), dabei ist aber der legitime Organizer-Fall mit untergegangen.

**Fix-Vorschlag:** In Check B denselben Personenkreis zulassen wie in Check A: das Feld `EmailTemplateOverrides` mit in das `$select` aufnehmen, `_coOrganizers` mit der bereits in `canRegisterForOthers` vorhandenen `splitEmails`-Logik (HTML strippen, an `;`/`,`/Zeilenumbruch splitten) parsen und zu `eventOrganizerEmails` addieren — oder schlicht `if (actorIsEventOrganizer) isEventOrganizer = true` setzen, da der Client dieselbe Datengrundlage hat wie die Button-Sichtbarkeit. Zusätzlich `sessionEmail` in Check B über dieselbe Identitäten-Menge (pageContext-E-Mail UND E-Mail aus `loginName`) matchen wie `canRegisterForOthers`, damit die Frist-Ausnahme nicht an zwei Schreibweisen derselben Person scheitert.

### 16. [hoch] E-Mail-Änderung im Organizer Center pflegt DEX_Participants nicht nach — Event verschwindet aus „Meine Events“ und die Doppel-Anmelde-Prüfung ist für die neue Adresse blind

**Wo:** `components/admin/logic/useEditModalHandlers.ts:150` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Organizer öffnet in der Teilnehmerliste „Bearbeiten“ und korrigiert die E-Mail einer Person (Tippfehler bei manueller Anlage, Namensänderung nach Heirat) von alt.name@deloitte.de auf neu.name@deloitte.de. Speichern meldet Erfolg. Danach: (1) `DEX_Participants` trägt die EventNumber weiterhin unter alt.name@deloitte.de, für neu.name@deloitte.de existiert kein Eintrag. (2) Die Person öffnet „Meine Events“ — MyEventsPage.tsx:421 lädt `getMyEventNumbers()`; hat sie noch irgendein anderes Event im Register, greift der Altdaten-Fallback (MyEventsPage.tsx:600 `} else {`) NICHT, und die Nachlese-Schleife ab Zeile 566 nimmt nur Zeilen mit Status 'Abgemeldet' auf. Das Event fehlt komplett, sie hält sich für nicht angemeldet. (3) Sie meldet sich erneut an: die v28.22-Prüfung auf unsichtbare Doppel-Anmeldungen (components/registration/submitFlow.ts:851 `getEventNumbersForEmail(participantEmail)`) fragt das Register unter der NEUEN Adresse, bekommt nichts zurück, warnt nicht — und es entsteht eine zweite Zeile, die einen zweiten Platz belegt. Genau diese Kette beschreibt die App selbst in AdminActionsCard.tsx:610 als Symptom eines fehlenden Register-Eintrags. (4) Bei der späteren Abmeldung ruft der Cancel-Pfad `removeParticipantEvent(neueAdresse, eventNumber)` — findet keinen Datensatz, die EventNumber bleibt für immer auf der alten Zeile stehen und taucht in `analyzeRegistryAgainstLists` als verwaister Verweis auf.

**Ursache:** Die E-Mail ist laut CLAUDE.md der einzige Schlüssel des Datenmodells; DEX_Participants schlüsselt ausschliesslich darüber (`getParticipantByEmail`, `$filter=Email eq '...'`). Alle Schreibpfade, die diesen Schlüssel VERGEBEN, machen einen Dual-Write (registerForEvent, registerTeam, addTeamMember, cancelRegistration). Der einzige Pfad, der den Schlüssel ÄNDERT, macht ihn nicht: `saveEdit` behandelt ParticipantEmail wie ein gewöhnliches Textfeld und schiebt es in dasselbe `patch`-Objekt wie Vorname/JobTitle. Dass der Wert die Identität der Person in einer zweiten Liste ist, kommt im Code nicht vor.

**Fix-Vorschlag:** In `saveEdit` nach erfolgreichem `adminUpdateRegistration` den Registerwechsel nachziehen, wenn `newEmail !== oldEmail` und `selectedEvent.eventNumber` gesetzt ist — in dieser Reihenfolge (prüfbare Nebenbuchhaltung zuerst, CLAUDE.md): `await eventServiceRef.upsertParticipant(newVorname, newNachname, newEmail, selectedEvent.eventNumber, editingReg.Status === 'Warteliste' ? 'Warteliste' : 'Angemeldet')`, danach `await eventServiceRef.removeParticipantEvent(oldEmail, selectedEvent.eventNumber)`. Bei einem Klammer-Event zusätzlich über alle aktiven Sub-Event-Zeilen der Person laufen (deren EventNumbers hängen an derselben Adresse). Schlägt einer der beiden Schritte fehl, den Organizer darauf hinweisen statt still Erfolg zu melden. Zweitens: Da die Zeile nach dem Wechsel der alten Person gehört (Zeilen-Autor), zusätzlich `trySetItemAuthor(subsiteUrl, REG_LIST_NAME, editingReg.Id, newEmail)` aufrufen — sonst sieht die neue Adresse ihre eigene Zeile wegen ReadSecurity=2 nicht.

### 17. [hoch] Fail-closed-Doppelanmelde-Prüfung (v30.14) ist wirkungslos: isUserAlreadyOnEvent und getMyRegistration liefern bei HTTP-Fehler „nicht angemeldet"

**Wo:** `services/EventService.ts:4340` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Anmeldewelle (Soft Opening). Person A klickt „Anmelden", der Insert läuft durch, die Bestätigung kommt aber verzögert; A klickt drei Minuten später erneut. Beim zweiten Klick antwortet SharePoint auf den Dup-Check-GET mit HTTP 429 (GETs laufen bewusst NICHT über withThrottleRetry, siehe EventService.ts:432ff „get geht wieder direkt raus"). `isUserAlreadyOnEvent` liefert `false`, direkt danach liefert `getMyRegistration` `null` — beide Guards sagen „nicht angemeldet". Es wird eine ZWEITE Zeile eingefügt (bei vollem Event eine doppelte Warteliste-Zeile). Genau der Vorfall, den der v30.14-Kommentar in EventContext.tsx:660ff beschreibt und zu verhindern behauptet.

**Ursache:** EventContext.tsx:688 schreibt `await eventService.isUserAlreadyOnEvent(...).catch(() => null)` und wertet `null` als „Prüfung gescheitert → fail-closed". `.catch` feuert aber nur bei einer abgelehnten Promise. `_sp.get` ist ein blanker `spHttpClient.get` (EventService.ts:432) und resolved bei 429/403/500 ganz normal mit `ok === false`; beide Service-Methoden fangen das intern ab und geben ihren „nichts gefunden"-Wert zurück. Der Unterschied zwischen „geprüft, nicht angemeldet" und „konnte nicht prüfen" geht damit im Service verloren, bevor der Aufrufer ihn sehen kann — dieselbe Falle wie bei getAllRegistrations (CLAUDE.md), nur ohne onHttpError-Ausweg.

**Fix-Vorschlag:** Beide Methoden müssen den Fehlstatus melden, statt ihn zu verschlucken. Minimal-invasiv analog zu getAllRegistrations: optionalen `onHttpError?: (_status: number) => void` ergänzen und in EventContext.tsx:686-694 auf dessen Auslösung fail-closed reagieren (`reason: 'dup-check-failed'`). Alternative ohne Signaturänderung: eine dritte Rückgabe (`'unknown'`) bzw. Werfen bei `!response.ok` in isUserAlreadyOnEvent — dann muss der Team-Pfad (EventContext.tsx:1396, 1747, 2061) das bewusst wieder auf „nicht blockieren" abfangen, weil er dort dokumentiert fail-open sein will.

### 18. [hoch] Fail-closed-Schutz gegen Doppel-Anmeldungen (v30.14) ist wirkungslos — die Prüfung wirft nie

**Wo:** `context/EventContext.tsx:685` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Während einer Anmeldewelle (oder beim Massen-Hinzufügen über „Teilnehmer hinzufügen", das je Person × Ziel sequentiell registriert) antwortet SharePoint auf die Duplikatsprüfung mit 429. `isUserAlreadyOnEvent` liefert dann still `false`, `getMyRegistration` gleich danach `null` — beide Prüfungen sagen „nicht angemeldet". Die Anmeldung läuft durch und legt eine ZWEITE aktive Zeile für dieselbe E-Mail an. Genau der Befund, den v30.14 beheben sollte (belegte Doppel-Warteliste im Soft Opening, zwei Zeilen im Abstand von Minuten). Der Organizer sieht die Person doppelt in der Teilnehmerliste, der Platzzähler ist um eins zu hoch.

**Ursache:** `isUserAlreadyOnEvent` (services/events/registration.ts:513-528) fängt alles selbst ab: `if (!response.ok) return false;` und `catch { return false; }` — sie wirft unter keinen Umständen. Das angehängte `.catch(() => null)` ist damit toter Code, `alreadyActive === null` tritt nie ein und der als „fail-CLOSED statt fail-open" kommentierte Zweig läuft nie. Der zweite Netz-Knoten `getMyRegistration` (services/events/registrationEdit.ts:440-461) verhält sich identisch (`if (!response.ok) return null; catch { return null; }`), sodass ein Lesefehler durchgängig als „keine Anmeldung vorhanden" interpretiert wird.

**Fix-Vorschlag:** `isUserAlreadyOnEvent` einen dreiwertigen Rückgabewert geben — z.B. `Promise<boolean | null>` mit `null` bei `!response.ok`/`catch` (die Doc-Zeile „false = frei, auch bei SP-Fehlern" entsprechend anpassen) — oder alternativ einen `onHttpError`-Rückruf wie bei `getAllRegistrations`. Erst dann greift der bestehende `alreadyActive === null`-Zweig. Die wenigen Aufrufer außerhalb von `registerForEvent` (Team-Add) müssen `null` ebenfalls als „nicht prüfbar → ablehnen" behandeln.

### 19. [hoch] Gruppen-Belegung auf der Anmeldeseite kommt aus getAllRegistrations — unter Element-Sicherheit sieht der Teilnehmer beide Gruppen als komplett frei

**Wo:** `components/RegistrationPage.tsx:1068` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Split-Event, Durchstarter (50) ist voll, Funstarter hat 40 Plätze frei. Ein normaler Teilnehmer öffnet die Anmeldeseite. Die Teilnehmerliste hat seit v26.87 ReadSecurity=2 („nur eigene Elemente") — getAllRegistrations liefert ihm 0 Zeilen und wirft bei 403 nicht, sondern gibt `[]` zurück. Die Gruppenkarten zeigen beide „50 / 50 frei", die Zusammenfassung zeigt keine Warteliste. Er wählt Durchstarter. In submitFlow.ts:497 ist `wunschFree === 50`, der Dialog „Wunsch-Gruppe voll — in die andere wechseln?" erscheint NICHT. reserveSeat liefert korrekt 'full', er landet ohne Vorwarnung auf der Warteliste, obwohl 40 Plätze in der anderen Gruppe frei gewesen wären. Genau die Regression, die v27.10 für nicht-geteilte Events mit dem für alle lesbaren Counter (getLiveCounterStats) behoben hat.

**Ursache:** Der Split-Zweig wurde bei der v27.10-Umstellung nicht mitgezogen: Für Events mit maxParticipants > 0 liest RegistrationPage.tsx:513 den Counter über getLiveCounterStats; bei geteilten Kapazitäten ist maxParticipants 0, der Effekt läuft gar nicht, und starterCounts hängt weiter an der zeilenweise gesicherten Teilnehmerliste. Der Fehler ist doppelt unsichtbar: getAllRegistrations wird ohne `onHttpError` gerufen (leer == verboten == leer, CLAUDE.md), und der umschließende `catch { /* ignore */ }` (Z. 1082) lässt starterCounts auf null, was in den Anzeigen über `?? 0` wieder als „niemand angemeldet" gelesen wird.

**Fix-Vorschlag:** starterCounts aus dem für alle lesbaren DEX_TeilnehmerCounter speisen: getCounterStats liefert bereits SeatsTakenDurch/SeatsTakenFun; einen Aufruf analog getLiveCounterStats bereitstellen, der die Gruppenfelder (inkl. seatsKnown) durchreicht. Solange das nicht steht, mindestens `getAllRegistrations(url, () => setStarterCounts(null))` nutzen und bei unbekannten Zahlen weder „X / Y frei" noch „ausgebucht" rendern, sondern einen Strich — eine erfundene 0 ist schlimmer als keine Zahl. Achtung: reserveSeat pflegt bei Split nur das Gruppenfeld, nicht SeatsTaken; getCounterStats Z. 5900ff bevorzugt aber `total`, wenn es > 0 ist — für die Gruppenanzeige daher gezielt durch/fun lesen.

### 20. [hoch] HotelSetupWizard liest Hotel-Fragen und -Antworten nur vom Hauptevent — bei Klammer-Events fallen Formular-Zeiträume und Zusatznächte still weg

**Wo:** `components/HotelSetupWizard.tsx:183` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Klammer-Event, die Hotel-Fragen stehen (wie üblich) im Formular der Sub-Events — entweder als `daterange` („Hotel: An-/Abreise") oder als Bedarfsfrage + „Hotel (additional nights)". Im Panel steht bei einer Person korrekt „Nächte gewünscht: 3 ~" (v29.6 löst über answerRowsOf und Parent+Child-Felder auf). Der Organizer öffnet den Einrichtungs-Assistenten: Schritt 1 zeigt weder die Tabelle „Zeiträume aus dem Anmeldeformular" noch „Zusatznächte — aus dem Anmeldeformular gelesen", Schritt 4 meldet „0 Extranächte" und schreibt allen den Standard-Zeitraum (2 Nächte). Direkt danach zeigt dasselbe Panel die rote Karte „N Person(en) mit abweichender Nächtezahl" für genau diese Leute — Assistent und Tabelle widersprechen sich.

**Ursache:** `fields` (Z. 182–192) durchsucht ausschließlich `event.eventSpecificFields`, `answersOf` (Z. 175–177) ausschließlich `p.CustomData` der übergebenen Zeile — und übergeben wird `people`, also die Klammer-Zeilen, die laut CLAUDE.md Schattenzeilen ohne CustomData sind. Die v29.3/v29.6-Auflösung (answerRowsOf + Parent- UND Child-`eventSpecificFields`) existiert nur im Panel und wurde beim Assistenten nie nachgezogen. Auffällig unauffällig wird der Fehler dadurch, dass `wishOf` als Prop hereingereicht wird und deshalb KORREKT auflöst: „wer braucht ein Zimmer" stimmt, nur die Zeiträume nicht — genau die Zahl, für die es den Assistenten gibt.

**Fix-Vorschlag:** Dem Wizard dieselbe Auflösung geben, die das Panel schon hat: `answerRowsOf` und eine gemergte Feldliste (`[...event.eventSpecificFields, ...childEvents.flatMap(c => c.eventSpecificFields)]`) als Props durchreichen (Panel Z. 2047 ff.) und in `fields` / `answersOf` verwenden — `answersOf` als Merge über `answerRowsOf(p).slice().reverse()` wie in `wishOf` (Panel Z. 884–893). Sauberer wäre, `formStayOf`/`wishNightsOf` einmal nach `utils/hotelAnswers` zu heben und in beiden Komponenten zu benutzen, damit die Deutung nicht ein drittes Mal auseinanderläuft.

### 21. [hoch] Leere Teilnehmerliste (403/404) wird im Check-in als „Person nicht angemeldet" behauptet

**Wo:** `components/CheckInPage.tsx:172` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Eine Co-Organizerin (oder ein Check-in-Helfer) öffnet Check-in und wählt im Picker einen Sub-Event-Termin, auf dessen Subsite sie keine Leserechte hat (genau der v30.37-Fall: Berechtigungen wurden nur auf der Klammer gesetzt). `getAllRegistrations` läuft in einen 403, bricht ab und liefert `[]`. Die Seite zeigt weder Fehler noch KPI-Kacheln, sondern „Keine Teilnehmer für dieses Event." Tippt sie die Teilnehmer-ID aus der QR-Mail ein, antwortet die App: „Keine Anmeldung mit der Teilnehmer-ID 42 bei diesem Event. Bitte die Nummer aus der QR-Mail prüfen" — eine positive Aussage über Daten, die nie gelesen wurden. Der Teilnehmer wird am Einlass abgewiesen bzw. auf eine falsche Fehlersuche geschickt.

**Ursache:** `getAllRegistrations` wirft nicht (services/events/registrationEdit.ts:652-684: `if (!response.ok) { if (onHttpError) onHttpError(response.status); break; }`) — der `catch`-Zweig hier ist tot, `searchLoadError` wird nie gesetzt. CheckInPage übergibt den seit v29.3 existierenden `onHttpError`-Rückruf nicht, obwohl der EventContext-Wrapper (context/EventContext.tsx:2360-2371) ihn durchreicht und sogar die fehlende Subsite als `onHttpError(0)` meldet. Genau der in CLAUDE.md zweimal dokumentierte Fehler („Ein leeres Ergebnis ohne geprüften Status ist keine Aussage über die Daten, sondern über gar nichts"), in der Check-in-Seite noch nicht nachgezogen.

**Fix-Vorschlag:** `let httpErr = 0; const regs = await getAllRegistrations(eventId, s => { httpErr = s; }); if (httpErr) { setSearchLoadError(httpErr === 403 ? 'Keine Leseberechtigung auf der Teilnehmerliste dieses Termins — bitte Organizer/Admin um Freigabe bitten.' : 'Teilnehmerliste konnte nicht geladen werden (HTTP ' + httpErr + ').'); return; }` — und den Cache in dem Fall NICHT mit `[]` füllen, damit weder „Keine Teilnehmer" noch „ID nicht gefunden" behauptet wird. `checkInByParticipantId` muss denselben Zustand unterscheiden (Liste gesperrt vs. ID unbekannt).

### 22. [hoch] Nachzug der fehlenden Klammer-Zeilen (shadowHeal) läuft nie: der 20-s-Timer wird von der eigenen Effect-Cleanup abgeräumt und nicht neu gesetzt

**Wo:** `context/EventContext.tsx:2809` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Bei einer Anmeldung ist die Klammer-Zeile an der Drosselung gescheitert und liegt als Merker in `utils/shadowHeal` (localStorage). Beim nächsten App-Start feuert der Effect beim ersten nicht-leeren `events`, setzt `shadowHealRanRef.current = true` und plant den Heal-Lauf für t+20 s. Wenige Sekunden später beendet `loadEvents` seinen nachgelagerten Teilnehmerzähl-Lauf und ruft `setEvents(withCounts)` mit einer NEUEN Array-Referenz (Zeile 316) → React führt die Cleanup aus (`clearTimeout`) und den Effect erneut aus, der wegen des bereits gesetzten Refs sofort zurückkehrt. Der Timer wird nie neu gesetzt: Die Klammer-Zeilen werden nie nachgeholt, die betroffenen Personen fehlen dauerhaft in der Klammer-Teilnehmerliste, bis jemand manuell repariert. Der Merker verfällt nach 14 Tagen still.

**Ursache:** Der Guard (`shadowHealRanRef`) und die Cleanup arbeiten gegeneinander: Die Cleanup gehört zum Effect-Lauf (wird bei jeder `events`-Änderung ausgeführt), der Guard verhindert aber das erneute Aufsetzen. Da `loadEvents` immer eine neue Array-Referenz liefert (`mapLimited` gibt ein neues Array zurück) und der Zähl-Lauf normalerweise deutlich unter 20 s dauert, trifft das praktisch bei jedem Start zu.

**Fix-Vorschlag:** Die Timer-Id in einem Ref halten und nur beim Unmount löschen: `const healTimerRef = React.useRef<number|null>(null); … healTimerRef.current = window.setTimeout(…);` und in einem separaten Effect mit `[]`-deps `return () => { if (healTimerRef.current) window.clearTimeout(healTimerRef.current); }`. Alternativ den Effect an ein einmalig kippendes Boolean (`events.length > 0`) statt an `events` hängen.

### 23. [hoch] Online-Meeting-Modus fehlt in `subTopGateKey` — die v29.57-Sperre verwirft die Änderung für ALLE Termine still

**Wo:** `components/EventCreationPage.tsx:7080` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Bestehendes Klammer-Event mit z.B. 10 Terminen, bisher ohne Online-Meeting (`onlineMeetingMode === 'none'`, teamsLink leer). Der Organizer öffnet den Wizard, stellt in Schritt Kommunikation NUR auf „DEX erzeugt den Link" (`'auto'`) und speichert. Das Hauptevent bekommt `OutlookIsOnlineMeeting = true` (Zeile 4579), aber KEIN einziger Termin: `persistSubEventsForParent` überspringt sie alle (Zeile 3825), weil `subGateUnchanged` true ist. Die Sub-Event-Zeilen behalten `OutlookIsOnlineMeeting = false`. Dasselbe umgekehrt beim Abschalten ('auto' → 'none'): die Termine erzeugen weiter Teams-Besprechungen. Im Update-Modal werden die Termine zwar als geändert angeboten (`onlineMeetingChanged()` in der Detect-Schleife, Zeile ~6160) und ein UpdateEvent wird gequeued — der Flow liest dann aber die unveränderte Spalte der Sub-Event-Zeile.

**Ursache:** `subGateUnchanged = editEvent && subTopGateInitialRef.current === subTopGateKey()` entscheidet einmal je Save, ob unveränderte Sub-Event-Entwürfe geschrieben werden. `onlineMeetingMode` fließt über `childPayload.outlookIsOnlineMeeting` (Zeile 3650) und `subUpdates['OutlookIsOnlineMeeting']` (Zeile 3902) in jedes Sub-Event, steht aber weder in `subTopGateKey` (dort nur `teams: teamsLink`) noch in `computeFormSnapshot` — und `teamsLink` bewegt sich beim Wechsel none↔auto gar nicht. Das ist exakt der Fall, vor dem der Doku-Kommentar über `subPersistKey` warnt: „WER HIER ETWAS ERGAENZT, das vom Hauptevent in ein Sub-Event fließt, muss es in `subTopGateKey` aufnehmen — sonst wird die Aenderung still verworfen." v30.26 hat den Modus eingeführt, die Schranke aber nicht nachgezogen. (Nebenwirkung derselben Lücke: der Ungespeichert-Wächter über `computeFormSnapshot` schlägt bei einer reinen Modus-Änderung ebenfalls nicht an — Verlassen ohne Warnung.)

**Fix-Vorschlag:** `onlineMeetingMode` in `subTopGateKey` aufnehmen, z.B. `teams: `${onlineMeetingMode}|${teamsLink}`` (deckt beide Richtungen ab, ohne einen neuen Schlüssel). Zusätzlich `onlineMeetingMode` in `computeFormSnapshot` ergänzen, damit auch der Ungespeichert-Wächter greift.

### 24. [hoch] Outlook-Update des Hauptevents wird am Sub-Event-Schalter vorbei verworfen — OutlookDirty wird trotzdem gelöscht

**Wo:** `components/wizard/logic/wizardSubmit.ts:1117` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Event mit Sub-Events. Der Organizer ändert in Schritt 1 den Titel, wechselt in Schritt „Kommunikation" auf den Reiter eines Sub-Events, für das „Outlook-Termin deaktivieren" gesetzt ist, und speichert von dort. `switchCommTab` (components/wizard/logic/commTabs.ts:133) hat `setDisableOutlook(true)` aus dem Sub-Slot gesetzt — der rohe State hält jetzt den SUB-Wert. Das Update-Modal listet das Hauptevent trotzdem (die Erkennung nutzt korrekt `resolveTopLevelCommState().disableOutlook`, siehe outlookChanges.ts:142), der Organizer hakt es an → `confirmOutlookSave` setzt `pendingOutlookDirtyWriteRef.current = !topChecked` = false (outlookActions.ts:263), das wird in Zeile 650 als `OutlookDirty:false` geschrieben. Danach greift Zeile 1117 nicht, weil `disableOutlook` (Sub-Wert) true ist: KEIN `UpdateEvent` in der Queue. Ergebnis: Der Kalender aller Teilnehmer behält den alten Titel, der Dirty-Marker ist gelöscht, der Wizard bietet die Aktualisierung nie wieder an — und die App meldet „gespeichert".

**Ursache:** Der Save-Pfad prüft an dieser einen Stelle den rohen `disableOutlook`-State statt `effDisableOutlook` (= `topComm.disableOutlook`), obwohl derselbe Wert zwei Zeilen weiter oben schon aufgelöst vorliegt und in Zeile 640 in die Spalte geschrieben wird. Genau diese Falle ist an zwei anderen Stellen bereits behoben (v18.45/v18.50 in `detectOutlookRelevantChangesImpl`); der Queue-Zweig wurde damals nicht mitgezogen.

**Fix-Vorschlag:** In Zeile 1117 (und in der Zählung in Zeile 1102) `disableOutlook` durch `effDisableOutlook` ersetzen. Zusätzlich absichern: `OutlookDirty:false` erst schreiben, wenn `queueOutlookEvent` `true` geliefert hat — sonst geht der Marker auch bei einem 429 verloren.

### 25. [hoch] Outlook-Update des Hauptevents wird beim Speichern von einem Sub-Reiter aus stumm übersprungen (roher `disableOutlook` statt `effDisableOutlook`)

**Wo:** `components/EventCreationPage.tsx:5580` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Bestehendes Event mit Sub-Events; das Hauptevent hat einen Outlook-Termin, ein Sub-Event (z.B. „Tag 2") hat in Schritt 6 den Schalter „Outlook-Termin deaktivieren" gesetzt. Der Organizer ändert in Schritt 1 auf der Klammer-Ebene Titel oder Startzeit, wechselt danach in Schritt 6 auf den Reiter „Tag 2" (dadurch steht `disableOutlook === true` im UI-State) und klickt „Speichern". `detectOutlookRelevantChanges()` löst korrekt über `resolveTopLevelCommState()` auf und listet das Hauptevent im Bestätigungs-Dialog; der Organizer hakt es an. `confirmOutlookSave` setzt `pendingOutlookUpdateForTopRef.current = true` UND `pendingOutlookDirtyWriteRef.current = false`. Im Save wird `OutlookDirty=false` geschrieben, der Zweig hier aber wegen `!disableOutlook === false` übersprungen: es landet KEIN 'UpdateEvent' in DEX_Outlook. Ergebnis: alle Teilnehmer behalten den alten Titel/die alte Uhrzeit im Kalender, die App meldet Erfolg, und weil OutlookDirty auf false steht, weist auch der nächste Wizard-Lauf nicht mehr darauf hin.

**Ursache:** Genau die Falle, für die v11.93 `resolveTopLevelCommState()` eingeführt hat: Auf einem Sub-Reiter tragen die Step-6-State-Variablen die Werte des Sub-Events. Alle anderen Stellen in `handleSubmitInner` nutzen deshalb die `eff*`-Konstanten (`effDisableOutlook` ist zwei Zeilen weiter oben definiert und im Scope), diese beiden Stellen wurden beim Umbau vergessen. Dieselbe Verwechslung steht in der Fortschritts-Rechnung darüber (`outlookTotal`, Zeile 5563: `(!disableOutlook && pendingOutlookUpdateForTopRef.current) ? 1 : 0`) — dort nur kosmetisch.

**Fix-Vorschlag:** `disableOutlook` an beiden Stellen durch `effDisableOutlook` ersetzen (Zeile 5580 und in `outlookTotal`, Zeile 5563). Zusätzlich sinnvoll: wenn der Zweig doch übersprungen wird, `pendingOutlookDirtyWriteRef.current` nicht auf `false` schreiben lassen, damit ein verschluckter Update-Auftrag beim nächsten Öffnen wieder sichtbar ist.

### 26. [hoch] Realtime-Push im Organizer Center überschreibt die geladene Teilnehmerliste bei jedem HTTP-Fehler mit []

**Wo:** `components/AdminPage.tsx:228` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer hat ein Event mit 77 Anmeldungen offen. Ein Teilnehmer meldet sich an → SharePoint-Push → `reload()`. Kommt dieser eine Request mit 429 (Drosselung während einer Anmeldewelle), 403 oder 500 zurück, liefert `getAllRegistrations` `[]` bzw. eine halb gelesene Seite, und `setRegistrations(r)` ersetzt die vollständige Liste. Tabelle, KPI-Kacheln, Warteliste, Duplikat- und Klammer-Boxen zeigen danach 0 — ohne Fehlertext, weil `regLoadError` in diesem Pfad nicht gesetzt wird. Wer daraufhin „Freie Plätze mit Warteliste füllen" oder eine Abmelde-Entscheidung beurteilt, tut das auf einer leeren Liste.

**Ursache:** `getAllRegistrations` wirft nie (registrationEdit.ts:667-683), meldet HTTP-Fehler nur über `onHttpError` — das `.catch(() => {})` hier ist toter Code und suggeriert eine Absicherung, die es nicht gibt. Der Auswahl-Pfad `components/admin/logic/useEventSelection.ts:83` nutzt `onHttpError` seit v30.37; der Realtime-Pfad und der Refresh-Knopf wurden dabei übersehen. Derselbe Defekt steht noch einmal in `AdminPage.tsx:163` (`handleRefresh`, „Aktualisieren"-Knopf).

**Fix-Vorschlag:** In `reload` den Status auswerten und bei Fehler NICHT überschreiben: `let ok = true; getAllRegistrations(ev.id, () => { ok = false; }).then(r => { if (cancelled) return; if (!ok) { setRegLoadError(ACCESS_DENIED_MSG); return; } setRegLoadError(''); setRegistrations(r); });`. Gleiches in `handleRefresh` (Z.163).

### 27. [hoch] Reservierter Sitzplatz wird nach fehlgeschlagenem Insert nie zurückgegeben

**Wo:** `context/EventContext.tsx:749` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Event mit maxParticipants=100, 60 belegt. Eine Person klickt „Anmelden". `reserveSeat` erhöht SeatsTaken per ETag-CAS auf 61. Der darauffolgende `registerForEvent`-POST scheitert (429 während einer Anmeldewelle, 403 auf einer frisch angelegten Subsite, 400 wegen einer fehlenden Custom-Field-Spalte) → `success=false`, die App meldet einen Fehler und legt KEINE Zeile an. SeatsTaken bleibt auf 61. Die Person versucht es dreimal → SeatsTaken 64 bei 60 echten Anmeldungen. Ab 100 im Counter geht jede weitere echte Anmeldung auf die Warteliste, obwohl real Plätze frei sind. Dasselbe in `registerTeam` (Zeile 1445/1451, dort werden gleich N Plätze reserviert und bei `{ok:false, reason:'insert-failed'}` alle N verbrannt) und in `addTeamMember` (Zeile 1790/1796).

**Ursache:** `reserveSeat` ist bewusst atomar und schreibt VOR dem Insert. Der Erfolgspfad ist gepflegt, der Fehlerpfad nicht: nach `success = r.ok` gibt es keinen `adjustSeatCounterField(..., -1)`-Rückweg. Dass das Muster bekannt ist, zeigt `services/events/registrationStatus.ts:133-140` — `switchSplitGroup` gibt den Platz bei fehlgeschlagenem MERGE explizit zurück. Geheilt wird die Drift nur durch `syncSeatsToActiveCount`, das ausschliesslich in privilegierten Organizer-Pfaden läuft; ein reiner Teilnehmer-Pfad heilt nie.

**Fix-Vorschlag:** In `registerForEvent` (EventContext) merken, ob und für welche Gruppe/Anzahl reserviert wurde (`reservedGroup`, `reservedCount`), und im Zweig `if (!success)` vor dem `return` `await eventService.adjustSeatCounterField(subsiteUrl, eventService.seatFieldFor(reservedGroup), -reservedCount)` aufrufen — nur wenn `status === 'Angemeldet'` (bei 'Warteliste' wurde nichts reserviert). Dasselbe in `registerTeam` (Rückgabe von N bzw. N minus erfolgreicher Inserts) und `addTeamMember`.

### 28. [hoch] Screenshot-Vorschau/Markieren zerstört sich selbst: revokeObjectURL läuft bei JEDER shots-Änderung

**Wo:** `components/QuestionButton.tsx:122` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Frage-Modal öffnen → einen Screenshot A anhängen (Live-Capture oder Datei) → ein zweites Bild B anhängen. `setShots(s => [...s, …])` ändert `shots`, React führt VOR dem neuen Effekt-Lauf die Cleanup des vorigen aus — und die schließt über das ALTE Array, ruft also `URL.revokeObjectURL(A.url)`, obwohl A weiter in der Liste steht und gerendert wird. Klick auf A → „Vergrößern & markieren": `ImageAnnotateModal` gibt bei `!open` null zurück (Zeile 37) und mountet das `<img src={A.url}>` daher frisch — die widerrufene Blob-URL lädt nicht, das Modal bleibt leer, `naturalWidth === 0`, und „Übernehmen" (`save()`, Zeile 68-71) schließt wortlos ohne zu speichern. Dasselbe nach `removeShot`: die Cleanup widerruft die URLs ALLER verbliebenen Screenshots.

**Ursache:** Der Kommentar sagt „beim Unmount", der Dependency-Array `[shots]` macht daraus „bei jeder Änderung". Effekt-Cleanups laufen in React vor jedem erneuten Lauf, nicht nur beim Unmount, und sehen den Closure-Stand des vorigen Renders. Das gezielte Einzel-Revoke existiert bereits an allen richtigen Stellen (removeShot Z.230, onAnnotateSave Z.237, resetForm Z.135, submit Z.264) — dieser Effekt ist zusätzlich und falsch. Der Upload selbst bleibt heil, weil `submit()` die `File`-Objekte (`s.file`) schickt; kaputt ist nur alles, was die URL noch braucht.

**Fix-Vorschlag:** Den aktuellen Stand über einen Ref halten und den Effekt wirklich nur beim Unmount laufen lassen: `const shotsRef = React.useRef(shots); shotsRef.current = shots;` und `React.useEffect(() => () => { shotsRef.current.forEach(s => { try { URL.revokeObjectURL(s.url); } catch {} }); }, []);`

### 29. [hoch] Shadow-Heal-Timer wird vom `events`-Update abgeraeumt und nie neu gestellt — die Nachheilung der Klammer-Zeilen laeuft nie

**Wo:** `context/EventContext.tsx:2809` · Quelle: Audit B (nach dem Umbau)

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

### 30. [hoch] Stellvertretende Anmeldung: Klammer-Zeile fehlt, weil parentAlreadyHasRow die Zeile des ANMELDENDEN prüft

**Wo:** `components/registration/submitFlow.ts:795` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Klammer-Event mit subEventsOnlyMode (z.B. Office-Tage). Organizerin O hat sich selbst für das Event angemeldet, also lädt der Effekt in RegistrationPage.tsx:729 ihre Klammer-Zeile nach `myParentReg`. Sie schaltet auf „Für andere Person anmelden" (oder öffnet die Seite direkt über „Register another person", navIntent='register-other'), wählt Person X und Tag B und sendet ab. Ergebnis: X bekommt die Tages-Zeile, aber KEINE Klammer-Zeile. Die übergreifenden Hauptevent-Antworten (Hotel, Verpflegung, Anreise aus `customData`) werden für X nirgends gespeichert, und im Organizer Center taucht X als „Fehlende Klammer-Anmeldung" auf. Der Selbstheilungs-Merker greift ebenfalls nicht, weil `addPendingShadowParent` (Zeile 817) an derselben Bedingung hängt.

**Ursache:** `myParentReg` wird ausschließlich im Selbst-Modus gefüllt — der Effekt in RegistrationPage.tsx:729 ruft `getMyRegistration(event.id)` für den EINGELOGGTEN User und läuft nur `if (!registerForOther)`; beim Umschalten auf den Stellvertreter-Modus wird der Wert weder im Effekt noch im Toggle (PersonalDataSection.tsx:79) zurückgesetzt. Mit v26.67 wurde `!registerForOther` aus `shouldShadowRegisterParent` entfernt („deckt jetzt Selbst- UND Fremd-Anmeldung ab"), die Herkunft von `parentAlreadyHasRow` aber nicht mitgezogen. Weil die Sub-Event-Anmeldung zusätzlich `skipShadowParent: true` setzt (Zeile 1054), zieht auch die zentrale Absicherung aus v30.42 die Klammer nicht mehr nach — niemand schreibt die Zeile. Bei der Deep-Link-Variante kommt eine Race dazu: der Effekt startet beim Mount noch mit `registerForOther === false`, sein `getMyRegistration` löst auf und setzt `myParentReg`, obwohl der Modus danach sofort auf „für andere" kippt.

**Fix-Vorschlag:** `parentAlreadyHasRow` an die Zielperson binden statt an den Anmeldenden: im Stellvertreter-Modus `checkRegistrationByEmail(event.id, participantEmail)` auswerten (der Wert liegt über `thirdPartyCheck.alreadyRegistered` bereits vor). Minimalfix: `const parentAlreadyHasRow = !registerForOther && !!myParentReg;` plus `setMyParentReg(null)` beim Umschalten auf `registerForOther` (Toggle und Effekt), damit auch `parentAlreadyRegistered` und der v30.61-Sammelmail-Zweig (Zeile 1169) nicht am fremden Zustand hängen.

### 31. [hoch] Team-Anmeldung: Promise.all über die Insert-Schreibvorgänge ohne Fehlerzähler — Teilerfolg wird als voller Erfolg gemeldet

**Wo:** `context/EventContext.tsx:1539` · Quelle: Audit B (nach dem Umbau)

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

### 32. [hoch] Wartelisten-Platz bei einem Sub-Event wird in „Meine Events" als „Angemeldet" angezeigt

**Wo:** `components/myEvents/MyEventSubEvents.tsx:672` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Teilnehmer klickt in „Meine Events" bei einem vollen Sub-Event auf „Anmelden". `registerForEvent` reserviert keinen Platz mehr und legt die Zeile mit Status 'Warteliste' an. Es kommt keinerlei Rückmeldung; nach dem `refresh()` steht der Termin grün umrandet da mit dem Zusatz „(Angemeldet)" und dem Knopf „Abmelden". Die Person hält sich für angemeldet, plant An-/Abreise und erscheint am Termin, obwohl sie nur auf der Warteliste steht. Dieselbe Person sieht in der Wartelisten-Mail etwas anderes als in der App.

**Ursache:** `refresh()` (Zeile 108) reduziert den ganzen Status auf `isActive = !!reg && reg.Status !== 'Abgemeldet'` und speichert nur eine Id-Menge (`registeredSet`) — 'Warteliste' fällt damit in denselben Topf wie 'Angemeldet'. Zusätzlich wird in `handleToggle` (Zeile 262) das Ergebnis `regRes` nur auf `regRes.ok` für die Schattenzeile geprüft; weder `regRes.status === 'Warteliste'` noch ein Fehlschlag (`reason: 'full'` bei abgeschalteter Warteliste, `'already-registered'`, `'dup-check-failed'`) erzeugt eine Meldung.

**Fix-Vorschlag:** Statt `Set<string>` einen `Record<string, string>` mit dem echten Status führen (`statusById[ce.id]`) und im Render zwischen 'Warteliste' (orange, „Warteliste") und 'Angemeldet' (grün) unterscheiden — analog zum Badge in `MyEventCard`. Zusätzlich in `handleToggle` nach `registerForEvent` das Ergebnis melden: bei `!regRes.ok` ein `showAlert` mit dem Grund, bei `regRes.status === 'Warteliste'` der Hinweis, dass der Termin voll war und die Person auf der Warteliste steht.

### 33. [hoch] Wizard-Speichern überschreibt Snapshots/Versandstempel mit dem Stand vom Öffnen (billingExtraRef friert beim Mount ein)

**Wo:** `components/EventCreationPage.tsx:292` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Admin öffnet die App und geht in den Wizard eines abrechnungsrelevanten Events, um die WBS-Nummer zu korrigieren. 11 s nach dem Boot startet `maybeSendBillingAutoMails` (DexEventPlatform.tsx:628-637) und arbeitet sequenziell alle Events durch — bei mehreren Events dauert das Minuten. Für genau dieses Event geht die Abrechnungsinfo-Mail raus; `_billing` bekommt in SharePoint `infoSentAt`, `autoInfoSentAt`, `infoSnapshot` und einen Log-Eintrag mit Mail-Volltext. Der Wizard hat seine `billingExtraRef` aber beim Mount eingefroren — sie kennt diese Schlüssel nicht. Der Admin klickt „Speichern": `EmailTemplateOverrides` wird komplett neu aufgebaut (wizardSubmit.ts:606 ff., `Object.assign` über die Piggyback-Liste), und `_billing` enthält danach nur noch relevant/sendMode/fields plus die ALTEN extras. Folge: die revisionssichere Versandhistorie und der Snapshot sind gelöscht, das F&A Center zeigt wieder „Es wurden bisher keine Abrechnungsinformationen an F&A versendet", und weil `autoInfoSentAt` weg ist, verschickt der nächste Admin-Boot dieselbe Mail ein zweites Mal an F&A. Dieselbe Löschung passiert bei zwei parallel arbeitenden Admins (einer im Wizard, einer im Organizer Center auf „Senden").

**Ursache:** Der Wizard schreibt `_billing` als GANZES Objekt zurück, obwohl er nur drei seiner Schlüssel besitzt. Der Schutz dafür ist ein Ref, das beim ersten Render einmalig aus `editEvent` gefüllt und nie mehr aktualisiert wird — auch dann nicht, wenn `applyBillingLocally` den Context-State während der offenen Wizard-Sitzung nachzieht und `editEvent` damit längst neuer ist als das Ref. Für exakt dieses Problem gibt es an anderer Stelle im selben Bauteil bereits die Lehre und die Lösung: `visAllSubsPiggyback` (v30.7, Zeile 285-291) liest bei unberührtem Zustand den Server-Stand erneut, statt dem eingefrorenen zu vertrauen.

**Fix-Vorschlag:** Die Extras beim Bauen des Piggybacks frisch aus dem aktuellen `editEvent` lesen statt aus dem Mount-Ref (Ref nur noch als Fallback), also in `billingPiggyback()`: `const live = (() => { try { const b = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._billing; if (b && typeof b === 'object') { const { relevant, sendMode, fields, ...extra } = b; return extra; } } catch {} return billingExtraRef.current; })();` und `...live` statt `...billingExtraRef.current` — den selbst angehängten Log-Eintrag auf `live.log` statt auf das Ref stapeln. Sauberer wäre, `_billing` aus dem Wizard gar nicht mehr im vollen Overrides-Blob zu schreiben, sondern nach dem Save per `patchEventOverridesValue(idNum,'_billing',…)` als Read-Modify-Write nur relevant/sendMode/fields zu setzen — dann gibt es keinen zweiten Schreiber auf diesem Schlüssel mehr.

### 34. [hoch] `audienceOnly` (Verteiler-Begrenzung von Personen-Feldern) wird beim Laden nie gemappt — Einstellung wirkt nie und wird beim nächsten Speichern gelöscht

**Wo:** `context/eventMapping.ts:543` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer hakt in Schritt „Eventfelder" bei einem Zimmerpartner-/Personen-Feld „nur Personen aus dem Verteilerkreis" an (FieldsStep.tsx:1128) und speichert. wizardSubmit.ts:987 schreibt `audienceOnly:true` korrekt ins CustomFields-JSON. Nach dem nächsten Laden ist `field.audienceOnly` undefined: RegistrationPage.tsx:2123 nimmt wieder `searchUsers` statt `searchUsersInAudience`, der Teilnehmer kann also erneut beliebige Kolleg:innen außerhalb des Verteilers als Zimmerpartner wählen (genau der Befund vom 21.08.2026, den v29.40 beheben sollte), und der Hinweistext (Zeile 2131) erscheint nicht. Öffnet der Organizer danach den Wizard, ist die Checkbox leer (useWizardEventFieldState.ts:287 liest denselben undefined-Wert) — der nächste Save entfernt das Flag endgültig aus dem JSON.

**Ursache:** `mapSPEventToDeloitteEvent` baut `eventSpecificFields` als explizite Property-Liste neu auf. `audienceOnly` (types/index.ts:640) ist als einzige Eigenschaft von `EventSpecificField` in dieser Liste nicht enthalten — dieselbe Drop-Klasse wie v29.20 (rangeStart/rangeEnd/maxNights), v17.19 (confirmLabel) und v11.16 (onlyForGroup), die im selben Block dokumentiert sind.

**Fix-Vorschlag:** In das Feld-Mapping aufnehmen, analog zu ccOnEmails: `audienceOnly: !!(cf as any).audienceOnly,`. Zusätzlich prüfen, ob bestehende Events das Flag inzwischen verloren haben (restoreCustomFieldDescriptions RESTORE_PROPS in maintenance.ts:218 um 'audienceOnly' ergänzen).

### 35. [hoch] `deleteEvent` räumt DEX_Participants mit genau dem Muster auf, das v29.3 in `deleteParticipantData` abgeschafft hat

**Wo:** `services/events/eventsCrud.ts:873` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer löscht ein Event mit 400 Anmeldungen. Reihenfolge im Code: (1) Zeile 822 recycelt die Subsite mit der Teilnehmerliste, (2) erst danach läuft dieser Block. `getAllParticipants()` liest nicht strikt (`readAllParticipants(svc, false)` in participantsRegistry.ts:524 bricht bei !response.ok einfach ab und liefert die bisher gelesenen Seiten) — bei einem 429 auf Seite 2 fehlen alle Personen ab dort und werden gar nicht erst als betroffen erkannt. Für die erkannten Personen feuern 400 MERGE-Requests gleichzeitig; jeder Fehlschlag wird von `.catch(() => null)` verschluckt, nicht gezählt und nicht gemeldet. `deleteEvent` liefert trotzdem `true`. Ergebnis: die EventNumber steht dauerhaft in `EventRegistered`/`EventOnWaitlist`, die Teilnehmerliste ist aber schon weg — nachrechnen unmöglich. Genau diese Verweise meldet später `analyzeRegistryAgainstLists` als „verwaist", und „Meine Events" versucht für die Betroffenen weiterhin ein Event zu laden, das es nicht mehr gibt.

**Ursache:** `deleteParticipantData` (services/events/eventStats.ts:217) wurde in v29.3 auf strikt lesen (`fetchAllParticipantsOrThrow`), sequentiell schreiben, Fehler zählen und bei Fehlern VOR der Löschung abbrechen umgestellt. `deleteEvent` — der häufigere Weg, weil er über die Danger Zone im Organizer Center läuft — ist bei der alten Implementierung stehen geblieben, inklusive der falschen Reihenfolge (unwiderrufliches Recycle vor der prüfbaren Nebenbuchhaltung).

**Fix-Vorschlag:** Den Block aus `deleteParticipantData` übernehmen: `fetchAllParticipantsOrThrow()` statt `getAllParticipants()`, sequentielle Schleife mit `registryFailed`-Zähler statt `Promise.all` + `.catch(() => null)`, und den ganzen Block VOR das Subsite-Recycle (Zeile 822) ziehen — bei `registryFailed > 0` oder Lesefehler `return false`, damit der Aufrufer den Löschversuch wiederholen kann, solange die Liste noch existiert.

### 36. [hoch] createEvent vergibt bei einem Lesefehler still EventNumber 1 — doppelte Event-Nummern in DEX_Events

**Wo:** `services/EventService.ts:2479` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Organizer legt ein Event an, während SharePoint drosselt (429) oder DEX_Events kurz mit 500/503 antwortet — beides tritt in dieser App nachweislich auf. enResp.ok ist false, der Code läuft ohne Meldung weiter und schreibt EventNumber = 1. Ab da tragen zwei Events dieselbe Nummer. Folge für Nutzer: DEX_Participants schlüsselt über die Nummer, also zeigt „Meine Events" jeder Person, die für das alte Event #1 angemeldet ist, zusätzlich das neue Event #1 (MyEventsPage.tsx:449 `myNumSet.has(e.eventNumber)`) — inklusive QR-Code `DEX|1|mail` für den falschen Termin. getEventByEventNumber ($top=1) trifft beim Self-Check-in das falsche Event. Und deleteEvent räumt beim Löschen des einen Events die Register-Einträge des anderen mit weg und löscht dessen Dokumentordner Event_1_*.

**Ursache:** Der Fehlerfall ist nicht vom Normalfall unterschieden: `let nextEventNumber = 1` dient gleichzeitig als „Liste ist leer"-Startwert UND als stiller Fallback für „Abfrage fehlgeschlagen". Ein leeres bzw. nicht gelesenes Ergebnis wird als Aussage über die Daten behandelt (die in CLAUDE.md beschriebene Falle). Zusätzlich liegt zwischen Lesen und Schreiben der Nummer die komplette Subsite- und Listen-Anlage (oft > 30 Sekunden) — zwei parallel anlegende Organizer bekommen deshalb auch ohne Fehlerfall dieselbe Nummer.

**Fix-Vorschlag:** Bei !enResp.ok / Exception abbrechen statt zu raten: `throw new Error('Event-Nummer konnte nicht ermittelt werden — bitte erneut versuchen.')` (der Wizard zeigt Errors aus createEvent bereits an). Zusätzlich die Nummer erst unmittelbar VOR dem Item-Insert lesen und nach dem Insert einmal auf Eindeutigkeit prüfen (Filter EventNumber eq N) und bei Kollision neu vergeben — analog zum Post-Insert-Dedup der TeilnehmerID.

### 37. [hoch] deleteEvent löscht das Hauptevent auch dann, wenn Kind-Events nicht gelöscht werden konnten — verwaiste Sub-Events ohne jede Bedienoberfläche

**Wo:** `context/EventContext.tsx:3410` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Admin löscht ein Klammer-Event mit 19 Terminen. Jedes Kind bedeutet Subsite-Recycle + Register-Aufräumen + Item-Recycle; nach etwa einem Dutzend greift die SharePoint-Drosselung, eventService.deleteEvent liefert für die restlichen Kinder false (es wirft nicht — bei 429 auf dem Item-Recycle ist response.ok false). Der Rückgabewert wird nicht geprüft, die Schleife läuft durch, und die Klammer wird trotzdem gelöscht. Die übrigen Kind-Zeilen bleiben mit einem ParentEventId auf ein nicht mehr existierendes Item stehen. Alle Übersichten filtern `!e.parentEventId` (LandingPage.tsx:173/465/1149), das Organizer Center wählt ebenfalls nur Top-Level-Events — die Termine sind damit in der App unerreichbar, ihre Subsites, Anmeldungen und Outlook-Termine existieren aber weiter (Teilnehmer bekommen Erinnerungen zu einem Event, das niemand mehr verwalten kann).

**Ursache:** Zwei Annahmen: (1) eventService.deleteEvent wirft bei Fehlschlag — tut es nicht, es liefert false (dieselbe Bug-Klasse wie v29.48/v29.21); das catch ist toter Code. (2) `events` (Client-State) enthält ALLE Kinder — getEvents() lädt aber nur `$top=100` nach StartDate desc, und loadEvents lässt Events aus, deren Mapping scheitert (v9.41). Ein nicht geladenes Kind wird gar nicht erst zum Löschen versucht.

**Fix-Vorschlag:** Ergebnis auswerten und bei Fehlschlag abbrechen: `const ok = await eventService.deleteEvent(Number(child.id)); if (!ok) failed.push(child.title);` — bei failed.length > 0 das Parent NICHT löschen, sondern die Termine namentlich melden („bitte später erneut löschen"). Die Kinderliste zusätzlich serverseitig ermitteln (Filter `ParentEventId eq '<id>'` gegen DEX_Events, strikt gelesen) statt aus dem gekappten Client-State.

### 38. [hoch] deleteEvent räumt DEX_Participants erst NACH dem Recyceln der Subsite auf — mit Promise.all und verschluckten Fehlern (v29.3-Fix nur in deleteParticipantData angekommen)

**Wo:** `services/EventService.ts:3037` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer löscht ein abgesagtes Event mit 300 Anmeldungen über „Event löschen". Schritt 1 (Z. 2996) recycelt die Subsite sofort. Erst Schritt 3 räumt das Register auf: getAllParticipants() liest NICHT strikt (eine abgebrochene Seite kommt still als vollständige Liste zurück), danach starten alle 300 MERGE-Requests gleichzeitig, SharePoint drosselt, und jeder Fehlschlag verschwindet in `.catch(() => null)`. Es bleiben Dutzende Register-Verweise auf eine EventNumber, deren Teilnehmerliste nicht mehr existiert — und nachrechnen kann es niemand mehr, weil die Subsite weg ist. Der Rückgabewert ist trotzdem `true`.

**Ursache:** CLAUDE.md beschreibt genau dieses Muster als Ursache der 1045 „verwaisten" Verweise; v29.3 hat es aber nur in `deleteParticipantData` umgedreht (jetzt services/events/eventStats.ts:243-262: strikt lesen, sequentiell, Fehlerzähler, Abbruch vor der Löschung). Der zweite, häufiger benutzte Löschpfad `deleteEvent` steht unverändert auf dem alten Stand. Hinzu kommt, dass der Fehlerzähler dort selbst dann nichts brächte, solange removeParticipantEvent unbedingt `true` liefert (siehe Befund 1).

**Fix-Vorschlag:** Schritt 3 vor Schritt 1 ziehen und die Implementierung aus eventStats.deleteParticipantData übernehmen: `fetchAllParticipantsOrThrow()` statt getAllParticipants(), sequentielle Schleife mit `registryFailed`-Zähler, und bei registryFailed > 0 mit `false` abbrechen, BEVOR Subsite/Bild/Liste recycelt werden. Der Aufrufer wertet `false` bereits als „später erneut versuchen". Setzt Befund 1 voraus, sonst zählt der Zähler weiterhin null.

### 39. [hoch] ensureOrganizerPermissionsMulti zählt fehlgeschlagene Rechtevergaben als Erfolg — „Organizer-Berechtigungen reparieren" meldet Vollzug, ohne etwas geändert zu haben

**Wo:** `services/EventService.ts:3717` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Co-Organizerin Carolin sieht bei einem Klammer-Event mit 19 Terminen überall „0 Teilnehmer" (der v30.37-Befund). Ein Organizer — nicht Site-Admin — klickt in AdminPage die Kachel „Organizer-Berechtigungen reparieren". Auf den Sub-Event-Subsites hat er selbst kein „Berechtigungen verwalten" (oder eine Subsite wurde vom Löschkonzept recycelt → 404). Jeder addroleassignment-POST antwortet 403/404. Die UI meldet trotzdem „3 Person(en) auf 20 Liste(n) berechtigt" (AdminPage.tsx:9270), der Admin hakt das Ticket ab — und Carolin sieht weiterhin auf jedem Termin 0 Teilnehmer, weil getAllRegistrations bei 403 `[]` liefert.

**Ursache:** this._post (Zeile ~8735) gibt bei HTTP-Fehlern eine Response mit ok=false zurück und wirft nur bei Netzwerkfehlern. Der Code prüft response.ok nirgends, das catch fängt also fast nie — der Kommentar „idempotent, Person hatte schon Rechte" beschreibt eine Annahme, die der Code nicht prüft. Der zweite POST (Teilnehmerliste) wertet gar nichts aus. Nebenbefund derselben Stelle: Der Listen-POST ist hart auf REG_LIST_NAME ('Teilnehmer') verdrahtet — Legacy-Events mit abweichendem RegistrationListName treffen dadurch immer ins Leere (404).

**Fix-Vorschlag:** Response auswerten: `const r = await this._post(...); if (r.ok || r.status === 200) result.grants++; else result.failed.push({site, userId, status: r.status});` und das Ergebnis um `failed` erweitern, damit AdminPage/EventContext den Fehlschlag benennen können statt „N Personen berechtigt" zu melden. Für Bestandslisten den tatsächlichen RegistrationListName des Events statt REG_LIST_NAME nutzen.

### 40. [hoch] findOrphanSubsites hält bei leerem Event-Read ALLE Subsites für verwaist — inklusive Löschknopf

**Wo:** `services/events/permissionsAudit.ts:326` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Admin klickt im Admin Hub „Verwaiste Subsites suchen". Der DEX_Events-Read scheitert einmalig (429 Drosselung — GETs laufen laut v29.50 bewusst OHNE Retry — oder 403/5xx). `getAllEventsForKpi` bricht die Schleife ab und liefert `[]`. `referenced` bleibt leer, damit gilt JEDE der bis zu 500 gefundenen Subsites als verwaist. Die Ergebnisliste zeigt alle Live-Event-Subsites mit „Endgültig löschen"-Knopf (AdminHubPage.tsx:258 → deleteSubsiteWeb) — ein Klick entfernt eine aktive Teilnehmerliste unwiderruflich.

**Ursache:** Genau die in CLAUDE.md dokumentierte Falle („Ein leeres Ergebnis ohne geprüften Status ist keine Aussage über die Daten, sondern über gar nichts"). `getAllEventsForKpi` (emailTemplatesList.ts) macht `if (!resp.ok) break;` und gibt die bis dahin gelesenen Zeilen zurück — bei Fehler auf der ersten Seite also `[]`. Die beiden anderen Aufrufer derselben Methode kennen das Problem und schützen sich (`getKpiTotals` und `recomputeEventKpiOnly` haben beide `if (all.length === 0) return null;`) — `findOrphanSubsites` hat diesen Wächter als einziger nicht.

**Fix-Vorschlag:** Wächter analog zu getKpiTotals einziehen: `if (events.length === 0) { throw new Error('DEX_Events konnte nicht gelesen werden — Prüfung abgebrochen'); }` bzw. ein Ergebnisfeld `incomplete: true` setzen, das die UI als Fehler statt als Ergebnisliste rendert. Sauberer: `getAllEventsForKpi` einen `onHttpError`-Rückruf spendieren (wie `getAllRegistrations` ihn seit v29.3 hat) und den Scan nur mit vollständig gelesener Event-Liste zulassen.

### 41. [hoch] getRoles() ohne $top — ab 100 Rollen-Zeilen verlieren Organizer ihre Rolle

**Wo:** `services/SharePointService.ts:588` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** DEX_Roles enthält 130 Zeilen. SharePoint liefert ohne `$top` die Standard-Seite von 100 Items, sortiert nach `Role,UserName`. Eine Organizerin namens „Zimmermann, Petra" fällt hinter Position 100 und ist im Ergebnis nicht enthalten. RoleContext.initRoles findet sie per `spRoles.find(...)` nicht → `setCurrentUserRole('User')`. Sie sieht kein Organizer Center, keine „Event anlegen"-Kachel und kann ihre eigenen Events nicht mehr bearbeiten — obwohl ihr Eintrag in DEX_Roles korrekt steht. Zweitwirkung: `addRole` dedupliziert über den (abgeschnittenen) `roles`-State, legt also für dieselbe Person eine zweite Zeile an — genau der Zustand, den der v29.63-Kommentar verhindern soll.

**Ursache:** SharePoint REST paginiert Item-Abfragen standardmäßig bei 100 und liefert den Rest nur über `odata.nextLink`. Die Methode setzt weder `$top` noch folgt sie dem nextLink; `data.value` wird 1:1 zurückgegeben. Jede andere Item-Abfrage der Service-Schicht setzt ein `$top` (DEX_Events durchgehend, `getRoleEmails`/`getRoleRecipients` in organizer.ts sogar `$top=5000`) — `getRoles` ist die einzige Ausnahme, und ausgerechnet an ihr hängt die gesamte Rechteermittlung.

**Fix-Vorschlag:** `&$top=5000` an beide Abfrage-Varianten anhängen (wie in organizer.getRoleEmails) und zusätzlich dem `odata.nextLink` folgen, solange einer kommt — analog zum Paging-Muster in `getAllEventsForKpi`. Ein Kappen ohne Fehlermeldung darf es bei einer Rechteliste nicht geben.

### 42. [hoch] onHttpError-Auswertung ignoriert 429/5xx — Drosselung wird zu „0 Teilnehmer" ohne Warnung

**Wo:** `components/AdminPage.tsx:670` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Klammer-Event mit 19 Terminen à ~300 Anmeldungen; die Listen werden in der Schleife sequentiell gelesen. SharePoint drosselt beim 12. Termin und antwortet mit HTTP 429 (oder 500/503/504). `EventService.getAllRegistrations` (Zeile 6024) ruft `onHttpError(429)`, bricht die Paging-Schleife ab und liefert die bis dahin gelesenen Zeilen — bei einem Fehler auf der ersten Seite also `[]`. Weil 429 in der `if`-Kette fehlt, wird NICHTS in `denied` geschrieben: Das rote Banner „Kein Zugriff auf N Teilnehmerliste(n)" erscheint nicht, die Summenzeile der Matrix zeigt für diesen Tag „0", die KPI-Kachel „Angemeldet" zählt zu niedrig, und die Geist-Box aus Fund 1 schlägt zusätzlich an. Für den Organizer sieht ein voller Termin aus wie ein leerer.

**Ursache:** Die Statusliste ist eine Aufzählung von Berechtigungs-/Nicht-vorhanden-Fällen, obwohl der Rückruf per Definition bei JEDEM nicht-ok-Status feuert. Die Referenz-Implementierung im selben Repo macht es richtig: `analyzeRegistryAgainstLists` (services/EventService.ts, Zeile 1068-1072) behandelt nur 404/410 als eindeutige Aussage und ALLES andere (403, 429, 500, Netz) als „übersprungen, sagt nichts über den Inhalt".

**Fix-Vorschlag:** Jeden Rückruf-Aufruf als „nicht lesbar" werten, nicht nur vier Codes: `const regs = await getAllRegistrations(ch.id, () => { denied.push(ch.title || ch.id); });` — und, wenn zwischen „Liste weg" und „gerade nicht lesbar" unterschieden werden soll, den Status wie im EventService mitführen und nur 404/410 als „Liste gelöscht" auszeichnen. Dasselbe gilt für `handleSelectEvent` (Zeile 3466-3470), das dieselbe zu enge Liste benutzt.

### 43. [hoch] reactivateRegistration schreibt StarterType/PreferredStarterType nicht — bei geteilten Kapazitäten wird der Platz in Gruppe B reserviert, die Zeile bleibt in Gruppe A

**Wo:** `services/EventService.ts:4649` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Split-Event (Durchstarter 50 / Funstarter 50). Person war Durchstarter und hat sich abgemeldet — cancelRegistration setzt nur Status/CancellationDate/TeilnehmerID, StarterType bleibt 'Durchstarter' auf der Zeile stehen. Zwei Wochen später meldet sie sich erneut an und wählt diesmal Funstarter. EventContext.tsx:737 reserviert per reserveSeat einen FUNSTARTER-Platz (SeatsTakenFun +1), setzt status='Angemeldet' und ruft in Zeile 810 `reactivateRegistration(...)` — ohne effectiveStarterType/preferredStarterType, die die Methode gar nicht entgegennimmt. Ergebnis: Die Zeile zählt weiter als Durchstarter (Teilnehmerliste, starterCounts, runManualPromote-Plan), der Zähler hat aber einen Funstarter-Platz verbraucht. Beim nächsten privilegierten syncSeatsToActiveCount werden die Ist-Werte aus den Zeilen geschrieben: Durchstarter steht auf 51/50 (überbucht), Funstarter hat einen Phantom-Platz frei. Die Person erscheint am Eventtag in der falschen Gruppe.

**Ursache:** registerForEvent (Z. 3839ff) nimmt starterType/preferredStarterType entgegen und schreibt sie in den Insert; reactivateRegistration ist der zweite, ältere Anmeldepfad und hat diese Parameter nie bekommen. EventContext.tsx:807-820 verzweigt zwischen beiden, reicht die Gruppe aber nur im else-Zweig durch. Da die Gruppenwahl auf der Anmeldeseite für jede nicht-aktive Person offen ist, ist der Wechsel beim Wiederanmelden ein normaler Bedienweg, kein Sonderfall.

**Fix-Vorschlag:** reactivateRegistration um `starterType?: string, preferredStarterType?: string` erweitern und in den MERGE-Body aufnehmen (bei status==='Warteliste' StarterType bewusst leeren und nur PreferredStarterType setzen — dasselbe Muster wie switchSplitGroup Z. 6931 und resolveOverbookToWaitlist). In EventContext.tsx:810 `effectiveStarterType`/`preferredStarterType` mit übergeben. Zusätzlich prüfen, ob bei einem Gruppenwechsel der alte Gruppenzähler additiv freigegeben werden muss (analog switchSplitGroup) oder bewusst dem Reconcile überlassen wird.

### 44. [hoch] reconcileCounters schließt genau die Split-Kapazitäts-Events aus — Filter über maxParticipants, wo die Kapazität in den Split-Feldern steht

**Wo:** `context/EventContext.tsx:488` · Quelle: Audit B (nach dem Umbau)

**Szenario:** B2Run-Event mit geteilten Kapazitäten (durchstarterCapacity 200, funstarterCapacity 150, `maxParticipants` = 0 — genau so legt der Wizard es an). Ein Admin startet die App: `reconcileCounters()` läuft, aber `(e.maxParticipants || 0) > 0` ist für dieses Event `false` → das Event ist NIE Ziel des Abgleichs. Damit läuft der einzige App-seitige Selbstheilungs-Lauf für `SeatsTaken`/`SeatsTakenDurch`/`SeatsTakenFun` an den Events vorbei, die als einzige ZWEI Zähler synchron halten müssen. Konkrete Folge: Ein Teilnehmer meldet sich selbst ab; `releaseSeatAfterCancel` findet `WaitlistTaken` = -1 („unbekannt", Bestands-Event ohne je gepflegtes Feld) und zählt fail-closed NICHTS herunter (seats.ts:511-515). Der Gruppenzähler bleibt dauerhaft zu hoch; die nächste Anmeldung für diese Gruppe bekommt von `reserveSeat` 'full' und landet auf der Warteliste, obwohl ein Platz frei ist — und die App repariert das von sich aus nie, weil das Event aus der Zielmenge gefiltert ist. Dass `isSplit` drei Zeilen tiefer berechnet wird, obwohl kein Split-Event den Filter passieren kann, zeigt die Absicht: der Lauf SOLLTE Split-Events erfassen.

**Ursache:** Der Filter benutzt `maxParticipants` als „hat dieses Event überhaupt eine Kapazität?"-Prüfung. Bei geteilten Kapazitäten ist `maxParticipants` per Konvention 0 (CLAUDE.md: „Bei geteilten Kapazitäten ist `maxParticipants` 0 … Wer `maxParticipants` als Obergrenze prüft, prüft dort **gar nichts**"). Die Split-Felder werden im Filter nicht berücksichtigt.

**Fix-Vorschlag:** Den Filter auf eine effektive Kapazität umstellen, z.B. `const effCap = (e.maxParticipants || 0) > 0 ? (e.maxParticipants || 0) : ((e.durchstarterCapacity || 0) + (e.funstarterCapacity || 0));` und dann `effCap > 0` prüfen. Das bereits vorhandene `isSplit` unterhalb bleibt unverändert und wird dadurch endlich wirksam.

### 45. [hoch] reloadSubEventRegs spiegelt den v30.37-Fix nicht: gesperrte Termine werden nach einer Abmeldung still zu 0

**Wo:** `components/admin/logic/createKlammerActions.ts:469` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Klammer-Event (`subEventsOnlyMode`) mit mehreren Terminen. Der Organizer meldet in der konsolidierten Matrix eine Person von einem Termin ab; danach läuft `reloadSubEventRegs()`. Ist einer der Termine für diese Person nicht lesbar (403 auf der Sub-Event-Subsite — genau der v30.37-Fall einer Co-Organizerin) oder antwortet gerade mit 429, wird `map[ch.id] = []` gesetzt: die Spalte dieses Termins fällt in der Matrix auf „—"/0, die KPI-Summen sinken. Der v30.37-Warnhinweis „gesperrte Termine" erscheint dabei NICHT, weil `setDeniedSubEventLists` hier gar nicht aufgerufen wird — er behält den Stand vom Initial-Load. Der Organizer sieht also unmittelbar nach der eigenen Aktion einen leeren Termin ohne jede Erklärung.

**Ursache:** Der Kommentar darüber sagt „Spiegelt den Lade-Effekt von oben" — der Lade-Effekt in `AdminPage.tsx:392-418` wurde in v30.37 auf `onHttpError` + `denied[]` umgestellt, diese Kopie nicht. Zusätzlich macht `catch { map[ch.id] = [] }` auch aus echten Ausnahmen ein „null Teilnehmer". Das ist die in CLAUDE.md fett markierte Regel: ein leeres Ergebnis ohne geprüften Status ist keine Aussage über die Daten.

**Fix-Vorschlag:** Die Duplikation auflösen und stattdessen den vorhandenen Effekt auslösen: `setSubRegReloadTick(t => t + 1)` (existiert bereits und ist in derselben Datei Z.452 in Gebrauch). Wenn die Funktion bleiben soll: denselben `denied`-Sammler wie der Effekt führen, `setDeniedSubEventLists(denied)` mitschreiben und bei Fehler den bisherigen Eintrag beibehalten statt `[]`.

### 46. [hoch] reloadSubEventRegs verschluckt Fehler komplett und lässt das Zugriffs-Banner stehen

**Wo:** `components/AdminPage.tsx:1882` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Organizer meldet über den konsolidierten Abmelde-Dialog eine Person von 6 Terminen ab. `runDeregModal` schreibt dabei pro Termin `cancelRegistration` + ChangeLog + Mail-Queue + Outlook-Queue + `queueIDReorder` — also ein Schwall Schreibzugriffe. Direkt danach ruft Zeile 2069 `reloadSubEventRegs()`. Drosselt SharePoint jetzt (429) oder ist eine Liste nicht lesbar, liefert `getAllRegistrations` `[]` OHNE zu werfen, `map[ch.id]` wird `[]`, und `deniedSubEventLists` bleibt unverändert. Ergebnis: Die Matrix zeigt für betroffene Termine 0 Anmeldungen, die KPI-Kacheln fallen entsprechend — und zwar ohne jeden Hinweis, unmittelbar nach einer Abmeldeaktion. Der Organizer liest das als „ich habe gerade alle abgemeldet". Umgekehrt bleibt ein vorher gesetztes `deniedSubEventLists` stehen, obwohl der neue Lauf sauber war — das Banner lügt dann in die andere Richtung.

**Ursache:** Zweite, schwächere Kopie des Lade-Pfads: Der Effekt in Zeile 647-687 hat seit v30.37 `onHttpError` + `denied` + `isLoadingSubEventRegs`, diese Handkopie hat nichts davon. Sie stammt aus v19.30 und wurde beim v30.37-Fix nicht mitgezogen.

**Fix-Vorschlag:** Die Kopie ersatzlos streichen und stattdessen den vorhandenen Reload-Trigger benutzen — `setSubRegReloadTick(t => t + 1)` (wird in Zeile 1710 und 1867 schon genau dafür verwendet). Dann laufen Fehlerbehandlung, Denied-Liste und Lade-Anzeige nur an EINER Stelle. Falls die Funktion wegen des `await` bleiben muss, mindestens `onHttpError` durchreichen und `setDeniedSubEventLists` mitschreiben.

### 47. [hoch] removeParticipantEvent meldet Erfolg, obwohl der MERGE abgelehnt wurde — der Fail-Closed-Riegel aus v29.3 greift dadurch nie

**Wo:** `services/EventService.ts:1236` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Das 3-Monats-Löschkonzept läuft (deleteParticipantData, Zeile 3331: `const ok = await this.removeParticipantEvent(...).catch(() => false); if (!ok) registryFailed += 1;`). Bei 300 Teilnehmern drosselt SharePoint nach einigen Dutzend sequentiellen MERGEs mit HTTP 429, oder ein einzelner MERGE scheitert mit 403/500. _merge liefert dafür eine Response mit ok=false und wirft nicht — removeParticipantEvent gibt trotzdem `true` zurück. registryFailed bleibt 0, die Prüfung „bei Fehlern GAR NICHT löschen" schlägt nicht an, und die Subsite wird direkt danach recycelt. Ergebnis: genau der Zustand, den v29.3 verhindern wollte — Verweise auf das Event bleiben personenbezogen in DEX_Participants stehen, die Teilnehmerliste ist unwiderruflich weg, nachrechnen geht nicht mehr. Die Register-Prüfung meldet sie später als „Verweis ohne Zeile".

**Ursache:** Der Rückgabewert von _merge wird nicht ausgewertet. Die Funktion kann nur zwei Dinge als Fehler erkennen: Teilnehmer nicht gefunden (getParticipantByEmail → null) oder eine geworfene Exception (nur bei Netzwerkfehlern). Jeder HTTP-Fehlerstatus — also der praktisch relevante Fall — kommt als Erfolg zurück. Der Aufrufer aus v29.3 ist korrekt geschrieben, verlässt sich aber auf eine Zusage, die diese Methode nicht einhält.

**Fix-Vorschlag:** `const resp = await this._merge(...); return resp.ok;` — und im Fehlerfall Status loggen. Damit greift der bestehende registryFailed-Zähler in deleteParticipantData wieder, und deleteEvent/AdminPage bekommen ebenfalls eine ehrliche Antwort.

### 48. [hoch] sendFAMail meldet Erfolg, obwohl der `_billing`-Patch fehlschlug — Versandhistorie weg, Auto-Mail wird wiederholt

**Wo:** `context/actions/billing.ts:113` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Event trägt ein eingebettetes Mail-Logo (`_eventLogo`, Base64) und schon ein paar F&A-Mails mit Volltext im Log. Beim Versand der Teilnehmerliste (300 Personen → ~50 KB HTML-Body, der als `entry.body` mitgespeichert wird) überschreitet `EmailTemplateOverrides` die 1,9-MB-Grenze; `patchEventOverridesValueEx` liefert `{ok:false,status:413}` (eventAssets.ts:222). Dasselbe passiert bei 429/5xx nach dem einen Retry. Die Mail ist da aber längst in DEX_Emails eingereiht. Der Nutzer sieht „Teilnehmerliste wurde an F&A gesendet." (BillingActionPanel:89), in der Versandhistorie steht nichts, `listSentAt` fehlt — und beim nächsten Admin-Boot sendet `maybeSendBillingAutoMails` dieselbe Mail erneut an F&A, weil der einzige verbleibende Schutz `hasQueuedEmail` auf der TRANSIENTEN DEX_Emails-Queue arbeitet (nach Versand + Archivierung leer, ausdrücklich dokumentiert in services/events/notificationLogs.ts:85-88). Ein Admin, der auf „Senden" doppelt klickt, weil die Historie leer bleibt, verschickt jedes Mal eine weitere Mail.

**Ursache:** Der Rückgabewert ist an den Queue-Erfolg gekoppelt, nicht an die Persistenz der Nebenbuchhaltung. `patched` wird ausgewertet, um den lokalen State nachzuziehen, aber nicht, um den Aufrufer zu informieren. Damit gilt in einer revisionssicheren Historie ein Versand als protokolliert, der es nicht ist — und der Doppelversand-Schutz hängt an genau diesem nicht geschriebenen Stempel.

**Fix-Vorschlag:** `return patched ? { ok: true } : { ok: false, reason: 'stamp-failed' };` und in `BillingActionPanel.doSend` sowie in `maybeSendBillingAutoMails` diesen Fall eigens behandeln: „Die Mail wurde versendet, der Vermerk in der Historie konnte NICHT gespeichert werden — bitte nicht erneut senden, bis der Vermerk steht." Sinnvoll zusätzlich: `patchEventOverridesValueEx` benutzen und den Grund (413 vs. 429) durchreichen, und den `body` bei sehr großen Mails gar nicht erst ins Log legen (der Größen-Deckel in `trimBillingLog` zählt Einträge, nicht Bytes).

### 49. [hoch] setWaitlistPosition kennt die getrennten Gruppen-Wartelisten nicht — „Platz ändern" verschiebt in die falsche Liste

**Wo:** `services/events/seats.ts:78` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Split-Kapazitäts-Event (B2Run) mit getrennten Wartelisten. Durchstarter-Warteliste: 10 Personen (TeilnehmerID 101–110), Funstarter-Warteliste: 5 Personen (TID 111–115). Der Organizer öffnet im Organizer Center die Tabelle „Warteliste Funstarter", sieht dort Person X auf **Platz 4**, klickt „Platz ändern" und trägt **2** ein (Wunsch: X soll Zweite der Funstarter-Warteliste werden). setWaitlistPosition baut aber EINE globale Warteliste [101…115], findet X bei fromIdx=13 und setzt toIdx=1 — X landet auf globaler Position 2, also zwischen dem ersten und zweiten DURCHSTARTER-Wartenden. Ergebnis: X ist innerhalb der Funstarter-Warteliste auf Platz 1 (nicht 2), alle zehn Durchstarter-Wartenden bekommen neue TeilnehmerIDs und rücken je eine Position nach hinten, und die Erfolgsmeldung nennt „jetzt Platz 2 (vorher 14)" — Zahlen, die der Organizer nirgends gesehen hat. Weil das Modal `max = wlPosModal.total = regs.length` (= 5) erzwingt, sind für die Funstarter-Liste überhaupt nur globale Positionen 1–5 erreichbar: JEDE Eingabe schiebt die Person an den Anfang ihrer Gruppe und mischt gleichzeitig die fremde Gruppe um. Die Funktion ist für die zweite Gruppe damit vollständig unbrauchbar.

**Ursache:** `setWaitlistPosition` liest nur `$select=Id,Status,TeilnehmerID` und filtert ausschließlich auf `Status === 'Warteliste'` — `PreferredStarterType` wird weder abgefragt noch berücksichtigt. Die UI (waitlistDurch/waitlistFun/waitlistUnassigned in AdminPage.tsx:1731-1739, Rang aus waitlistTruePos, das je Gruppe rankt) und `promoteFirstWaitlistItem` (filtert `PreferredStarterType eq …`, waitlist.ts:78-81) rechnen dagegen PRO GRUPPE. `targetPosition` wird als Index in die globale Liste interpretiert, kommt aber als Rang innerhalb der Gruppe herein.

**Fix-Vorschlag:** In `setWaitlistPosition` die Gruppe mitlesen und mitfiltern: `$select=Id,Status,TeilnehmerID,PreferredStarterType`, einen optionalen Parameter `group?: string` ergänzen und die Zielliste als `allItems.filter(i => i.Status === 'Warteliste' && (!group || (i.PreferredStarterType || '') === group))` bilden. Die Umnummerierung muss dann die ANDEREN Wartelisten-Zeilen in ihrer bisherigen TeilnehmerID-Reihenfolge unangetastet mitschreiben (globale Sequenz weiterhin 1..N Aktive, dann alle Wartenden — nur die Reihenfolge INNERHALB der Zielgruppe ändert sich). Aufrufer: `WaitlistPositionModal` reicht `reg.PreferredStarterType` durch, wenn `selectedEvent.durchstarterCapacity/funstarterCapacity` gesetzt sind und `splitSharedWaitlist` NICHT aktiv ist.

### 50. [hoch] „Alle markieren" ignoriert die aktiven Filter — Massenzuordnung trifft unsichtbare Zeilen

**Wo:** `components/HotelPlanningPanel.tsx:1705` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Event mit 300 aktiven Anmeldungen, 270 davon sind bereits auf „Hotel A"/„Hotel B" verteilt. Der Organizer stellt den Filter auf „Nur ohne Hotel" (oder tippt einen Suchbegriff), sieht 30 Zeilen, klickt den Haken in der Kopfzeile — optisch sind genau diese 30 markiert — und klickt dann in der Markierungs-Leiste auf „→ Hotel C". Ergebnis: ALLE 300 Personen werden geschrieben, die 270 bereits zugeordneten werden auf Hotel C umgebucht. Dasselbe gilt für „Zuordnung aufheben" (Zeile 1673) und für die Zeitraum-Knöpfe. Ist die Freigabe für Teilnehmer an, sehen alle 300 die falsche Umbuchung sofort unter „Meine Events".

**Ursache:** Die drei Filter (filterHotel, hideNoWish, search) stehen als Inline-Kette direkt im <tbody> (Zeilen 1786–1794). Es gibt keine gemeinsame Variable für „die gerade sichtbaren Zeilen": Der Kopf-Haken und alle Massenaktionen arbeiten auf dem ungefilterten `people`-Memo (Zeile 866). Auch `checked` vergleicht gegen `people.length`, ist also nach dem Klick korrekt „an", obwohl viel mehr markiert ist als zu sehen. Der einzige Hinweis ist die Zahl in „N markiert:".

**Fix-Vorschlag:** Die gefilterte + sortierte Liste EINMAL als Memo ziehen (`const visiblePeople = React.useMemo(...)`) und sowohl im <tbody> als auch im Kopf-Haken verwenden: `onChange={e => setSelected(e.target.checked ? new Set(visiblePeople.map(p => p.Id)) : new Set())}` und `checked={selected.size > 0 && selected.size === visiblePeople.length}`. Zusätzlich in `assignSelectedTo`/`applyStayToSelected` die Zeilen aus `visiblePeople` statt aus `people` ziehen, damit eine alte Markierung nach einem Filterwechsel nicht nachwirkt.

### 51. [hoch] „Kommunikation auf alle Sub-Events übernehmen" überspringt genau den gerade offenen Termin-Reiter

**Wo:** `components/EventCreationPage.tsx:7796` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Event mit z.B. 5 Terminen. Der Organizer steht in Schritt „Kommunikation" auf dem Reiter von Termin 3 (Scope-Leiste → activeCommTabIdx = 3) und klickt dort den Knopf „Kommunikation des Haupt-Events auf alle Sub-Events übernehmen". Der Erfolgs-Hinweis meldet „gilt jetzt für alle 5 Termine". Danach speichert er direkt. Ergebnis: Termine 1, 2, 4 und 5 bekommen die Hauptevent-Kommunikation (Mail-Sprache, Logo, Outlook-Text, Überschriften, Betreff, alle Mail-Schalter, emailTemplateOverrides), Termin 3 behält seine alten Werte — ohne jede Rückmeldung. Derselbe Effekt beim Reiterwechsel statt Speichern.

**Ursache:** Die Kommunikationsfelder liegen NICHT laufend im Draft (CLAUDE.md), sondern im gespiegelten UI-State (emailLanguage, outlookBody, emailLogoPreview, emailTemplateOverrides …). `applyCommToAllSubEvents` schreibt die Hauptevent-Werte per `setSubEvents` in ALLE Draft-Slots, lässt den sichtbaren UI-State aber unangetastet — der zeigt weiter die alten Werte von Termin 3. Der nächste `flushActiveCommTabToState()` (in `attemptSubmit` die allererste Anweisung, ebenso in `switchCommTab`) schreibt genau diesen stale UI-State wieder in `subEventsRef.current[activeCommTabIdx-1]` und überschreibt damit die eben kopierten Werte. Der Knopf ist in `wizard/steps/CommunicationStep.tsx` (Block ab `{subEvents.length > 0 && (`, Zeile 113) NICHT auf `activeCommTabIdx === 0` gegated — der Fall ist also der Normalfall, sobald jemand von einem Termin-Reiter aus klickt.

**Fix-Vorschlag:** Am Ende von `applyCommToAllSubEvents` den sichtbaren Reiter neu laden, damit UI-State und Draft wieder übereinstimmen: nach dem `setSubEvents(...)` bei `activeCommTabIdx > 0` die UI-States aus `src` setzen (setEmailLanguage(src.emailLanguage), setEmailLogoPreview, setOutlookLogoPreview, setOutlookBody, setOutlookHeading, setOutlookSubheading, setOutlookSubject, setDisableEmails, setDisableRegistrationEmail, setDisableCancellationEmail, setAutoDeregisterOnDecline, setInactiveHandling, setDisableOutlook, setEmailTemplateOverrides({...src.emailTemplateOverrides})). Alternativ (weniger Code, gleiche Wirkung): vor dem Kopieren `switchCommTab(0)` rufen und den Knopf nur auf Reiter 0 anbieten.

### 52. [hoch] „Kommunikation des Haupt-Events auf alle Sub-Events übernehmen" wird für den gerade offenen Termin-Reiter sofort wieder zurückgerollt

**Wo:** `components/EventCreationPage.tsx:8119` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Klammer-Event mit 5 Terminen, Schritt 6 (Kommunikation). Der Organizer schaltet über die Scope-Leiste auf „Tag 3", schaut dort die Mailtexte an, und klickt danach den Knopf „Kommunikation des Haupt-Events auf alle Sub-Events übernehmen" (der Block ist nur mit `subEvents.length > 0` gegattert, siehe components/wizard/steps/CommunicationStep.tsx:110/171 — er ist auf einem Sub-Reiter sichtbar). Die Erfolgsmeldung sagt „gilt jetzt für alle 5 Termine". Beim Speichern (oder schon beim nächsten Reiterwechsel) läuft `flushActiveCommTabToState()`, schreibt die noch unveränderten UI-States (outlookBody, outlookHeading, outlookSubject, emailLogoBase64, disableEmails, emailTemplateOverrides …) in `subEventsRef.current[2]` — Tag 3 hat danach wieder seine ALTEN Texte/Logos, alle anderen vier die neuen. Der Organizer sieht das erst an der abweichenden Mail bzw. am abweichenden Outlook-Termin bei den Teilnehmern von Tag 3.

**Ursache:** Die Kommunikationsfelder liegen nicht laufend im Draft, sondern im UI-State und werden nur beim Reiterwechsel gespiegelt (CLAUDE.md-Falle „Kommunikationsfelder der Sub-Events liegen nicht laufend im Draft"). `applyCommToAllSubEvents` flusht zwar korrekt VOR dem Kopieren, schreibt danach aber nur in den `subEvents`-State und lässt die UI-States unberührt. Damit hält die UI für den aktiven Sub-Reiter einen Stand, der dem frisch geschriebenen Draft widerspricht — und der nächste Flush gewinnt. `applySubTransfer` hat das Problem nicht, weil dort die Quelle immer der aktive Reiter ist und die Ziele ihn ausschließen (`targets: … filter(i => i !== activeCommTabIdx - 1)`).

**Fix-Vorschlag:** Nach dem `setSubEvents(...)` in `applyCommToAllSubEvents` die Step-6-UI-States neu aus `src` laden, wenn `activeCommTabIdx > 0` — also dieselben Setter wie im Lade-Zweig von `switchCommTab` (setEmailLanguage, setEmailLogoPreview, setOutlookLogoPreview, setOutlookBody, setOutlookHeading, setOutlookSubheading, setOutlookSubject, setDisableEmails, setDisableRegistrationEmail, setDisableCancellationEmail, setAutoDeregisterOnDecline, setInactiveHandling, setDisableOutlook, setEmailTemplateOverrides mit geklontem Objekt). Alternativ vor dem Kopieren `switchCommTab(0)` rufen und den Scope danach über `setScope(0)` konsistent halten — dann ist der aktive Slot nie ein Ziel.

### 53. [hoch] „Offene Teams" liest die item-level-gesicherte Teilnehmerliste — die Box bleibt für Teilnehmer immer leer

**Wo:** `context/EventContext.tsx:2249` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Organizer aktiviert bei einem Team-Event „Offene Slots öffentlich sichtbar" (`teamOpenSlotsVisible`). Ein noch nicht angemeldeter Teilnehmer öffnet die Anmeldeseite. `RegistrationPage.tsx:701` ruft `listOpenTeamsForEvent(event.id)`; die Teilnehmerliste läuft mit ReadSecurity=2 („nur eigene Elemente", gesetzt in `services/events/subsiteProvisioning.ts:419`), der Nutzer hat in dieser Liste noch gar keine Zeile → `all` = `[]` → `byTeam` leer → Rückgabe `[]`. Die Render-Bedingung `RegistrationPage.tsx:2441` (`openTeamsLoaded && openTeams.length > 0 && !parentAlreadyRegistered`) ist damit nie erfüllt: Die Box mit den beitrittsfähigen Teams erscheint nie, obwohl offene Teams existieren. Der Organizer sieht sie in der Vorschau (Full Control) und hält das Feature für funktionsfähig.

**Ursache:** `getAllRegistrations` liefert HTTP 200 mit security-getrimmtem Ergebnis — ein leeres Ergebnis ist hier keine Aussage über die Daten (genau der Merksatz aus CLAUDE.md). `services/events/seats.ts:231-247` dokumentiert dieselbe Regression bereits (v27.10) und prüft dort gegen den ungefilterten `ItemCount`; `listOpenTeamsForEvent` hat diese Absicherung nie bekommen und ist ausgerechnet der einzige Pfad, der ausschliesslich von NICHT angemeldeten Personen aufgerufen wird.

**Fix-Vorschlag:** Die Team-Belegung nicht aus der Teilnehmerliste ableiten. Entweder (a) die offenen Teams über eine für alle lesbare Quelle bereitstellen (Analog zu `DEX_TeilnehmerCounter` / `getCounterStats` eine Team-Zeile je Event pflegen), oder (b) `listOpenTeamsForEvent` wie `getActiveCounts` gegen `getListItemCount(subsiteUrl, REG_LIST_NAME)` verifizieren und bei unvollständiger Sicht einen erkennbaren Fehler werfen, den die Anmeldeseite als Hinweis („Teams konnten nicht geladen werden") rendert — statt „es gibt keine offenen Teams" zu behaupten.

### 54. [hoch] „Organizer werden"-Freigabe stuft eine F&A-Person still auf Organizer herab

**Wo:** `components/OrganizerRequestsBanner.tsx:107` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Eine Person hat die Rolle **F&A**. Ein Organizer trägt sie in Schritt „Team" als Co-Organizer ein und speichert. `wizardSubmit` ruft `requestCoOrganizerApprovals` (context/actions/organizerRoles.ts:77); dessen `elevated`-Menge kommt aus `getRoleEmails('Organizer')` + `getRoleEmails('Admin')` und enthält F&A NICHT → es wird ein „Organizer werden"-Antrag angelegt. Der Admin sieht das Banner, klickt „Freigeben". Der v29.63-Schutz greift nicht (F&A fehlt in `alreadyEntitled`), also läuft `addRole(mail, name, 'Organizer', …)`; dort findet `addRole` die vorhandene Zeile mit Rolle 'F&A' ≠ 'Organizer' und ruft `updateRole(existing.id, 'Organizer')` (RoleContext.tsx:233). Ergebnis: Die Rolle F&A ist überschrieben. Die Person verliert das F&A Center (`isFA` wird false) und den Abrechnungs-Schritt im Wizard (`canEditBilling` liefert im Piloten `FA_BILLING_STEP_FOR_ORGANIZERS === false`). Weder Admin noch Betroffene bekommen einen Hinweis — der Admin liest nur „Freigegeben". Nebenbefund derselben Ursache: `getRoleEmails('Admin')` deckt über `roleFilter` nur 'Admin'/'SuperAdmin' ab, nicht 'IT-Admin' — jeder als Co-Organizer benannte IT-Admin erzeugt bei JEDEM Save einen überflüssigen Antrag samt Admin-Mail.

**Ursache:** Der Personenkreis „hat schon Organizer-Rechte" ist an drei Stellen unabhängig voneinander ausformuliert: `roleFilter` in services/events/organizer.ts (kennt nur die Legacy-Namen), `elevated` in organizerRoles.ts:82-84 und `alreadyEntitled` in OrganizerRequestsBanner.tsx:107. Als v30.60 F&A zu einem Organizer-Superset machte (`isOrganizer = 'Organizer' || 'F&A' || isAdmin` in RoleContext.tsx:341), wurde nur die eine Ableitung im RoleContext nachgezogen. Der Kommentar direkt über `alreadyEntitled` benennt genau diese Gefahr („das wäre eine Herabstufung durch die Hintertür") — nur eben für Admin/IT-Admin, nicht für die neue Rolle.

**Fix-Vorschlag:** Eine einzige exportierte Ableitung einführen, z.B. `hasOrganizerRights(role: UserRole)` = `role === 'Organizer' || role === 'F&A' || role === 'Admin' || role === 'IT-Admin'`, und sie in `alreadyEntitled` sowie in `requestCoOrganizerApprovals` verwenden (dort zusätzlich `getRoleEmails('F&A')` und `getRoleEmails('IT-Admin')` einsammeln, bzw. `roleFilter` um die beiden Werte erweitern). Ergänzend in `RoleContext.addRole` eine Herabstufung nie implizit ausführen: Wenn `existing.role` mehr Rechte trägt als `role`, nicht `updateRole` aufrufen, sondern `true` zurückgeben (nichts zu tun) — genau wie beim gleichnamigen Fall.

### 55. [hoch] „Organizer-Berechtigungen reparieren" meldet Erfolg, ohne eine einzige Antwort zu prüfen

**Wo:** `services/events/subsiteProvisioning.ts:344` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Eine Co-Organizerin sieht ein Klammer-Event mit 19 Terminen als vollständig leer (0 Teilnehmer) — der Fall, für den die Aktion v30.37 gebaut wurde. Der Haupt-Organizer klickt im Organizer Center „Organizer-Berechtigungen reparieren". Auf den Sub-Event-Subsites hat er selbst kein „Manage Permissions" (z.B. weil ein Termin über den Recreate-Pfad neu angelegt wurde, oder die Subsite recycelt ist): jeder POST kommt mit 403 bzw. 404 zurück. Trotzdem meldet die Kachel „3 Person(en) auf 20 Liste(n) berechtigt" — grün, ohne Fehler. Die Co-Organizerin sieht danach weiter überall 0. Dasselbe im Admin Hub: `repairAllOrganizerPermissions` summiert `r.grants` und AdminHubPage.tsx:350 zeigt „…, N Zuweisung(en) gesetzt" mit `variant: 'success'`, weil `r.errors` nur geworfene Exceptions zählt — und `ensureOrganizerPermissionsMulti` wirft nie.

**Ursache:** `svc._post` gibt die `SPHttpClientResponse` zurück und wirft bei HTTP-Fehlern NICHT (services/EventService.ts:1973-1983). Der `try/catch` fängt also nur Netzwerkfehler; `result.grants++` läuft bei 403/404 genauso. Die zweite, eigentlich entscheidende Zuweisung (Leserecht auf der Teilnehmerliste) wird gar nicht erst gezählt. Das Ergebnis-Objekt trägt damit nur die EINGABEN (`sites`, `users`) zurück, und AdminActionsCard.tsx:1490 formuliert genau daraus die Erfolgsmeldung. Das ist die in CLAUDE.md notierte Falle in umgekehrter Richtung: nicht ein leeres Ergebnis wird als Aussage gelesen, sondern ein ungeprüftes.

**Fix-Vorschlag:** In `ensureOrganizerPermissionsMulti` beide `_post`-Antworten auswerten: `const r = await svc._post(...); if (r.ok || r.status === 200) result.grants++; else result.failed.push({site, userId, status: r.status});` — SharePoint antwortet auf ein bereits vorhandenes Recht mit 200, ein echtes Scheitern ist also unterscheidbar. Das Ergebnis um `failed`/`failedSites` erweitern und in AdminActionsCard.tsx:1488-1492 sowie AdminHubPage.tsx:350 anzeigen (`resultIsError` bzw. `variant: 'error'`, wenn `failed.length > 0`), inklusive der betroffenen Subsite-URLs — genauso wie die Teilnehmer-Ansicht seit v30.37 gesperrte Termine namentlich aufzählt statt sie als 0 zu rendern.

### 56. [hoch] „Organizer-Mails reparieren" schreibt verschobene Name↔E-Mail-Paare zurück

**Wo:** `components/AdminPage.tsx:8901` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Admin → Aktionen → „Organizer-Mails reparieren (alle Events)". Ein Event hat organizers = [Anna Meier, Bernd Schulz, Carla Weber] und organizerEmails = ['', '', 'carla.weber@…']. Die Graph-Suche löst Anna auf, Bernd bleibt mehrdeutig (genau der vom Tile selbst dokumentierte Fall „bleiben mit leerem Email-Slot"). Geschrieben wird Organizer='Anna Meier; Bernd Schulz; Carla Weber' und OrganizerEmail='anna.meier@…;;carla.weber@…'. Beim nächsten Laden filtert context/eventMapping.ts:49 die leeren Segmente heraus → organizerEmails = [anna, carla]. Ab jetzt zeigt der Wizard (components/wizard/steps/DetailsStep.tsx:197, `organizerEmails[i]`) neben „Bernd Schulz" Carlas Foto und Adresse, und `remove(idx)` auf Bernds Chip löscht Carlas E-Mail. Dieselbe Verschiebung entsteht auch mit nur EINEM Slot „E-Mail ohne Namen" (dann fehlt ein Name-Segment). Genau die Fehlausrichtung, die die Aktion beheben soll, wird also neu erzeugt — und der zweite Lauf zementiert sie, weil Index 1 dann name+mail hat und nicht mehr als Mismatch erkannt wird.

**Ursache:** Organizer und OrganizerEmail sind zwei parallele, POSITIONSGEBUNDENE Strings, aber die Lesekante (`eventMapping.ts`, `.split(';').filter(s => s)`) verwirft leere Segmente. Ein „leerer Slot" überlebt den Round-Trip also nicht. Der Wizard weiß das und wirft in `sanitizeOrganizerPairs` (EventCreationPage.tsx:3765) jedes unvollständige Paar bewusst weg (`if (n && e) pairs.push(...)`, sonst `dropped++`). Die Admin-Reparatur macht das Gegenteil (`.filter(p => p.n || p.e)`) und erzeugt damit Lücken im E-Mail-String. Dieselbe Schreibweise steckt auch im Tile „Login-Tokens in Namen reparieren" (`updateEvent(ev.id, { 'Organizer': names.join('; '), 'OrganizerEmail': emails.join(';') })`), dort zusätzlich mit einem sparse Array, wenn `emails` kürzer als `names` ist.

**Fix-Vorschlag:** Nur vollständige Paare persistieren — analog zu `sanitizeOrganizerPairs`: `.filter(p => p.n && p.e)`. Die unauflösbaren Namen NICHT stillschweigend mitschreiben, sondern (wie bereits vorbereitet) über `unresolvedNames` melden, damit der Admin sie im Wizard nachträgt. Alternativ, falls Slots wirklich erhalten bleiben sollen: beide Spalten gemeinsam als JSON-Paarliste speichern oder die Lesekante auf `.split(';').map(s => s.trim())` OHNE `filter(Boolean)` umstellen — dann aber an ALLEN Lesestellen. Dasselbe gilt für das Tile „Login-Tokens in Namen reparieren".

### 57. [hoch] „Spalten fixen (alle Events)" löst pro Event einen kompletten loadEvents aus — das ist der v29.77-Anfragensturm mit 429-Sperre

**Wo:** `context/actions/maintenance.ts:113` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Admin klickt im Admin Hub „Spalten fixen (alle Events)". Die Schleife läuft über alle ~94 Events; `fixRegistrationListColumns` liefert für JEDES Event mit mindestens einem Custom-Field eine nicht-leere `customFieldMap` (regListRepair.ts:310 nimmt auch bereits bestehende Zuordnungen auf), also wird `updateEvent` ohne `skipReload` gerufen. Jeder dieser Aufrufe zieht ein volles `loadEvents()` nach sich (getEvents über alle Events plus ein nachgelagerter Teilnehmerzähl-Lauf über alle Subsites, 6 parallel). Ergebnis: dutzende überlappende Komplett-Reloads, SharePoint drosselt, der Lauf bricht mit Fehlern ab und der ausführende Account kann in die Nutzer-Sperre laufen.

**Ursache:** Die deps-Signatur (maintenance.ts:19) sieht `opts?: { skipReload?: boolean }` ausdrücklich vor, die Aufrufstelle nutzt es aber nicht. Genau dieselbe Konstellation ist in EventContext.tsx:2440ff (v29.77) als Ursache der 429-Drossel dokumentiert: nicht die kleinen POSTs, sondern die Voll-Reloads pro Schleifendurchlauf.

**Fix-Vorschlag:** In der Schleife `updateEvent(ev.id, {...}, { skipReload: true })` verwenden und NACH der Schleife (vor dem return, neben `onProgress(total,total,'')`) genau einmal `loadEvents()`/`refreshEvents()` aufrufen. Dasselbe für restoreCustomFieldDescriptions (Zeile 280).

### 58. [mittel] Anmelde-/Abmeldefrist je Sub-Event: UTC-ISO wird als Browser-Lokalzeit angezeigt, aber als Berliner Zeit zurückgeschrieben

**Wo:** `components/wizard/steps/CapacityStep.tsx:433` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Ein Organizer arbeitet auf einer Maschine, deren Zeitzone nicht Europe/Berlin ist (Citrix-/VDI-Bilder stehen häufig auf UTC, Reise/UK-Kollegen auf Europe/London). Ein Sub-Event hat die Frist 07.09.2026 23:59 Berlin, gespeichert als `2026-09-07T21:59:00.000Z`. Der Picker zeigt (UTC-Browser) 21:59 statt 23:59 — die Termin-Liste daneben zeigt über `isoToLocal()` weiterhin 23:59, zwei Ansichten widersprechen sich. Fasst der Organizer das Feld an (auch nur denselben Tag neu anklicken), wird 21:59 als BERLINER Zeit interpretiert und als `2026-09-07T19:59:00.000Z` geschrieben: Die Anmeldung schließt zwei Stunden früher als angezeigt und als vorher gültig. Jedes weitere Anfassen verschiebt sie erneut um den Offset. Identisch für die Abmeldefrist (Zeile 484).

**Ursache:** Die Sub-Event-Spalten liegen als UTC-ISO vor; die Datei hat dafür `isoToLocal()`/`berlinLocalToUtcIso()` und in Schritt 1 die zentralen Helfer `subIsoToDate`/`subDateToIso`. Diese beiden Picker umgehen den Lese-Helfer und nehmen `new Date(iso)`, was die Anzeige an die Browser-Zeitzone koppelt, während der Schreibpfad korrekt `berlinLocalToUtcIso()` benutzt. Hin- und Rückweg laufen dadurch über zwei verschiedene Zeitzonen — die Asymmetrie ist bei Browser=Berlin unsichtbar und genau deshalb bisher nicht aufgefallen. Der Top-Level-Picker ist unauffällig, weil `registrationDeadline` dort schon als Berliner Lokal-String im State liegt.

**Fix-Vorschlag:** In beiden Pickern die vorhandenen Helfer benutzen: `selected={subIsoToDate(se.registrationDeadline)}` bzw. `selected={se.registrationDeadline ? localStrToDate(isoToLocal(se.registrationDeadline)) : null}` und im onChange weiterhin `berlinLocalToUtcIso(dateToLocalStr(date))`. Damit sind Anzeige und Schreibweg wieder dieselbe Zeitzone — wie bei Start/Ende in Schritt 1. Dasselbe Muster gilt für die „Regel wäre …"-Ausgaben daneben (`new Date(ruleIso).toLocaleString('de-DE', …)` ohne `timeZone: 'Europe/Berlin'`).

### 59. [mittel] Assistenz kann eine abgemeldete Person nicht wieder für ein Sub-Event anmelden

**Wo:** `components/AssistantPage.tsx:657` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Eine Assistenz meldet ihren Chef versehentlich für den falschen Tag an und meldet ihn über „Abmelden" wieder ab. Nach dem `reload()` zeigt die Zeile des Termins das graue Badge „Abgemeldet" — und keinen einzigen Knopf mehr: „Anmelden" fehlt (weil eine Zeile existiert), „Angaben anpassen"/„Abmelden" fehlen (weil die Zeile nicht aktiv ist). Für den richtigen Tag geht es, für den einmal abgemeldeten Tag ist die Assistenz dauerhaft ausgesperrt und muss die Organizer bitten.

**Ursache:** `getProxyRegistrationsByActor` (services/EventService.ts:3280-3305) filtert nicht nach Status und liefert auch 'Abgemeldet'-Zeilen; `cancelRegistration` setzt nur den Status und lässt `RegisteredByEmail` stehen (services/events/registrationStatus.ts:182-186). In `buildGroups` landet die abgemeldete Zeile deshalb in `group.regs`, `item` ist gesetzt. `RegRow` blendet Edit/Cancel über `active = ACTIVE_STATUSES.indexOf(status) >= 0` aus (Zeile 733/750/755), der Register-Knopf hängt aber an `!item` statt an `!active`.

**Fix-Vorschlag:** Die Register-Bedingung an den Status koppeln statt an die Existenz der Zeile: in AssistantPage `const itemActive = item && ACTIVE_STATUSES.indexOf(item.registration.Status) >= 0;` und `onRegister={!itemActive ? () => openRegister(group, child) : undefined}`; in `RegRow` entsprechend `{!active && onRegister && !over && (…)}`. `registerForEvent` reaktiviert eine abgemeldete Zeile bereits korrekt (EventContext: `existing.Status !== 'Abgemeldet'` blockt nur aktive Zeilen).

### 60. [mittel] Belegungsanzeige der Sub-Events zählt nur die eigene Zeile (Item-Level-Security)

**Wo:** `components/myEvents/MyEventSubEvents.tsx:128` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Ein Teilnehmer öffnet „Meine Events" bei einem Klammer-Event mit Terminen à 20 Plätzen. Neben jedem Termin steht „0/20" bzw. „1/20", obwohl der Termin ausgebucht ist; das rote „(voll)" erscheint nie und der Knopf „Anmelden" bleibt aktiv. Er klickt, landet still auf der Warteliste (siehe vorheriger Fund) und liest weiterhin „1/20 frei". Im Kalender-Raster (Zeile 419-421) gilt dasselbe: der Tooltip verspricht „19 von 20 Plätzen frei".

**Ursache:** Auf den Teilnehmerlisten der Subsites steht seit v26.87 `ReadSecurity: 2` („nur eigene Elemente", gesetzt in services/EventService.ts:2311). `getAllRegistrations` liefert einem normalen Teilnehmer deshalb höchstens die eigene Zeile — genau die Falle, die in CLAUDE.md/v27.10 für `switchSplitGroup` schon einmal behoben wurde („die frühere Zählung über getAllRegistrations ist für normale User unbrauchbar"). Die richtige Quelle wäre `ce.currentParticipants` (types/index.ts:142), die der Rest der App (`RegistrationPage`, `RegistrationActionBar`) benutzt.

**Fix-Vorschlag:** `counts` nicht mehr aus `getAllRegistrations` ableiten, sondern `ce.currentParticipants` verwenden (bei Bedarf über `getLiveCounterStats` aufgefrischt, wie auf der Anmeldeseite). Damit entfallen zugleich N vollständige Listen-Reads pro Karten-Render. Falls die Zählung für Organizer/Admins genauer sein soll, den bisherigen Weg nur für `isAdmin || isParentOrganizer` behalten.

### 61. [mittel] Das interne Originalbild __eventimgorig__ erscheint Teilnehmern als Event-Dokument

**Wo:** `services/events/eventAssets.ts:351` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer lädt in Schritt 1 ein Querformat-Bild hoch und schneidet es in der App rund zu. `uploadEventOrigImageAsAttachment` legt zusätzlich das Attachment `__eventimgorig__m9k2z1.jpg` am DEX_Events-Item ab. `getEventAttachments` filtert nur das Präfix `__eventimage__`, das Original rutscht durch und landet in `event.documents` (EventContext.tsx:348). Teilnehmer sehen auf der Anmeldeseite und der Organizer in Schritt 7 „Dokumente" eine Datei namens `__eventimgorig__m9k2z1.jpg` zum Download — neben den echten Anhängen wie der Agenda.

**Ursache:** Die beiden Präfixe wurden bewusst unterschiedlich gewählt (Kommentar Zeile 17-21: „bewusst KEIN '__eventimage__'-Präfix-Match, sonst würde der normale Bild-Upload es mitlöschen"). Derselbe Nicht-Match sorgt aber dafür, dass der Dokument-Filter das Original nicht erkennt — beim Einführen des Originals in v28.11 wurde nur die Lösch-Schleife angepasst, nicht der Lese-Filter. Kein anderer Ort im Repo filtert `__eventimgorig__` (grep über alle .ts/.tsx bestätigt das).

**Fix-Vorschlag:** Beide internen Präfixe ausfiltern und die Konstante dafür verwenden statt des Literals: `.filter((f: any) => { const n = String(f.FileName || ''); return n.indexOf('__eventimage__') !== 0 && n.indexOf(ORIG_IMAGE_PREFIX) !== 0; })`. Bestandsevents brauchen keine Migration — der Filter wirkt beim Lesen.

### 62. [mittel] Demo-Vorlagen schreiben eine 10-Zeichen-Frist („YYYY-MM-DD") in einen State, den der DatePicker als UTC-Mitternacht liest — die Frist springt von 23:59 auf 02:00

**Wo:** `components/wizard/hooks/useWizardOptionState.tsx:497` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Nachgerechnet mit TZ=Europe/Berlin: `new Date("2026-09-05")` ist laut ECMAScript ein Date-Only-Form und wird als UTC-Mitternacht geparst → `Sat Sep 05 2026 02:00:00 GMT+0200`. Ein Organizer klickt im Wizard auf „Demo" (loadDemoStandard) — im Feld „Anmeldung bis" steht daraufhin „05.09.2026, 02:00", obwohl die Vorlage 23:59 berechnet hatte. Zwei Folgen: (a) Speichert er ohne Anfassen, schreibt `deadlineToEndOfDayIso` (EventCreationPage.tsx:392-398) wegen der 10 Zeichen 23:59 — Anzeige und gespeicherter Wert gehen um 22 Stunden auseinander. (b) Korrigiert er die offensichtlich falsche „02:00" durch erneutes Anklicken desselben Tages im Kalender, liefert `onChange` `date.getHours() === 2` und schreibt „2026-09-05T02:00"; jetzt sind es 16 Zeichen, `deadlineToEndOfDayIso` lässt sie unverändert und die Anmeldung schließt am 05.09. um 02:00 statt 23:59. Teilnehmer, die am 05.09. tagsüber buchen wollen, bekommen „Anmeldefrist abgelaufen" — einen ganzen Tag zu früh gesperrt. Im Winter ist es 01:00 statt 02:00, der Effekt bleibt.

**Ursache:** Der State `registrationDeadline`/`lastDeregisterDate` trägt überall sonst das 16-Zeichen-Format „YYYY-MM-DDTHH:MM" (Berliner Lokalzeit) — so laden ihn `useWizardEventFieldState.ts:66/75` über `isoToLocal`, so schreiben ihn alle DatePicker. Nur die vier Demo-Loader legen die 10-Zeichen-Form hinein. `new Date()` behandelt beide Formen unterschiedlich (Date-Only = UTC, Date-Time ohne Offset = Browser-Lokalzeit), und `deadlineToEndOfDayIso` verzweigt zusätzlich auf `dateStr.length === 10`. Damit hängt an einem State-Wert zweierlei Bedeutung.

**Fix-Vorschlag:** In allen vier Demo-Loadern `fmtDate(deadline)` durch `fmtDatetime(deadline)` ersetzen (useWizardOptionState.tsx:497, :498, :521, :522 und wizardTemplates.ts:348, :349, :425, :426). Das Date trägt die 23:59 bereits (`beforeNextSaturday(n, 23, 59)`), es geht nur die Uhrzeit-Komponente verloren. Danach ist der State durchgängig 16-stellig; `deadlineToEndOfDayIso` behält seinen 10-Zeichen-Zweig als Alt-Pfad, wird aber nicht mehr getroffen.

### 63. [mittel] Deutsche Job-Titel („Assistenz") sehen den Knopf „Für andere anmelden" nie

**Wo:** `components/RegistrationPage.tsx:191` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Eine Person mit dem Profil-Titel „Assistenz", „Teamassistenz" oder „Assistentin" öffnet die Anmeldeseite eines Events, um ihre Partnerin/ihren Director anzumelden. `isAssistant` ist false (in „assistenz" steckt kein „assistant"), also ist `canRegisterForOther` false und der Umschalter „Für eine andere Person anmelden" wird in PersonalDataSection.tsx:75/138 gar nicht gerendert. Die Person kann die Anmeldung nicht durchführen und hat keinen Anhaltspunkt, warum — sie sieht schlicht ein normales Selbst-Anmeldeformular. Serverseitig wäre sie zugelassen: `canRegisterForOthers` matcht seit v23.9 ausdrücklich beide Schreibweisen. Der Fix von v23.9 kann für genau die Zielgruppe, für die er gebaut wurde, nie greifen.

**Ursache:** Die Job-Titel-Heuristik existiert in der Codebasis dreimal mit drei verschiedenen Mustern: `services/events/profileData.ts:118` prüft `indexOf('assisten') || indexOf('assistan')`, `components/EventListPage.tsx:152` prüft `/assisten|assistant/i`, und nur diese Stelle prüft das englische `includes('assistant')`. Beide Seiten lesen denselben Wert (SP-Profil-Property `Title`, vgl. context/UserContext.tsx:208 und services/events/profileData.ts:160), die Abweichung ist also rein im Muster. Bei der v23.9-Korrektur wurde nur der Service angefasst, nicht das UI-Gate, das über die Sichtbarkeit entscheidet.

**Fix-Vorschlag:** Die Heuristik in einen Helfer ziehen (z.B. `utils/fieldHeuristics.ts` → `looksLikeAssistantJobTitle(jt: string)` mit `/assisten|assistan/i`) und an allen drei Stellen verwenden; hier konkret `const isAssistant = /assisten|assistan/i.test(currentJobTitleLc);`. Das Ziel-Kriterium (`ALLOWED_TARGET_TITLES = ['partner','director']`) bleibt unverändert, es ist im Tenant ohnehin englisch.

### 64. [mittel] Die in v30.63 dokumentierte Organizer-Selbstheilung des Platzzählers ist nicht verdrahtet — `onlyMine` wird nirgends übergeben

**Wo:** `components/DexEventPlatform.tsx:584` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Ein Organizer (kein Admin) betreut ein volles Event. Der Zähler ist nach oben gedriftet (Selbst-Abmeldung bei unbekanntem Wartelisten-Stand zählt fail-closed nicht herunter, seats.ts:511-515). Laut Release Notes v30.63 („Jetzt läuft er auch für Organizer, per neuem `onlyMine`-Filter auf die selbst betreuten Events beschränkt") soll der Abgleich beim Start des Organizers laufen und das heilen. Er läuft nicht: Der Effect bricht bei `if (!isAdmin) return;` ab, und der einzige Aufruf im ganzen Repo (`grep -rn reconcileCounters`) übergibt kein `opts`. Bis sich ein echter Admin anmeldet UND die 6-h-Drossel (`dex_counter_reconcile_lastrun`) abgelaufen ist, bleibt der Zähler falsch — die Anmeldeseite zeigt seit v30.62 genau diesen Zähler an, meldet also „ausgebucht", während Plätze frei sind. `canCreateEvents` steht bereits im Deps-Array, wird im Rumpf aber nie benutzt: der Organizer-Zweig war vorgesehen und ist nie eingebaut worden.

**Ursache:** `reconcileCounters(opts?: { onlyMine?: boolean })` hat den Parameter bekommen (EventContext.tsx:462, samt ausführlicher Begründung im Kommentar ab Zeile 464), aber die einzige Aufrufstelle in DexEventPlatform.tsx wurde nicht nachgezogen — weder das `isAdmin`-Gate noch das Argument.

**Fix-Vorschlag:** Gate und Aufruf zusammenführen: `if (!isAdmin && !canCreateEvents) return;` und `reconcileCounters(isAdmin ? undefined : { onlyMine: true })`. Die Drossel sollte dabei pro Rolle getrennt geschlüsselt werden (z.B. `dex_counter_reconcile_lastrun_org`), sonst blockiert ein Organizer-Lauf den nächsten Admin-Lauf für sechs Stunden.

### 65. [mittel] Endgültige Löschungen zählen Versuche statt Erfolge — deleteRegistration liefert `false`, das nirgends geprüft wird

**Wo:** `components/admin/participants/CancelledList.tsx:182` · Quelle: Audit B (nach dem Umbau)

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

### 66. [mittel] F&A-Excel: In der Spalte „Participent Type" steht der Anzeigename der Person

**Wo:** `utils/faBilling.ts:485` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** F&A lädt im F&A Center die Teilnehmerliste als Excel. In Zeile 14 steht die Kopfzeile der F&A-Vorlage, ab Zeile 15 die Personen. Spalte A ist mit „Participent Type" beschriftet, enthält aber „Mustermann, Max" — bei 300 Teilnehmern also 300 verschiedene „Typen". Der Name steht ohnehin schon zerlegt in „First Name"/„Last Name" (Spalte C/D). Wer die Datei in der gewohnten Weise weiterverarbeitet (Pivot/Filter auf den Teilnehmertyp), bekommt Unsinn; wer sie prüft, hält die Datei für falsch aufgebaut.

**Ursache:** `buildFASheetAoa` schreibt die Werte positionsweise gegen `FA_SHEET_PARTICIPANT_HEADERS`, und für Spalte A gab es in DEX keinen passenden Wert — statt die Zelle leer zu lassen (so wie es v30.50 für `Personalnummer`/`kostenstelle` bewusst getan hat: „Ein ‚—' wäre für F&A ein Wert") wurde der Anzeigename eingesetzt. Die Release Note zu v30.50 begründet ausdrücklich, warum die Beschriftungen die der F&A-Vorlage sind, sagt aber zu dieser Spalte nichts — der Name ist dort offensichtlich nicht gemeint, sonst hieße die Spalte „Name".

**Fix-Vorschlag:** Entweder die Zelle leer lassen (wie vor v30.60 bei Personalnummer/Kostenstelle) und F&A sie ausfüllen lassen, oder einen echten Typ ableiten, den DEX kennt — intern/extern lässt sich an der E-Mail-Domäne bzw. am Vorhandensein eines Verzeichnis-Treffers (`getEmployeeData`) festmachen: `const type = map[email] ? 'Employee' : 'Guest'`. Vorher mit F&A abklären, welche Werte die Vorlage dort erwartet; bis dahin ist die leere Zelle das ehrlichere Ergebnis.

### 67. [mittel] Hotelplanung: der Warnkasten „Unterkunftsbedarf ohne Hotel“ liest nur die Klammer-Zeile — bei einem Klammer-Event bleibt er leer, obwohl Personen ohne Zimmer sind

**Wo:** `components/HotelPlanningPanel.tsx:991` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Klammer-Event, bei dem die Hotel-Frage im Formular des Termins steht (der Normalfall laut CLAUDE.md: „Wer sich über ein Sub-Event angemeldet hat, hat auch die Hotel-Frage in dessen Formular beantwortet, und die Klammer-Zeile ist bei einer Klammer nur eine Schattenzeile ohne Antworten“). Verschärft für jede Person, die NICHT über die Anmeldeseite kam: das zentrale Netz in EventContext.tsx:1331-1336 legt die Klammer-Zeile mit `registerForEvent(parentEv.id, {}, ...)` an — CustomData ist buchstäblich `{}`. Der Organizer öffnet „Hotels & Übernachtungen“, verteilt manuell und verlässt sich auf den orangen Kasten „N Person(en) mit Unterkunftsbedarf ohne Hotel“ (Zeile 1468-1483, Text: „genau der Fall, der in einer Excel-Liste durchrutscht“). Der Kasten rendert gar nicht, weil `JSON.stringify` der leeren Schattenzeile weder 'hotel' noch 'yes' enthält — obwohl `autoDistribute` dieselben Personen über `wishOf` sehr wohl als Kandidaten führt. Die Ansicht sagt „alles versorgt“, und genau die Person ohne Zimmer reist an. Zweite Lücke im selben Ausdruck: wer den Zeitraum über ein `daterange`-Feld angegeben hat (v28.63), wird nicht erkannt, weil `formStayOf` nicht abgefragt wird.

**Ursache:** v29.3 hat die Antwort-Auflösung `answerRowsOf(p)` (Klammer-Zeile + alle Sub-Event-Zeilen derselben Adresse) eingeführt und `wishOf`/`formStayOf` darauf umgestellt — dieser dritte Leser derselben Frage wurde übersehen und arbeitet weiter auf `p` allein. Damit ist er die von CLAUDE.md beschriebene Ausnahme („Antworten stehen dort, wo angemeldet wurde — nicht auf der Klammer“), nur an einer anderen Stelle als 2029.2.

**Fix-Vorschlag:** Den Filter über dieselbe Quelle laufen lassen wie `autoDistribute`, dann verschwindet auch die Heuristik: `const wantsHotelWithout = React.useMemo(() => people.filter(p => !(p.Hotel || '').trim() && wishOf(p) === true), [people, wishOf]);`. `wishOf` deckt beide Ebenen ab, kennt die Zusatznächte-Ausnahme (v28.59) und wertet `formStayOf` mit aus; ausserdem entfällt der Blob-Scan, der heute jede Zeile mit irgendeinem 'yes' irgendwo im JSON als Hotel-Wunsch liest. Der useMemo braucht `wishOf` in den Deps (ist bereits ein useCallback).

### 68. [mittel] KPI-Kachel „Warteliste" fehlt genau bei Events mit geteilten Kapazitäten

**Wo:** `components/AdminPage.tsx:9404` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** B2Run-/Split-Event (Durchstarter 80 / Funstarter 40), Warteliste aktiv, 25 Personen warten. Organizer öffnet das Event im Organizer Center: Die KPI-Reihe zeigt „Angemeldet 120 (● Durchstarter 80/80 ● Funstarter 40/40)", „QR versendet", „Eingecheckt", „Abgemeldet" — aber KEINE Wartelisten-Kachel. Erst weit unten stehen die Tabellen „Warteliste Durchstarter (18)" / „Warteliste Funstarter (7)". Die Kopfzahlen behaupten damit, es gebe keine Warteliste, obwohl 25 Personen warten; wer nach „steht noch jemand an?" sucht, muss scrollen bzw. hält das Event für ausgebucht ohne Nachrücker.

**Ursache:** Bei geteilten Kapazitäten ist `maxParticipants` per Definition 0 — der Wizard schreibt `'MaxParticipants': useSplitCapacities ? 0 : …` (components/wizard/logic/wizardSubmit.ts:365), die Kapazität steht in `durchstarterCapacity`/`funstarterCapacity`. Die Bedingung `maxParticipants > 0` sollte nur „unbegrenzte" Events ausschließen, schließt aber ungewollt alle Split-Events mit aus, weil beide Fälle denselben Wert 0 benutzen. Dieselbe Verwechslung steckt in der Sichtbarkeits-Bedingung der Nachrück-Audit-Spalten in `availableColumns` („Nachgerückt am", „Hat ersetzt"), die dadurch bei Split-Events ebenfalls nie erscheinen.

**Fix-Vorschlag:** Die Kapazitäts-Prüfung um den Split-Fall erweitern: `const hasWaitlistKPI = !!(selectedEvent?.waitlistEnabled && ((selectedEvent?.maxParticipants || 0) > 0 || isSplitCapacity));`. Besser noch einen gemeinsamen Helfer `hasCapacityLimit(ev)` = `maxParticipants > 0 || durchstarterCapacity > 0 || funstarterCapacity > 0` einführen und ihn AUCH in der `availableColumns`-Bedingung für `promotedDate`/`replaced` verwenden — sonst laufen die beiden Stellen wieder auseinander.

### 69. [mittel] KPI-Zähler „Events" wird beim Löschen um Sub-Events dekrementiert, die beim Anlegen nie mitgezählt wurden

**Wo:** `context/EventContext.tsx:2520` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer legt ein Klammer-Event mit 8 Terminen an: `createEvent` erhöht den KPI-Zähler genau EINMAL (Zeile 579: `if (!input.isFictive && !input.parentEventId)`). Wird dasselbe Event später gelöscht, zieht `deleteEvent` 9 ab (8 Kinder + Klammer). Die Startseiten-Kachel „Events durchgeführt" zeigt danach 8 zu wenig; bei mehreren gelöschten Kalender-Events kann der Wert negativ werden. Sichtbar bleibt der falsche Wert für alle normalen Nutzer, bis zufällig ein Admin die App öffnet — nur dort läuft `recomputeEventKpiOnly` (DexEventPlatform.tsx:255, zusätzlich einmal pro Browser-Session gedrosselt).

**Ursache:** Die beiden Buchungsstellen benutzen unterschiedliche Zählregeln. v23.38 hat die Erhöhung bewusst auf Top-Level-Events beschränkt (Kommentar Zeile 576-578), die Gegenbuchung im Löschpfad wurde dabei nicht nachgezogen. Auch `recomputeEventKpiOnly` zählt nur `!e.parentEventId` — die Erhöhungsseite ist also die korrekte.

**Fix-Vorschlag:** Im Löschpfad dieselbe Regel anwenden: nur das Top-Level-Event zählen, `const childEventsToDecrement = (ev && !ev.isFictive) ? 1 : 0;` — die Kinder-Schleife bleibt für `bumpKpiParticipants` unverändert.

### 70. [mittel] Kalender-Tageskachel wird in Browser-Zeitzone berechnet und rutscht bei UTC-Browsern auf den Vortag

**Wo:** `components/registration/EventSpecificSection.tsx:420` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Der Wizard erzeugt für den Office-Tag Di, 16.09.2026 ein Sub-Event mit Start 00:00 Europe/Berlin. `subDateToIso` → `berlinLocalToUtcIso` speichert das als `2026-09-15T22:00:00.000Z` (MESZ). Ein Teilnehmer öffnet die Anmeldeseite in einem Browser auf UTC (in EventCreationPage.tsx:335 ausdrücklich als real vorgekommener Fall benannt: „ein Browser auf UTC oder in einer VM/Citrix mit falscher TZ") oder in einer westlicheren Zeitzone. `new Date(iso).getDate()` liefert dort 15 → Schlüssel „2026-09-15". Die Kachel für den Dienstag erscheint im Monatsraster auf dem Montag; ein Termin am 1. eines Monats wandert sogar in das Raster des Vormonats. Der Klick bucht `entry.ce`, also den technisch richtigen Termin — der Teilnehmer glaubt aber, den Montag gebucht zu haben, und erscheint am falschen Tag. Dieselbe Rechnung steht in components/myEvents/MyEventSubEvents.tsx:355, also zeigt auch „Meine Events" den verschobenen Tag.

**Ursache:** Die App interpretiert Event-Zeiten laut Konvention IMMER als Europe/Berlin; die dafür gebauten Intl-Helfer (`berlinOffsetMs`, `isoToLocal`) stehen ausschliesslich in `EventCreationPage.tsx` — `rg 'Europe/Berlin'` trifft im gesamten Quellbaum nur diese eine Datei. Alle teilnehmersichtbaren Tages-Rechnungen benutzen die browserlokalen Getter, und da eine Tagesgrenze in Berlin in UTC noch im Vortag liegt, kippt der Tag um eins.

**Fix-Vorschlag:** Einen Berlin-festen Tages-Helfer nach `utils/eventFormat.ts` ziehen (`berlinDayKey(iso)` über `Intl.DateTimeFormat('en-CA', { timeZone: 'Europe/Berlin', year:'numeric', month:'2-digit', day:'2-digit' })` bzw. `formatToParts`) und ihn in `EventSpecificSection.dayOf`, `MyEventSubEvents.dayOf` sowie in der `openFrom`-Rechnung beider Dateien (dort ebenfalls `new Date(ce.startDate).getFullYear()/getMonth()/getDate()`) verwenden. Die vorhandenen Helfer aus `EventCreationPage` dabei auf dasselbe Modul umstellen, damit es weiterhin nur EINE Umrechnungsstelle gibt.

### 71. [mittel] Konsolidierte Matrix: `hasParentReg` und die Parent-Feld-Auflösung ignorieren den Status und nehmen die älteste Zeile — abgemeldete Klammer-Zeile schluckt die Fehlermeldung und fängt die Bearbeitung ab

**Wo:** `components/admin/sections/ConsolidatedView.tsx:138` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Ausgangslage wie im vorigen Fund: eine Person hat aktive Termin-Zeilen, ihre Klammer-Zeile steht auf 'Abgemeldet' (gezielte Klammer-Abmeldung über den Dereg-Dialog, oder der Anmeldeseiten-Fehler oben). In der konsolidierten Matrix: (1) `hasParentReg` liefert true, weil kein Status geprüft wird → der rote Kasten „Fehlende Klammer-Anmeldung“ (Zeile 361) listet sie NICHT, und die Aktion „Zur Klammer hinzufügen“ ist ausgeblendet (Zeile 1068, `!hasParentReg`). Der Organizer hat keinen Weg, den Zustand zu erkennen oder zu reparieren. (2) Die Aktion „Felder“ (Zeile 1067/1080 → `openMainFieldsEdit`) ist dagegen sichtbar, und createKlammerActions.ts:313 nimmt wieder dieselbe Zeile: der Organizer trägt Hotel/Verpflegung in eine ABGEMELDETE Zeile ein. HotelPlanningPanel liest nur ACTIVE_STATI, sieht die Werte nie; die Person bleibt ohne Zimmer. (3) Analog schreibt „Assistenz zuordnen“ (createKlammerActions.ts:266) den RegisteredBy/Autor auf die abgemeldete Zeile. Zweite Variante desselben Fehlers: Bei ZWEI Klammer-Zeilen zur selben Adresse — laut DuplicateRegHintBox.tsx:58 ein ausdrücklich toleriertes „Aufräumen ist optional“-Szenario — liefert `getAllRegistrations` mit `$orderby=Id asc` (registrationEdit.ts:665) immer die ÄLTESTE; steht die aktuelle Antwort auf der neueren Zeile, zeigt die Matrix die veraltete, weil der Sub-Event-Fallback (Zeile 959) nur bei LEEREM Wert greift.

**Ursache:** Die Matrix schliesst aus dem E-Mail-Schlüssel auf genau eine Klammer-Zeile — mit `some`/`find` ohne Status-Filter und ohne definierte Reihenfolge. `registrations` kommt aus `getAllRegistrations` und enthält bewusst ALLE Zeilen inkl. 'Abgemeldet' (die Abmeldungen-Liste braucht sie). Dass genau dieselbe Suche eine Zeile weiter unten im Dereg-Dialog korrekt auf ACTIVE gefiltert wird, zeigt, dass es Vergessen und keine Absicht ist.

**Fix-Vorschlag:** Eine gemeinsame Auflösung einführen und an allen fünf Stellen benutzen — `const activeParentRegOf = (emailKey: string) => registrations.filter(r => (r.ParticipantEmail || '').toLowerCase().trim() === emailKey).sort((a,b) => (b.Id||0)-(a.Id||0)).find(r => ACTIVE.indexOf(r.Status||'') >= 0)`, also: nur aktive Zeilen, davon die neueste. `hasParentReg` = `!!activeParentRegOf(key)` — dann meldet der rote Kasten den Fall wieder und „Zur Klammer hinzufügen“ repariert ihn (registerForEvent geht über den Reaktivierungs-Pfad, es entsteht keine Dublette). `openMainFieldsEdit`/`submitAssignAssistant` in createKlammerActions.ts auf dieselbe Funktion umstellen.

### 72. [mittel] Nach dem Einchecken bleibt die Trefferliste auf „Angemeldet" stehen (Cache wird nicht nachgeführt)

**Wo:** `components/CheckInPage.tsx:713` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Am Einlass wird über die Teilnehmer-ID eingecheckt (seit v30.33 der Weg, der überall trägt). Nach dem Bestätigen zeigt die Liste die Person weiter mit blauem Badge „Angemeldet", die KPI-Kachel „Eingecheckt" bleibt auf dem Wert vom Laden stehen, und mit aktivem Filter „Nur offene" steht die Person weiter unter den offenen Anmeldungen. Zwei Helfer an zwei Geräten checken dieselbe Person doppelt ein bzw. suchen sie erneut; niemand kann an der Liste ablesen, wer schon durch ist.

**Ursache:** `searchHits` und `checkInKpis` lesen ausschließlich `searchRegsCache[nameSearchEventId]` (Zeilen 197 und 213). `confirmCheckIn` schreibt den neuen Status nur in die Ergebnismeldung und in `checkedInCount`, aber nicht in den Cache — anders als `markNoShowFromSearch`, das den Cache in Zeile 360-363 korrekt patcht. Ein Reload der Liste findet nicht statt, weil `loadRegsForSearch` bei vorhandenem Cache sofort zurückkehrt (Zeile 168).

**Fix-Vorschlag:** In `confirmCheckIn` nach erfolgreichem Schreiben denselben Patch fahren wie bei No-Show: `setSearchRegsCache(prev => { const list = prev[nameSearchEventId] || []; return { ...prev, [nameSearchEventId]: list.map(r => r.Id === pendingCheckIn.regId ? { ...r, Status: 'Eingecheckt' } : r) }; });` (nur im Erfolgsfall, siehe Fund zum ignorierten Rückgabewert).

### 73. [mittel] Nach einem erfolgreichen Outlook-Recreate meldet der Wizard fälschlich „Termin konnte nicht gelöscht werden" und lädt die Eventliste je Termin komplett neu

**Wo:** `components/wizard/logic/persistSubEvents.ts:583` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer klickt „Fehlende Termine jetzt anlegen" für 19 Termine. Alle Recreates GELINGEN. Die alten dbIds werden bewusst nicht in keptDbIds aufgenommen (Kommentar Zeile 336: „das alte Item wurde gelöscht"), stehen aber weiterhin im Mount-Snapshot initialSubEventDbIds (EventCreationPage.tsx:2054, useState-Initializer). Die Aufräum-Schleife ruft deshalb für jede alte Id deleteEvent auf: EventService.deleteEvent macht getEvent → 404 → return false. Der Organizer bekommt nach einem vollständig erfolgreichen Speichern die rote Meldung „19 abgewählte Termine konnten nicht gelöscht werden: … Sie stehen deshalb weiterhin in der Liste — bitte speichere erneut." Zusätzlich ruft jeder dieser Läufe EventContext.deleteEvent, das am Ende immer `await loadEvents()` macht — 19 vollständige Neuladungen der Eventliste (laut Kommentar in v29.77 rund 28 MB je Lauf) nach einem Save, der nichts zu löschen hatte.

**Ursache:** Zwei Mengen laufen auseinander: initialSubEventDbIds ist ein Snapshot beim Mounten und kennt die ersetzte Id weiterhin, keptDbIds soll aber „nicht löschen" bedeuten. Für einen ersetzten Termin ist beides falsch modelliert — er ist weder „behalten" (die Id ist tot) noch „abgewählt" (der Termin existiert unter neuer Id).

**Fix-Vorschlag:** Eine dritte Menge `replacedDbIds` führen: bei jedem erfolgreichen Recreate `replacedDbIds.add(draft.dbId)` und die Aufräum-Schleife auf `if (!keptDbIds.has(oldId) && !replacedDbIds.has(oldId))` einschränken. Alternativ initialSubEventDbIds nach einem Recreate um die alte Id bereinigen.

### 74. [mittel] Nach-Diagnose in fixAllEventColumns liest aus der veralteten Closure — frisch angelegte Spalten werden als „fehlt weiterhin" gemeldet

**Wo:** `context/actions/maintenance.ts:118` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Event hat Custom-Felder ohne `spInternalName` (typischer Fall für den Fix-Lauf). Der Admin startet „Spalten fixen (alle Events)": Die Spalten werden angelegt und die Zuordnung per updateEvent nach DEX_Events geschrieben. Direkt danach liest Zeile 118 aus `events` — dem Array, das beim Erzeugen der Closure gebunden wurde. Das enthält weiterhin die alten Felder mit leerem `spInternalName`. `diagnoseRegistrationList` wertet ein leeres `spInternalName` als „(noch nicht zugeordnet)" (regListRepair.ts:64), also meldet der Ergebnis-Bericht für jedes gerade reparierte Feld „fehlt weiterhin" und `fixedColumns` bleibt leer. Der Admin, der nach der Ursache einer scheiternden Klammer-Anmeldung sucht, bekommt genau die falsche Auskunft — und das ist laut v30.58 der einzige Zweck dieses Berichts.

**Ursache:** `makeMaintenanceActions({ … events … })` bindet das `events`-Array des aktuellen Renders. `updateEvent` löst zwar ein `setEvents` aus, die bereits laufende async-Schleife sieht davon aber nichts — die Closure-Variable ist unveränderlich. `evFresh` ist damit immer identisch zu `ev`.

**Fix-Vorschlag:** Nicht aus dem State lesen, sondern die gerade geschriebene Zuordnung verwenden: `const afterFields = upd.map(f => ({ id: f.id, label: f.label, spInternalName: f.spInternalName || '' }))` (bzw. `res.customFieldMap` direkt), sonst auf die `diagFields` von vorher zurückfallen. Alternativ die Umgebung als Ref (`eventsRef`) statt als Wert hereinreichen.

### 75. [mittel] Nachgeladene Event-Dokumente gehen bei jedem loadEvents verloren und werden wegen des nicht invalidierten `documentsLoadedRef` nie erneut geholt

**Wo:** `context/EventContext.tsx:338` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Teilnehmer öffnet „Meine Events", die Dokument-Liste eines Events wird sichtbar (MyEventCard.tsx:952 rendert nur bei `event.documents.length > 0`). Auf derselben Seite ändert er seine Antworten („updateMyRegistration") oder meldet sich von einem Sub-Event ab — beide Pfade rufen am Ende `loadEvents()`. `loadEvents` ersetzt die Liste komplett durch frisch gemappte Objekte, und `mapSPEventToDeloitteEvent` setzt immer `documents: []` (eventMapping.ts:491). Die Download-Box verschwindet sofort. `ensureEventDocuments` würde sie nachladen, überspringt das Event aber, weil `documentsLoadedRef` es weiterhin als „geladen" führt — und der Effect in MyEventsPage.tsx:268 hängt an `[topLevelEvents.length]`, die sich nicht ändert. Die Dokumente sind bis zu einem echten Seiten-Reload (F5) weg.

**Ursache:** Zwei Zustände über dieselben Daten, die nicht synchronisiert sind: der Merker (Ref) sagt „geladen", die Nutzlast im State wird von jedem Reload auf `[]` zurückgesetzt. Zusätzlich überschreibt der nachgelagerte Zähl-Lauf (Zeile 316, `setEvents(withCounts)`) den State mit einem Snapshot, der VOR einem parallel eingetroffenen `ensureEventDocuments` gezogen wurde — dieselbe Verlust-Situation auch ohne Nutzeraktion. (Verwandt, aber eigene Baustelle: der Wizard initialisiert `documents`/`initialDocumentNames` im useState-Initializer, bevor sein eigener ensureEventDocuments-Effect fertig ist — EventCreationPage.tsx:749/1313 — und zeigt Schritt „Dokumente" deshalb leer.)

**Fix-Vorschlag:** In `loadEvents` bereits geladene Anhänge übernehmen statt sie wegzuwerfen: `setEvents(prev => mapped.map(e => { const old = prev.find(p => p.id === e.id); return old?.documents?.length ? { ...e, documents: old.documents } : e; }))` — analog für `setEvents(withCounts)`. Ersatzweise `documentsLoadedRef.current.clear()` (und die Inflight-Map) zu Beginn jedes `loadEvents`.

### 76. [mittel] Neue Sub-Events erben die Mail-Sprache des zuletzt geöffneten Termin-Reiters statt die des Hauptevents

**Wo:** `components/EventCreationPage.tsx:3522` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Hauptevent auf DE, ein bestehender Termin ist auf EN gestellt. Der Organizer legt in Schritt 1 einen neuen Termin an (neue Drafts setzen `emailLanguage` nicht — weder `toggleDaySubEvent` noch der „Sub-Event hinzufügen"-Knopf in wizard/steps/SubEventsSection.tsx), wechselt danach in Schritt Kommunikation auf den EN-Reiter, schaut etwas nach und klickt von dort aus „Speichern". Der neue Termin wird mit `EmailLanguage = 'EN'` angelegt: Anmelde-/Abmelde-Mail und Outlook-Termin-Standardtext kommen auf Englisch, obwohl das Event deutsch ist.

**Ursache:** `emailLanguage` ist einer der Zustände, die `switchCommTab` zwischen Hauptevent-Slot und Sub-Event-Slot hin- und herspiegelt; auf einem Sub-Reiter hält er den SUB-Wert. Genau dafür gibt es `resolveTopLevelCommState()`, und 24 Zeilen weiter oben wird es für die Logo-Vererbung (`parentComm`) auch benutzt. Der Sprach-Fallback ist bei dieser Umstellung (v11.93/v14.4-Serie) übersehen worden und liest weiter den rohen State. Dieselbe Verwechslung steht auch in `buildOutlookBody(..., (emailLanguage || '').toUpperCase() !== 'EN')` im Hauptevent-Save-Zweig (Zeilen ~4585 und ~5493) — dort betrifft sie nur die Beschriftung des Teams-Teilnahme-Blocks, ist also harmloser, hat aber dieselbe Ursache.

**Fix-Vorschlag:** `const subEmailLang = draft.emailLanguage || parentComm.emailLanguage;` — `parentComm` steht in derselben Funktion bereits zur Verfügung. Analog in den beiden `buildOutlookBody`-Aufrufen des Hauptevent-Zweigs `emailLanguage` durch `effEmailLanguage` ersetzen (die restlichen Argumente dort nutzen bereits die `eff*`-Werte).

### 77. [mittel] Null-Zugriff auf starterCounts im Submit — der „Anmelden"-Knopf reagiert dann gar nicht

**Wo:** `components/registration/submitFlow.ts:492` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Split-Event mit nur einer wählbaren Gruppe: der Effekt in RegistrationPage.tsx:1031 setzt preferredStarterType automatisch (singleStarterType). Der Nutzer füllt das Formular aus und klickt „Anmelden", während der asynchrone Belegungs-Effekt (dynamischer import der EventService-Datei + SharePoint-GET, RegistrationPage.tsx:1059-1084) noch läuft oder in seinen `catch { /* ignore */ }` gefallen ist. starterCounts ist dann `null`, Zeile 492 wirft „Cannot read properties of null (reading 'durch')". Da `setIsSubmitting(true)` erst in Zeile 574 kommt, gibt es keine Fehlermeldung und kein Overlay: Der Klick verpufft, in der Konsole steht eine unhandled rejection. Der Nutzer klickt erneut — und je nach Timing passiert wieder nichts.

**Ursache:** `starterCounts` ist in RegistrationPage.tsx:286 als `{...} | null` deklariert und startet auf null, der Props-Vertrag von createSubmitFlow (submitFlow.ts:103) deklariert ihn dagegen ohne `| null`. Weil tsconfig.json `"strict": false` (und damit strictNullChecks aus) hat, meldet der Compiler die Zuweisung in RegistrationPage.tsx:1520 nicht — das einzige Netz dieser Codebasis (tsc/ESLint/Build) ist an genau dieser Stelle blind. Die Anzeige-Stellen in EventSpecificSection.tsx nutzen konsequent `starterCounts?.durch ?? 0`, der Submit-Pfad als einziger nicht.

**Fix-Vorschlag:** In submitFlow.ts den Prop-Typ auf `{ durch: number; fun: number; durchWait: number; funWait: number } | null` ziehen und in Zeile 491 abfangen: Bei `!starterCounts` den Fallback-Dialog-Block überspringen (reserveSeat entscheidet ohnehin serverseitig, der Dialog ist nur ein Vorab-Komfort) — kein Rechnen mit erfundenen Zahlen. Sauberer wäre zusätzlich, den Submit zu blockieren und kurz auf die Zahlen zu warten, solange der Effekt noch läuft.

### 78. [mittel] QR-Mail-Konfiguration eines Sub-Events geht beim nächsten Wizard-Speichern verloren (headerImage, blockLang, blockNote)

**Wo:** `components/wizard/hooks/useWizardVisibilityState.tsx:367` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Organizer Center → QR-Mail eines Sub-Events gestalten: Kopf-Bild auf „Event-Foto" + volle Breite stellen, Block-Sprache auf DE, eigene Notiz unter dem Code. `saveQrMailOverride` (components/admin/logic/createQrMailActions.tsx:181-186) schreibt das als `QRCode.headerImage` / `.blockLang` / `.blockNote` in `EmailTemplateOverrides` der Sub-Event-Zeile. Danach öffnet derselbe Organizer das Event im Wizard und speichert (aus irgendeinem Grund — Titel, Frist, egal). Die Sub-Event-Hydration kopiert aus dem `QRCode`-Eintrag nur die Text- und Formatfelder, `persistSubEvents.ts:198-205` schreibt genau diesen beschnittenen Stand zurück. Ergebnis: Die QR-Mails dieses Termins zeigen wieder das DEX-Orb statt des Event-Fotos, die Block-Sprache kippt auf die Mail-Sprache und die Notiz ist weg — ohne Meldung, ohne dass jemand die QR-Mail angefasst hat. Beim HAUPTEVENT passiert das nicht: dort landet der komplette Eintrag ungefiltert in `rest` (Zeile 116) und wird unverändert zurückgeschrieben. Genau diese Asymmetrie macht den Fall schwer auffindbar.

**Ursache:** `EmailOverrideEntry` (components/wizard/emailOverrideEntry.ts) kennt nur Text- und Formatfelder; die Sub-Event-Hydration baut den Eintrag Feld für Feld daraus neu auf, statt ihn durchzureichen. Die v30.52/v30.60-Erweiterungen des QR-Overrides (`headerImage`, `blockLang`, `blockNote`) wurden im Organizer-Center-Pfad ergänzt, aber nicht im Wizard-Rundlauf. Zusätzlich fehlt `headerImage` in der Aufnahme-Bedingung: Ein Override, der NUR das Kopf-Bild ändert, fiele komplett durch.

**Fix-Vorschlag:** Den Eintrag durchreichen statt neu zusammensetzen — `filtered[key] = { ...val, subject: val.subject || '', heading: val.heading || '', bodyHtml: val.bodyHtml || '' }` — und `val.headerImage || val.blockLang || val.blockNote` in die Aufnahme-Bedingung mit aufnehmen. `EmailOverrideEntry` um die drei optionalen Felder ergänzen, damit der Typ die Erweiterung künftig erzwingt.

### 79. [mittel] Reservierter Sitzplatz wird bei fehlgeschlagenem Insert nicht zurückgegeben — die Kapazität schrumpft still

**Wo:** `context/EventContext.tsx:813` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Event mit MaxParticipants=50. Der Organizer hat nachträglich ein Abfragefeld ergänzt, dessen SP-Spalte auf DIESER Teilnehmerliste fehlt (der in v30.58 dokumentierte Fall: „bringt den ganzen Insert zu Fall, und zwar nur bei den Personen, die das Feld ausfüllen"). Person füllt das Feld aus und klickt „Anmelden": reserveSeat erhöht SeatsTaken auf 41, der Insert wird mit HTTP 400 abgelehnt, registerForEvent liefert `{ok:false, reason:'insert-failed'}`. Der Platz wird nicht freigegeben. Die Person versucht es dreimal — SeatsTaken steht auf 44, real sind 40 angemeldet. Später melden sich vier Personen an, die auf der Warteliste landen, obwohl Plätze frei sind. Dasselbe passiert bei reason 'not-allowed'/'deadline' und bei einem gescheiterten reactivateRegistration.

**Ursache:** reserveSeat (EventService.ts:5555ff) ist bewusst atomar und schreibt VOR dem Insert. Auf der Fehlerseite gibt es keinen Gegenpart: EventContext.tsx:807-820 wertet nur `r.ok` aus und springt bei false direkt zum `return`. Der v18.8-Floor in reserveSeat heilt ausdrücklich nur nach OBEN (`if (realActive > current) current = realActive`), ein zu hoher Zähler bleibt stehen. Geheilt wird das erst durch den privilegierten syncSeatsToActiveCount — beim Admin-Boot, gedrosselt auf 1×/6 h (DexEventPlatform.tsx:589) — oder durch die Sync-Kette des IDReorder-Flows, die nur nach einer Abmeldung läuft.

**Fix-Vorschlag:** In EventContext.tsx nach `success === false` den reservierten Platz zurückgeben, wenn vorher tatsächlich reserviert wurde (Merker `seatReserved` + Gruppe setzen, wo reserveSeat 'reserved' lieferte). Freigabe über den vorhandenen atomaren Pfad: `adjustSeatCounterField(subsiteUrl, seatFieldFor(group), -1)` — die Methode ist derzeit private, für switchSplitGroup existiert der Rollback bereits (EventService.ts:6952ff) und kann als Vorlage dienen. Ohne Rollback bleibt die Kapazität bis zum nächsten Admin-Boot falsch.

### 80. [mittel] Rollierende Fristen rechnen mit 24-h-Tagen statt Kalendertagen — nach der Sommerzeit-Umstellung fällt die Frist auf den Vortag

**Wo:** `components/EventCreationPage.tsx:256` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Office-Tage-Reihe (Kalender-Termine, Ganztags, Start 00:00 Berlin) über Ende März, Regel „Anmeldung bis 1 Tag vor dem jeweiligen Termin". Nachgerechnet mit TZ=Europe/Berlin: Termin 27.03.2026 → Frist 26.03. 00:00 (Muster: Vortag 00:00). Termin 30.03.2026 (der Montag nach der Zeitumstellung) → Start 2026-03-29T22:00Z, minus 86400000 ms = 2026-03-28T22:00Z = **28.03. 23:00 Berlin** statt 29.03. 00:00. Für diesen einen Termin endet das Buchungsfenster damit 25 Stunden früher als bei allen Geschwistern; wer am 29.03. den 30.03. buchen will, bekommt in EventSpecificSection.tsx:517-518 (`deadlinePassed`) einen gesperrten Kalendertag mit „Anmeldefrist abgelaufen". Die Liste im Wizard (SubEventsSection.tsx:934) und die Teilnehmeransicht zeigen dazu „Frist 28.03." — ein Datum, das aus der Reihe fällt, während CapacityStep.tsx:463-467 weiterhin „Entspricht der rollierenden Regel" meldet, weil der Abgleich dieselbe fehlerhafte Rechnung benutzt. Im Herbst kippt es andersherum: Termin 26.10.2026 → Frist 25.10. **01:00** statt 00:00.

**Ursache:** `amount * 86400000` verschiebt einen absoluten UTC-Zeitpunkt um exakt 24 h. Über eine DST-Grenze hinweg entspricht das keinem Kalendertag mehr: der Berliner Offset wechselt zwischen +1 und +2, das Ergebnis liegt eine Stunde daneben — und weil die erzeugten Termine bewusst auf 00:00 stehen (toggleDaySubEventImpl in wizard/logic/wizardMisc.ts:288-289), rutscht diese Stunde über Mitternacht auf den Vortag. Der Kommentar über der Funktion beansprucht ausdrücklich „glatte Tagesgrenzen", was nur ohne DST-Wechsel stimmt.

**Fix-Vorschlag:** Bei `unit === 'days'` im Berliner Kalender rechnen statt in Millisekunden: Start über `isoToLocal(startIso)` in „YYYY-MM-DDTHH:MM" zerlegen, nur den Datumsteil um `amount` Tage verschieben (z.B. über `Date.UTC(y, m-1, d)` + `setUTCDate`) und das Ergebnis mit derselben Uhrzeit wieder durch `berlinLocalToUtcIso` schicken. `unit === 'hours'` darf weiter mit 3600000 rechnen (Stunden sind absolut gemeint). Die Rechnung gehört danach EINMAL in `utils/eventFormat.ts` und wird von EventCreationPage.tsx importiert — die Kopie in `subEventRegDeadline` (utils/eventFormat.ts:74) muss zwingend mitgezogen werden, sonst widersprechen sich der Wizard-Wert und der Anmeldeseiten-Fallback zusätzlich.

### 81. [mittel] Split-Kapazität: „Anmelden" stürzt still ab, solange starterCounts noch null ist

**Wo:** `components/registration/submitFlow.ts:492` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Event mit geteilter Kapazität (durchstarterCapacity und funstarterCapacity beide > 0). Der Ladeeffekt für die Belegung (RegistrationPage.tsx:1059) steigt bei leerem `event.subsiteUrl` oder fehlendem `window.__dexSpfxContext` sofort aus bzw. ist beim Klick noch nicht durch. Der Nutzer wählt seine Gruppe, füllt das Formular aus und klickt „Anmelden": `starterCounts` ist null → TypeError „Cannot read properties of null (reading 'durch')". Der Fehler fliegt aus der async-Funktion in eine unbehandelte Promise-Rejection; es erscheint kein setError, kein Overlay, es wird nichts gespeichert. Der Button wirkt schlicht tot, und der Nutzer klickt erneut.

**Ursache:** `starterCounts` ist `useState<{...} | null>(null)` (RegistrationPage.tsx:286) und wird nur im Erfolgsfall des Effekts gesetzt; der Effekt hat drei Ausstiege ohne Wert (`!isSplitGroup || !event?.subsiteUrl`, `if (!ctx) return`, `catch { /* ignore */ }`). Der Kontext-Typ in SubmitFlowCtx (submitFlow.ts:100) deklariert das Feld nicht-nullable, und wegen `"strict": false` in tsconfig.json meldet der Compiler die Zuweisung von null nicht. Alle Anzeigestellen greifen defensiv zu (`starterCounts?.durch ?? 0`, EventSpecificSection.tsx:112 und 156) — nur der Submit-Pfad nicht.

**Fix-Vorschlag:** Vor der Rechnung entschärfen: `const dTaken = starterCounts?.durch ?? 0; const fTaken = starterCounts?.fun ?? 0;` und damit rechnen — oder den ganzen Fallback-Block überspringen, solange `starterCounts` null ist (ohne belastbare Zahlen ist der Umsteige-Dialog ohnehin keine Aussage; die harte Grenze zieht serverseitig `reserveSeat`). Zusätzlich `starterCounts` im SubmitFlowCtx als `| null` deklarieren, damit der Compiler die Stelle künftig sieht.

### 82. [mittel] Sub-Event-Outlook-Update: Queue-Ergebnis wird ignoriert, OutlookDirty wird trotzdem auf false gesetzt

**Wo:** `components/wizard/logic/wizardSubmit.ts:1144` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Terminreihe mit vielen Sub-Events. Der Organizer ändert Zeiten und hakt im Update-Modal alle Termine an. Nach rund 20 Schreibzugriffen drosselt SharePoint (429) — `queueOutlookEvent` fängt das intern ab und liefert `false` (services/events/outlookQueue.ts:169-172), es WIRFT nicht. Der `catch` greift also nie, `OutlookDirty:false` wird für jeden dieser Termine geschrieben. Ergebnis: Für die betroffenen Termine steht kein Eintrag in DEX_Outlook, der Kalender bleibt auf dem alten Stand, und weil der Dirty-Marker weg ist, meldet der Wizard beim nächsten Öffnen keine offene Änderung mehr. Derselbe Fehler in der Aktion „Alle Termine aktualisieren" (components/wizard/logic/outlookActions.ts:189-190): dort zählt `done += 1` jeden fehlgeschlagenen POST als Erfolg und die Meldung sagt „N Termine angestoßen".

**Ursache:** `queueOutlookEvent` meldet Fehlschlag als Rückgabewert, nicht als Ausnahme — dieselbe Klasse, die in v29.21 (`updateEvent`) und v29.48 (`deleteEvent`) schon zweimal zugeschlagen hat. Hier wurde der Rückgabewert nicht ausgewertet und der Dirty-Marker vorbehaltlos gelöscht.

**Fix-Vorschlag:** Rückgabewert prüfen: `const ok = await svc.queueOutlookEvent(...); if (ok) await updateEvent(subId, { OutlookDirty: false }, { skipReload: true }); else failedOutlookTitles.push(subTitle);` — und die gescheiterten Termine am Ende namentlich melden (wie `failedSubTitles` in persistSubEvents.ts). In `triggerOutlookUpdateAllImpl` `done`/`failed` am Rückgabewert statt am `catch` festmachen.

### 83. [mittel] Vollständigkeits-Check „Unterkunftsbedarf ohne Hotel" arbeitet mit einer Volltextsuche über die ganze Zeile und meldet Unbeteiligte

**Wo:** `components/HotelPlanningPanel.tsx:993` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Ein Event hat im Formular eine beliebige Ja/Nein-Frage („Fotoeinwilligung: Yes"). Sobald die Hotel-Spalten auf der Teilnehmerliste existieren (spätestens nach der ersten Zuordnung, `ensureHotelColumns`), liefert `$select=*` bei JEDER Zeile `"Hotel":null,"HotelFrom":null,"HotelTo":null` mit — die Themen-Sperre in Zeile 997 ist damit immer erfüllt. Übrig bleibt die Suche nach „yes": Jede nicht zugeordnete Person mit irgendeinem „Yes" in irgendeiner Antwort landet in der orangen Karte „N Person(en) mit Unterkunftsbedarf ohne Hotel". Dasselbe über den Namen: „Anja Müller" enthält „ja ". Der Organizer bucht daraufhin Zimmer für Leute, die keins wollten. Umgekehrt fehlt bei einer Klammer, wer die Hotelfrage auf der Sub-Event-Zeile beantwortet hat — der Blob ist nur die (antwortlose) Klammer-Zeile.

**Ursache:** Die Heuristik stammt aus v28.39 und war als „bewusst breit" gedacht, weil es damals noch keine belastbare Deutung gab. Seit v28.48/v29.3 gibt es `wishOf(p)` (Z. 878), das Parent- und Child-Zeilen zusammenlegt, die Zusatznächte-Frage ausschließt und Ja/Nein sauber trennt — nur diese eine Auswertung wurde nie umgestellt. Zusätzlich prüft der „hotel"-Teil faktisch nicht das Formular, sondern die Spaltennamen der Zeile selbst.

**Fix-Vorschlag:** Den Block durch die vorhandene, präzise Auswertung ersetzen: `const wantsHotelWithout = React.useMemo(() => people.filter(p => !(p.Hotel || '').trim() && wishOf(p) === true), [people, wishOf]);` — damit stimmt die Karte automatisch mit der Spalte „Hotel-Wunsch" und mit `autoDistribute` überein.

### 84. [mittel] Wartelisten-Platz-Dialog rechnet mit der gefilterten Trefferzahl statt mit der echten Wartelistenlänge

**Wo:** `components/admin/participants/WaitlistTables.tsx:132` · Quelle: Audit B (nach dem Umbau)

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

### 85. [mittel] Zeitraum auf markierte Personen anwenden schreibt nichts, wenn (noch) kein Hotel gesetzt ist

**Wo:** `components/HotelPlanningPanel.tsx:925` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Der Organizer markiert 20 noch nicht zugeordnete Personen und klickt in der Markierungs-Leiste auf den Zeitraum-Knopf „⏱ Mit Vorabend" (der genau dafür da ist: „Zeiträume einmal anlegen, dann markieren und mit einem Klick zuweisen"). Der Fortschritt läuft durch, danach steht in der Spalte „Zeitraum" weiterhin „— kein Zeitraum —" und in „Nächte" ein „—". Es wurde nichts gespeichert. Hatte eine Zeile bereits Daten ohne Hotel, sind sie danach gelöscht.

**Ursache:** `applyStayToSelected` reicht das (leere) aktuelle Hotel durch. In `services/events/hotelPlanning.ts:56-60` bedeutet ein leeres Hotel ausdrücklich „Zuordnung aufheben": `'HotelFrom': hotel ? (fromIso || null) : null` — die Daten werden also verworfen statt geschrieben. Der optimistische Overlay in `writeAssignment` (Z. 810–813) macht dieselbe Rechnung, deshalb sieht die Tabelle nicht einmal kurz richtig aus, und ein Fehler wird nirgends gemeldet.

**Fix-Vorschlag:** In `applyStayToSelected` Zeilen ohne Hotel überspringen und benennen (z.B. „N Person(en) übersprungen — ohne Hotel lässt sich kein Zeitraum speichern"), oder — besser zur Bedienerwartung passend — `setHotelAssignment` so ändern, dass es nur dann löscht, wenn Hotel UND Daten leer sind, und ein leeres Hotel mit Daten die Spalten HotelFrom/HotelTo trotzdem setzt.

### 86. [mittel] Zielgruppen-Suche: langsamere ältere Antwort überschreibt die neuere Trefferliste

**Wo:** `components/AudiencePicker.tsx:532` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer tippt in der Zielgruppen-Suche „mül", macht ≥300 ms Pause → Request A startet (mit `audienceIncludeIntl` ist die Graph-Suche merklich langsam). Er tippt weiter zu „müller", pausiert erneut → Request B startet und kommt schneller zurück, die Dropdown-Liste zeigt korrekt die Müller-Treffer. Danach trifft A ein und `setAudienceResults([...g, ...u])` überschreibt die Liste mit den Treffern zu „mül" — angezeigt werden Namen, die nicht zur Eingabe im Feld passen, und ein Klick nimmt die falsche Person/Gruppe in die Zielgruppe auf. Zusätzlich blendet `setIsSearchingAudience(false)` aus A den „Suche..."-Hinweis aus, während B noch läuft.

**Ursache:** `audienceTimerRef` entprellt nur das STARTEN: ein bereits abgeschickter Request wird beim nächsten `clearTimeout` nicht verworfen. Beim Eintreffen der Antwort fehlt jede Prüfung, ob sie noch zur aktuellen Eingabe gehört (kein Sequenz-Zähler, kein Query-Vergleich, kein cancelled-Flag).

**Fix-Vorschlag:** Sequenz-Ref einführen und die Antwort verwerfen, wenn sie überholt ist: `const seq = ++audienceSeqRef.current; … if (seq !== audienceSeqRef.current) return; setAudienceResults([...g, ...u]);` — und dieselbe Prüfung in den `InternationalSearchToggle`-Zweig darunter (Z. ~559), der denselben State schreibt.

### 87. [mittel] `URL.revokeObjectURL`-Cleanup mit `[shots]`-Dependency zerstoert die Vorschau-URLs der bereits angehaengten Screenshots

**Wo:** `components/QuestionButton.tsx:122` · Quelle: Audit B (nach dem Umbau)

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

### 88. [mittel] `_requireSubEventSelection` bleibt für immer gesetzt, sobald ein Event einmal im Modus „Nur Sub-Events" war

**Wo:** `components/EventCreationPage.tsx:4625` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Klammer-Event wurde mit „Nur Sub-Events" angelegt → gespeichert werden `_subEventsOnlyMode: true` UND `_requireSubEventSelection: true`. Später stellt der Organizer in Schritt 1 den Anmelde-Modus per Radio auf „Haupt-Event + Sub-Events" um (wizard/steps/SubEventsSection.tsx:396, `onChange={() => setSubEventsOnlyMode(modeVal)}`) und speichert. `_requireSubEventSelection` wird erneut geschrieben. Folge auf der Anmeldeseite: (a) über der Sub-Event-Liste steht dauerhaft der orange Pflicht-Hinweis „Pflicht: bitte mindestens ein X auswählen" (RegistrationPage.tsx:4232), obwohl das Submit-Gate ihn seit v24.64 gar nicht mehr prüft — die Meldung ist schlicht falsch; (b) die Team-Anmeldung ist hart ausgeblendet (`isTeamCapable` … `&& !(event?.requireSubEventSelection && childEvents.length > 0)`, RegistrationPage.tsx:798) — der Organizer schaltet Team-Anmeldung ein, auf der Anmeldeseite erscheint der Toggle trotzdem nie; (c) MyEventsPage behandelt das Event weiter als `isSectionedEvent`. Es gibt keine Bedienung im Wizard, mit der man das wieder loswird.

**Ursache:** Der State wird aus `editEvent.requireSubEventSelection` initialisiert (Zeile 917) — und dieser Getter in context/EventContext.tsx:1439-1443 liefert bewusst `!!(ov._requireSubEventSelection || ov._subEventsOnlyMode)`. Damit wird ein ABGELEITETER Wert als eigener Zustand geladen und beim nächsten Save als eigener Schlüssel zurückgeschrieben. Der Kommentar an Zeile 913 sagt selbst, dass der Flag „beim Save aus dem subEventsOnlyMode-Toggle abgeleitet" werden soll — genau das tut `||` aber nicht mehr, sobald der Load ihn schon true gemacht hat. Zurückgesetzt wird er nur an einer einzigen Stelle (SubEventsSection.tsx:158, beim kompletten Deaktivieren der Sub-Events), nicht beim Moduswechsel.

**Fix-Vorschlag:** Zwei Möglichkeiten, beide klein: (1) In EventContext.tsx:1439 das `|| ov._subEventsOnlyMode` streichen — `subEventsOnlyMode` wird ohnehin separat gelesen und der Wizard/die Anmeldeseite ODERt selbst, wo es gebraucht wird. Oder (2) im Wizard `setRequireSubEventSelection(false)` an den Modus-Radio hängen (SubEventsSection.tsx:396: `onChange={() => { setSubEventsOnlyMode(modeVal); if (!modeVal) setRequireSubEventSelection(false); }}`). Variante (1) ist die saubere, weil sie die Ursache (abgeleiteter Wert wird als gespeicherter gelesen) beseitigt.

### 89. [mittel] `activeScopeIdx` wird als einziger Reiter-Index nie in Range geklemmt — Schritt 1 schreibt dann ins Hauptevent, während die Scope-Leiste ein Sub-Event als aktiv anzeigt

**Wo:** `components/EventCreationPage.tsx:6722` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Neues Event anlegen: Sub-Events aktivieren, zwei Sub-Events anlegen, beim zweiten auf „Bearbeiten" klicken (activeScopeIdx = 2). Danach in der Schritt-1-Kopfzeile auf „Demo" klicken (der Knopf steht bei `currentStep === 0` VOR dem `activeScopeIdx === 0`-Gate, Zeile ~9200 vs. ~9236) und die Variante „Sub-Event" wählen: `resetDemoVariantBaseState()` setzt `setSubEvents([])`, der Loader danach genau EIN Sub-Event. activeScopeIdx bleibt 2. Folge: `renderGlobalScopeBar` rendert mit `Math.min(2, 1) = 1`, hebt also den Sub-Event-Reiter als aktiv hervor, `scopeSub = subEvents[1]` ist aber `undefined` — Titel, Start, Ende, Beschreibung und Bild in Schritt 1 hängen dadurch wieder am TOP-LEVEL. Wer jetzt den Titel tippt, benennt das Hauptevent um, obwohl die Leiste „Networking-Dinner" als gewählte Ebene zeigt. Mit der Variante „Standard" (0 Sub-Events) ist es noch härter: die Scope-Leiste rendert gar nicht mehr (`subEvents.length === 0` → null), gleichzeitig ist der gesamte event-weite Block in Schritt 1 (Opt-in, Anmelde-Modus, Sub-Event-Liste, Entwurf/Aktivierung) an `activeScopeIdx === 0` gebunden und bleibt display:none — es gibt keine Bedienung mehr, um auf die Klammer zurückzukommen, während der Hinweis „Du bearbeitest die Grundlagen eines Sub-Events … Reiter oben" auf eine Leiste verweist, die es nicht gibt.

**Ursache:** v28.89 hat `activeScopeIdx` als fünften, gemeinsamen Index eingeführt, ihn aber nicht in die Range-Garantien von v11.57/v15.0 aufgenommen. Die vier Alt-Indizes werden geklemmt, `activeScopeIdx` nicht. Der Kommentar über `scopeSub` behauptet „Absichern gegen einen Index, der auf einen gelöschten Sub-Event zeigt" — `subEvents[i] === undefined` sichert aber nur gegen den Crash, nicht gegen den falschen Schreibort: der Fallback ist ausgerechnet der Top-Level. Die üblichen Löschwege (`removeSubEventDraft`, `toggleDaySubEvent`) rufen vorher `setScope(0)` und verdecken die Lücke; der Demo-Knopf ist der Pfad, der das nicht tut.

**Fix-Vorschlag:** `if (activeScopeIdx > subEvents.length) setScope(0);` in denselben Effect aufnehmen (und `activeScopeIdx` in die Deps). Zusätzlich `scopeSub` defensiv machen: statt stillem Fallback auf den Top-Level bei `activeScopeIdx > 0 && !subEvents[activeScopeIdx - 1]` den Scope hart auf 0 zurücksetzen. Und `resetDemoVariantBaseState()` sollte `setScope(0)` mitmachen, weil es das gesamte Sub-Event-Array ersetzt.

### 90. [mittel] activeScopeIdx wird als einziger der fünf Reiter-Indizes nie in Range geklemmt — Demo-Variante wechseln sperrt Schritt 1 aus

**Wo:** `components/EventCreationPage.tsx:6722` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Neu-Anlage (kein Edit): 1) „Demo" oben rechts im grünen Schritt-1-Kopf → Variante „Sub-Event" wählen (legt ein Sub-Event an). 2) In der Scope-Leiste auf den Sub-Event-Reiter klicken → `setScope(1)`. 3) Erneut „Demo" (der Knopf steht in BasicsStep.tsx Zeile 119, AUSSERHALB des `activeScopeIdx === 0`-Blocks und ist deshalb auch auf einem Sub-Reiter sichtbar) → Variante „Standard" wählen. `loadDemoStandard` → `resetDemoVariantBaseState()` → `setSubEvents([])`. Ergebnis: `subEvents.length === 0`, aber `activeScopeIdx` bleibt 1. `renderGlobalScopeBar()` gibt bei 0 Sub-Events null zurück (Zeile 6782) — es gibt also keinen Reiter mehr, über den man zurück auf die Klammer käme. Gleichzeitig sind `SubEventsSection visible={currentStep === 0 && activeScopeIdx === 0}` (Zeile 9319) und der gesamte event-weite Block in BasicsStep (Zeilen 154–445: eigene Vorlagen, Entwurfs-Kachel, Vorlagen-Picker, Entwurf/Aktivierung, Sub-Event-Opt-in) ausgeblendet. Der Organizer sieht in Schritt 1 nur noch Titel/Datum/Beschreibung/Bild plus den Hinweis „Du bearbeitest die Grundlagen eines Sub-Events" — obwohl es kein Sub-Event gibt — und kommt ohne Neuladen der Seite nicht mehr an Entwurfs-Schalter und Sub-Event-Sektion.

**Ursache:** `setScope` (Zeile 8160) hält activeScopeIdx, activeLocationTabIdx, activeCapacityTabIdx, activeFieldsTabIdx und activeCommTabIdx synchron. Für die letzten vier gibt es Range-Garantien (Zeilen 6712–6725), für activeScopeIdx nicht. Alle Pfade, die subEvents verkleinern, setzen bisher selbst zurück — `removeSubEventDraft` (8436) und `toggleDaySubEvent` (8346, ohnehin nur auf Scope 0 erreichbar). Die Demo-Loader (`resetDemoVariantBaseState`, Zeile 3350, `setSubEvents([])` in Zeile 3372, aufgerufen von loadDemoStandard/Groups/SubEvent/SubEventTeam, Zeilen 3508/3532/3556/3597) tun das nicht — sie setzen nur `setCurrentStep(0)`. Nebeneffekt derselben Ursache: die Comm-Garantie in Zeile 6714 feuert dann `switchCommTab(0)` und überschreibt die frisch geladenen Demo-Kommunikationswerte mit dem `topLevelCommSnapshot` von VOR dem Variantenwechsel.

**Fix-Vorschlag:** Zwei Zeilen, unabhängig voneinander sinnvoll: (a) in den Effect ab Zeile 6721 `if (activeScopeIdx > subEvents.length) setScope(0);` aufnehmen (activeScopeIdx in die Dependencies) — das ist die generische Absicherung, die den fünften Index endlich genauso behandelt wie die anderen vier; (b) zusätzlich in `resetDemoVariantBaseState()` (Zeile 3350) vor `setSubEvents([])` ein `setScope(0)` setzen, damit auch die Comm-Spiegelung sauber auf die Klammer zurückfällt, bevor die Demo-Werte geschrieben werden.

### 91. [mittel] deleteEvent räumt DEX_Participants nach dem Recyclen der Subsite auf — nicht strikt gelesen, parallel geschrieben, Fehler verworfen

**Wo:** `services/EventService.ts:3030` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Admin löscht ein Event mit 300 Anmeldungen. Schritt 1 recycelt bereits die Subsite (Zeile ~2995), erst danach läuft dieser Block. getAllParticipants liest NICHT strikt: bricht eine Seite mit HTTP 429/500 ab, liefert readAllParticipants(false) die bis dahin gelesenen Zeilen — die betroffenen Personen stehen dann gar nicht in der Liste. Die verbleibenden 300 MERGEs gehen gleichzeitig raus (Promise.all) und provozieren genau die Drosselung; jeder Fehler wird per `.catch(() => null)` weggeworfen. Ergebnis: Personen behalten die Event-Nummer in EventRegistered/EventOnWaitlist, die Teilnehmerliste ist weg, niemand erfährt davon. Die Register-Prüfung meldet die Verweise später als „verwaist" — der Rückstand, dessentwegen v29.0/v29.1 1045 Falschmeldungen produzierten.

**Ursache:** Die in v29.3 für deleteParticipantData festgelegte Reihenfolge (prüfbare Nebenbuchhaltung zuerst, sequentiell, mit Fehlerzähler, bei Fehlern abbrechen BEVOR gelöscht wird) wurde in deleteEvent nie nachgezogen — dort steht noch der Original-Code von vor v29.2. Verstärkt wird das durch Fund #2: removeParticipantEvent meldet auch abgelehnte MERGEs als Erfolg, ein Fehlerzähler würde hier also selbst nach dem Umbau noch 0 zeigen.

**Fix-Vorschlag:** Den Block vor die Recycle-Schritte ziehen und wie in deleteParticipantData bauen: fetchAllParticipantsOrThrow() (strikt), sequentielle Schleife mit registryFailed-Zähler, bei Lesefehler oder registryFailed > 0 mit `false` abbrechen, bevor Subsite/Item recycelt werden. Setzt Fund #2 voraus, sonst zählt der Zähler nichts.

### 92. [mittel] getInvitedRecipients liefert bei HTTP-Fehler stillschweigend eine Teilliste — Doppel-Einladung an den ganzen Verteiler

**Wo:** `services/events/outlookQueue.ts:126` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Ein Event mit großem Verteiler; die erste Einladungsrunde ist raus, die DEX_Emails-Queue hat mehrere hundert Zeilen (Bcc-Chunks à 450). Beim Öffnen des Einladungs-Dialogs läuft `getInvitedRecipients` in ein 429/503 (Throttling) auf Seite 1. Die Funktion wirft nicht — sie liefert `[]`. `openInviteModal` (useMailComposers.tsx:350) setzt daraufhin `invitedLc` auf ein leeres Set, der Dialog zeigt „Nur an noch nicht Eingeladene (1200) — 0 Adresse(n) fallen raus“, und der Organizer verschickt die Einladung ein zweites Mal an alle 1200 Personen. Das `.catch(...)` im Aufrufer feuert nie, weil kein Fehler geworfen wird.

**Ursache:** Exakt das in CLAUDE.md beschriebene Muster („getAllRegistrations wirft nicht“): Ein leeres/unvollständiges Ergebnis ohne geprüften Status wird als Aussage über die Daten gelesen. Der Doc-Kommentar der Funktion warnt nur vor der Archivierung alter Queue-Zeilen, nicht vor dem HTTP-Abbruch; der `guard < 20`-Abbruch bei >10.000 Zeilen hat denselben Effekt. Der Aufrufer kann „nichts gefunden“ und „Abfrage fehlgeschlagen“ nicht unterscheiden.

**Fix-Vorschlag:** Signatur auf `Promise<{ ok: boolean; emails: string[] }>` ändern (oder einen `onHttpError`-Rückruf ergänzen, wie es v29.3 für `getAllRegistrations` getan hat) und bei `ok === false` den Radio-Modus „Nur an noch nicht Eingeladene“ im Dialog sperren bzw. deutlich als „Abgleich nicht möglich“ ausweisen, statt „0 Adresse(n) fallen raus“ anzuzeigen.

### 93. [mittel] recomputeEventKpiOnly nullt den Teilnehmer-Zähler, wenn der KPI-Cache einmal nicht lesbar ist

**Wo:** `services/events/emailTemplatesList.ts:709` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Admin öffnet die App; der Effekt in DexEventPlatform.tsx:255 ruft `recomputeEventKpiOnly()`. Der Lesezugriff auf die _Config-Zeile scheitert einmalig (429/403/Netz) → `getKpiCache()` liefert `null` → `updateKpiCache` schreibt `TotalParticipantsCount = 0` in die gemeinsame _Config-Zeile. Ab sofort zeigt die Startseite allen Nutzern „0 Teilnehmer". Der Wert erholt sich nicht: Er wird laut Architektur nur noch per ±1-Bump bei An-/Abmeldungen fortgeschrieben, ein Voll-Recompute läuft nicht mehr.

**Ursache:** `getKpiCache` unterscheidet in der Rückgabe zwar zwischen `null` (Fehler) und `{participants: 0, events: 0}` (Zeile fehlt), der Aufrufer wirft die Unterscheidung mit `?? 0` aber sofort weg. Der Kommentar im Aufrufer (DexEventPlatform.tsx:236) sagt ausdrücklich, der Teilnehmer-Live-Zähler sei „FÜHREND und wird NICHT mehr überschrieben" — geschrieben wird er hier trotzdem, nur eben mit einem geratenen Wert.

**Fix-Vorschlag:** Bei nicht lesbarem Cache gar nicht schreiben: `const cache = await svc.getKpiCache(); if (cache === null) return null;` — oder `updateKpiCache` um eine Variante erweitern, die nur `TotalEventsCount` merged und `TotalParticipantsCount` unangetastet lässt (das ist ohnehin die Absicht der Methode).

### 94. [mittel] reloadSubEventRegs zeigt nach einer Abmeldung nicht lesbare Sub-Event-Listen als leer — ohne den v30.37-Warnhinweis

**Wo:** `components/admin/logic/createKlammerActions.ts:470` · Quelle: Audit B (nach dem Umbau)

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

### 95. [mittel] {{Address}} im Sub-Event-Outlook-Body ist immer leer

**Wo:** `components/wizard/logic/persistSubEvents.ts:169` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Organizer wechselt im Wizard über die Scope-Leiste auf ein Sub-Event, öffnet im Schritt Kommunikation den Outlook-Termin-Editor und klickt den angebotenen Variablen-Knopf „Adresse“ ({{Address}}, WizardModals.tsx:548). Die Live-Vorschau zeigt korrekt „Musterstr. 1, 80331 München“ (WizardModals.tsx:528 baut `Address` aus addrStreet/HouseNo/Zip/City). Nach dem Speichern steht im Outlook-Termin ALLER Teilnehmer dieses Sub-Events an dieser Stelle nichts — z. B. „Treffpunkt: “. Der Fehler ist beim Speichern eingebrannt und in der App nicht mehr sichtbar.

**Ursache:** Reine Auslassung im Sub-Event-Zweig: Der Hauptevent-Pfad (wizardSubmit.ts:392) baut `Address` aus den vier Adressteilen, dieser Zweig setzt sie hart auf ''. Das Sub-Event HAT eine eigene Adresse — sie wird 40 Zeilen weiter unten als `draftAddr = draft.locationAddress` gelesen und nach `LocationAddress` serialisiert; sie steht beim Aufbau der `vars` nur noch nicht bereit, weil `draftAddr` erst später deklariert wird.

**Fix-Vorschlag:** `draftAddr`/`draftHasAddress` vor den `vars`-Block ziehen und dieselbe Formatierung wie im Hauptevent-Pfad verwenden: `Address: [draftAddr.street, draftAddr.houseNo].filter(Boolean).join(' ') + ((draftAddr.zip || draftAddr.city) ? ', ' + [draftAddr.zip, draftAddr.city].filter(Boolean).join(' ') : '')`. Falls das Sub-Event keine eigene Adresse hat, auf die Hauptevent-Adresse zurückfallen statt auf ''.

### 96. [mittel] {{Organizer}} in Massen-/Einladungsmail rendert „Nachname, Vorname“-Mus

**Wo:** `components/admin/modals/MassmailComposerModal.tsx:79` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Event hat zwei Organisatorinnen; `event.organizers` ist `['Sathasivam, Philipp', 'Oesterle, Ines']` (Roh-Split der SP-Spalte, context/eventMapping.ts:48). Der Organizer schreibt in die Massenmail „Bei Rückfragen wende dich an {{Organizer}}.“ (Variablen-Knopf „Organizer“, MassmailComposerModal.tsx:202). Alle Teilnehmer bekommen: „…wende dich an Sathasivam, Philipp, Oesterle, Ines.“ — Vor- und Nachnamen sind nicht mehr auseinanderzuhalten. Auch mit nur EINER Person steht dort „Sathasivam, Philipp“ statt „Philipp Sathasivam“. Identisch in components/admin/modals/InviteComposerModal.tsx:158 (Einladungs-Rundmail an den ganzen Verteiler).

**Ursache:** Beide Composer bauen die Organizer-Liste selbst mit `join(', ')`, statt den dafür vorhandenen Helfer `formatOrganizerList(organizers, lang)` (context/eventTextHelpers.ts) zu nutzen, der genau diese Umsortierung („Nachname, Vorname“ → „Vorname Nachname“, verbunden mit „und“/„and“) macht. Der Helfer wird an allen anderen Mail-Stellen verwendet (EventContext.tsx:872, wizardSubmit.ts:390, useWaitlistActions.ts:214) und sogar in der Wizard-Vorschau (WizardModals.tsx:525) — nur diese beiden Versandwege sind daran vorbeigelaufen.

**Fix-Vorschlag:** In beiden Dateien `const orgNames = formatOrganizerList(selectedEvent.organizers || [], selectedEvent.emailLanguage || 'EN') || (selectedEvent.organizers || []).join(', ');` verwenden (Import aus `context/eventTextHelpers` bzw. `context/EventContext`, wie in useWaitlistActions.ts).

### 97. [mittel] „Aktuell registriert" zählt im Klammer-Modus die Warteliste mit, die KPI-Kachel daneben nicht

**Wo:** `components/AdminPage.tsx:7464` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Klammer-Event mit 100 angemeldeten und 20 nur-auf-Warteliste-stehenden Personen. In der Detail-Karte steht „Aktuell registriert: 120", in der KPI-Kachel direkt darunter steht „Angemeldet: 100". Zwei Zahlen für dieselbe Aussage, auf demselben Bildschirm, 20 auseinander — der Organizer, der die Zahl an F&A oder ans Catering meldet, greift die falsche ab.

**Ursache:** `consolidatedFiltered` leitet sich von `consolidatedRows` ab, und dort enthält die Statusliste ACTIVE (Zeile 5248) bewusst auch `'Warteliste'` — die Matrix soll Wartende ja als „W" zeigen. Die KPI-Kachel füllt dagegen `consolidatedActiveByEmail` (Zeile 9301) nur aus `Angemeldet | QR versendet | Eingecheckt`. Die Detail-Karte hat einfach die vorhandene Zeilenmenge der Matrix benutzt, statt dieselbe Statusliste wie die Kachel anzuwenden. Im Nicht-Klammer-Fall stimmen beide, weil `activeRegs` die Warteliste ausschließt — deshalb fällt es nur bei Klammer-Events auf.

**Fix-Vorschlag:** In der Detail-Karte dieselbe Zählung wie die KPI-Kachel verwenden, also im Klammer-Modus die aktiven Zeilen entdoppelt zählen: `isConsolidatedMode ? consolidatedFiltered.filter(r => consolidatedChildren.some(ch => { const cr = r.perChild[ch.id]; return !!cr && ['Angemeldet','QR versendet','Eingecheckt'].indexOf(cr.Status) >= 0; })).length : activeRegs.length` — oder, einfacher und robuster, die bereits gebildete Menge `consolidatedActiveByEmail` in eine Variable außerhalb des KPI-IIFE heben und an beiden Stellen benutzen.

### 98. [mittel] „Automatisch füllen" verwirft Personalnummern, die während des Verzeichnis-Abrufs von Hand eingetippt wurden

**Wo:** `components/FACenterPage.tsx:434` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** F&A öffnet ein Event mit 100 Teilnehmern und klickt „Automatisch füllen". Der Abruf läuft als 5 Graph-Batches und dauert spürbar; die Eingabefelder bleiben derweil bedienbar (nur die beiden Knöpfe sind über `pnAutoBusy` gesperrt). Die Sachbearbeiterin tippt in der Wartezeit zwei Nummern von Hand ein. Wenn die Promise auflöst, ersetzt `setPnDraft(next)` den Entwurf durch eine Kopie des Standes VOM KLICK plus die Verzeichnistreffer — die beiden getippten Nummern sind weg, ohne Meldung. Sie stehen auch nicht in `next`, weil das Verzeichnis für diese Personen gerade nichts geliefert hat (deshalb wurden sie ja von Hand nachgetragen).

**Ursache:** Klassische veraltete Closure: `pnDraft` wird beim Rendern des Buttons gelesen und im `.then()` als Basis benutzt, statt den funktionalen Updater zu verwenden. Der Merge-Zweig darunter denkt bereits an „von Hand Korrigiertes nicht überschreiben" (`cur.personalNr !== undefined ? …`), greift aber nur auf denselben veralteten Stand zu.

**Fix-Vorschlag:** Funktional aktualisieren: `setPnDraft(prev => { const next = { ...prev }; for (const r of rows) { … } return next; });` — dann ist `prev` immer der Stand zum Zeitpunkt des Schreibens. Ergänzend die beiden Eingabefelder während `pnAutoBusy` auf `readOnly` setzen, damit gar nicht erst der Eindruck entsteht, die Eingabe sei sicher.

### 99. [mittel] „Felder reparieren" verdoppelt Auswahloptionen bei jedem Klick (Guard prüft anderen Text, als er einfügt)

**Wo:** `components/AdminPage.tsx:9189` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Event mit einem Select-Feld „T-Shirt-Größe" (Optionen S/M/L). 1. Klick auf Aktionen → „Felder reparieren": Optionen werden zu ['Ohne T-Shirt','S','M','L']. 2. Klick: `hasNo` sucht nach 'kein' — 'ohne t-shirt' enthält kein 'kein' → false → ['Ohne T-Shirt','Ohne T-Shirt','S','M','L']. Jeder weitere Klick hängt eine weitere Kopie vorn an; Teilnehmer sehen im Anmeldeformular ein Dropdown mit mehrfach identischen Einträgen. Noch schlimmer bei einem B2Run-Event: das Feld „Deloitte-Laufshirt" erfüllt `isShirt` ('deloitte-laufshirt' enthält 'shirt') UND den Laufshirt-Zweig. Schon im SELBEN Lauf bekommt es erst 'Ohne T-Shirt' vorangestellt, dann — weil weiterhin kein 'kein' vorkommt — zusätzlich noch einmal 'Habe bereits ein Laufshirt', obwohl das bereits Option 1 war. Ergebnis nach EINEM Klick: ['Habe bereits ein Laufshirt','Ohne T-Shirt','Habe bereits ein Laufshirt','XS','S',…]. Parallel flippt derselbe Lauf `required` erst auf false (Shirt-Zweig) und dann zurück auf true (Laufshirt-Zweig) und meldet beides widersprüchlich in `changes` („… -> optional" und „… als Pflichtfeld markiert").

**Ursache:** Der Idempotenz-Guard prüft ein Schlüsselwort ('kein'), das in keinem der beiden eingefügten Optionstexte vorkommt ('Ohne T-Shirt', 'Habe bereits ein Laufshirt'). Der Guard kann seinen eigenen Effekt also nie erkennen — die Aktion ist nicht idempotent, obwohl sie als wiederholbares Aufräumen gedacht ist. Zusätzlich überlappen die Bedingungen `isShirt` (Substring 'shirt') und der `b2run_laufshirt`-Zweig, weil 'Laufshirt' das Wort 'shirt' enthält; beide schreiben nacheinander in dasselbe `nf.options`/`nf.required`.

**Fix-Vorschlag:** Gegen den TATSÄCHLICH eingefügten Wert prüfen, z.B. `const NO_SHIRT = 'Ohne T-Shirt'; if (!opts.some(o => o.trim().toLowerCase() === NO_SHIRT.toLowerCase() || o.toLowerCase().indexOf('kein') >= 0)) opts.unshift(NO_SHIRT);` — analog mit 'Habe bereits ein Laufshirt' in Zeile 9233. Außerdem den `isShirt`-Zweig gegen Laufshirt-Felder abgrenzen (`&& !/laufshirt/i.test(label) && nf.id !== 'b2run_laufshirt'`), damit `required` nicht zweimal umgeschrieben wird und ein Laufshirt-Feld keine „Ohne T-Shirt"-Option bekommt. Zur Sicherheit vor dem Speichern zusätzlich `nf.options = Array.from(new Set(opts))`.

### 100. [mittel] „Freie Plätze mit Warteliste füllen" überbucht bei gemeinsamer Warteliste eine einzelne Gruppe

**Wo:** `components/admin/logic/useWaitlistActions.ts:288` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Split-Event mit `splitSharedWaitlist = true`. Gruppe A (Durchstarter) cap 30, 30 aktiv (voll). Gruppe B (Funstarter) cap 30, 25 aktiv. Auf der gemeinsamen Warteliste stehen 8 Personen, die ersten 6 davon mit `PreferredStarterType = 'Durchstarter'`. Der Organizer klickt „Freie Plätze mit Warteliste füllen": `cap = 60`, `active = 55`, `free = 5`, `count = 5`. Fünfmal wird `promoteFirstWaitlistItem(sub, undefined, undefined, undefined)` gerufen — ohne Typfilter, also die ersten fünf nach TeilnehmerID. Jede bekommt `StarterType = firstWaiting.PreferredStarterType = 'Durchstarter'` (waitlist.ts:112-115). Ergebnis: Gruppe A steht auf 35/30, Gruppe B unverändert auf 25/30. Die Anmeldeseite zeigt für A „ausgebucht", die Überbuchungs-Box im Organizer Center flaggt fünf Personen als „über Kapazität", und `detectOverbooking` (overbooking.ts:54) markiert sie mit `OverbookReview='Pending'` — durch eine Aktion, die der Organizer gerade selbst ausgelöst hat. Die Gruppengrenze wird bei der regulären Anmeldung sehr wohl erzwungen (EventContext.tsx:731-734 reserviert immer gegen die EINZELNE Gruppenkapazität, unabhängig von `splitSharedWaitlist`) — nur dieser Nachrück-Pfad kennt sie nicht.

**Ursache:** Bei `splitSharedWaitlist` wird die Warteliste zu einem Topf zusammengefasst und als Obergrenze nur die SUMME beider Kapazitäten geprüft. Die Gruppenzuordnung der nachrückenden Person bleibt aber erhalten (`StarterType` wird aus ihrem `PreferredStarterType` gesetzt) und die Gruppenkapazitäten gelten weiter — an jeder anderen Stelle (`reserveSeat`, `detectOverbooking`, Gruppen-Karten der Anmeldeseite). Gemeinsame Warteliste heißt „gemeinsame Reihenfolge", nicht „gemeinsamer Kapazitätstopf".

**Fix-Vorschlag:** Im `!perGroup`-Zweig bei `isSplitCapacity` nicht über die Summe rechnen, sondern die freien Plätze je Gruppe ermitteln und die Reihenfolge aus der gemeinsamen Warteliste nehmen: freie Plätze `freeA = capA - activeA`, `freeB = capB - activeB` bestimmen und je Nachrück-Schritt `promoteFirstWaitlistItem(sub, undefined, undefined, <Gruppe, die noch frei ist>)` aufrufen — d.h. den Typfilter auch bei gemeinsamer Warteliste setzen, solange nur eine der beiden Gruppen noch Platz hat. Alternativ (falls fachlich gewünscht) müsste `reserveSeat` bei `splitSharedWaitlist` ebenfalls gegen die Summe prüfen — dann aber konsistent an allen Stellen, inklusive der Gruppen-Karten der Anmeldeseite.

### 101. [mittel] „Nur Sub-Events"-Modus: die letzte gebuchte Anmeldung lässt sich nie abwählen

**Wo:** `components/registration/RegistrationActionBar.tsx:100` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Office-Tage-Event (subEventsOnlyMode, Kalender oder Liste). Teilnehmer hat Tag A und Tag B gebucht; beide sind beim Öffnen vorausgewählt. Wählt er NUR A ab und sendet ab, wird A korrekt abgemeldet (Cancel-Pfad in submitFlow.ts:1069). Wählt er danach auch B ab, um sich komplett abzumelden, wird der Anmelden-Button DEAKTIVIERT und trägt den Text „Bitte mindestens ein Event auswählen" — obwohl der Nutzer bewusst nichts mehr wählen will. Die letzte Anmeldung ist über die Anmeldeseite nicht kündbar; in einem Event OHNE subEventsOnlyMode funktioniert dasselbe Abwählen dagegen.

**Ursache:** `nothingPicked` prüft ausschließlich `selectedSessions.size === 0` (Stand v15.11) und kennt den mit v28.88 eingeführten Fall „leere Auswahl ist eine Änderung" nicht. Die Erkennung dafür existiert bereits (`sessionsChanged` / `nothingToSubmit`, RegistrationPage.tsx:1344/1350) und wird für `alreadyDone` auch benutzt — aber `nothingPicked` läuft daran vorbei und ist in `isDisabled` mit ODER verknüpft. Dieselbe Lücke ein zweites Mal im Submit selbst: submitFlow.ts:256 bricht bei `selectedSessions.size === 0` mit derselben Meldung ab, ohne `sessionsChanged` zu berücksichtigen. Der Guard nennt damit einen Grund („du musst etwas auswählen"), der auf die Absicht des Nutzers gar nicht zutrifft.

**Fix-Vorschlag:** `sessionsChanged` in die Props der ActionBar aufnehmen und `const nothingPicked = isSubOnly && selectedSessions.size === 0 && !sessionsChanged;` setzen; in submitFlow.ts:256 die Bedingung um `&& !sessionsChanged` ergänzen. Damit greift der Guard weiter für eine frische Anmeldung ohne Auswahl, blockiert aber nicht mehr die vollständige Abmeldung.

### 102. [mittel] „Prüfen"-Dialog zeigt die Mail-Einstellungen des aktiven Sub-Event-Reiters als die des Hauptevents

**Wo:** `components/EventCreationPage.tsx:10258` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Neues Event mit einem Sub-Event. Der Organizer geht in Schritt 6 (Kommunikation), klickt in der Scope-Leiste auf den Reiter des Sub-Events und schaltet dort „Bestätigungs-Mails" aus (oder der Reiter hat schon abweichende Werte gespeichert). Danach klickt er unten auf „Prüfen". Der Dialog meldet unter der Überschrift „Schritt 6 — Kommunikation" — also für das HAUPTEVENT — „Bestätigungs-Mails: deaktiviert", obwohl für das Hauptevent die Mails aktiv sind. Dasselbe gilt für Mail-Sprache, Outlook-Termin, Auto-Abmeldung und „Person nicht mehr bei Deloitte". Umgekehrt genauso: Wer die Mails am Hauptevent abgeschaltet hat und vor dem Prüfen versehentlich auf einen Sub-Reiter geklickt hat, bekommt „aktiv" gemeldet und legt das Event im guten Glauben an.

**Ursache:** Die Kommunikations-States (emailLanguage, disableEmails, disableRegistrationEmail, disableCancellationEmail, disableOutlook, autoDeregisterOnDecline, inactiveHandling) sind Spiegel-States: `switchCommTab(idx)` (Zeile 4338 ff.) schreibt beim Reiterwechsel den bisherigen Stand in den Slot und lädt die Werte des Ziel-Reiters IN DIESELBEN Variablen. Steht ein Sub-Reiter aktiv, halten sie die Sub-Event-Werte. Genau dafür existiert `resolveTopLevelCommState()` (Zeile 4480), das `handleSubmitInner` und `persistSubEventsForParent` benutzen — der Prüfen-Dialog liest die Variablen dagegen roh. Er ist nach v22.36 dazugekommen und wurde nie an den v11.93-Resolver angeschlossen. (Genau die Falle „Kommunikationsfelder der Sub-Events liegen nicht laufend im Draft" aus CLAUDE.md.)

**Fix-Vorschlag:** Im IIFE des Prüfen-Dialogs einmal `const topComm = resolveTopLevelCommState();` bilden (die Funktion ist rein, sie liest nur State + `topLevelCommSnapshot.current`) und die sechs Zeilen der Kommunikations-Sektion auf `topComm.emailLanguage`, `topComm.disableEmails`, `topComm.disableRegistrationEmail`, `topComm.disableCancellationEmail`, `topComm.disableOutlook`, `topComm.autoDeregisterOnDecline`, `topComm.inactiveHandling` umstellen. Optional zusätzlich je Sub-Event eine Zeile aus `subEventsRef.current` (dort stehen die geflushten Sub-Werte) — dann prüft der Dialog wirklich das ganze Event.

### 103. [niedrig] Ausgeblendete Spalten der Teilnehmer-Tabelle gehen dauerhaft verloren, weil die Config vor dem Laden der Anmeldungen beschnitten zurückgeschrieben wird

**Wo:** `components/admin/logic/useColumnConfig.ts:233` · Quelle: Audit A (Gesamtcode) · Einschaetzung: wahrscheinlich

**Szenario:** Organizer blendet in der Teilnehmer-Tabelle eine der BEDINGTEN Spalten aus, z.B. „Nachgerückt am" (`promotedDate`, hängt an `hasWaitlistActivity`) oder „Startnummer" (hängt an `registrations.some(r => r.Startnummer)`). Beim nächsten Öffnen des Events ist `registrations` noch leer (`handleSelectEvent` setzt erst `selectedEvent` und lädt danach), also fehlt die Spalte in `availableColumns`. Der Lade-Effekt filtert sie deshalb aus `order` UND aus `hidden` heraus (Z.233), und der Persist-Effekt (Z.240-245) schreibt den beschnittenen Stand sofort nach localStorage. Wenige Sekunden später kommen die Anmeldungen, `availableColumns` wächst, der Lade-Effekt läuft erneut, findet die Spalte nun in `missing` und hängt sie SICHTBAR vor „Aktion" ein. Die Ausblendung ist damit weg — und lässt sich nicht dauerhaft setzen, weil derselbe Ablauf bei jedem Öffnen greift.

**Ursache:** Lade- und Persist-Effekt hängen beide an `availableColumns`, das seinerseits von den noch nicht geladenen `registrations` abhängt. Der Persist-Effekt kann nicht unterscheiden zwischen „der User hat eine Spalte umgestellt" und „der Loader hat gerade auf einen unvollständigen Spaltensatz reduziert", und schreibt den Zwischenstand als Wahrheit fest.

**Fix-Vorschlag:** Unbekannte IDs beim Laden NICHT verwerfen (order/hidden unverändert übernehmen und erst beim Rendern gegen `availableColumns` filtern), oder den Persist-Effekt an ein Dirty-Flag hängen, das erst `hideColumn`/`showColumn`/`moveColumn` setzt — dann schreibt nur eine echte Benutzeraktion nach localStorage.

### 104. [niedrig] DocumentsViewer: Blob-URL-Cleanup greift wegen leerem dep-Array nie

**Wo:** `components/myEvents/DocumentsViewer.tsx:109` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Teilnehmer öffnet in „Meine Events" nacheinander mehrere Event-Dokumente (PDFs, teils mehrere MB). Jedes `toggleDoc` lädt die Datei per XHR und ruft `setBlobUrl(URL.createObjectURL(correctBlob))` (Z.98) — die vorherige URL wird dabei nicht freigegeben. Beim Verlassen der Seite läuft zwar die Cleanup, sie liest `blobUrl` aber aus der Closure des ERSTEN Renders, wo der Wert noch leer ist; also wird auch die letzte URL nicht freigegeben. Der Blob-Speicher wächst mit jedem geöffneten Dokument bis zum Schließen des Tabs — auf dem Handy, für das diese Komponente gerade gebaut wurde (react-pdf-Pfad Z.95), ist das der schmerzhafte Fall.

**Ursache:** `[]` friert die Closure auf den Initialwert ein — der Wert, den die Cleanup prüfen soll, entsteht erst später. Zusätzlich fehlt das Freigeben der alten URL beim Ersetzen.

**Fix-Vorschlag:** `}, [blobUrl]);` — hier ist das genau richtig: die Cleanup läuft dann beim Wechsel mit der ALTEN URL, die ab diesem Moment tatsächlich niemand mehr rendert (im Unterschied zum QuestionButton-Fall, wo das alte Array noch benutzte URLs enthält).

### 105. [niedrig] Doppelter Default-Eintrag OutlookDeclineReminder_OnBehalfOf/EN beim Erstanlegen der Template-Liste

**Wo:** `services/events/emailTemplatesList.ts:147` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Auf einem frischen Tenant legt `ensureEmailTemplatesList` die Liste DEX_EmailTemplates an und schreibt das `defaults`-Array ohne Existenzprüfung Zeile für Zeile. Der EN-Eintrag für `OutlookDeclineReminder_OnBehalfOf` steht zweimal darin (Zeile 120 und Zeile 147) — es entstehen zwei identische Zeilen. Im globalen Vorlagen-Editor (getAllEmailTemplates, ohne Dedup) sieht der Admin die Vorlage doppelt; ändert er die eine, bleibt die andere alt, und welche gewinnt, entscheidet `$top=1` in `getEmailTemplate` — also die kleinere Item-Id, nicht die zuletzt bearbeitete.

**Ursache:** Copy-Paste im `defaults`-Array: an Position 147 sollte offenkundig ein anderer Eintrag stehen (die DE-Variante folgt separat in Zeile 149), stattdessen wurde der EN-Block aus Zeile 120 wiederholt. Die beiden anderen Arrays derselben Datei (`newTemplates` ab Zeile 250, `standards` ab Zeile 333) enthalten den Eintrag jeweils korrekt genau einmal — und sie prüfen zusätzlich auf Existenz, weshalb der Fehler nur beim Erstanlegen zuschlägt.

**Fix-Vorschlag:** Zeile 147-148 ersatzlos streichen. Zusätzlich absichern: die Seeding-Schleife für `defaults` auf `TemplateType+Language` deduplizieren, bevor sie POSTet (z.B. über ein `Set` auf `${t.TemplateType}_${t.Language}`), damit ein künftiger Doppeleintrag keine Doppelzeile mehr erzeugt.

### 106. [niedrig] Fehlender englischer Übersetzungsschlüssel: Englische Nutzer sehen „EVENTLIST.VIEW" über dem Ansichts-Umschalter

**Wo:** `context/LanguageContext.tsx:38` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Nutzer mit Locale „en" (Default der App!) öffnet die Event-Übersicht. Über den Karten/Liste-Umschaltern steht statt „View" der rohe Schlüssel: EventListPage.tsx:386 rendert `{t('eventlist.view') || 'Ansicht'}` in Großbuchstaben (textTransform:'uppercase') — also „EVENTLIST.VIEW".

**Ursache:** `t()` liefert bei fehlendem Schlüssel `translations[locale][key] || translations['en'][key] || key`, für locale='en' also den Schlüssel selbst. Der ist ein truthy String, deshalb greift der `|| 'Ansicht'`-Fallback an der Aufrufstelle nie. Ein Schlüsselabgleich beider Blöcke zeigt genau diesen einen Ausreißer (de: 224 Schlüssel, en: alle bis auf 'eventlist.view').

**Fix-Vorschlag:** Im en-Block `'eventlist.view': 'View',` ergänzen (neben 'eventlist.onlyactive'). Zusätzlich erwägen, `t()` bei fehlendem Schlüssel im Debug-Modus zu loggen, damit solche Lücken nicht erst im Bild auffallen.

### 107. [niedrig] Fortschrittsbalken der Massenzuordnung wird nie angezeigt bzw. verschwindet nach der ersten Person

**Wo:** `components/HotelPlanningPanel.tsx:1650` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Der Organizer klickt „Automatisch verteilen" und bestätigt 200 Personen. `setBulkProgress` läuft, aber die Markierung ist leer, also rendert der einzige Fortschrittsbalken der Seite gar nicht — 200 sequentielle MERGE-Requests laufen minutenlang ohne jede Rückmeldung, die Seite wirkt eingefroren, der Organizer klickt erneut. Dasselbe bei „Ganzes Sub-Event zuordnen". Und bei „N markiert → Hotel X" ist der Balken genau eine Person lang zu sehen: `writeAssignment` ruft in Zeile 831 `setSelected(new Set())`, damit fällt der komplette Markierungs-Block samt Balken aus dem Baum, während die Schleife weiterläuft.

**Ursache:** Der Balken wurde in v28.48 in die Markierungs-Leiste gebaut, als es nur „markieren → zuordnen" gab. Die späteren Massenpfade (autoDistribute v28.56, assignWholeSub v28.51) setzen zwar `bulkProgress`, ihr Container hängt aber weiter an `selected.size > 0`; zusätzlich leert das Schreib-Primitiv die Markierung schon beim ersten Durchlauf.

**Fix-Vorschlag:** Den `{bulkProgress && …}`-Block aus der Markierungs-Leiste herausziehen und direkt unter die Überschrift „Zuordnung" (bzw. in die Kopfkarte) hängen, so dass er allein von `bulkProgress` abhängt. Zusätzlich `setSelected(new Set())` aus `writeAssignment` entfernen und in die aufrufenden Einzel-/Massen-Handler verschieben (nach Abschluss der Schleife).

### 108. [niedrig] Kalender-Kachel: Tooltip erfindet „null von N Plätzen frei"

**Wo:** `components/registration/EventSpecificSection.tsx:561` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Kalender-Event (subEventCalendar) mit Kapazität je Tag. Für einen Teilnehmer, der weder den Platzzähler (SeatsTaken nie geschrieben) noch die Teilnehmerliste lesen darf, liefert `occupancyOf` bewusst null. Die Kachel zeigt korrekt „—", der Hover-Tooltip daneben behauptet aber „null von 80 Plätzen frei".

**Ursache:** `free` ist absichtlich `number | null` (Zeile 555, v30.62: „unbekannt ist keine Zahl"). Im `title`-Array wird der Wert jedoch ohne Null-Prüfung in den Template-String interpoliert; nur die sichtbare Beschriftung darunter prüft `free !== null` (Zeile 633). Genau der Widerspruch, den v30.62 beseitigen sollte — nur eine Ebene tiefer stehen geblieben.

**Fix-Vorschlag:** Im title-Array denselben Dreiweg ziehen wie bei der Beschriftung: `hasCap && free !== null` → „X von N Plätzen frei"; `hasCap && free === null` → „N Plätze · Belegung nicht ermittelbar"; sonst „Unbegrenzte Plätze".

### 109. [niedrig] Meldung „x von y Personen übernommen" bleibt nach dem Eventwechsel stehen, obwohl der Entwurf verworfen wurde

**Wo:** `components/FACenterPage.tsx:221` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** F&A füllt bei Event A die Personalnummern automatisch („7 von 12 Personen aus dem Verzeichnis übernommen. Nicht vergessen: unten speichern."), geht ohne zu speichern über „Zurück zur Übersicht" in Event B. Dort steht dieselbe grüne Meldung unter der Teilnehmertabelle von B — die Tabelle zeigt aber die gespeicherten (leeren) Werte, weil `pnDraft` geleert wurde. Die Aufforderung „unten speichern" bezieht sich auf Daten, die es nicht mehr gibt; der Speichern-Knopf ist zudem deaktiviert (`Object.keys(pnDraft).length === 0`), was den Widerspruch komplett macht.

**Ursache:** Der Aufräum-Effect setzt `pnEventId` und `pnDraft` zurück, `pnAutoNote` aber nicht — der Hinweistext gehört zum selben Entwurf und hat denselben Lebenszyklus.

**Fix-Vorschlag:** Im selben Effect `setPnAutoNote('')` ergänzen (und der Vollständigkeit halber `setOpenMailIdx(null)`, das beim Wechsel über die Tabelle ebenfalls stehen bleibt).

### 110. [niedrig] Organizer-Tour: letzter Wizard-Schritt zeigt Admins ins Leere (Anlegen-Knopf existiert dort nicht mehr)

**Wo:** `components/tutorial/tutorialTours.ts:243` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Ein Admin (oder eine F&A-Rolle) startet das Organizer-Tutorial. Bei der letzten Wizard-Karte stellt die Tour den Assistenten auf Index 8 (Fun-Zone) und sucht `[data-tour="wizard-submit"]`. Für diese Person hat `steps` aber zehn Einträge (Schritt 10 „Abrechnung"), und der Anlegen-Knopf rendert nur im else-Zweig von `currentStep < steps.length - 1` (EventCreationPage.tsx Zeile 9577/9587) — auf Index 8 steht dort „Weiter". Das Element existiert also gar nicht: TutorialGuide pollt 25 × 160 ms (~4 s) vergeblich, der Spotlight bleibt aus, die Karte „Und was passiert beim Anlegen?" erscheint zentriert über einem Knopf, den sie nicht zeigt — und Schritt 10 kommt in der Tour überhaupt nicht vor. Für Nicht-Admins (9 Schritte) funktioniert dieselbe Karte korrekt, deshalb fällt es beim Testen mit einem Organizer-Konto nicht auf.

**Ursache:** v29.66 hat Schritt 10 ans Ende des `steps`-Arrays gehängt und dabei bewusst alle festen Indizes stabil gelassen; `steps.length` wurde in EventCreationPage überall mitgezogen (Weiter/Anlegen-Umschaltung, Speichern-Knopf, Fortschrittslinie), die Tour-Definition in tutorialTours.ts aber nicht. Die Karte hängt an der ANNAHME „letzter Schritt = Index 8", die seit v29.66 nur noch für Personen ohne Abrechnungs-Schritt gilt. Die Klemmung in EventCreationPage (`Math.min(detail, canBilling ? 9 : 8)`, Zeile 7146 f.) korrigiert nur nach oben, nicht nach unten.

**Fix-Vorschlag:** In tutorialTours.ts Zeile 243 `wizardStep: 8` → `wizardStep: 9`. Das trägt beide Fälle: Admins landen auf Index 9 (Abrechnung, letzter Schritt → Anlegen-Knopf ist da), bei allen anderen klemmt der Empfänger-Effect die 9 auf 8 (Fun-Zone, dort ebenfalls letzter Schritt → Knopf ist da). Ergänzend: die Texte „alle 9 Schritte" (Zeilen 133/134, 156/157 und `descDe`/`descEn`) rollenabhängig formulieren oder neutral („alle Schritte").

### 111. [niedrig] Outlook-Body des Hauptevents bekommt die Sprache des offenen Sub-Reiters (roher `emailLanguage` statt `effEmailLanguage`)

**Wo:** `components/EventCreationPage.tsx:4890` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Event mit Teams-Link und Sub-Events; Hauptevent-Mailsprache DE, ein Sub-Event auf EN gestellt. Der Organizer steht beim Speichern in Schritt 6 auf dem EN-Sub-Reiter. In den `OutlookBody` des HAUPTEVENTS wird dann der Teams-Beitreten-Block (`teamsJoinBlockHtml(teamsLink, teamsIsDe)`, services/EmailTemplates.ts:1135) auf Englisch gebacken, während der restliche Termin-Text deutsch bleibt — der Teilnehmer bekommt einen zweisprachigen Kalendereintrag. Umgekehrt genauso (Hauptevent EN, Sub DE). Betrifft sowohl den Update-Pfad (Zeile 4890) als auch den Create-Pfad (Zeile 5854).

**Ursache:** Dieselbe v11.93-Falle wie oben, hier nur beim letzten Argument stehengeblieben: `resolvedBody`, `resolvedOlHeading`, `effOutlookLogo` und `effOutlookSubject` in derselben Zeile kommen bereits aus dem Top-Level-Resolver, die Sprache aber direkt aus der State-Variable, die auf einem Sub-Reiter den Sub-Wert trägt. Der erklärende Kommentar steht unmittelbar darunter — die Zeile wurde beim Umbau schlicht nicht mitgezogen.

**Fix-Vorschlag:** An beiden Stellen `(emailLanguage || '')` durch `(effEmailLanguage || '')` ersetzen; `effEmailLanguage = topComm.emailLanguage` ist in beiden Zweigen bereits definiert und wird direkt daneben für `defaultOutlookBody` und `formatOrganizerList` verwendet.

### 112. [niedrig] Sommerzeit-Ende: Auto-Korrektur der Hotel-Abreise landet auf demselben Tag → 0 Nächte

**Wo:** `components/StayRangePicker.tsx:109` · Quelle: Audit A (Gesamtcode) · Einschaetzung: sicher

**Szenario:** Teilnehmer füllt auf der Anmeldeseite das Hotel-Feld aus und wählt als Anreise Sonntag, 25.10.2026 (die Nacht, in der die Sommerzeit endet). Weil noch keine Abreise gesetzt ist, soll der Picker automatisch den Folgetag eintragen. `d` ist 25.10. 00:00 Berlin (= 24.10. 22:00 UTC); +86 400 000 ms ergibt 25.10. 23:00 Berlin, weil dieser Kalendertag 25 Stunden hat. `toLocalDay` liefert also wieder „2026-10-25". Anreise = Abreise → `stayNights` gibt 0 → das Feld zeigt „0 Nächte", der Wert ist als Zeitraum unbrauchbar, und die Hotelplanung (HotelPlanningPanel/`nightsBetween`) zählt für diese Person null Nächte. Verifiziert mit TZ=Europe/Berlin: 25.10.2026 → from 2026-10-25, to 2026-10-25 (zum Vergleich 15.09.2026 → 2026-09-16, korrekt).

**Ursache:** `d` ist ein Date auf lokaler Mitternacht (aus dem react-datepicker bzw. `dayToDate`, das `new Date("YYYY-MM-DDT00:00:00")` ohne Z parst — also Browser-Lokalzeit). Auf einen LOKALEN Zeitpunkt exakt 24 h zu addieren ist nur außerhalb der DST-Wechsel dasselbe wie „ein Kalendertag weiter": Am Rückstelltag hat der Tag 25 h, am Umstelltag im März 23 h. Der Rest der Codebasis rechnet Tagesschritte bewusst in UTC und ist damit immun — `HotelPlanningPanel.tsx:86` und `HotelSetupWizard.tsx:87` bauen `addDays` über `Date.parse(day + 'T00:00:00Z') + n*86400000`. Nur diese eine Stelle mischt lokalen Zeitpunkt und Millisekunden-Arithmetik.

**Fix-Vorschlag:** Den Tagesschritt kalendarisch statt in Millisekunden rechnen — entweder wie im Rest der Datei über UTC (`new Date(Date.parse(from + 'T00:00:00Z') + 86400000).toISOString().slice(0,10)`, d.h. den bereits vorhandenen `addDays`-Helfer aus HotelPlanningPanel wiederverwenden statt einer zweiten Variante) oder lokal über `const cand = new Date(d); cand.setDate(cand.getDate() + 1);`, was den DST-Sprung korrekt überspringt. Der UTC-Weg ist vorzuziehen, weil `stayNights`/`nightsBetween` ohnehin über `T00:00:00Z` rechnen.

### 113. [niedrig] StayRangePicker: automatische Abreise per +86400000 ms erzeugt in der Nacht der Zeitumstellung einen 0-Nächte-Zeitraum

**Wo:** `components/StayRangePicker.tsx:109` · Quelle: Audit B (nach dem Umbau)

**Szenario:** Nachgerechnet mit TZ=Europe/Berlin: Ein Teilnehmer öffnet auf der Anmeldeseite (bzw. der Organizer in der Hotelplanung) den Hotel-Zeitraum und wählt als Anreise den 25.10.2026 — die Nacht, in der die Uhr von 3 auf 2 zurückgestellt wird. `dayToDate('2026-10-25')` = 25.10. 00:00 MESZ; `+ 86400000 ms` landet auf 25.10. **23:00** MEZ, also demselben Kalendertag. `toLocalDay(cand)` liefert „2026-10-25", `to === from`, `stayNights` = 0. Das Feld zeigt „0 Nächte", und gespeichert wird der Wert „2026-10-25 – 2026-10-25" — genau der Zustand, den der Kommentar zwei Zeilen darüber verhindern soll. Für jeden anderen Anreisetag (auch für die Frühjahrs-Umstellung 29.03.) funktioniert es korrekt.

**Ursache:** Die Tages-Arithmetik läuft über eine feste Millisekunden-Konstante auf einem Date, das lokale Mitternacht darstellt. In der Nacht der Rückstellung hat der Kalendertag 25 Stunden, +24 h bleibt deshalb im selben Tag.

**Fix-Vorschlag:** Den Folgetag über die Kalender-Komponenten bilden statt über Millisekunden: `const cand = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);` — `Date` normalisiert den Überlauf und der Konstruktor löst DST korrekt auf. `stayNights`/`nightsBetween` selbst bleiben unverändert richtig, weil sie über `T00:00:00Z` verankert rechnen. Dieselbe Konstante steht auch in HotelPlanningPanel.tsx:86 und HotelSetupWizard.tsx:87 (`addDays`), dort aber auf `T00:00:00Z` verankert und damit unkritisch.

### 114. [niedrig] Unmount-Cleanup in DocumentsViewer haelt `blobUrl` aus dem ERSTEN Render fest ('') und gibt deshalb nie etwas frei

**Wo:** `components/myEvents/DocumentsViewer.tsx:109` · Quelle: Audit B (nach dem Umbau)

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

