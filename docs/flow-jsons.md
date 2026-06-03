# Power Automate Flow JSONs

Dieses Dokument enthält die vollständigen Flow-Definitionen aller **6 DEX-Flows**.
Wird aktualisiert wenn Flows geändert werden.

Übersicht:
1. **DEX_IDReorder_TeilnehmerIDs** — TeilnehmerIDs renummerieren + Warteliste nachrücken
2. **DEX_SEND_MAIL** — Mail-Versand aus DEX_Emails-Queue
3. **DEX_CreateOutlookEvent** — Outlook-Termin initial anlegen
4. **DEX_Outlook_Einladungen** — Teilnehmer zum Outlook-Termin hinzufügen/entfernen, Termin aktualisieren/löschen
5. **DEX_OutlookDeclineHandler** — Decline-Mails abfangen → Reminder-Mail queuen (inkl. weitergeleitete Declines FW:/WG:)
6. **DEX_OutlookForwardHandler** (NEU) — Meeting-Forward-Notifications abfangen → FYI-Mail an Organizer wenn weitergeleiteter Empfänger nicht registriert

---

## 1. DEX_IDReorder_TeilnehmerIDs

**Trigger:** Neuer Eintrag in DEX_IDReorder
**Zweck:** TeilnehmerIDs neu vergeben (Aktive + Warteliste lückenlos sortiert) + Nachrücken von Warteliste (seit v6.7 inkl. typ-bewusster Promotion für B2Run-Split-Wartelisten; seit v10.20 mit optionalem Shared-Waitlist-Modus)
**Letztes Update:** 2026-06-02 (v18.66 — OrgNachruecker-Mail + Paginierung as-implemented; siehe offene Korrektur unten)

### ✅ GELÖST 2026-06-02 (v18.66) — leeres `from` + unquoted `Warteliste` + Self-Reference

> **Status: umgesetzt & verifiziert.** Alle drei unten beschriebenen Fehler
> sind im Tenant behoben (Filter-`from` → `variables('AllParticipants')`,
> `'Warteliste'` gequotet, `Merge_Pages`-Compose gegen die Self-Reference).
> Der finale Flow-JSON steht weiter unten im json-Block. Der folgende Abschnitt
> bleibt als Fehlerbild-Doku stehen.

Beim Review des as-implemented Flow-Exports (nach dem Paginierungs-Umbau)
sind **zwei Fehler** in den **Filter-array-Actions** (Query) aufgefallen.
Beide entstehen dadurch, dass diese Actions vor dem Umbau auf
`@body('Get_Enrolled_Participants')?['value']` zeigten — beim Löschen dieser
Quelle wurde die `from`-Bindung geleert und in der Paginierungs-Anleitung
**nicht** auf `AllParticipants` umgebogen.

**Betroffene Actions:** `Filter_Non_Waitlist`, `Filter_Active`,
`Filter_Waitlist`, `Filter_Active_Durchstarter`, `Filter_Active_Funstarter`.

1. **Leeres `From` (kritisch).** Alle fünf Filter-Actions haben aktuell
   `from = ""` → sie filtern ein **leeres** Array. Folge:
   - `Count_Active` (= `length(Filter_Non_Waitlist)`) ist immer **0**.
   - Damit ist `Check_Nachrücken` (`Count_Active < MaxParticipants`) **immer
     wahr** → der Flow rückt bei **jeder** Abmeldung jemanden nach,
     **auch wenn das Event voll ist** → schleichende Überbuchung.
   - Im Split-Pfad analog: `Count_Active_Durchstarter` / `_Funstarter` = 0 →
     `Check_*_Free` immer wahr → Überbuchung pro Gruppe.
   - Die reine **ID-Neuvergabe** ist NICHT betroffen (sie liest
     `GenerateSPData` direkt aus `variables('AllParticipants')`).
   - **Fix:** In jeder der fünf Actions das **From**-Feld auf den
     fx-Ausdruck `variables('AllParticipants')` setzen.
2. **`Warteliste` ohne Anführungszeichen** (in den drei Nicht-Split-Filtern
   `Filter_Non_Waitlist`, `Filter_Active`, `Filter_Waitlist`). Die Where-
   Bedingung lautet z. B. `@not(equals(item()?['Status'],Warteliste))` —
   der Vergleichswert `Warteliste` ist ein **String** und muss in
   einfache Anführungszeichen: `'Warteliste'`. Am einfachsten über den
   **Basismodus** der Filter-array-Action neu eintippen (Status / *is not
   equal to* / `Warteliste`) — Power Automate quotet den Wert dann selbst.
   Die Split-Filter (`Filter_Active_Durchstarter` / `_Funstarter`) haben die
   Anführungszeichen bereits korrekt.
3. **Self-reference in `Append_Page`** (Paginierungs-Schleife). Die
   `Set variable AllParticipants`-Action darf `variables('AllParticipants')`
   nicht im eigenen Set-Wert referenzieren (Fehler beim Speichern:
   `WorkflowRunActionInputsInvalidProperty` — „Self reference is not
   supported…"). **Fix:** einen **Compose** `Merge_Pages` =
   `union(variables('AllParticipants'), body('Get_Page')?['value'])` zwischen
   `Get_Page` und `Append_Page` einschieben, und `Append_Page` auf
   `outputs('Merge_Pages')` setzen. Siehe Schritt b/c in der
   Paginierungs-Anleitung unten.

**Verifikation nach dem Fix:** Event mit Cap = 2, 2 Aktiven, 1 Wartelistler;
eine **dritte** Person zusätzlich auf die Warteliste setzen; jetzt eine
Abmeldung auslösen → es darf **genau eine** Person nachrücken (auf den frei
gewordenen Platz), nicht zwei. `Count_Active` muss nach dem Lauf der echten
Aktiven-Zahl entsprechen, nicht 0.

### UI-Anleitung 2026-06-03 (v18.69) — Renummerierung: Diff-Verfahren + selbst-prüfende Schleife (garantiert 1..N, keine Duplikate)

**Status:** Struktur im Tenant umgesetzt und Action-für-Action verifiziert
(2026-06-03). Finaler Lauf-Test (sauberes 1..N nach Abmeldung) durch den
Maintainer ausstehend; danach wird der finale JSON-Block hier ersetzt.

**Ziel/Motivation:** Die bisherige Renummerierung schrieb **stur alle 1..N
neu** (bei 576 Teilnehmern 576 Schreibvorgänge — egal wer sich abmeldet) und
prüfte sich über `max == N`, was bei gleichzeitigen Anmeldungen unzuverlässig
ist und Duplikate aus stillen Batch-Teilfehlern nicht erkennt. Das neue
Verfahren:

- **Diff statt „alles neu":** Lade alle Teilnehmer sortiert nach
  `TeilnehmerID asc`, berechne pro Position die Soll-ID (Position i → i+1) und
  **schreibe nur, wo Soll ≠ Ist**. Meldet sich Nr. 123 ab, bleiben 1–122
  unberührt, nur ab 123 wird verschoben (124→123, 125→124, …).
- **Selbst-prüfende Schleife (garantiert lückenlos, ohne Duplikate):** Die
  Renummerierung läuft in einer **Do-until**, die nach jedem Durchlauf **frisch
  lädt** und stoppt, **sobald der Diff leer ist**. Ein Duplikat erzwingt
  zwangsläufig einen Positions-Versatz → Diff nicht leer → die Schleife
  korrigiert weiter (bis max. 5×). Sie **kann nicht aufhören**, solange
  irgendwo eine Lücke/ein Duplikat ist.
- **Manuelle Paginierung entfällt:** Statt der fragilen
  `Load_All_Pages`-Schleife (nextLink/`union`/`$skiptoken`) lädt die Standard-
  Action **„Elemente abrufen" (Get items)** mit eingebauter Paginierung. Die
  ist **keine** Until-Action und darf daher in der Retry-Schleife stehen
  (Do-until lässt sich nicht in Do-until schachteln). **Wichtig:** Get items
  liefert das Item-Id-Feld als **`ID`** (groß), nicht `Id`.

**Vorbereitung:** Falls eine Teilnehmerliste je **> 5.000** Einträge (inkl.
Abgemeldete) hat, in den Listeneinstellungen die Spalte **`TeilnehmerID`**
indizieren (sonst List-View-Threshold). Bei ~1.500 unkritisch.

**Phase 1 — Obsolete Actions löschen:** `Init_NextPageUri`, `Load_All_Pages`,
`Filter_Active`, `Filter_Waitlist`, `Sort_ByStatusPriority`, die **oberste**
`Generate_Indices`, die **oberste** `GenerateSPData`, sowie `Verify_Max` (in
`Batch_Until_Clean`). **Behalten:** `Update_item`, `Settings`,
`Get_ListItemType`, `Init_AllParticipants`, `BatchGuids`, `batchTemplate`,
`Filter_Non_Waitlist`, `Count_Active`, `Batch_Until_Clean` (mit `Loop_Batches`),
`Get_EventDetails`/`Is_B2RunSplit`/`DEX_IDReorder`/`Error_Handler`, Counter-Abgleich.

**Phase 2 — `Batch_Until_Clean` zur Renummerier-Schleife ausbauen** (alle
Actions INNERHALB der Schleife, VOR `Loop_Batches`, in dieser Reihenfolge):

1. **`Load_Participants`** (Elemente abrufen / Get items):
   - Site Address (custom/fx): `outputs('Settings')?['siteAddress']`
   - List Name (custom/fx): `outputs('Settings')?['listName']`
   - Filter Query: `Status ne 'Abgemeldet'`
   - Order By: `TeilnehmerID asc`
   - Top Count: `5000`
   - ⋮ → Settings → **Pagination: On**, Threshold **5000**
2. **`Set_AllParticipants`** (Variable festlegen), runAfter `Load_Participants`:
   - Name `AllParticipants`, Value (fx): `body('Load_Participants')?['value']`
3. **`Generate_Indices`** (Compose), runAfter `Set_AllParticipants`:
   `range(0, length(variables('AllParticipants')))`
4. **`GenerateSPData_Full`** (Auswählen/Select), runAfter `Generate_Indices`:
   - From (fx): `outputs('Generate_Indices')`
   - Map: `ID` → `variables('AllParticipants')[item()]?['ID']` ·
     `TeilnehmerID` → `add(item(), 1)` ·
     `Old` → `variables('AllParticipants')[item()]?['TeilnehmerID']`
5. **`GenerateSPData`** (Array filtern), runAfter `GenerateSPData_Full`:
   - From (fx): `body('GenerateSPData_Full')`
   - Bedingung (advanced): `@not(equals(item()?['TeilnehmerID'], item()?['Old']))`
6. **`Loop_Batches`**: runAfter `GenerateSPData`. Inhalt unverändert
   (`Select_map` → `batchData` → `SendBatch`); `Select_map` nutzt
   `item()?['ID']` + `item()?['TeilnehmerID']` aus dem Diff.
7. **Schleifen-Bedingung** (`Batch_Until_Clean` → Loop until, advanced):
   `@equals(length(body('GenerateSPData')), 0)` · Count **5** · Timeout **PT30M**.

**Phase 3 — `Count_Active` für die Nachrück-Logik:**
- `Filter_Non_Waitlist`: runAfter **`Batch_Until_Clean` Succeeded** (From bleibt
  `@variables('AllParticipants')`, where `@not(equals(item()?['Status'], 'Warteliste'))`).
- `Count_Active`: runAfter `Filter_Non_Waitlist` (`@length(body('Filter_Non_Waitlist'))`).
- `Get_EventDetails` (bzw. `Process_Batch_Scope`): runAfter **`Count_Active` Succeeded**.
- Split-Filter `Filter_Active_Durchstarter`/`_Funstarter` lesen weiter
  `@variables('AllParticipants')` — unverändert.

**Ziel-Reihenfolge:**
```
Update_item → Settings → Get_ListItemType → Init_AllParticipants
→ BatchGuids → batchTemplate
→ Batch_Until_Clean (Load_Participants→Set_AllParticipants→Generate_Indices
   →GenerateSPData_Full→GenerateSPData→Loop_Batches ; until Diff leer, 5×)
→ Filter_Non_Waitlist → Count_Active
→ Get_EventDetails → Is_B2RunSplit → DEX_IDReorder
→ Counter-Abgleich
```

**Test:** Person abmelden → 1. Durchlauf: `GenerateSPData` = nur Schwanz ab der
Lücke; 2. Durchlauf: `GenerateSPData` leer → Schleife endet. Liste muss
lückenlos 1..N ohne Duplikate sein.

### UI-Anleitung 2026-06-02 (v18.63) — Organizer-Mail bei Abmeldung mit Nachrücker

**Ziel:** Sobald der Flow nach einer Abmeldung jemanden von der Warteliste
nachrückt, soll **zusätzlich** eine Standard-Mail an die **Organizer** des
Events gehen („Es gab eine Abmeldung, X ist nachgerückt"). Die App legt dafür
das Template **`OrgNachruecker`** (DE+EN, pre-wrapped) in `DEX_EmailTemplates`
an — bitte einmalig **Settings → „Default-Mail-Templates re-seed"** klicken,
damit es vorhanden ist. Platzhalter: `{{EventTitle}}`, `{{PromotedName}}`.

**Überblick — wo überall:** Der Flow hat bis zu **drei** Nachrück-Zweige. In
**jeden** kommen drei neue Actions (Template laden, Abgemeldeten laden, Org-Mail
in Queue). Die Action-Namen MÜSSEN flow-weit eindeutig sein, daher pro Zweig
eigene Namen (Suffix `_N` / `_D` / `_F`). Wenn dein Event **kein** B2Run-Split
ist (keine Durchstarter/Funstarter-Kapazität), existieren die Zweige 2 + 3 gar
nicht — dann reicht **Zweig 1**.

Gemeinsame Bausteine (in allen Zweigen identisch, nur der Action-Name ändert sich):
- **Org-Template-URI** = `@concat('_api/web/lists/getbytitle(''DEX_EmailTemplates'')/items?$filter=TemplateType eq ''OrgNachruecker'' and Language eq ''', coalesce(first(outputs('Get_EventDetails')?['body/value'])?['EmailLanguage'], 'EN'), '''&$select=Subject,BodyHtml&$top=1')`
- **CancelledName** = `@coalesce(triggerOutputs()?['body/CancelledName'], 'eine Person')` — kommt **direkt aus dem Trigger**: die App schreibt seit v18.65 den Namen der abgemeldeten Person beim Anlegen des DEX_IDReorder-Items mit (kein separater Abfrage-Schritt nötig, race-sicher).
- **Empfänger** (item/Recipient) = `@first(outputs('Get_EventDetails')?['body/value'])?['OrganizerEmail']`
- **EventTitle** = `@first(outputs('Get_EventDetails')?['body/value'])?['Title']`
- **EventId** = `@triggerOutputs()?['body/EventId']`

> Hinweis: dadurch entfällt der frühere `Get_Cancelled_Person`-Schritt komplett —
> pro Zweig bleiben nur **zwei** neue Actions (Template laden + Mail queuen).

---

#### ZWEIG 1 — Normaler Nachrücker (NEIN-Zweig: `Check_Nachrücken` → `Promote_Waitlist`)

Diese drei Actions **innerhalb** des `Condition_1`-TRUE-Zweigs, NACH dem
bestehenden `Queue_Email`:

1. **`Get_Org_Template_N`** — Send an HTTP request to SharePoint, GET, URI = *Org-Template-URI* (oben). Header `Accept: application/json;odata=verbose`. runAfter: `Queue_Email` → Succeeded.
2. **`Queue_Org_Email_N`** — Create item in **DEX_Emails**. runAfter: `Get_Org_Template_N` → Succeeded. Felder:
   - **item/Title:** `@replace(replace(coalesce(first(body('Get_Org_Template_N')?['d']?['results'])?['Subject'], concat('Abmeldung mit Nachrücker: ', first(outputs('Get_EventDetails')?['body/value'])?['Title'])), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title']), '{{PromotedName}}', concat(first(body('Get_Waitlist_First')?['d']?['results'])?['Vorname'], ' ', first(body('Get_Waitlist_First')?['d']?['results'])?['Nachname']))`
   - **item/Recipient:** *Empfänger* (oben)
   - **item/RecipientName:** `Organizer`
   - **item/Body:** `@replace(replace(replace(coalesce(first(body('Get_Org_Template_N')?['d']?['results'])?['BodyHtml'], ''), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title']), '{{PromotedName}}', concat(first(body('Get_Waitlist_First')?['d']?['results'])?['Vorname'], ' ', first(body('Get_Waitlist_First')?['d']?['results'])?['Nachname'])), '{{CancelledName}}', coalesce(triggerOutputs()?['body/CancelledName'], 'eine Person'))`
   - **item/EmailType → Value:** `OrgNachruecker`
   - **item/EventTitle:** *EventTitle* (oben)
   - **item/EventId:** *EventId* (oben)
   - **item/Status → Value:** `Pending`

---

#### ZWEIG 2 — B2Run Durchstarter (`Promote_Durchstarter`)

NUR falls B2Run-Split. Drei Actions im Durchstarter-Promote-Zweig, NACH
`Queue_Email_Durchstarter`. **Waitlist-Quelle:** `Get_Waitlist_First_Durchstarter`.

1. **`Get_Org_Template_D`** — GET, URI = *Org-Template-URI*. Header `Accept: application/json;odata=verbose`. runAfter: `Queue_Email_Durchstarter` → Succeeded.
2. **`Queue_Org_Email_D`** — Create item in **DEX_Emails**. runAfter: `Get_Org_Template_D` → Succeeded. Felder:
   - **item/Title:** `@replace(replace(coalesce(first(body('Get_Org_Template_D')?['d']?['results'])?['Subject'], concat('Abmeldung mit Nachrücker: ', first(outputs('Get_EventDetails')?['body/value'])?['Title'])), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title']), '{{PromotedName}}', concat(first(body('Get_Waitlist_First_Durchstarter')?['d']?['results'])?['Vorname'], ' ', first(body('Get_Waitlist_First_Durchstarter')?['d']?['results'])?['Nachname']))`
   - **item/Recipient:** *Empfänger*
   - **item/RecipientName:** `Organizer`
   - **item/Body:** `@replace(replace(replace(coalesce(first(body('Get_Org_Template_D')?['d']?['results'])?['BodyHtml'], ''), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title']), '{{PromotedName}}', concat(first(body('Get_Waitlist_First_Durchstarter')?['d']?['results'])?['Vorname'], ' ', first(body('Get_Waitlist_First_Durchstarter')?['d']?['results'])?['Nachname'])), '{{CancelledName}}', coalesce(triggerOutputs()?['body/CancelledName'], 'eine Person'))`
   - **item/EmailType → Value:** `OrgNachruecker`
   - **item/EventTitle:** *EventTitle* · **item/EventId:** *EventId* · **item/Status → Value:** `Pending`

---

#### ZWEIG 3 — B2Run Funstarter (`Promote_Funstarter`)

NUR falls B2Run-Split. Drei Actions im Funstarter-Promote-Zweig, NACH
`Queue_Email_Funstarter`. **Waitlist-Quelle:** `Get_Waitlist_First_Funstarter`.

1. **`Get_Org_Template_F`** — GET, URI = *Org-Template-URI*. Header `Accept: application/json;odata=verbose`. runAfter: `Queue_Email_Funstarter` → Succeeded.
2. **`Queue_Org_Email_F`** — Create item in **DEX_Emails**. runAfter: `Get_Org_Template_F` → Succeeded. Felder:
   - **item/Title:** `@replace(replace(coalesce(first(body('Get_Org_Template_F')?['d']?['results'])?['Subject'], concat('Abmeldung mit Nachrücker: ', first(outputs('Get_EventDetails')?['body/value'])?['Title'])), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title']), '{{PromotedName}}', concat(first(body('Get_Waitlist_First_Funstarter')?['d']?['results'])?['Vorname'], ' ', first(body('Get_Waitlist_First_Funstarter')?['d']?['results'])?['Nachname']))`
   - **item/Recipient:** *Empfänger*
   - **item/RecipientName:** `Organizer`
   - **item/Body:** `@replace(replace(replace(coalesce(first(body('Get_Org_Template_F')?['d']?['results'])?['BodyHtml'], ''), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title']), '{{PromotedName}}', concat(first(body('Get_Waitlist_First_Funstarter')?['d']?['results'])?['Vorname'], ' ', first(body('Get_Waitlist_First_Funstarter')?['d']?['results'])?['Nachname'])), '{{CancelledName}}', coalesce(triggerOutputs()?['body/CancelledName'], 'eine Person'))`
   - **item/EmailType → Value:** `OrgNachruecker`
   - **item/EventTitle:** *EventTitle* · **item/EventId:** *EventId* · **item/Status → Value:** `Pending`

> **Betreff:** `{{CancelledName}}` ist im Title bewusst nicht ersetzt (Betreff
> bleibt kurz, zeigt nur den Nachrücker). Wer es im Betreff will, ergänzt im
> jeweiligen `item/Title` ein weiteres `replace(…, '{{CancelledName}}', …)`.
> **Headers** für alle Template-GET-Actions: `Accept: application/json;odata=verbose`
> (damit `?['d']?['results']` greift — wie bei den bestehenden Waitlist-GETs).

**Hinweis:** Der `OrganizerEmail`-Wert in `DEX_Events` ist `;`-getrennt — der
`DEX_SEND_MAIL`-Flow verarbeitet mehrere Empfänger genauso wie bei den
Organizer-BCCs. Die fertige Mail wird (wie alle Queue-Mails) von `DEX_SEND_MAIL`
versendet inkl. Logo/ORB-Ersetzung. In den B2Run-Split-Zweigen die
`Get_Waitlist_First`-Referenzen oben durch `Get_Waitlist_First_Durchstarter`
bzw. `Get_Waitlist_First_Funstarter` ersetzen.

Sobald im Tenant umgesetzt: bitte den Flow-JSON zurückschicken, dann pflege ich
den finalen Stand hier ein.

### UI-Anleitung 2026-06-02 (v18.55) — Paginierung: alle Teilnehmer laden (>1.000 / ILS-Listen)

**Problem:** `Get_Enrolled_Participants` ist ein **einzelner** HTTP-GET
(`…/items?$select=Id,TeilnehmerID,Status,RegistrationDate&$filter=Status ne 'Abgemeldet'&$orderby=RegistrationDate asc&$top=5000`).
SharePoint liefert auf **großen** bzw. **Item-Level-Security**-Listen oft nur
eine **Teil-Seite** (~1.000 Zeilen) zurück plus einen `@odata.nextLink` — eine
einzelne „Send an HTTP request"-Action **folgt diesem nextLink NICHT**. Folge:
Ab ~1.000 Teilnehmern werden nur die ersten ~1.000 renummeriert/gezählt, der
Rest fällt still weg → ID-Vergabe und Nachrück-Logik „nicht mehr sauber".

**Lösung:** den einzelnen GET durch eine **Paginierungs-Schleife** ersetzen, die
solange weiter-GETtet, wie ein `@odata.nextLink` zurückkommt, und alle Zeilen in
einer Array-Variable sammelt. Feld-Casing (`Id`, `Status`, …) bleibt identisch,
deshalb bewusst **NICHT** auf den „Get items"-Connector wechseln (der liefert
`ID` statt `Id` und würde `GenerateSPData` brechen).

**UI-Schritte:**

1. **Zwei Variablen anlegen** (direkt nach `Get_ListItemType`, VOR der bisherigen
   `Get_Enrolled_Participants`-Action):
   - **Initialize variable** `AllParticipants`, Typ **Array**, Wert `[]`.
   - **Initialize variable** `NextPageUri`, Typ **String**, Wert (fx-Tab):
     `concat('_api/web/lists/getbytitle(''', outputs('Settings')?['listName'], ''')/items?$select=Id,TeilnehmerID,Status,RegistrationDate&$filter=Status ne ''Abgemeldet''&$orderby=RegistrationDate asc&$top=5000')`
2. **Do until** „Load_All_Pages" einfügen (nach den beiden Init-Variablen):
   - Bedingung: `NextPageUri` **is equal to** `` (leerer String).
   - Change limits: Count `5000`, Timeout `PT1H`.
   - **Inhalt der Schleife (in dieser Reihenfolge):**
     a. **Send an HTTP request to SharePoint** `Get_Page`:
        - Site Address (fx): `outputs('Settings')?['siteAddress']`
        - Method: **GET**
        - Uri (fx): `variables('NextPageUri')`
        - Headers: `Accept` = `application/json;odata=nometadata`
     b. **Compose** `Merge_Pages` (fx):
        `union(variables('AllParticipants'), body('Get_Page')?['value'])`
        — runAfter `Get_Page` Succeeded.
        **WICHTIG (v18.66):** dieser Compose-Zwischenschritt ist Pflicht. Eine
        **Set variable**-Action darf die eigene Variable **nicht** im Set-Wert
        referenzieren (Power Automate: „Self reference is not supported when
        updating the value of variable 'AllParticipants'" →
        `WorkflowRunActionInputsInvalidProperty`). Compose darf die Variable
        lesen, Set variable liest danach nur den Compose-Output.
        **WICHTIG (v18.67):** zum Zusammenführen `union(...)` benutzen, **nicht**
        `concat(...)`. `concat` ist eine **String**-Funktion und klebt zwei
        Arrays als Text zusammen → `Append_Page` schlägt mit
        „The variable 'AllParticipants' of type 'Array' cannot be … updated
        with value of type 'String'" fehl. `union` führt zwei Arrays zu einem
        Array zusammen (Duplikate werden entfernt — bei Paginierung unkritisch,
        da Seiten-Items unterschiedliche `Id` haben).
     c. **Set variable** `AllParticipants` (fx): `outputs('Merge_Pages')`
        — runAfter `Merge_Pages` Succeeded.
     d. **Set variable** `NextPageUri` (fx):
        `if(empty(coalesce(body('Get_Page')?['@odata.nextLink'], body('Get_Page')?['odata.nextLink'])), '', concat('_api', last(split(coalesce(body('Get_Page')?['@odata.nextLink'], body('Get_Page')?['odata.nextLink']), '_api'))))`
        — runAfter Schritt c Succeeded. (Schneidet den relativen Pfad **ab `_api`**
        aus dem nextLink heraus; bei letzter Seite → leerer String → Schleife endet.)
        **WICHTIG (v18.67) — Endlosschleifen-Fix:** die frühere Variante baute
        den relativen Pfad per `replace(nextLink, siteAddress + '/', '')`. Bei
        `$filter`/`$orderby` auf **nicht-indizierten** Spalten (hier `Status` /
        `RegistrationDate`) macht SharePoint ein **Listen-Scan-Paging** und
        liefert einen `nextLink` mit `$skiptoken` — **auch bei winzigen Listen**.
        Passte `SubsiteUrl` nicht **zeichengenau** auf das nextLink-Präfix
        (Groß-/Kleinschreibung, Slash), blieb die URI **absolut** → `Get_Page`
        holte dieselbe Seite erneut → der `$skiptoken` advancte nie →
        **Endlosschleife** bis count 400 / Timeout 3 h. Der `split('_api')`-Ansatz
        ist robust gegen Host/Casing/Slash und deckt zusätzlich beide
        nextLink-Schreibweisen ab (`@odata.nextLink` bzw. `odata.nextLink` unter
        `odata=nometadata`).
3. **Alte Action entfernen:** `Get_Enrolled_Participants` **löschen**. Die Action,
   die bisher danach lief (`Count_Active`), per **Configure run after** auf
   **`Load_All_Pages` Succeeded** setzen.
4. **Drei Referenzen umbiegen** von `body('Get_Enrolled_Participants')?['value']`
   auf `variables('AllParticipants')`:
   - `Count_Active` (Compose): `@length(filter(variables('AllParticipants'), not(equals(item()?['Status'], 'Warteliste'))))`
   - `Generate_Indices` (Compose): `@range(0, length(variables('AllParticipants')))`
   - `GenerateSPData` (Select) → „From" bleibt `@outputs('Generate_Indices')`, im
     Map-Feld **ID**: `@variables('AllParticipants')[item()]?['Id']` (TeilnehmerID
     bleibt `@add(item(), 1)`).
5. **Speichern.** Test mit einem Event > 1.000 aktiven Teilnehmern: nach einer
   Abmeldung müssen ALLE TeilnehmerIDs lückenlos 1..N neu vergeben sein (nicht nur
   die ersten ~1.000).

**Hinweis Listen-View-Threshold:** Bei **> 5.000** Items pro Event greift zusätzlich
die SharePoint-5000er-Grenze — `Id` ist indiziert, `$orderby=RegistrationDate`
und `$filter=Status` sind es i. d. R. nicht. Eure Events liegen bei ~1.500, daher
unkritisch; falls je > 5.000, müssten `RegistrationDate`/`Status` als
Listen-Indizes angelegt werden.

Sobald im Tenant umgesetzt: bitte den exportierten Flow-JSON zurückspielen, dann
aktualisiere ich den Flow-Block unten in dieser Datei.

### UI-Anleitung 2026-09-XX (v17.15) — Nachrueck-Audit (PromotedDate / ReplacedParticipantEmail / ReplacedByParticipantEmail)

**Hintergrund:** seit App-Version v17.15 gibt es drei neue Spalten in jeder Subsite-Teilnehmerliste, mit denen wir nachvollziehen wann ein Wartelistler nachgerueckt ist und wessen Abmeldung den Promote ausgeloest hat. Die App-Logik in `EventService.promoteFirstWaitlistItem` setzt diese Felder bereits — fuer den **User-Self-Cancel-Pfad**, der durch den Power-Automate-Flow promotet, muessen die gleichen Felder hier zusaetzlich gesetzt werden. Sonst sind die App-Spalten nur halb gefuellt (nur Admin-Cancels haben das Audit).

**Was sich aendert (auf Hoch-Level):**

1. Im Flow gibt es bereits eine Variable mit der Item-Id des cancelnden Eintrags (`triggerOutputs()` → DEX_IDReorder enthaelt `ParticipantEmail`, oder via `Get_Item` der Teilnehmerliste). Falls nicht: einen `Get_item`-Step ergaenzen, der die cancelnde Person aus der Subsite-Teilnehmerliste laed (Filter `Status eq 'Abgemeldet' and modified desc top 1` — am bequemsten anhand der `Cancel_ItemId` aus der DEX_IDReorder-Trigger-Zeile).
2. In den drei `Promote_*`-Update-Actions (Promote_Waitlist, Promote_Durchstarter, Promote_Funstarter) zwei neue Felder mitschreiben:
   - `PromotedDate = utcNow()` (Expression: `utcNow()`)
   - `ReplacedParticipantEmail = <ParticipantEmail der cancelnden Person>`
3. Eine neue Update-Action **nach** dem Promote: PATCH auf das cancelnde Item mit `ReplacedByParticipantEmail = <ParticipantEmail der promoteten Person>`. Die promotete Person ist die, die in Schritt 2 ge-MERGEt wurde — ihre E-Mail kommt aus `body('Get_Waitlist_First')?['ParticipantEmail']` (bzw. `Get_Waitlist_First_Durchstarter` / `_Funstarter`).

**UI-Schritte (Power Automate Maker, Cloud-Flow-Editor):**

1. Oeffne den Flow `DEX_IDReorder_TeilnehmerIDs`. Scrolle zur Scope-Action `Process_Batch_Scope`.
2. **Vorbereitung — Cancelling-Item E-Mail abrufen.** Direkt nach `Get_EventDetails` einen neuen `Get items`-Step einfuegen mit Settings:
   - Site Address: dynamisch aus `outputs('Settings')?['siteUrl']`
   - List Name: dynamisch aus `outputs('Settings')?['listName']`
   - Filter Query: `Id eq @{triggerOutputs()?['body/Cancel_ItemId']}` (oder analog, je nach Trigger-Spaltennamen)
   - Top Count: 1
   Diesen Step **`Get_Cancelled_Item`** nennen.
3. **Promote-Update erweitern.** In den drei Update-Actions
   - `Promote_Waitlist` (NEIN-Zweig — normales Event)
   - `Promote_Durchstarter` (B2Run-Split-Zweig)
   - `Promote_Funstarter` (B2Run-Split-Zweig)
   im Action-Body unter „Advanced parameters" zwei neue Felder hinzufuegen:
   - `PromotedDate` → Expression `utcNow()`
   - `ReplacedParticipantEmail` → Expression `first(outputs('Get_Cancelled_Item')?['body/value'])?['ParticipantEmail']`
4. **Zweiter PATCH auf die cancelnde Person.** Direkt **nach** dem Promote-Step (im jeweiligen Zweig) eine neue `Update item`-Action einfuegen:
   - Item Id: `first(outputs('Get_Cancelled_Item')?['body/value'])?['Id']`
   - Felder:
     - `ReplacedByParticipantEmail` → Expression `body('Get_Waitlist_First')?['ParticipantEmail']` (bzw. `Get_Waitlist_First_Durchstarter` im B2Run-Split-Zweig)
   - Die Action **`Mark_Cancelled_Replaced`** nennen.
5. Run-after von `Mark_Cancelled_Replaced` auf `Succeeded` des jeweiligen `Promote_*`-Steps setzen (nicht parallel — sonst Race-Condition wenn der Promote fehlschlaegt).
6. Speichern. Teste mit einem Event das eine Warteliste hat: melde dich als Wartelistler an, lasse einen Aktiven sich abmelden. In der App-Teilnehmer-Tabelle muessen jetzt die Spalten `Nachgerueckt am` und `Ersetzt` gesetzt sein, in der Abgemeldet-Liste die Spalte `Ersetzt durch`.

**Verifikation:**
- Event mit 2 Plaetzen Cap, 2 Aktiven, 1 Wartelistler.
- Wartelistler ist Person A. Beide Aktiven sind Person B (Id=2) und C (Id=3).
- B meldet sich SELBST in der App ab. Der Flow promotet A.
- In der App-Teilnehmer-Tabelle: A hat `PromotedDate` = jetzt, `Ersetzt` = B.
- In der Abgemeldet-Section: B hat `Ersetzt durch` = A.

**Caveat:** wenn der Flow nicht angepasst wird, funktioniert das Audit nur beim Admin-Cancel-Pfad (Organizer/Admin meldet jemanden im Admin-Center ab — diese App-Code-Logik ist seit v17.15 fertig). Beim User-Self-Cancel laeuft der PA-Flow, der ohne diese Anpassung die drei Audit-Felder leer laesst.

Sobald die Aenderung im Tenant gespeichert ist: bitte den exportierten Flow-JSON (Maker → drei Punkte rechts oben am Flow → **Export → Code view** oder per `Get-AdminFlow` PowerShell) zurueckspielen, dann aktualisieren wir den Promote-Block weiter unten in dieser Datei.

### UI-Anleitung 2026-05-06 (v10.20) — Shared-Waitlist-Modus respektieren

**Hintergrund:** seit App-Version v10.20 kann der Organizer pro Event entscheiden, ob bei aktiver Split-Capacity (DurchstarterCapacity > 0 AND FunstarterCapacity > 0) **eine gemeinsame Warteliste** (FIFO über beide Gruppen) oder **zwei getrennte Wartelisten** (typ-bewusst, alter B2Run-Stil) gelten sollen. Die Information steckt im neuen Boolean-Feld `SplitSharedWaitlist` in `DEX_Events`. Aktuell ignoriert der Flow das Feld — er promotet immer typ-bewusst, sobald beide Kapazitäten > 0 sind. Damit der Shared-Modus auch beim User-Self-Cancel greift, muss der Flow um eine zusätzliche Bedingungs-Zeile erweitert werden.

**Was sich ändert:** in der Bedingungs-Action `Is_B2RunSplit` kommt eine dritte Bedingungs-Zeile hinzu, die prüft ob `SplitSharedWaitlist` **nicht** auf `true` steht. Nur wenn alle drei Bedingungen JA sind (DurchCap > 0 UND FunCap > 0 UND SharedWaitlist = false), läuft der typ-bewusste Zweig. Sobald der Organizer für ein Event die gemeinsame Warteliste aktiviert, fällt der Flow in den NEIN-Zweig (`Check_Nachrücken`) und nimmt den ältesten Wartelistler ohne Typ-Filter — exakt der gewünschte FIFO-Modus.

**UI-Schritte (Power Automate Maker, Cloud-Flow-Editor):**

1. Öffne <https://make.powerautomate.com> → links **Meine Flows** → Flow **DEX_IDReorder_TeilnehmerIDs** anklicken → oben rechts **Bearbeiten** klicken.
2. Im Editor scrollst du zur Scope-Action **Process_Batch_Scope**. Innerhalb des Scope findest du die Bedingungs-Action **Is_B2RunSplit** (sie steht direkt nach **Get_EventDetails** und vor den Promote-Zweigen). Klick einmal auf die Header-Leiste der Bedingung, damit sie aufgeklappt ist.
3. In der Bedingung siehst du aktuell zwei Zeilen, mit `And` verknüpft:
   - Zeile 1: `DurchstarterCapacity` `is greater than` `0`
   - Zeile 2: `FunstarterCapacity` `is greater than` `0`

   Klick rechts neben Zeile 2 auf die drei Punkte (`⋮`) → **Add → Add row** (deutsch: **Hinzufügen → Zeile hinzufügen**). Eine dritte leere Zeile erscheint. Stelle sicher, dass die Verknüpfung oben weiterhin `And` ist (nicht `Or`).
4. **Linkes Feld** der neuen Zeile: klick rein → es öffnet sich der Picker mit zwei Tabs **Dynamic content** und **Expression**. Wechsle auf den **Expression**-Tab (oder das `fx`-Symbol). In das Eingabefeld tippst du folgende Expression — zeichengenau, eine einzige Zeile:

   ```
   coalesce(first(outputs('Get_EventDetails')?['body/value'])?['SplitSharedWaitlist'], false)
   ```

   Klicke **OK** / **Add**. Das `coalesce` sorgt dafür, dass alte Events ohne die `SplitSharedWaitlist`-Spalte (Wert ist `null`) wie `false` behandelt werden — sonst würde die Bedingung bei Legacy-Events brechen.
5. **Operator** (mittleres Dropdown der Zeile): wähle **is equal to** (deutsch: **ist gleich**).
6. **Rechtes Feld** der neuen Zeile: klick rein → Tab **Expression** (`fx`) → tippe genau:

   ```
   false
   ```

   Klicke **OK** / **Add**. Wichtig: das muss die Expression `false` sein (Boolean-Literal), nicht der String `"false"`. Wenn du nur `false` ohne Anführungszeichen einträgst und auf OK klickst, ist es korrekt.
7. Oben rechts klick auf **Save** (deutsch: **Speichern**). Der Save-Dialog kann ein paar Sekunden brauchen, weil PA den Flow gegen den SP-Connector validiert.

**Verifikation:**

- Lege im Tenant ein Test-Event mit Split-Capacity an, setze in Schritt 3 (Kapazität & Sichtbarkeit) den Toggle **Eine gemeinsame Warteliste**, gib z.B. 1 Platz Gruppe A + 1 Platz Gruppe B + Warteliste aktiv. Melde dich + zwei Test-Personen an, sodass jede Gruppe voll ist und mindestens einer auf der Warteliste steht. Melde dann den Angemeldeten der "falschen" Gruppe ab. Erwartet: **die Person auf der Warteliste rückt nach, egal welche Gruppe sie ursprünglich gewählt hatte** (also auch wenn der freie Platz in einer anderen Gruppe entsteht).
- Bei einem Event ohne den Toggle (oder bei einem Legacy-B2Run-Event ohne `SplitSharedWaitlist`-Spalte) bleibt das Verhalten unverändert: typ-bewusste Promotion, Funstarter rücken nicht in Durchstarter-Plätze und umgekehrt.

**Was passiert mit dem `StarterType` des Nachrückers?** Im NEIN-Zweig (Check_Nachrücken → Promote_Waitlist) setzt der Flow nur `Status = Angemeldet`, der `StarterType` bleibt unverändert. Das heißt: die nachrückende Person behält ihren `PreferredStarterType` als `StarterType`. Praktisch heißt das für den Organizer: bei aktiver gemeinsamer Warteliste kann es passieren, dass ein Funstarter-Slot leer bleibt, während ein Durchstarter aus der Warteliste nachrückt — die zwei Capacity-Werte sind dann nur noch Sum-Indikator. Genau das ist Sinn der Option.

**Caveat:** wenn du den Flow nicht anpasst, funktioniert der Shared-Waitlist-Modus nur beim **Admin-Cancel** (Organizer/Admin meldet jemanden im Admin Center ab) — die App-Logik in `AdminPage.tsx` respektiert `splitSharedWaitlist` clientseitig. Beim **User-Self-Cancel** läuft der PA-Flow, der ohne diese Anpassung weiter typ-bewusst promotet.

Sobald die Änderung im Tenant gespeichert ist: bitte den exportierten Flow-JSON (Maker → drei Punkte rechts oben am Flow → **Export → Code view** oder per `Get-AdminFlow` PowerShell) zurückspielen, dann aktualisieren wir den `Is_B2RunSplit`-Block weiter unten in dieser Datei.

### Änderungen 2026-04-22 (v6.7) — Typ-bewusster Promote für B2Run-Split

Die App (EventContext.cancelRegistration) macht seit v6.7 kein Client-seitiges Nachrücken mehr — der komplette Promote-Prozess läuft hier im Flow. Gleichzeitig wurde der Promote-Zweig typ-bewusst ausgebaut: bei B2Run-Split-Events (DurchstarterCapacity > 0 AND FunstarterCapacity > 0) gibt es zwei getrennte Promote-Pässe, einer pro Typ mit eigenem Warteliste-Filter.

**Neue Flow-Struktur (nach Loop_Batches):**

```
Get_EventDetails
 └── Is_B2RunSplit (DurchstarterCapacity > 0 AND FunstarterCapacity > 0)
     ├── WENN JA (B2Run-Split-Event):
     │   ├── Filter_Active_Durchstarter (Query / Filter array)
     │   ├── Count_Active_Durchstarter (Compose: length(body('Filter_Active_Durchstarter')))
     │   ├── Check_Durchstarter_Free (Condition: aktive Durchstarter < DurchstarterCapacity)
     │   │   └── WENN JA:
     │   │       ├── Get_Waitlist_First_Durchstarter (SharePoint GET: $filter=Status eq 'Warteliste' and PreferredStarterType eq 'Durchstarter')
     │   │       └── Has_Durchstarter_Waitlist (Condition: length > 0)
     │   │           └── WENN JA:
     │   │               ├── Promote_Durchstarter (MERGE: Status=Angemeldet + StarterType=Durchstarter)
     │   │               ├── Get_Email_Template_Durchstarter
     │   │               ├── Queue_Email_Durchstarter (Nachrücken-Mail)
     │   │               └── Queue_Outlook_Durchstarter (Einladen)
     │   ├── Filter_Active_Funstarter (Query / Filter array, runAfter Check_Durchstarter_Free)
     │   ├── Count_Active_Funstarter (Compose: length(body('Filter_Active_Funstarter')))
     │   └── Check_Funstarter_Free (Condition: aktive Funstarter < FunstarterCapacity)
     │       └── WENN JA:
     │           ├── Get_Waitlist_First_Funstarter (SharePoint GET: $filter=PreferredStarterType eq 'Funstarter')
     │           └── Has_Funstarter_Waitlist (Condition: length > 0)
     │               └── WENN JA:
     │                   ├── Promote_Funstarter (MERGE: Status=Angemeldet + StarterType=Funstarter)
     │                   ├── Get_Email_Template_Funstarter
     │                   ├── Queue_Email_Funstarter
     │                   └── Queue_Outlook_Funstarter
     └── WENN NEIN (normales Event):
         └── Check_Nachrücken (unverändert alte Logik: Count_Active < MaxParticipants)
             └── Get_Waitlist_First → Condition_1 → Promote_Waitlist + Queue_Email + Queue_Outlook + Get_Email_Template
```

`DEX_IDReorder.runAfter = Is_B2RunSplit: Succeeded` — läuft nach beiden Zweigen.

**Wichtige Details:**
- Im B2Run-Split-Zweig werden **beide Typen nacheinander** geprüft (Durchstarter zuerst, dann Funstarter). Wenn beide freie Plätze + passende Warteliste-Einträge haben, werden in einem einzigen Flow-Run zwei Teilnehmer nachgerückt.
- Filterung pro Starter-Typ läuft seit 2026-05-08 (v11.25-Iteration des Flow-Cleanups) über **zwei `Filter array`-Actions** (`Filter_Active_Durchstarter` / `Filter_Active_Funstarter`) statt über inline-`filter()`-Expressions im Compose. Hintergrund: das Template-Function `filter` ist in vielen Power-Automate-Tenants nicht verfügbar (`The template function 'filter' is not defined or not valid`). Die `Filter array`-Action ist eine echte Action und funktioniert ueberall.
- Die Filter-Bedingung ist:
  - **Durchstarter:** `@and(equals(item()?['StarterType'], 'Durchstarter'), not(equals(item()?['Status'], 'Warteliste')))`
  - **Funstarter:** `@and(equals(item()?['StarterType'], 'Funstarter'), not(equals(item()?['Status'], 'Warteliste')))`
  - **Wichtig:** der Vergleich nutzt `item()?['StarterType']` direkt — NICHT `?['Value']`. Grund: die Items kommen über `Send HTTP request to SharePoint` mit `odata=nometadata` zurück, dabei werden Choice-Felder als plain Strings serialisiert (nicht als `{ Value: '...' }`-Objekt wie beim nativen SharePoint-Connector).
- `Count_Active_Durchstarter` / `Count_Active_Funstarter` (Compose) zählen einfach `length(body('Filter_Active_<typ>'))`.
- `Promote_Durchstarter` / `Promote_Funstarter` setzen **Status=Angemeldet UND StarterType=<typ>** — der nachgerückte Warteliste-Teilnehmer bekommt automatisch den freigewordenen Typ zugewiesen.
- `Queue_Email_*` füllen `{{Name}}` und `{{EventTitle}}` aus dem SharePoint-Template (DEX_EmailTemplates, TemplateType=`Nachruecken`, Language=EN/DE).
- Die internen Slot-IDs `Durchstarter` / `Funstarter` in den Filter- und Promote-Actions sind **nicht** die User-Labels (die sind frei via `splitLabelA` / `splitLabelB`), sondern fixe Choice-Werte der `StarterType`-Spalte auf der Teilnehmer-Liste — funktional wie ein Enum mit zwei Werten („Slot A" und „Slot B"). Der Flow funktioniert unverändert für jedes Event, egal wie der Organizer die Gruppen labelt.

### Änderungen 2026-04-22 (v6.6) — Zwei-Pass-Sortierung der Teilnehmer-IDs

Neuer Compose-Step `Sort_ByStatusPriority` zwischen `Count_Active` und `Generate_Indices`.
Er liefert die Enrolled-Items in der gewünschten Reihenfolge: erst alle Angemeldeten
(Status ∈ Angemeldet / QR versendet / Eingecheckt, sortiert nach RegistrationDate), dann
alle Warteliste-Teilnehmer (auch nach RegistrationDate). Die Expression ist:

```
@union(
  filter(body('Get_Enrolled_Participants')?['value'], not(equals(item()?['Status'], 'Warteliste'))),
  filter(body('Get_Enrolled_Participants')?['value'], equals(item()?['Status'], 'Warteliste'))
)
```

`Generate_Indices` zählt die Länge von `Sort_ByStatusPriority` statt von
`Get_Enrolled_Participants`. `GenerateSPData` greift bei `ID` auf
`outputs('Sort_ByStatusPriority')[item()]?['Id']` zu — also in Status-Priority-
Reihenfolge. Die `TeilnehmerID` bleibt `add(item(), 1)`, d.h. Index+1.

Ergebnis: Angemeldete bekommen IDs 1..N, Warteliste bekommt IDs N+1..N+M — saubere,
lückenlose Sortierung. Beispiel mit 100 Plätzen: Wenn #98 (Angemeldet) abmeldet,
werden #99 → #98, #100 → #99, #101 (alter Warteliste-Erster, wird durch Nachrücken
auch Angemeldet) → #100, #102 (bleibt Warteliste) → #101 usw.

**Zusätzlich `Filter_Non_Waitlist`** (Query-Action) — zählt Nicht-Warteliste-Items als
Vorstufe zu `Count_Active`. Funktional identisch zum alten Inline-`length(filter(...))`,
aber lesbarer.

**Vorherige Änderungen 2026-04-20:**

1. `Get_Active_Participants` umbenannt zu `Get_Enrolled_Participants`. Filter geändert von
   `(Status eq 'Angemeldet') or (Status eq 'QR versendet') or (Status eq 'Eingecheckt')`
   auf `Status ne 'Abgemeldet'`. Vorher wurden Warteliste-Einträge beim Renummerieren
   übersprungen — führte zu Lücken in der TeilnehmerID-Sequenz. Jetzt bekommen Aktive
   + Warteliste gemeinsam fortlaufende IDs nach RegistrationDate.
2. Neuer Compose-Step `Count_Active` zählt Enrolled-Items mit `Status ≠ 'Warteliste'`.
   `Check_Nachrücken`-Bedingung vergleicht jetzt `Count_Active < MaxParticipants` statt
   `length(GenerateSPData) < MaxParticipants`. Nötig weil `GenerateSPData` nach Fix #1
   auch Warteliste-Einträge enthält und die alte Bedingung sonst nie mehr triggern würde.
3. `Promote_Waitlist.body` enthält nur noch `Status: Angemeldet` — KEIN `TeilnehmerID`
   mehr. Die korrekte TID wurde bereits im Batch-Update gesetzt (erste Warteliste in
   RegistrationDate-Reihenfolge = Count_Active + 1). Alte Logik `TID = length(GenerateSPData) + 1`
   hätte nach Fix #1 eine zu hohe TID erzeugt (= EnrolledCount + 1 statt Count_Active + 1).
4. `Queue_Email` (Nachrücker-Mail): Subject/Body/EventTitle-Spalte nutzen jetzt
   `first(outputs('Get_EventDetails')?['body/value'])?['Title']` statt
   `triggerOutputs()?['body/Title']` — letzteres war der DEX_IDReorder-Queue-Title
   ("Reorder: <EventName>"), nicht der echte Event-Name. Anrede + RecipientName nutzen
   jetzt `Vorname` statt `ParticipantName` (nur Vorname in der Begrüßung).
5. `Queue_Outlook.item/Title`: ebenfalls auf EventDetails.Title umgestellt — vorher
   "Einladen: Reorder: <Name>", jetzt "Einladen: <Name>".
6. Nachrücken-Template in `DEX_EmailTemplates` wird jetzt pre-wrapped gespeichert
   (komplettes Deloitte-Design inklusive Logo/Header/Footer), weil der PA-Flow den
   BodyHtml raw verwendet. Client-Code erkennt pre-wrapped Templates in
   `buildEmailFromTemplate()` und skippt den zweiten Wrap. App-seitig v5.33.0+ nötig.

> **Stand v18.66 (umgesetzt & verifiziert 2026-06-02):** Der folgende JSON-Block ist
> der **finale** Flow inklusive Paginierung (`Init_AllParticipants` /
> `Init_NextPageUri` / `Load_All_Pages` mit `Merge_Pages`-Compose gegen die
> Self-Reference), gefixten Filter-Actions (`from = variables('AllParticipants')`,
> `'Warteliste'` gequotet) und den `OrgNachruecker`-Mail-Actions je
> Nachrück-Zweig (`Get_Org_Template_N/_D/_F` + `Queue_Org_Email_N/_D/_F`,
> `{{CancelledName}}` direkt aus `triggerOutputs()?['body/CancelledName']`).

```json
TRIGGER:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "9d46ff77-5fe2-4e1d-9b93-14b9dca1a360"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetOnNewItems"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "runtimeConfiguration": { "concurrency": { "runs": 1, "maximumWaitingRuns": 100 } },
  "splitOn": "@triggerOutputs()?['body/value']"
}

ACTIONS (Reihenfolge laut runAfter):

Update_item (PatchItem → Status=Processing)
Settings (Compose: siteAddress/listName/batchSize=250)
Get_ListItemType (HTTP GET ListItemEntityTypeFullName)

Init_AllParticipants (InitializeVariable, Array)  runAfter Get_ListItemType
Init_NextPageUri (InitializeVariable, String):
  @concat('_api/web/lists/getbytitle(''', outputs('Settings')?['listName'], ''')/items?$select=Id,TeilnehmerID,Status,RegistrationDate&$filter=Status ne ''Abgemeldet''&$orderby=RegistrationDate asc&$top=5000')
  runAfter Init_AllParticipants

Load_All_Pages (Until @equals(variables('NextPageUri'),''), count 400, timeout PT3H)  runAfter Init_NextPageUri
  Get_Page (HTTP GET @variables('NextPageUri'), Accept nometadata)
  Merge_Pages (Compose): @union(variables('AllParticipants'), body('Get_Page')?['value'])   runAfter Get_Page
  Append_Page (SetVariable AllParticipants): @outputs('Merge_Pages')                          runAfter Merge_Pages
  Set_NextPageUri (SetVariable NextPageUri):
    @if(empty(coalesce(body('Get_Page')?['@odata.nextLink'], body('Get_Page')?['odata.nextLink'])), '', concat('_api', last(split(coalesce(body('Get_Page')?['@odata.nextLink'], body('Get_Page')?['odata.nextLink']), '_api'))))  runAfter Append_Page

Filter_Non_Waitlist (Query)  runAfter Load_All_Pages
  from:  @variables('AllParticipants')
  where: @not(equals(item()?['Status'], 'Warteliste'))
Count_Active (Compose: @length(body('Filter_Non_Waitlist')))  runAfter Filter_Non_Waitlist
Filter_Active (Query)  runAfter Count_Active
  from:  @variables('AllParticipants')
  where: @not(equals(item()?['Status'], 'Warteliste'))
Filter_Waitlist (Query)  runAfter Filter_Active
  from:  @variables('AllParticipants')
  where: @equals(item()?['Status'], 'Warteliste')
Sort_ByStatusPriority (Compose: @union(body('Filter_Active'), body('Filter_Waitlist')))  runAfter Filter_Waitlist
Generate_Indices (Compose: @range(0, length(variables('AllParticipants'))))  runAfter Sort_ByStatusPriority
GenerateSPData (Select)  runAfter Generate_Indices
  from: @outputs('Generate_Indices')
  select: { "ID": "@variables('AllParticipants')[item()]?['Id']", "TeilnehmerID": "@add(item(), 1)" }
BatchGuids (Compose: batchGUID/changeSetGUID)  runAfter GenerateSPData
batchTemplate (Compose: PATCH-Changeset-Template)  runAfter BatchGuids

Process_Batch_Scope (Scope)  runAfter batchTemplate
  Loop_Batches (Foreach @chunk(body('GenerateSPData'), outputs('Settings')?['batchSize']), repetitions 1)
    Select_map → batchData → SendBatch (_api/$batch)
  Get_EventDetails (GetItems DEX_Events, $filter ID eq EventId)  runAfter Loop_Batches
  Is_B2RunSplit (If: DurchstarterCapacity>0 AND FunstarterCapacity>0)  runAfter Get_EventDetails
    [JA — Split]
      Filter_Active_Durchstarter (Query)
        from:  @variables('AllParticipants')
        where: @and(equals(item()?['StarterType'], 'Durchstarter'), not(equals(item()?['Status'], 'Warteliste')))
      Count_Active_Durchstarter (Compose @length(body('Filter_Active_Durchstarter')))
      Check_Durchstarter_Free (If Count<DurchstarterCapacity)
        Get_Waitlist_First_Durchstarter (HTTP GET Status=Warteliste & PreferredStarterType=Durchstarter, $orderby TeilnehmerID asc top 1)
        Has_Durchstarter_Waitlist (If results>0)
          Promote_Durchstarter (MERGE Status=Angemeldet, StarterType=Durchstarter)
          Get_Email_Template_Durchstarter (GET Nachruecken-Template)  runAfter Promote_Durchstarter
          Queue_Email_Durchstarter (DEX_Emails, EmailType Nachruecken)  runAfter Get_Email_Template_Durchstarter
          Get_Org_Template_D (GET OrgNachruecker-Template, Accept verbose)  runAfter Queue_Email_Durchstarter
          Queue_Org_Email_D (DEX_Emails, EmailType OrgNachruecker, Recipient OrganizerEmail,
            Body replace {{EventTitle}}/{{PromotedName}}/{{CancelledName}})  runAfter Get_Org_Template_D
          Queue_Outlook_Durchstarter (DEX_Outlook Einladen)  runAfter Queue_Org_Email_D
      Filter_Active_Funstarter (Query)  runAfter Check_Durchstarter_Free
        from:  @variables('AllParticipants')
        where: @and(equals(item()?['StarterType'], 'Funstarter'), not(equals(item()?['Status'], 'Warteliste')))
      Count_Active_Funstarter (Compose @length(body('Filter_Active_Funstarter')))
      Check_Funstarter_Free (If Count<FunstarterCapacity)
        Get_Waitlist_First_Funstarter (HTTP GET Status=Warteliste & PreferredStarterType=Funstarter, $orderby TeilnehmerID asc top 1)
        Has_Funstarter_Waitlist (If results>0)
          Promote_Funstarter (MERGE Status=Angemeldet, StarterType=Funstarter)
          Get_Email_Template_Funstarter (GET Nachruecken-Template)  runAfter Promote_Funstarter
          Queue_Email_Funstarter (DEX_Emails, EmailType Nachruecken)  runAfter Get_Email_Template_Funstarter
          Get_Org_Template_F (GET OrgNachruecker-Template, Accept verbose)  runAfter Queue_Email_Funstarter
          Queue_Org_Email_F (DEX_Emails, EmailType OrgNachruecker, Recipient OrganizerEmail,
            Body replace {{EventTitle}}/{{PromotedName}}/{{CancelledName}})  runAfter Get_Org_Template_F
          Queue_Outlook_Funstarter (DEX_Outlook Einladen)  runAfter Queue_Org_Email_F
    [NEIN — kein Split]
      Check_Nachrücken (If Count_Active<MaxParticipants AND MaxParticipants>0)
        Get_Waitlist_First (HTTP GET Status=Warteliste, $orderby TeilnehmerID asc top 1)
        Condition_1 (If results>0)
          Promote_Waitlist (MERGE Status=Angemeldet)
          Get_Email_Template (GET Nachruecken-Template)  runAfter Promote_Waitlist
          Queue_Email (DEX_Emails, EmailType Nachruecken)  runAfter Get_Email_Template
          Get_Org_Template_N (GET OrgNachruecker-Template, Accept verbose)  runAfter Queue_Email
          Queue_Org_Email_N (DEX_Emails, EmailType OrgNachruecker, Recipient OrganizerEmail,
            Body replace {{EventTitle}}/{{PromotedName}}/{{CancelledName}})  runAfter Get_Org_Template_N
          Queue_Outlook (DEX_Outlook Einladen)  runAfter Queue_Org_Email_N
  DEX_IDReorder (PatchItem → Status=Done)  runAfter Is_B2RunSplit
  Error_Handler (Scope: Set_Failed → Status=Failed)  runAfter DEX_IDReorder [Failed]

Get_Counter_Item (HTTP GET DEX_TeilnehmerCounter items(1))  runAfter Batch_Update*
Get_Max_TeilnehmerID (HTTP GET Teilnehmer $orderby TeilnehmerID desc top 1)  runAfter Get_Counter_Item
Compute_MaxValue (Compose max TeilnehmerID)  runAfter Get_Max_TeilnehmerID
Compute_CurrentCounter (Compose @coalesce(body('Get_Counter_Item')?['NextValue'],0))  runAfter Compute_MaxValue
If_Counter_Stale (If Current != Max)  runAfter Compute_CurrentCounter
  Patch_Counter (MERGE DEX_TeilnehmerCounter NextValue=Compute_MaxValue, retry exp 3)
```

> *(\*) `Batch_Update` ist der Anzeigename des `Process_Batch_Scope` im Tenant —
> die drei Counter-Reconcile-Actions laufen nach Abschluss des Scopes.*

> **Wichtige Ausdrücke verbatim** (für Copy-Paste in den fx-Tab):
>
> - **Init_NextPageUri:** `concat('_api/web/lists/getbytitle(''', outputs('Settings')?['listName'], ''')/items?$select=Id,TeilnehmerID,Status,RegistrationDate&$filter=Status ne ''Abgemeldet''&$orderby=RegistrationDate asc&$top=5000')`
> - **Merge_Pages (v18.67):** `union(variables('AllParticipants'), body('Get_Page')?['value'])` — **nicht** `concat` (String-Funktion → Typfehler in Append_Page)
> - **Append_Page (Wert):** `outputs('Merge_Pages')`
> - **Set_NextPageUri (v18.67, Endlosschleifen-Fix):** `if(empty(coalesce(body('Get_Page')?['@odata.nextLink'], body('Get_Page')?['odata.nextLink'])), '', concat('_api', last(split(coalesce(body('Get_Page')?['@odata.nextLink'], body('Get_Page')?['odata.nextLink']), '_api'))))`
> - **Org-Template-URI (alle drei Zweige):** `concat('_api/web/lists/getbytitle(''DEX_EmailTemplates'')/items?$filter=TemplateType eq ''OrgNachruecker'' and Language eq ''', coalesce(first(outputs('Get_EventDetails')?['body/value'])?['EmailLanguage'], 'EN'), '''&$select=Subject,BodyHtml&$top=1')` — Header `Accept: application/json;odata=verbose`
> - **Org-Mail Body (Beispiel Zweig N):** `replace(replace(replace(coalesce(first(body('Get_Org_Template_N')?['d']?['results'])?['BodyHtml'], ''), '{{EventTitle}}', first(outputs('Get_EventDetails')?['body/value'])?['Title']), '{{PromotedName}}', concat(first(body('Get_Waitlist_First')?['d']?['results'])?['Vorname'], ' ', first(body('Get_Waitlist_First')?['d']?['results'])?['Nachname'])), '{{CancelledName}}', coalesce(triggerOutputs()?['body/CancelledName'], 'eine Person'))`

---

## 2. DEX_SEND_MAIL

**Trigger:** Neuer Eintrag in DEX_Emails
**Zweck:** E-Mails aus Queue versenden über Shared Mailbox (no_reply.events@deloitte.de)
**Letztes Update:** 2026-04-29 (v8.5/v8.6: Cc + Bcc-Support)

Ablauf: Trigger → Config laden (Logo + Default-Bild aus DEX_EmailTemplates via GetItems) → Event laden → Compose_Logo (aus Config) → Compose_Image (Event-Bild oder Default) → Platzhalter ersetzen → Email senden (mit Cc + Bcc) → Status=Sent

### UI-Anleitung 2026-06-01 (v18.30) — Wichtigkeit aus der Queue lesen (hohe Wichtigkeit / rotes „!")

**Hintergrund:** Die App schreibt jetzt eine neue Spalte `Importance` in
`DEX_Emails` (leer/`Normal` = normal, `High` = rotes Ausrufezeichen). Aktuell
steht im Flow die Wichtigkeit fest auf `Normal`. Damit z.B. die DEX-Anfrage-
Mail von der Landing-Page mit hoher Wichtigkeit ankommt, muss die Versand-
Aktion den Wert aus dem Trigger-Item lesen.

**Schritte (alles über die UI, kein Code-View):**

1. Flow `DEX_SEND_MAIL` öffnen → **Bearbeiten**.
2. Die Aktion **„Send an email from a shared mailbox (V2)"** aufklappen.
3. Falls **Importance** nicht sichtbar ist: unten auf **„Show advanced
   options"** / „Erweiterte Optionen anzeigen" klicken — dort steht das Feld
   **Importance** (aktuell auf `Normal`).
4. In das **Importance**-Feld klicken → den festen Wert `Normal` entfernen →
   rechts auf das **Blitz-/fx-Symbol** für dynamischen Inhalt. Im
   **Expression-Tab (fx)** folgenden Ausdruck eingeben und mit **OK**
   bestätigen:

   ```
   if(equals(triggerBody()?['Importance'], 'High'), 'High', 'Normal')
   ```

   (Liest die `Importance`-Spalte des neuen DEX_Emails-Eintrags; ist sie
   `High`, geht die Mail mit hoher Wichtigkeit raus, sonst normal.)
5. **Speichern**.

Danach den **aktuellen Flow-JSON** (Code View → kopieren) hier in
`docs/flow-jsons.md` einpflegen. Bis dahin gilt: die App setzt das Feld
korrekt, der Flow ignoriert es aber noch (alle Mails normal) — nicht
destruktiv.

### Änderungen 2026-04-29 (v8.5)

`SEND_EMAIL`-Aktion wurde um zwei Header erweitert:

- `emailMessage/Cc` ← `triggerBody()?['Cc']` (war schon länger vorhanden, jetzt offiziell dokumentiert)
- `emailMessage/Bcc` ← `triggerOutputs()?['body/Bcc']` **(NEU)**

Hintergrund: Seit App-Version v8.5 schreibt der DEX-Client beim Anmelde-/Abmelde-Versand optional Organizer-Mails ins `Bcc`-Feld der `DEX_Emails`-Queue (event-spezifisch konfigurierbar im Reiter „Kommunikation" der Event-Erstellung — Modi `Never` / `Always` / `FromDate` für Anmeldungen, `Never` / `Always` / `AfterDeadline` für Abmeldungen). Der Flow muss dieses Feld an den Send-Mail-Connector durchreichen, sonst gehen die Bestätigungs-Mails ohne BCC raus, obwohl die App es so vorsieht.

`Cc` und `Bcc` sind im SP-Schema **Multi-Line-Plain-Text** (RichText=false), Werte semikolon-separiert. Wenn das Feld leer ist, gibt der Connector keinen Header aus — die Mail geht ganz normal nur an `To` raus.

```json
TRIGGER:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline-1",
      "operationId": "GetOnNewItems"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "splitOn": "@triggerOutputs()?['body/value']"
}

GET_CONFIG (Logo + Default-Bild aus DEX_EmailTemplates via GetItems):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "2c428d35-e6fb-42f9-8a20-580acd6d05f4",
      "$filter": "TemplateType eq '_Config'",
      "$top": 1
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetItems"
    }
  },
  "runAfter": {}
}

GET_EVENT (Event-Daten für EventId):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "$filter": "@concat('ID eq ', triggerBody()?['EventId'])",
      "$top": 1
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetItems"
    }
  },
  "runAfter": { "Get_Config": ["Succeeded"] }
}

COMPOSE_LOGO (Logo Base64 aus Config):
{
  "type": "Compose",
  "inputs": "@first(body('Get_Config')?['value'])?['LogoBase64']",
  "runAfter": { "Get_Event": ["Succeeded"] }
}

COMPOSE_IMAGE (Event-Bild oder Default-Bild):
{
  "type": "Compose",
  "inputs": "@if(empty(first(outputs('Get_Event')?['body/value'])?['EmailImageBase64']), first(body('Get_Config')?['value'])?['DefaultImageBase64'], first(outputs('Get_Event')?['body/value'])?['EmailImageBase64'])",
  "runAfter": { "Compose_Logo": ["Succeeded"] }
}

SEND_EMAIL (Shared Mailbox, mit Cc + Bcc seit v8.5):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "emailMessage/MailboxAddress": "no_reply.events@deloitte.de",
      "emailMessage/To": "@triggerBody()?['Recipient']",
      "emailMessage/Subject": "@triggerBody()?['Title']",
      "emailMessage/Body": "<p class=\"editor-paragraph\">@{replace(replace(triggerBody()?['Body'], '{{LOGO_URL}}', outputs('Compose_Logo')), '{{ORB_URL}}', outputs('Compose_Image'))}</p>",
      "emailMessage/Cc": "@triggerBody()?['Cc']",
      "emailMessage/Bcc": "@triggerOutputs()?['body/Bcc']",
      "emailMessage/Importance": "Normal"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "SharedMailboxSendEmailV2"
    }
  },
  "runAfter": { "Compose_Image": ["Succeeded"] }
}

SET_SENT:
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/Status/Value": "Sent",
      "item/SentDate": "@utcNow()"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline-1",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Send_an_email_from_a_shared_mailbox_(V2)": ["Succeeded"] }
}

SET_FAILED (Email-Versand fehlgeschlagen):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/Status/Value": "Failed"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline-1",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Send_an_email_from_a_shared_mailbox_(V2)": ["Failed"] }
}
```

---

## 3. DEX_CreateOutlookEvent

**Trigger:** Neuer Eintrag in DEX_Events
**Zweck:** Outlook-Kalendereintrag im Deloitte-Design erstellen (Logo + Event-Bild aus DEX_EmailTemplates) und iCalUId zurückschreiben
**Letztes Update:** 2026-04-09

### UI-Anleitung 2026-06-02 (v18.44) — Abweichendes Outlook-Datum (Start/Ende)

**Hintergrund:** Der Organizer kann im Outlook-Editor (und im Reiter „Ort &
Programm" für den Ort) ein vom Event abweichendes **Outlook-Datum** setzen
(neue Spalten `OutlookStart` / `OutlookEnd`, DateTime). Leer = der Termin nutzt
weiter `StartDate` / `EndDate` des Events. Der Flow soll bevorzugt
`OutlookStart`/`OutlookEnd` nehmen und nur bei leer auf `StartDate`/`EndDate`
zurückfallen.

**`DEX_CreateOutlookEvent` → „Create event (V4)":**
- Feld **Start time** → fx:
  ```
  convertFromUtc(coalesce(triggerBody()?['OutlookStart'], triggerBody()?['StartDate']), 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')
  ```
- Feld **End time** → fx:
  ```
  convertFromUtc(coalesce(triggerBody()?['OutlookEnd'], triggerBody()?['EndDate']), 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')
  ```

**`DEX_Outlook_Einladungen` → Compose `Build_Update_Body`:** die `start.dateTime`
und `end.dateTime` analog auf `coalesce(...OutlookStart, ...StartDate)` bzw.
`coalesce(...OutlookEnd, ...EndDate)` umstellen (jeweils `first(outputs('Get_Event_Details')?['body/value'])?['…']`).

Danach den aktuellen Flow-JSON hier einpflegen.

### UI-Anleitung 2026-06-02 (v18.42) — Betreff bearbeitbar (eigener Termin-Titel)

**Hintergrund:** Bisher war der Betreff des Outlook-Termins fest der Event-Titel
(`item/subject` ← `Title`). Mit v18.42 kann der Organizer im Outlook-Body-Editor
einen eigenen **Betreff** setzen (neue Spalte `OutlookSubject`). Leer = weiter
Event-Titel. Der Flow soll `OutlookSubject` nehmen und nur bei leer auf `Title`
zurückfallen.

**Schritte:**

1. Flow `DEX_CreateOutlookEvent` → **Bearbeiten** → Aktion **„Create event (V4)"**.
2. Feld **Subject** anklicken → bestehenden `Title`-Token entfernen → **fx /
   Expression** → eintragen:
   ```
   coalesce(triggerBody()?['OutlookSubject'], triggerBody()?['Title'])
   ```
   → **OK** → **Speichern**.
3. Für **Aktualisierungen** zusätzlich im Flow `DEX_Outlook_Einladungen`, Compose
   **`Build_Update_Body`**, die `"subject"`-Zeile auf denselben Fallback umstellen:
   ```
   "subject": "@{coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookSubject'], first(outputs('Get_Event_Details')?['body/value'])?['Title'])}"
   ```

Danach den aktuellen Flow-JSON hier einpflegen.

### UI-Anleitung 2026-06-02 (v18.34) — Ort in den Outlook-Termin übernehmen

**Hintergrund:** Das „Ort"-Feld des Outlook-Termins blieb bisher leer, weil
die `Create event (V4)`-Action keinen Location-Parameter setzt. Die SPFx-App
schreibt jetzt eine neue Spalte **`OutlookLocation`** in `DEX_Events` (lesbarer
Ort = Veranstaltungsort + Adresse, z.B. „Rheinterrasse Düsseldorf,
Schwannstraße 6, 40476 Düsseldorf"). Der Flow muss diese Spalte nur ins
Location-Feld des Termins mappen — ein einziger Klick, keine Expression.

**Schritte (alles über die UI):**

1. Flow `DEX_CreateOutlookEvent` öffnen → **Bearbeiten**.
2. Die Aktion **„Create event (V4)"** (Outlook-Termin anlegen) aufklappen.
3. Falls das Feld **Location** nicht sichtbar ist: unten auf **„Show advanced
   options"** / „Erweiterte Optionen anzeigen" klicken.
4. In das Feld **Location** klicken → rechts auf das **Blitz-Symbol**
   (dynamischer Inhalt) → in der Liste den Trigger-Wert **`OutlookLocation`**
   (aus „When an item is created", DEX_Events) auswählen.
5. **Speichern**.

Damit bekommt **jeder neu angelegte Termin** automatisch den Ort. Für
**bestehende Events** siehe die UI-Anleitung in Flow 4
(`DEX_Outlook_Einladungen`, UpdateEvent) — dort wird der Ort beim Aktualisieren
nachgezogen; die App füllt `OutlookLocation` bei Bestands-Events automatisch
nach, sobald eine Outlook-Aktion (Einladen / Aktualisieren) ansteht.

Danach den aktuellen Flow-JSON hier einpflegen. Der zu ergänzende Parameter:
`"item/location": "@triggerBody()?['OutlookLocation']"`.

### Hinweis 2026-05-21 (v11.58): KEINE Flow-Änderung für OutlookDirty nötig

Mit v11.57 gibt es auf der `DEX_Events`-Liste die Boolean-Spalte
`OutlookDirty`. Die SPFx-App pflegt sie **vollständig selbst** — keine
Power-Automate-Anpassung erforderlich:

- Save **mit** aktiviertem „Outlook aktualisieren"-Haken → die App ruft
  `queueOutlookEvent('UpdateEvent')` auf UND setzt `OutlookDirty=false`.
- Save **ohne** Haken → die App setzt `OutlookDirty=true`, damit der
  Wizard beim nächsten Aufruf die Hinweis-Box „Outlook-Synchronisation
  steht aus" zeigt.

Solange der Organizer nicht speichert, passiert nichts. Die ursprünglich
in v11.57 vorgeschlagene UI-Anleitung zur Doppel-Absicherung im Flow
ist gestrichen — sie war unnötig.

Ablauf: Trigger (neues Event) → Config laden (Logo + Default-Bild) → Compose_Logo → Compose_Image → Platzhalter in OutlookBody ersetzen → Outlook-Termin mit HTML-Body erstellen (UTC-Zeit wird per `convertFromUtc` nach Europe/Berlin konvertiert) → CalendarLink in DEX_Events speichern

**Hinweis:** Der OutlookBody wird bereits in der SPFx-App im Deloitte-HTML-Template gewrappt (mit `{{LOGO_URL}}` und `{{ORB_URL}}` Platzhaltern). Der Flow ersetzt diese Platzhalter durch Base64-Bilder.

```json
TRIGGER (Neues Item in DEX_Events):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetOnNewItems"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "splitOn": "@triggerOutputs()?['body/value']"
}

GET_CONFIG (Logo + Default-Bild aus DEX_EmailTemplates):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$select=LogoBase64,DefaultImageBase64&$top=1"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "HttpRequest"
    }
  },
  "runAfter": {}
}

COMPOSE_LOGO (Logo Base64 aus Config):
{
  "type": "Compose",
  "inputs": "@first(body('Get_Config')?['value'])?['LogoBase64']",
  "runAfter": { "Get_Config": ["SUCCEEDED"] }
}

COMPOSE_IMAGE (Event-Bild oder Default-Bild):
{
  "type": "Compose",
  "inputs": "@if(empty(triggerBody()?['EmailImageBase64']), first(body('Get_Config')?['value'])?['DefaultImageBase64'], triggerBody()?['EmailImageBase64'])",
  "runAfter": { "Compose_Logo": ["SUCCEEDED"] }
}

CREATE_EVENT_V4 (Outlook-Termin mit Deloitte-Design Body) — Stand 2026-06-02 (v18.42/v18.44):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "table": "AAMkADU5YjlkMDBiLWU2MDktNGViMy1iNGIwLTI0YWFkNDkyN2VjMABGAAAAAABjJcNB5xJWS7D2nCeePixeBwAbtMj6YVUGQJroN6O--ImBAAAAAAEGAAAbtMj6YVUGQJroN6O--ImBAAKF4fCpAAA=",
      "item/subject": "@coalesce(triggerBody()?['OutlookSubject'], triggerBody()?['Title'])",
      "item/start": "@convertFromUtc(coalesce(triggerBody()?['OutlookStart'], triggerBody()?['StartDate']), 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')",
      "item/end": "@convertFromUtc(coalesce(triggerBody()?['OutlookEnd'], triggerBody()?['EndDate']), 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')",
      "item/timeZone": "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna",
      "item/requiredAttendees": "@last(split(replace(coalesce(triggerBody()?['OrganizerEmail'], ''), '</div>', ''), '\">'))",
      "item/body": "<p class=\"editor-paragraph\">@{replace(replace(coalesce(triggerBody()?['OutlookBody'], ''), '{{LOGO_URL}}', outputs('Compose_Logo')), '{{ORB_URL}}', outputs('Compose_Image'))}</p>",
      "item/location": "@triggerBody()?['OutlookLocation']",
      "item/showAs": "busy",
      "item/responseRequested": false,
      "item/sensitivity": "private"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "V4CalendarPostItem"
    }
  },
  "runAfter": { "Compose_Image": ["Succeeded"] }
}

UPDATE_EVENT (CalendarLink zurückschreiben):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/CalendarLink": "@outputs('Create_event_(V4)')?['body/iCalUId']"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Create_event_(V4)": ["Succeeded"] }
}

PATCH_DONOTFORWARD (Stand 2026-05-22):
Nach Update_DEX_Event laeuft eine PATCH-Action auf Graph, die die
Extended Property „DoNotForward=true" auf den frisch erzeugten
Outlook-Termin setzt. Damit ist „Weiterleitung" im Outlook-Client
deaktiviert (in modernen Outlook-Versionen wird der Forward-Button
ausgegraut). `sensitivity: "private"` alleine reicht dafuer nicht —
die Extended Property ist die offizielle Mechanik.

WICHTIG zur URI: der Termin liegt im Kalender des Connection-Users
(der Power-Automate-Connection-Owner), nicht in der Shared-Mailbox
no_reply.events@deloitte.de. Letztere ist nur die Send-As-Identitaet
fuer `organizer`. Deshalb funktioniert `/me/events/{id}` — beim
ersten Versuch mit `/users/no_reply.events@.../events/{id}` kam HTTP
404 „The specified object was not found in the store."

{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "Uri": "@concat('https://graph.microsoft.com/v1.0/me/events/', outputs('Create_event_(V4)')?['body/id'])",
      "Method": "PATCH",
      "Body": "{\n  \"singleValueExtendedProperties\": [\n    {\n      \"id\": \"Boolean {00062008-0000-0000-C000-000000000046} Name DoNotForward\",\n      \"value\": \"true\"\n    }\n  ]\n}",
      "ContentType": "application/json"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "HttpRequest"
    }
  },
  "runAfter": { "Update_item": ["SUCCEEDED"] }
}

SET_FAILED (Outlook-Termin Erstellung fehlgeschlagen):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "id": "@triggerBody()?['ID']",
      "item/Title": "@triggerBody()?['Title']",
      "item/OutlookEventId": "FAILED"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "PatchItem"
    }
  },
  "runAfter": { "Create_event_(V4)": ["Failed"] }
}
```

---

## 4. DEX_Outlook_Einladungen

**Trigger:** Neuer Eintrag in DEX_Outlook (alle 5 Minuten, **Concurrency 100** seit v18.48, vorher 1)
**Zweck:** Outlook-Termin verwalten: Teilnehmer einladen/ausladen, Event-Daten aktualisieren, Termin löschen (via Graph API).
**Letztes Update:** 2026-06-02 (v18.48, im Tenant verifiziert): **Option-B-Pro-Event-Lock** ergänzt — Trigger-Concurrency von 1 auf 100 erhöht (Parallelität über verschiedene Events), abgesichert durch einen Lock pro EventId via Liste `DEX_OutlookLocks` (Spalte `EventId` mit „Eindeutige Werte erzwingen"). Davor: 2026-04-22 (v6.4): Sub-Event-Sonderlogik komplett entfernt. Sub-Events sind seit v6.4 eigene DEX_Events-Items mit gesetztem `ParentEventId` — sie laufen durch denselben Einladen/Ausladen/Update/Delete-Pfad wie Top-Level-Events, **keine** Override-Felder, **kein** separater Branch.

### Flow-Struktur (v18.48 — mit Pro-Event-Lock)

```
Trigger (DEX_Outlook, alle 5 Min, Concurrency 100)
├── Initialize_variable (LockAcquired = false, boolean)
├── Lock_erwerben (Until: LockAcquired == true, Count 400 / Timeout PT3H)
│   ├── Create_item (DEX_OutlookLocks, EventId = triggerBody EventId — Eindeutigkeit = atomarer Lock)
│   ├── Set_Lock_Acquired (LockAcquired = true)        ← runAfter Create_item SUCCEEDED
│   └── Warte_und_Retry (Wait 30 Sek)                  ← runAfter Create_item FAILED
├── Is_DeleteEvent (ActionType == DeleteEvent)          ← runAfter Lock_erwerben SUCCEEDED
│   └── TRUE: Find_Outlook_Event_For_Delete (Graph GET by CalendarLink)
│             └── Outlook_Event_Found (length > 0)
│                 └── TRUE: Delete_Outlook_Event (Graph DELETE) → Set_Sent_DeleteEvent (SP Status=Sent)
├── Get_Event_Details (SharePoint Get Items DEX_Events by EventId, runAfter Is_DeleteEvent)
├── Init_RealEventId (string)
├── Init_Attendees (array)
├── Has_OutlookEventId / Check_CalendarLink (Event.CalendarLink gesetzt?)
    ├── TRUE: Find_Outlook_Event (Graph GET by iCalUId)
    │         → Set_variable (varRealEventId)
    │         → Check_EventFound (varRealEventId nicht leer?)
    │           ├── TRUE: Get_Existing_Event → Set_variable_1 (var_Attendees)
    │           │         → Is_UpdateEvent (ActionType == UpdateEvent)
    │           │           ├── TRUE:  Build_Update_Body (Compose) → Send_an_HTTP_request (Graph PATCH Titel/Start/Ende/Body)
    │           │           └── FALSE: Check_ActionType (ActionType == Einladen)
    │           │                      ├── TRUE:  Add_Attendee + Update_Event_Einladen (PATCH attendees)
    │           │                      └── FALSE: Filter_Attendees + Update_Event_Ausladen (PATCH attendees)
    │           │         → Set_Sent (Status=Sent)
    │           └── FALSE: Set_Failed_1 (Event nicht in Outlook gefunden)
    └── FALSE: Set_Failed (kein CalendarLink im DEX_Events-Item)
├── Find_Lock (Get items DEX_OutlookLocks, Filter EventId eq '…')   ← runAfter Has_OutlookEventId [SUCCEEDED, FAILED, SKIPPED, TIMEDOUT]
└── Apply_to_each (über Find_Lock) → Delete_item (Lock IMMER freigeben)  ← runAfter Find_Lock SUCCEEDED
```

### Pro-Event-Lock — finale Action-JSONs (v18.48, im Tenant verifiziert)

Locks-Liste `DEX_OutlookLocks` (table `5682f015-49e1-4a8e-b269-2b98f9bcea54`),
Spalte `EventId` = indiziert + „Eindeutige Werte erzwingen" (von der App
provisioniert). Die vier neuen/Lock-Actions:

```json
"Initialize_variable": {
  "type": "InitializeVariable",
  "inputs": { "variables": [ { "name": "LockAcquired", "type": "boolean", "value": false } ] },
  "runAfter": {}
}

"Lock_erwerben": {
  "type": "Until",
  "expression": "@equals(variables('LockAcquired'),true)",
  "limit": { "count": 400, "timeout": "PT3H" },
  "actions": {
    "Create_item": {
      "type": "OpenApiConnection",
      "inputs": { "parameters": {
        "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
        "table": "5682f015-49e1-4a8e-b269-2b98f9bcea54",
        "item/Title": "@triggerBody()?['EventId']",
        "item/EventId": "@triggerBody()?['EventId']",
        "item/LockedAt": "@utcNow()"
      }, "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PostItem" } }
    },
    "Set_Lock_Acquired": {
      "type": "SetVariable",
      "inputs": { "name": "LockAcquired", "value": true },
      "runAfter": { "Create_item": [ "SUCCEEDED" ] }
    },
    "Warte_und_Retry": {
      "type": "Wait",
      "inputs": { "interval": { "count": 30, "unit": "Second" } },
      "runAfter": { "Create_item": [ "FAILED" ] }
    }
  },
  "runAfter": { "Initialize_variable": [ "SUCCEEDED" ] }
}

"Find_Lock": {
  "type": "OpenApiConnection",
  "inputs": { "parameters": {
    "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
    "table": "5682f015-49e1-4a8e-b269-2b98f9bcea54",
    "$filter": "EventId eq '@{triggerBody()?['EventId']}'"
  }, "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "GetItems" } },
  "runAfter": { "Has_OutlookEventId": [ "SUCCEEDED", "TIMEDOUT", "SKIPPED", "FAILED" ] }
}

"Apply_to_each_Lock": {
  "type": "foreach",
  "foreach": "@outputs('Find_Lock')?['body/value']",
  "actions": {
    "Delete_item": {
      "type": "OpenApiConnection",
      "inputs": { "parameters": {
        "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
        "table": "5682f015-49e1-4a8e-b269-2b98f9bcea54",
        "id": "@item()?['ID']"
      }, "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "DeleteItem" } }
    }
  },
  "runAfter": { "Find_Lock": [ "SUCCEEDED" ] }
}
```

**Korrektheit (warum das atomar ist):** Der Lock-Erwerb ist **nicht** ein
Get-dann-Create (das hätte eine Race-Lücke), sondern der `Create_item` selbst —
nur **ein** Lauf kann die Zeile mit einer gegebenen `EventId` anlegen, alle
anderen scheitern an der Eindeutigkeit (HTTP 4xx, nicht retrybar). `Warte_und_Retry`
(runAfter **FAILED**) fängt den Fehlschlag ab, damit die `Until` weiterläuft statt
abzubrechen — **ohne** diesen Zweig würde die Schleife beim ersten Konflikt mit
Failed abbrechen, der Lauf die Verarbeitung überspringen und am Ende über
`Find_Lock` sogar das Lock des aktiven Laufs löschen. `Find_Lock` läuft in **allen
vier** Zuständen, damit der Lock auch bei Fehler/Skip/Timeout immer freigegeben
wird (sonst bliebe das Event dauerhaft blockiert).

### ActionTypes

- `Einladen` — einzelnen Teilnehmer zum Outlook-Termin hinzufügen
- `Ausladen` — einzelnen Teilnehmer aus dem Outlook-Termin entfernen
- `UpdateEvent` — Titel, Start, Ende **und Body** des Outlook-Termins aktualisieren (seit 2026-04-17: `OutlookBody` aus DEX_Events mit aufgelöstem `{{ORB_URL}}` per Graph PATCH; vorher blieb der Body vom initialen Create unverändert)
- `DeleteEvent` — kompletten Outlook-Termin löschen. Das DEX_Outlook-Queue-Item enthält `CalendarLink` (iCalUId) direkt, weil das zugehörige DEX_Events-Item bereits gelöscht ist wenn der Flow läuft.

### Warum keine Sub-Event-Sonderlogik (seit v6.4)

Sub-Events sind eigene DEX_Events-Items mit `ParentEventId` ≠ leer. Dadurch:
- Der **`DEX_CreateOutlookEvent`**-Flow legt für jedes Sub-Event automatisch einen eigenen Kalendertermin an (wie für Top-Level-Events).
- Anmeldungen werden in der **eigenen Subsite** des Sub-Events gespeichert (eigene Teilnehmerliste, QR-Codes, Warteliste).
- **`DEX_Outlook_Einladungen`** (dieser Flow) bekommt für jede Sub-Event-Anmeldung eine ganz normale `Einladen`-Queue-Zeile mit EventId=Sub-Event-ID. Er findet das Calendar-Event über den CalendarLink im Sub-Event-DEX_Events-Item und fügt den Attendee hinzu — keine Sonderwege nötig.
- Bei **Parent-Löschung** löscht die App kaskadierend alle Child-Events (`EventContext.deleteEvent` iteriert `events.filter(e.parentEventId === parentId)`). Jedes Child-Delete queued sein eigenes `DeleteEvent`-Item → Is_DeleteEvent-Branch räumt den Kalender auf.

### Parent-Event-Update (seit 2026-04-17, unverändert)

Ablauf:
1. Trigger (neues DEX_Outlook-Item)
2. **Is_DeleteEvent?** → Ja: Outlook-Termin per `triggerBody()?['CalendarLink']` finden (iCalUId) → DELETE → Status=Sent. Nein: weiter.
3. Event-Details laden (DEX_Events via EventId) → CalendarLink vorhanden? → Outlook-Event per iCalUId finden → Event-ID speichern → Event gefunden?
4. Bestehende Attendees laden → Is_UpdateEvent? → Ja: PATCH Titel/Start/Ende **+ Body** → Nein: Einladen/Ausladen → Status=Sent

**ActionTypes:**
- `Einladen` — einzelnen Teilnehmer zum Outlook-Termin hinzufuegen
- `Ausladen` — einzelnen Teilnehmer aus dem Outlook-Termin entfernen
- `UpdateEvent` — Titel, Start, Ende **und Body** des Outlook-Termins aktualisieren
  (seit 2026-04-17 wird `OutlookBody` aus DEX_Events mit aufgeloestem
  `{{ORB_URL}}` per Graph PATCH gesetzt — vorher blieb der Body vom
  initialen Create unveraendert)
- `DeleteEvent` — kompletten Outlook-Termin loeschen. Das DEX_Outlook-Queue-Item
  enthaelt `CalendarLink` (iCalUId) direkt, weil das zugehoerige DEX_Events-Item
  bereits geloescht ist, wenn der Flow laeuft.

**Concurrency:** 1 (sequentielle Verarbeitung, max 100 wartende Runs) —
**mit v18.48 (Option B) auf 25 erhöht**, siehe UI-Anleitung unten.

### UI-Anleitung 2026-06-02 (v18.48) — Option B: Pro-Event-Lock für parallele Outlook-Läufe

**Problem:** Der Flow lief mit **Concurrency 1** streng seriell. Bei Grossevents
patcht jeder Lauf die **komplette** Teilnehmerliste (bis zu 1500 Personen) an
Graph — ein einzelner Lauf dauert dann 13–40 Minuten. Anmeldungen für völlig
**unterschiedliche** Events stauten sich dahinter stundenlang in der Queue.

**Lösung Option B:** Die Trigger-Concurrency wird hochgesetzt (25 parallele
Läufe), damit Läufe für **verschiedene** Events gleichzeitig laufen. Ein
**Pro-Event-Lock** verhindert, dass zwei Läufe für **dasselbe** Event
gleichzeitig die Teilnehmerliste lesen-und-schreiben (sonst gehen Einträge
verloren). Der Lock nutzt die neue SP-Liste **`DEX_OutlookLocks`** (wird von
der App automatisch beim nächsten Laden angelegt — Spalte `EventId` mit
„Eindeutige Werte erzwingen"). **Falls die Liste nicht automatisch erscheint:**
manuell anlegen (Liste „DEX_OutlookLocks", Spalte `EventId` Einzeltext → in den
Spalteneinstellungen „Eindeutige Werte erzwingen" = Ja).

**So wird der Lock erworben:** Der Flow legt ein Lock-Item mit `EventId` an.
Gelingt das → er hat den Lock. Schlägt es fehl (Eindeutigkeits-Prüfung, weil
ein anderer Lauf für dasselbe Event schon ein Item hat) → kurz warten und
erneut versuchen. Am Ende löscht der Flow das Lock-Item wieder.

**Schritt 1 — Concurrency erhöhen:**
1. Flow öffnen → Bearbeiten → auf den **Trigger** (Neues Item in DEX_Outlook)
   klicken.
2. Oben rechts auf die **drei Punkte (⋮)** → **Settings**.
3. **Concurrency Control** auf **On** stellen, **Degree of Parallelism** auf
   **25** setzen → **Done**.

**Schritt 2 — Lock erwerben (ganz am Anfang, direkt nach dem Trigger,
VOR `Is_DeleteEvent`):**
1. **Variable initialisieren** „Init_LockAcquired": **+ Neuer Schritt** →
   **Initialize variable** → Name `LockAcquired`, Typ **Boolean**, Wert `false`.
   (Diese Action ans erste Position direkt nach dem Trigger ziehen.)
2. **Do until** „Lock_erwerben" hinzufügen. Bedingung: Variable
   `LockAcquired` **is equal to** `true`. Unter **Change limits**: Count `400`,
   Timeout `PT3H`.
   **Wichtig:** In einer Do-until zählen **Count UND Timeout gleichzeitig** —
   die Schleife endet, sobald das **erste** Limit erreicht ist. Bei `Count 60`
   × 15-Sek-Delay wäre also schon nach **~15 Minuten** Schluss, egal wie hoch
   der Timeout steht. Da ein einzelner Lauf fürs selbe Event 13–40 Minuten
   dauert und sich bei Grossevents mehrere stapeln können, muss der wartende
   Lauf **Stunden** tolerieren. Faustregel: Count × Delay ≳ Timeout, hier
   `400 × 30 Sek ≈ 3,3 h` ≥ `PT3H`. Wer noch mehr Puffer will, setzt z.B.
   `Count 600` + `PT5H`. (Ein dauerhaft hängender Lock ist trotzdem
   unwahrscheinlich, weil die Release-Action unten **immer** läuft — der
   Timeout ist nur das Sicherheitsnetz gegen einen abgestürzten Lauf.)
3. **In** die Do-until-Schleife eine **Scope**-Action „Try_Claim_Lock" legen.
   Darin eine **Create item**-Action (SharePoint) auf die Liste
   **`DEX_OutlookLocks`**:
   - `EventId` = im **Expression-Tab (fx)**: `triggerBody()?['EventId']`
   - `Title` = im **fx**: `triggerBody()?['EventId']`
   - `LockedAt` = im **fx**: `utcNow()`
4. **Nach** der Scope (noch innerhalb Do until) eine **Set variable**-Action
   „Set_Lock_Acquired": `LockAcquired` = `true`.
   → Bei dieser Action über **⋮ → Configure run after** **nur** „Try_Claim_Lock
   **is successful**" anhaken (die anderen Häkchen entfernen).
5. **Daneben** (ebenfalls nach der Scope) eine **Delay**-Action
   „Warte_und_Retry": **30 Sekunden** (passend zur Count/Timeout-Rechnung oben).
   → Bei dieser Action über **⋮ → Configure run after** **nur** „Try_Claim_Lock
   **has failed**" anhaken. (So wird nur gewartet, wenn der Lock gerade von
   einem anderen Lauf desselben Events gehalten wird; danach läuft die Schleife
   erneut, weil `LockAcquired` noch `false` ist.)

**Schritt 3 — bestehende Logik anhängen:** Die erste bestehende Action
(`Is_DeleteEvent`) über **⋮ → Configure run after** so setzen, dass sie nach
**`Lock_erwerben` is successful** läuft (statt nach dem Trigger). Damit läuft
die komplette bisherige Verarbeitung erst, **nachdem** der Lock erworben wurde.

**Schritt 4 — Lock freigeben (ganz am Ende, IMMER ausführen):**
1. Als **letzte** Action eine **Get items** „Find_Lock" auf
   **`DEX_OutlookLocks`** mit Filter (Expression):
   `EventId eq '@{triggerBody()?['EventId']}'`.
2. Danach **Apply to each** über `Find_Lock` → darin **Delete item**
   „Lock_freigeben" (Id = `ID` aus Find_Lock).
3. **Wichtig:** Bei **`Find_Lock`** über **⋮ → Configure run after** **alle vier**
   Häkchen setzen — **is successful, has failed, is skipped, has timed out** —
   damit der Lock auch dann gelöscht wird, wenn die Verarbeitung mittendrin
   fehlschlägt. Sonst bliebe ein „toter" Lock liegen und das Event wäre
   dauerhaft blockiert.

**Caveat:** Läufe für dasSELBE Event serialisieren weiterhin (per Lock) — das
ist gewollt (Race-Schutz). Der Gewinn ist die **Parallelität über
verschiedene Events hinweg**. Falls ein Lock durch einen Flow-Abbruch doch mal
hängenbleibt, kann der Eintrag in `DEX_OutlookLocks` manuell gelöscht werden
(die `LockedAt`-Spalte zeigt das Alter).

**Bitte nach dem Umbau:** den aktuellen Flow-JSON (Code View → kopieren)
schicken, damit der vollständige Stand hier in `docs/flow-jsons.md` eingepflegt
werden kann.

**UpdateEvent-Pattern** (seit 2026-04-17): String-Concat mit `json()` ist nicht
robust genug fuer beliebige HTML-Inhalte (Quotes/Newlines/Sonderzeichen brechen
das Parsing). Stattdessen wird der Body via Compose-Action vorgebaut und im
HTTP-PATCH referenziert — Logic Apps escaped die `@{...}`-Tokens automatisch.

**Compose `Build_Update_Body`** (vor `Send_an_HTTP_request` ausfuehren) — Stand
2026-06-02 (v18.42/v18.44, im Tenant verifiziert): Betreff/Start/Ende/Ort nutzen
`coalesce(Outlook<X>, <Original>)` — leerer Override ⇒ Event-Wert:
```json
{
  "subject": "@{coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookSubject'], first(outputs('Get_Event_Details')?['body/value'])?['Title'])}",
  "start": {
    "dateTime": "@{convertFromUtc(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookStart'], first(outputs('Get_Event_Details')?['body/value'])?['StartDate']), 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')}",
    "timeZone": "W. Europe Standard Time"
  },
  "end": {
    "dateTime": "@{convertFromUtc(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookEnd'], first(outputs('Get_Event_Details')?['body/value'])?['EndDate']), 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')}",
    "timeZone": "W. Europe Standard Time"
  },
  "showAs": "busy",
  "responseRequested": false,
  "sensitivity": "private",
  "location": {
    "displayName": "@{coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookLocation'], '')}"
  },
  "body": {
    "contentType": "html",
    "content": "@{replace(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookBody'], ''), '{{ORB_URL}}', coalesce(first(outputs('Get_Event_Details')?['body/value'])?['EmailImageBase64'], ''))}"
  }
}
```

**UI-Anleitung 2026-06-02 (v18.34) — Ort beim Aktualisieren übernehmen
(bestehende Events):** Damit auch **bestehende** Outlook-Termine den Ort
bekommen, wird im `UpdateEvent`-Zweig der Ort mit-gepatcht. In der
Compose-Action **`Build_Update_Body`** im JSON die Zeile
`"location": { "displayName": "@{coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookLocation'], '')}" }`
ergänzen (siehe oben). Bestands-Events bekommen den Ort dann, sobald der
Organizer das Event bearbeitet und das Outlook-Update bestätigt — die SPFx-App
(v18.34) wertet eine reine **Ort-Änderung** jetzt ebenfalls als
Outlook-relevant und füllt `OutlookLocation` in `DEX_Events` automatisch nach.

**Stand 2026-05-22 (v11.88):** der Body schreibt zusaetzlich `showAs: busy`
+ `responseRequested: false` + `sensitivity: private`. Damit landet der
Termin direkt im Kalender der Teilnehmer (kein Akzeptieren-Klick noetig),
ist als Beschaeftigt markiert UND kann nicht an Dritte weitergeleitet
werden (Outlook deaktiviert den Forward-Button bei privater
Vertraulichkeitsstufe). Im `DEX_CreateOutlookEvent`-Flow sind die gleichen
drei Parameter in der `Create_event_(V4)`-Action gesetzt.

**Hinweis fuer Teilnehmer:** „Privat" beschraenkt nur Weiterleitung und
Free/Busy-Anzeige fuer Dritte (Kollegen sehen nur „Privat — Beschaeftigt",
keine Titel). Der eingeladene Teilnehmer selbst sieht den Termin ganz
normal mit allen Details.

**Vollstaendiger Stand der Compose-Action `Build_Update_Body` (Code-View 2026-05-22):**
```json
{
  "type": "Compose",
  "inputs": {
    "subject": "@{first(outputs('Get_Event_Details')?['body/value'])?['Title']}",
    "start": {
      "dateTime": "@{convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['StartDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')}",
      "timeZone": "W. Europe Standard Time"
    },
    "end": {
      "dateTime": "@{convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['EndDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss')}",
      "timeZone": "W. Europe Standard Time"
    },
    "showAs": "busy",
    "responseRequested": false,
    "sensitivity": "private",
    "body": {
      "contentType": "html",
      "content": "@{replace(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['OutlookBody'], ''), '{{ORB_URL}}', coalesce(first(outputs('Get_Event_Details')?['body/value'])?['EmailImageBase64'], ''))}"
    }
  }
}
```

**`Send_an_HTTP_request`-Body** (PATCH zur Graph API):
```
@outputs('Build_Update_Body')
```

### Is_DeleteEvent-Branch (ganz am Anfang, vor Get_Event_Details)

```json
{
  "type": "If",
  "expression": {
    "and": [
      { "equals": ["@triggerBody()?['ActionType']?['Value']", "DeleteEvent"] }
    ]
  },
  "actions": {
    "Find_Outlook_Event_For_Delete": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events?$filter=iCalUId eq ''', triggerBody()?['CalendarLink'], '''')",
          "Method": "GET",
          "ContentType": "application/json"
        },
        "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365", "connection": "shared_office365", "operationId": "HttpRequest" }
      }
    },
    "Outlook_Event_Found": {
      "type": "If",
      "expression": { "and": [ { "greater": ["@length(body('Find_Outlook_Event_For_Delete')?['value'])", 0] } ] },
      "actions": {
        "Delete_Outlook_Event": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', first(body('Find_Outlook_Event_For_Delete')?['value'])?['id'])",
              "Method": "DELETE",
              "ContentType": "application/json"
            },
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365", "connection": "shared_office365", "operationId": "HttpRequest" }
          }
        },
        "Set_Sent_DeleteEvent": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
              "table": "d794655b-c950-416c-a478-5dbae285e46d",
              "id": "@triggerBody()?['ID']",
              "item/Title": "@triggerBody()?['Title']",
              "item/Status/Value": "Sent",
              "item/SentDate": "@utcNow()"
            },
            "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connection": "shared_sharepointonline", "operationId": "PatchItem" }
          },
          "runAfter": { "Delete_Outlook_Event": ["Succeeded", "Failed"] }
        }
      },
      "else": { "actions": {} },
      "runAfter": { "Find_Outlook_Event_For_Delete": ["Succeeded"] }
    }
  },
  "else": { "actions": {} },
  "runAfter": {}
}
```

Wichtig: `Get_Event_Details` hat nach Einfuegen dieser Condition `runAfter = { "Is_DeleteEvent": ["Succeeded"] }`, damit die Haupt-Logik nur laeuft wenn es kein DeleteEvent ist. Der Else-Zweig von `Is_DeleteEvent` ist leer (alle weiteren Actions kommen sowieso nach der Condition).

### SharePoint-Liste DEX_Outlook

Fuer `DeleteEvent` muessen folgende Schema-Aenderungen vorgenommen werden (werden bei neuen Listen automatisch von `ensureOutlookList()` angelegt, bei bestehenden muss der Admin sie manuell ergaenzen):

- **Choice `ActionType`** erweitern um `DeleteEvent`
- **Neue Spalte** `CalendarLink` (Multiple lines of text, plain) — enthaelt die iCalUId, damit der Flow das Outlook-Event auch ohne Zugriff auf DEX_Events finden kann.


```json
TRIGGER (Neues Item in DEX_Outlook, Concurrency: 1):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "d794655b-c950-416c-a478-5dbae285e46d"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetOnNewItems"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "runtimeConfiguration": { "concurrency": { "runs": 1, "maximumWaitingRuns": 100 } },
  "splitOn": "@triggerOutputs()?['body/value']"
}

GET_EVENT_DETAILS (Event-Daten für EventId):
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
      "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
      "$filter": "@concat('ID eq ', triggerBody()?['EventId'])",
      "$top": 1
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "GetItems"
    }
  },
  "runAfter": {}
}

INIT_REALEVENTID (Variable für Outlook Event-ID):
{
  "type": "InitializeVariable",
  "inputs": { "variables": [{ "name": "varRealEventId", "type": "string" }] },
  "runAfter": { "Get_Event_Details": ["Succeeded"] }
}

INIT_ATTENDEES (Variable für Teilnehmer-Array):
{
  "type": "InitializeVariable",
  "inputs": { "variables": [{ "name": "var_Attendees", "type": "array" }] },
  "runAfter": { "Init_RealEventId": ["Succeeded"] }
}

CHECK_CALENDARLINK (CalendarLink vorhanden? + Einladen/Ausladen/UpdateEvent):
{
  "type": "If",
  "expression": {
    "and": [{ "greater": ["@length(coalesce(first(outputs('Get_Event_Details')?['body/value'])?['CalendarLink'], ''))", 0] }]
  },
  "actions": {
    "Set_Sent": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
          "table": "d794655b-c950-416c-a478-5dbae285e46d",
          "id": "@triggerBody()?['ID']",
          "item/Title": "@triggerBody()?['Title']",
          "item/Status/Value": "Sent"
        },
        "host": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
          "connection": "shared_sharepointonline",
          "operationId": "PatchItem"
        }
      },
      "runAfter": { "Check_EventFound": ["Succeeded"] }
    },
    "Set_variable": {
      "type": "SetVariable",
      "inputs": { "name": "varRealEventId", "value": "@first(body('Find_Outlook_Event')?['value'])?['id']" },
      "runAfter": { "Find_Outlook_Event": ["Succeeded"] }
    },
    "Find_Outlook_Event": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events?$filter=iCalUId eq ''', first(outputs('Get_Event_Details')?['body/value'])?['CalendarLink'], '''')",
          "Method": "GET",
          "ContentType": "application/json"
        },
        "host": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
          "connection": "shared_office365",
          "operationId": "HttpRequest"
        }
      }
    },
    "Check_EventFound": {
      "type": "If",
      "expression": {
        "and": [{ "greater": ["@length(coalesce(variables('varRealEventId'), ''))", 0] }]
      },
      "actions": {
        "Get_Existing_Event": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
              "Method": "GET",
              "ContentType": "application/json"
            },
            "host": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
              "connection": "shared_office365",
              "operationId": "HttpRequest"
            }
          }
        },
        "Set_variable_1": {
          "type": "SetVariable",
          "inputs": { "name": "var_Attendees", "value": "@body('Get_Existing_Event')?['attendees']" },
          "runAfter": { "Get_Existing_Event": ["Succeeded"] }
        },
        "Set_Failed_1_1": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
              "table": "d794655b-c950-416c-a478-5dbae285e46d",
              "id": "@triggerBody()?['ID']",
              "item/Title": "@triggerBody()?['Title']",
              "item/Status/Value": "Failed"
            },
            "host": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
              "connection": "shared_sharepointonline",
              "operationId": "PatchItem"
            }
          },
          "runAfter": { "Is_UpdateEvent": ["Failed"] }
        },
        "Is_UpdateEvent": {
          "type": "If",
          "expression": {
            "and": [{ "equals": ["@triggerBody()?['ActionType']?['Value']", "UpdateEvent"] }]
          },
          "actions": {
            "Send_an_HTTP_request (PATCH Event-Daten)": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
                  "Method": "PATCH",
                  "Body": "@json(concat('{\"subject\":\"', first(outputs('Get_Event_Details')?['body/value'])?['Title'], '\",\"start\":{\"dateTime\":\"', convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['StartDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss'), '\",\"timeZone\":\"W. Europe Standard Time\"},\"end\":{\"dateTime\":\"', convertFromUtc(first(outputs('Get_Event_Details')?['body/value'])?['EndDate'], 'W. Europe Standard Time', 'yyyy-MM-ddTHH:mm:ss'), '\",\"timeZone\":\"W. Europe Standard Time\"}}'))",
                  "ContentType": "application/json"
                },
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                  "connection": "shared_office365",
                  "operationId": "HttpRequest"
                }
              }
            }
          },
          "else": {
            "actions": {
              "Check_ActionType (Einladen oder Ausladen)": {
                "type": "If",
                "expression": {
                  "and": [{ "equals": ["@triggerBody()?['ActionType']?['Value']", "Einladen"] }]
                },
                "actions": {
                  "Add_Attendee": {
                    "type": "AppendToArrayVariable",
                    "inputs": {
                      "name": "var_Attendees",
                      "value": "@json(concat('{\"type\":\"required\",\"status\":{\"response\":\"none\",\"time\":\"0001-01-01T00:00:00Z\"},\"emailAddress\":{\"name\":\"', triggerBody()?['Attendee'], '\",\"address\":\"', triggerBody()?['Attendee'], '\"}}'))"
                    }
                  },
                  "Update_Event_Einladen": {
                    "type": "OpenApiConnection",
                    "inputs": {
                      "parameters": {
                        "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
                        "Method": "PATCH",
                        "Body": "@json(concat('{\"attendees\":', string(variables('var_Attendees')), '}'))",
                        "ContentType": "application/json"
                      },
                      "host": {
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                        "connection": "shared_office365",
                        "operationId": "HttpRequest"
                      }
                    },
                    "runAfter": { "Add_Attendee": ["Succeeded"] }
                  }
                },
                "else": {
                  "actions": {
                    "Filter_Attendees": {
                      "type": "Query",
                      "inputs": {
                        "from": "@variables('var_Attendees')",
                        "where": "@not(equals(toLower(item()?['emailAddress']?['address']),toLower(triggerBody()?['Attendee'])))"
                      }
                    },
                    "Update_Event_Ausladen": {
                      "type": "OpenApiConnection",
                      "inputs": {
                        "parameters": {
                          "Uri": "@concat('https://graph.microsoft.com/v1.0/users/no_reply.events@deloitte.de/events/', variables('varRealEventId'))",
                          "Method": "PATCH",
                          "Body": "@json(concat('{\"attendees\":', string(body('Filter_Attendees')), '}'))",
                          "ContentType": "application/json"
                        },
                        "host": {
                          "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                          "connection": "shared_office365",
                          "operationId": "HttpRequest"
                        }
                      },
                      "runAfter": { "Filter_Attendees": ["Succeeded"] }
                    }
                  }
                }
              }
            }
          },
          "runAfter": { "Set_variable_1": ["Succeeded"] }
        }
      },
      "else": {
        "actions": {
          "Set_Failed_1 (Event nicht in Outlook gefunden)": {
            "type": "OpenApiConnection",
            "inputs": {
              "parameters": {
                "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                "table": "d794655b-c950-416c-a478-5dbae285e46d",
                "id": "@triggerBody()?['ID']",
                "item/Title": "@triggerBody()?['Title']",
                "item/Status/Value": "Failed"
              },
              "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                "connection": "shared_sharepointonline",
                "operationId": "PatchItem"
              }
            }
          }
        }
      },
      "runAfter": { "Set_variable": ["Succeeded"] }
    }
  },
  "else": {
    "actions": {
      "Set_Failed (Kein CalendarLink)": {
        "type": "OpenApiConnection",
        "inputs": {
          "parameters": {
            "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
            "table": "d794655b-c950-416c-a478-5dbae285e46d",
            "id": "@triggerBody()?['ID']",
            "item/Title": "@triggerBody()?['Title']",
            "item/Status/Value": "Failed"
          },
          "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
            "connection": "shared_sharepointonline",
            "operationId": "PatchItem"
          }
        }
      }
    }
  },
  "runAfter": { "Init_Attendees": ["Succeeded"] }
}
```

---

## Listen-GUIDs (aktuell, Stand 2026-04-07)

| Liste | GUID |
|-------|------|
| DEX_IDReorder | 9d46ff77-5fe2-4e1d-9b93-14b9dca1a360 |
| DEX_Events | 28457815-1163-4e92-8b08-3ae43f477d9e |
| DEX_Emails | 57aa0840-df98-41ae-a39b-323c0b80ae3b |
| DEX_Outlook | d794655b-c950-416c-a478-5dbae285e46d |
| DEX_EmailTemplates | 2c428d35-e6fb-42f9-8a20-580acd6d05f4 |

---

# FLOW: DEX_OutlookDeclineHandler (Mail-basiert + Decline-Digest, Stand 2026-05-05)

**Zweck:** Zwei verzahnte Aufgaben in einem Flow:

1. **Reminder an den Decliner** — wenn jemand einen Outlook-Kalender-Termin im `no_reply.events@deloitte.de`-Postfach ablehnt, aber in der DEX-Teilnehmerliste noch als angemeldet steht, bekommt er eine Erinnerungsmail mit Action-Button „Anmeldung stornieren". Der Link öffnet die DEX App im Abmelde-Modus (Deep-Link `?action=cancel&event=<eventNumber>`).
2. **Digest an die Organizer** *(seit 2026-05-05, app-seitig ab v9.38 mit `OutlookDeclineDigest`-Template)* — direkt nach dem Reminder fragt der Flow den **echten Outlook-Termin** via `V3CalendarGetItem` ab und liest pro Attendee den `status.response`-Wert aus. Alle mit `response == 'declined'` werden gegen die Teilnehmerliste gegen-gecheckt; jeder, der dort noch `Status == 'Angemeldet'` steht, landet als Zeile in einer HTML-Tabelle. Die Tabelle geht als `OutlookDeclineDigest`-Mail an die Organizer.

**Trigger:** `When a new email arrives (V3)` auf der Shared Mailbox — nicht `When an event is modified`, weil der Event-Trigger keine pro-Attendee-`status.response`-Felder liefert und damit nicht erkennbar ist, WER abgelehnt hat.

## Bekannte Einschränkungen

| Fall | Verhalten | Abdeckung |
|------|-----------|-----------|
| User lehnt ab + sendet Antwort (Default) | ✅ Mail kommt → Flow feuert → Reminder wird gequeued | ~85% |
| User lehnt ab + "keine Antwort senden" | ❌ Keine Mail → kein Reminder (silent decline) | Lücke |
| User hat Outlook auf DE/EN/FR/IT | ✅ Subject-Filter deckt alle ab | OK |
| User hat Outlook auf anderer Sprache (PL/TR/ES/...) | ❌ Subject wird nicht erkannt | Lücke |
| Zwei Events mit identischem Titel | ⚠️ Flow nimmt das erste in DEX_Events | Edge Case |

Für 100% Abdeckung wäre ein Graph-API-Polling-Flow nötig (kein Ziel für jetzt).

## UI-Anleitung zum Anlegen (Schritt für Schritt)

### 1. Neuer Cloud-Flow anlegen

1. https://make.powerautomate.com öffnen.
2. Links **+ Create** → **Automated cloud flow**.
3. **Flow name:** `DEX_OutlookDeclineHandler`.
4. Trigger: `When a new email arrives V3` eintippen → **Office 365 Outlook — When a new email arrives (V3)** auswählen.
5. **Create**.

### 2. Trigger konfigurieren

- **Folder:** `Inbox`.
- **Show advanced options** klicken:
  - **Original Mailbox Address:** `no_reply.events@deloitte.de`.
  - **Include Attachments:** `No`.
  - **Subject Filter:** leer lassen (Sprachabhängigkeit wird in Schritt 3 über Condition gelöst).
  - **Importance, From, To, CC, Has Attachment:** leer lassen.

### 3. Condition: Ist das eine Decline-Mail? (`Is_Decline_Mail`)

- **+ New step** → **Control — Condition**.
- Linke Seite: **Expression-Tab (fx)** →
  ```
  or(
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined:'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined '),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt:'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt '),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'refusé'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'rifiutato:'),
    and(
      or(
        startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw:'),
        startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg:')
      ),
      or(
        contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'declined:'),
        contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'abgelehnt:')
      )
    )
  )
  ```
  Der letzte `and(...)`-Block deckt **weitergeleitete** Decline-Mails ab:
  Subject startet mit `FW:` / `WG:` UND der `bodyPreview` enthaelt `Declined:`
  oder `Abgelehnt:` (typisches Outlook-Forward-Format).
- Operator: `is equal to`.
- Rechte Seite: **Expression (fx)** → `true`.
- Rename **(⋮)** → `Is_Decline_Mail`.

Alle weiteren Schritte im **If yes**-Zweig; **If no** bleibt leer.

### 4. `Cleaned_Subject` (Compose)

- **Data Operation — Compose**.
- **Inputs (fx):**
  ```
  trim(
    replace(
      replace(
        replace(
          replace(
            replace(
              replace(
                replace(
                  replace(triggerOutputs()?['body/subject'], 'FW:', ''),
                  'WG:', ''),
                'Declined:', ''),
              'Declined ', ''),
            'Abgelehnt:', ''),
          'Abgelehnt ', ''),
        'Refusé :', ''),
      'Rifiutato:', '')
  )
  ```
  Zusaetzlich zu den sechs Decline-Prefixen werden auch `FW:` und `WG:`
  abgeschnitten, damit weitergeleitete Decline-Mails den reinen Event-Titel
  in `Cleaned_Subject` haben.
- Rename → `Cleaned_Subject`.

### 4a. `Real_Sender` (Compose)

- **Data Operation — Compose** (NACH `Cleaned_Subject`).
- **Inputs (fx):**
  ```
  if(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg:')), trim(first(split(first(skip(split(coalesce(triggerOutputs()?['body/bodyPreview'], ''), '<'), 1)), '>'))), triggerOutputs()?['body/from'])
  ```
  Fuer direkte Decline-Mails = `body/from` wie bisher. Fuer Forwards extrahiert
  die Expression die erste `<email>`-Adresse aus dem `bodyPreview`. Bei
  "On Behalf Of"-Forwards ist das die Adresse der **Assistenz**, nicht des
  Principals — darum braucht's unten zusaetzlich die Name-basierte Suche.
- Rename → `Real_Sender`.

### 4b. `Decliner_Lastname` (Compose)

- **Data Operation — Compose** (NACH `Real_Sender`).
- **Inputs (fx):**
  ```
  if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'On Behalf Of '), trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'On Behalf Of ')), ','))), '')
  ```
  Extrahiert den Nachnamen aus Outlook's "On Behalf Of <Nachname>, <Titel>
  <Vorname>"-Pattern. Leer bei direkten Decline-Mails.
- Rename → `Decliner_Lastname`.

### 4c. `Decliner_Firstname` (Compose)

- **Data Operation — Compose** (NACH `Decliner_Lastname`).
- **Inputs (fx):**
  ```
  if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'On Behalf Of '), trim(replace(replace(replace(last(split(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'On Behalf Of ')), decodeUriComponent('%0D%0A'))), ',')), 'Prof. ', ''), 'Dr. ', ''), 'Dipl.-Ing. ', '')), '')
  ```
  Extrahiert den Vornamen nach dem Komma, strippt die haeufigsten Titel
  (`Prof. `, `Dr. `, `Dipl.-Ing. `). Beispiel: `"Nibler, Dr. Marcus"` →
  `"Marcus"`. Leer bei direkten Decline-Mails.
- Rename → `Decliner_Firstname`.

### 5. `Get_DEX_Event` (SharePoint Get items)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Events`.
- **Show advanced options** → **Filter Query (fx):**
  ```
  concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')
  ```
- **Top Count:** `1`.
- **Configure run after** → `Decliner_Firstname` → `Succeeded`.
- Rename → `Get_DEX_Event`.

### 6. Condition `Event_Found`

- Linke Seite **(fx):** `length(outputs('Get_DEX_Event')?['body/value'])`
- Operator: `is greater than`
- Rechte Seite: `0` (Plain Text).
- Rename → `Event_Found`.

### 7. `Get_Teilnehmer_Entry` (Send HTTP request to SharePoint, im If yes)

- **SharePoint — Send an HTTP request to SharePoint**.
- **Site Address (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl']`
- **Method:** `GET`.
- **Uri (fx):** Konditional, filtert je nach verfuegbaren Daten nach E-Mail
  oder Name:
  ```
  concat('_api/web/lists/getbytitle(''Teilnehmer'')/items?$filter=', if(empty(outputs('Decliner_Lastname')), concat('ParticipantEmail eq ''', replace(outputs('Real_Sender'), '''', ''''''), ''''), if(empty(outputs('Decliner_Firstname')), concat('Nachname eq ''', replace(outputs('Decliner_Lastname'), '''', ''''''), ''''), concat('Vorname eq ''', replace(outputs('Decliner_Firstname'), '''', ''''''), ''' and Nachname eq ''', replace(outputs('Decliner_Lastname'), '''', ''''''), ''''))), ' and Status ne ''Abgemeldet''&$top=1&$select=Id,Status,Vorname,Nachname,ParticipantEmail')
  ```
  Drei Faelle:
  1. **Kein OnBehalfOf** → `ParticipantEmail eq '<Real_Sender>'` (direkter
     Decline oder regulaerer Forward)
  2. **OnBehalfOf mit Vor- und Nachname** → `Vorname eq '<x>' and Nachname eq '<y>'`
     (praeziseste Variante, vermeidet Kollision bei mehreren Niblers im Event)
  3. **OnBehalfOf nur Nachname** (Vorname-Parsing gescheitert) →
     `Nachname eq '<y>'` (Fallback)

  **Wichtig:** `ParticipantEmail` MUSS mit im `$select` stehen, damit
  `Final_Recipient_Email` die echte E-Mail-Adresse des Principals (z.B.
  Dr. Nibler) laden kann — `Real_Sender` ist bei OnBehalfOf-Faellen die
  Assistenz-Adresse und darf NICHT als Reminder-Empfaenger verwendet werden.
- **Headers:**
  - `Accept: application/json;odata=nometadata`
- Rename → `Get_Teilnehmer_Entry`.

### 7a. `Final_Recipient_Email` (Compose)

- **Data Operation — Compose** (NACH `Get_Teilnehmer_Entry`, VOR
  `Still_Registered`).
- **Inputs (fx):**
  ```
  coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['ParticipantEmail'], outputs('Real_Sender'))
  ```
  Nimmt die ParticipantEmail aus dem gefundenen Teilnehmer-Eintrag. Fallback
  auf `Real_Sender` wenn kein Teilnehmer gefunden (dann geht der Reminder eh
  nicht raus, weil `Still_Registered=false`, aber verhindert null-errors).
- **Configure run after** → `Get_Teilnehmer_Entry` → `Succeeded`.
- Rename → `Final_Recipient_Email`.

### 8. Condition `Still_Registered`

- Linke Seite **(fx):** `length(body('Get_Teilnehmer_Entry')?['value'])`
- Operator: `is greater than`
- Rechte Seite: `0`.
- **Configure run after** → `Final_Recipient_Email` → `Succeeded`.
- Rename → `Still_Registered`.

### 9. `Get_Existing_Reminder` (SharePoint Get items, im Still_Registered/yes)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Emails`.
- **Filter Query (fx):**
  ```
  concat(
    'EmailType eq ''OutlookDeclineReminder'' and EventId eq ''',
    first(outputs('Get_DEX_Event')?['body/value'])?['ID'],
    ''''
  )
  ```
  Kein `Recipient`-Filter im OData (Multi-Line-Note-Feld, `eq` nicht moeglich)
  — der nachgeschaltete `Filter_By_Recipient`-Schritt pickt den aktuellen
  Sender heraus.
- **Top Count:** `20`.
- Rename → `Get_Existing_Reminder`.

### 9a. `Filter_By_Recipient` (Data Operation — Filter array)

- **Data Operation — Filter array**.
- **From (fx):** `body('Get_Existing_Reminder')?['value']`
- **Condition (fx):**
  ```
  contains(concat(';', replace(item()?['Recipient'], ' ', ''), ';'), concat(';', outputs('Final_Recipient_Email'), ';'))
  ```
  Vergleicht gegen `outputs('Final_Recipient_Email')` (die echte Empfaenger-
  Adresse aus dem Teilnehmer-Eintrag), damit bei OnBehalfOf-Forwards die
  bereits versendete Reminder-Mail korrekt gefunden wird.
- Rename → `Filter_By_Recipient`.

### 10. Condition `No_Reminder_Yet`

- Linke Seite **(fx):** `length(body('Filter_By_Recipient'))`
- Operator: `is equal to`
- Rechte Seite: `0`.
- Rename → `No_Reminder_Yet`.

### 11. `Get_Reminder_Template` (SharePoint Get items, im No_Reminder_Yet/yes)

- **SharePoint — Get items**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_EmailTemplates`.
- **Filter Query (fx):** Konditional — Direkt-Decliner bekommen die schlichte
  Variante, OnBehalfOf bekommen die Variante mit Assistant-Forward-Button:
  ```
  concat('TemplateType eq ''', if(empty(outputs('Decliner_Lastname')), 'OutlookDeclineReminder', 'OutlookDeclineReminder_OnBehalfOf'), ''' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')
  ```
- **Top Count:** `1`.
- Rename → `Get_Reminder_Template`.

### 11a. `Assistant_Forward_Mailto` (Compose, nach `Get_Reminder_Template`)

- **Data Operation — Compose**.
- **Inputs (fx):** Baut die `mailto:`-URL fuer den Assistant-Forward-Button im
  OnBehalfOf-Template — vorausgefuellt mit Event-Organizer-Adressen, Subject
  und Body, der den Partner-Namen + Event-Titel enthaelt:
  ```
  concat('mailto:', first(outputs('Get_DEX_Event')?['body/value'])?['OrganizerEmail'], '?subject=', encodeUriComponent(concat('Bitte um Abmeldung: ', first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], ' ', first(body('Get_Teilnehmer_Entry')?['value'])?['Nachname'], ' — ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])), '&body=', encodeUriComponent(concat('Hallo,', decodeUriComponent('%0D%0A%0D%0A'), 'ich (Assistenz) habe in Vertretung den Outlook-Termin für ', first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], ' ', first(body('Get_Teilnehmer_Entry')?['value'])?['Nachname'], ' (', outputs('Final_Recipient_Email'), ') abgelehnt.', decodeUriComponent('%0D%0A%0D%0A'), 'Bitte storniere die Anmeldung für die Person für das Event "', first(outputs('Get_DEX_Event')?['body/value'])?['Title'], '" über das Admin Center der Event Experience Platform.', decodeUriComponent('%0D%0A%0D%0A'), 'Danke!')))
  ```
  Das Standard-Decliner-Template enthaelt keinen `{{AssistantForwardUrl}}`-
  Platzhalter, dort wird der Compose-Output einfach nicht verwendet.
- **Configure run after** → `Get_Reminder_Template` → `Succeeded`.
- Rename → `Assistant_Forward_Mailto`.

### 12. `Create_Reminder_Queue_Item` (SharePoint Create item, nach `Get_Reminder_Template`)

- **SharePoint — Create item**.
- **Site Address:** `DOL-c-DE-EventExperiencePlatform`.
- **List Name:** `DEX_Emails`.
- **Title (fx):**
  ```
  replace(coalesce(first(outputs('Get_Reminder_Template')?['body/value'])?['Subject'], concat('Outlook-Abmeldung-Reminder: ', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title'])
  ```
- **Recipient (fx):** `outputs('Final_Recipient_Email')`
- **RecipientName (fx):**
  ```
  coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], outputs('Final_Recipient_Email'))
  ```
  Vorname aus dem gefundenen Teilnehmer-Eintrag, damit die Mail mit
  "Dear Marcus," statt "Dear mparschalk@deloitte.de," beginnt. Fallback auf
  `Final_Recipient_Email` falls Vorname-Feld leer ist.
- **EmailType Value:** `OutlookDeclineReminder`.
- **EventTitle (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['Title']`
- **EventId (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['ID']`
- **Status Value:** `Pending`.
- **Body (fx):** Template laden und vier Platzhalter ersetzen — `DEX_SEND_MAIL`
  ersetzt danach nur noch `{{LOGO_URL}}` / `{{ORB_URL}}`:
  ```
  replace(replace(replace(replace(coalesce(first(outputs('Get_Reminder_Template')?['body/value'])?['BodyHtml'], ''), '{{Name}}', coalesce(first(body('Get_Teilnehmer_Entry')?['value'])?['Vorname'], outputs('Final_Recipient_Email'))), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title']), '{{CancelUrl}}', concat('https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView&action=cancel&event=', string(first(outputs('Get_DEX_Event')?['body/value'])?['EventNumber']))), '{{AssistantForwardUrl}}', outputs('Assistant_Forward_Mailto'))
  ```
  **Wichtig beim `{{Name}}`-Replace:** `coalesce(...Vorname, Final_Recipient_Email)`
  nutzen (nicht `Real_Sender`), sonst landet bei OnBehalfOf-Faellen die
  Assistenz-Mail-Adresse in der Anrede statt des Principal-Vornamens. Der
  vierte Replace `{{AssistantForwardUrl}}` ist NUR im OnBehalfOf-Template als
  Platzhalter vorhanden — bei Direkt-Decliner-Mails wirkt er als No-op.
- **Configure run after** → `Assistant_Forward_Mailto` → `Succeeded`.
- Rename → `Create_Reminder_Queue_Item`.

## Ablauf-Diagramm

```
Trigger: When a new email arrives (V3) — no_reply.events@deloitte.de / Inbox
   │
   ├── Is_Decline_Mail? (Subject startsWith Declined/Abgelehnt/Refusé/Rifiutato)
   │
   ├── Yes: Cleaned_Subject (strip prefix)
   │        │
   │        ├── Get_DEX_Event (Title eq Cleaned_Subject, mit Apostroph-Escape)
   │        │
   │        ├── Event_Found? (length > 0)
   │        │
   │        └── Yes: Get_Teilnehmer_Entry (ParticipantEmail = Sender, Status != Abgemeldet)
   │                   │
   │                   ├── Still_Registered? (length > 0)
   │                   │
   │                   └── Yes: Get_Existing_Reminder (Dedup check)
   │                            │
   │                            ├── No_Reminder_Yet? (length == 0)
   │                            │
   │                            └── Yes: Create_Reminder_Queue_Item in DEX_Emails
   │                                     (EmailType='OutlookDeclineReminder')
   │                                     │
   │                                     ├── Get_Outlook_EventId (Compose: DEX_Events.CalendarLink → iCalUId)
   │                                     │
   │                                     └── Has_Outlook_Event? (length > 0)
   │                                            │
   │                                            └── Yes: Get_Outlook_Event (V3CalendarGetItem)
   │                                                     │
   │                                                     ├── Set_variable (StillRegisteredCount = 0)
   │                                                     ├── Filter_Declined_Attendees (status.response == 'declined')
   │                                                     │
   │                                                     ├── Apply_To_Each_Decliner:
   │                                                     │     └── Get_Decliner_Status (Teilnehmerliste, ParticipantEmail + Status='Angemeldet')
   │                                                     │           └── Condition_Still_Registered? (length > 0)
   │                                                     │                 └── Yes: Increment_Count
   │                                                     │                          + Append_Decliner_Row to DeclineRowsHtml
   │                                                     │
   │                                                     ├── Compose_DeclineList_Table (HTML <table>)
   │                                                     │
   │                                                     └── Has_Decliners? (StillRegisteredCount > 0)
   │                                                            └── Yes: Get_Digest_Template (DEX_EmailTemplates,
   │                                                                              OutlookDeclineDigest, Lang)
   │                                                                     + Create_Digest_Queue_Item in DEX_Emails
   │                                                                              (EmailType='OutlookDeclineDigest',
   │                                                                              Recipient=OrganizerEmail)
   │
   └── DEX_SEND_MAIL picks up Pending Mails und sendet mit Template,
       {{CancelUrl}} zeigt auf DEX.aspx?action=cancel&event=<eventNumber>
       (Digest-Mail rendert {{DeclineCount}} + {{DeclineList}} aus dem
        OutlookDeclineDigest-Template, Recipients = Organizer-Mails.)
```

## Decline-Digest (Approach B via Outlook-Connector, Stand 2026-05-05)

**Hintergrund:** Vor der Digest-Erweiterung gab es nur den Reminder an den einzelnen Decliner. Die Organizer hatten **keine konsolidierte Sicht** darüber, wer aktuell zwar in der Kapazität zählt (Status `Angemeldet` in der Teilnehmerliste), den Outlook-Termin aber abgelehnt hat. Das ist relevant für Catering-/Hotel-/Bus-Planung.

**Zwei Implementierungs-Optionen waren auf dem Tisch:**

| Approach | Quelle der Decliner-Liste | Catched silent declines | Reflektiert Re-Accepts |
|----------|---------------------------|--------------------------|-------------------------|
| A — über `DEX_Emails` | Eigene OutlookDeclineReminder-Records dieses Events | ❌ Nein (nur Decliner mit „Antwort senden") | ❌ Nein (Reminder-Eintrag bleibt) |
| **B — Outlook-Connector** | `attendees[].status.response` aus dem echten Outlook-Termin | ✅ Ja | ✅ Ja |

→ **Approach B implementiert.** Approach A wurde verworfen und nicht deployed.

**Voraussetzung für B:** das Feld `CalendarLink` in `DEX_Events` ist gesetzt — wird beim erfolgreichen `DEX_CreateOutlookEvent`-Run zurückgeschrieben (enthält die `iCalUId` des Outlook-Termins, die `V3CalendarGetItem` als Lookup-Schlüssel akzeptiert). Falls leer (alte Events vor Outlook-Integration oder DEX_CreateOutlookEvent-Fehler), springt der Flow über `Has_Outlook_Event?` in den No-Branch und überspringt den Digest stillschweigend.

> **Hinweis zur Spalten-Wahl:** in der `DEX_Events`-Liste gibt es zwei Outlook-bezogene Spalten — `OutlookEventId` und `CalendarLink`. Der Decline-Digest nutzt **`CalendarLink`** (= iCalUId), weil `V3CalendarGetItem` damit über alle Mailbox-Kalender hinweg sucht und nicht an die Calendar-ID der Mailbox gekoppelt ist. `OutlookEventId` wird vom App-Code primär genutzt (siehe `EventService.findOutlookEvent`), kann aber auf alten Events fehlen — `CalendarLink` ist die robustere Quelle.

### UI-Anleitung — Schritte hinzufügen (komplett, Stand 2026-05-05)

Alle folgenden Schritte werden **innerhalb des `No_Reminder_Yet`-If-yes-Branches** angefügt, **nach** `Create_Reminder_Queue_Item`. Voraussetzung sind die zwei Top-Level-`InitializeVariable`-Actions außerhalb der äußeren If-Bedingung:

- `Init_DeclineRows_HTML` — String, default leer.
- `Init_StillRegistered_Count` — Integer, default `0`. **Run after** `Init_DeclineRows_HTML` Succeeded.

Diese zwei Actions müssen **vor** dem `Is_Decline_Mail`-If liegen (siehe finaler Flow-JSON unten).

#### 1. `Get_Outlook_EventId` (Compose)

- **+ New step** → **Data Operation — Compose**.
- **Inputs (fx):** `coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['CalendarLink'], '')`
- **Configure run after** → `Create_Reminder_Queue_Item` → `Succeeded`.
- Rename → `Get_Outlook_EventId`.

#### 2. `Has_Outlook_Event` (Condition)

- **+ New step** → **Control — Condition**.
- Linke Seite (fx): `length(outputs('Get_Outlook_EventId'))`
- Operator: `is greater than`
- Rechte Seite: `0`
- Rename → `Has_Outlook_Event`.

Alle weiteren Schritte (3–9) gehen in den **`If yes`-Branch**. `If no` bleibt leer (kein Outlook-Termin → kein Digest).

#### 3. `Get_Outlook_Event` (Office 365 Outlook — Get event V3)

- **+ Add an action** → **Office 365 Outlook — Get event (V3)**.
- Connection: gleiche, die bereits den Trigger bedient (Account mit Zugriff auf `no_reply.events@deloitte.de`).
- **Calendar Id:** der Default-Kalender des Postfachs. In unserem Tenant aktuell:
  ```
  AAMkADU5YjlkMDBiLWU2MDktNGViMy1iNGIwLTI0YWFkNDkyN2VjMABGAAAAAABjJcNB5xJWS7D2nCeePixeBwAbtMj6YVUGQJroN6O--ImBAAAAAAEGAAAbtMj6YVUGQJroN6O--ImBAAKF4fCpAAA=
  ```
  *(Bei Tenant-Wechsel: über die Connector-Dropdowns neu picken — der Wert ändert sich.)*
- **Event Id (fx):** `outputs('Get_Outlook_EventId')`

#### 4. `Set_variable` (Set variable — StillRegisteredCount = 0)

*Technisch redundant (PA-Variables resetten beim Run-Start), aber explizit für die Lesbarkeit.*

- **+ Add an action** → **Variables — Set variable**.
- Name: `StillRegisteredCount`
- Value: `0`
- Configure run after → `Get_Outlook_Event` → `Succeeded`.

#### 5. `Filter_Declined_Attendees` (Data Operation — Filter array)

- **+ Add an action** → **Data Operation — Filter array**.
- **From (fx):** `coalesce(first(body('Get_Outlook_Event')?['value'])?['attendees'], json('[]'))`
  *(`first(body('...')?['value'])` weil `Get events (V4)` ein Array zurückliefert — wir picken das erste/einzige Match. `json('[]')` als Fallback statt `createArray()`, weil `createArray()` ohne Args einen `Unable to process template language expressions`-Fehler wirft.)*
  *(coalesce: Events ohne Attendees → leeres Array, kein Null-Pointer.)*
- **Edit in advanced mode** und Filter-Expression:
  ```
  @equals(toLower(coalesce(item()?['status']?['response'], '')), 'declined')
  ```
- Run after `Set_variable` Succeeded.

#### 6. `Apply_To_Each_Decliner` (Apply to each)

- **+ Add an action** → **Control — Apply to each**.
- Output (fx): `body('Filter_Declined_Attendees')`
- Run after `Filter_Declined_Attendees` Succeeded.

**Innerhalb der Loop:**

##### 6.a — `Get_Decliner_Status` (SharePoint — Get items)

- Site Address (fx): `concat('https://deudeloitte.sharepoint.com', replace(first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl'], 'https://deudeloitte.sharepoint.com', ''))`
- List Name: leer lassen oder `Teilnehmer` — die Connection nutzt das Filter-Query, der List-Name ist nur kosmetisch.
- Filter Query (fx):
  ```
  concat('ParticipantEmail eq ''', items('Apply_To_Each_Decliner')?['emailAddress']?['address'], ''' and Status ne ''Abgemeldet''')
  ```
- Top Count: `1`

##### 6.b — `Condition_Still_Registered` (Condition)

- Linke Seite (fx): `length(body('Get_Decliner_Status')?['value'])`
- Operator: `is greater than`
- Rechte Seite: `0`

**Im `If yes`-Branch:**

###### 6.b.i — `Increment_Count` (Variables — Increment variable)

- Name: `StillRegisteredCount`
- Value: `1`

###### 6.b.ii — `Append_Decliner_Row` (Variables — Append to string variable)

- Name: `DeclineRowsHtml`
- Value (fx, eine lange Zeile):
  ```
  concat('<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;">', coalesce(first(body('Get_Decliner_Status')?['value'])?['Vorname'], ''), ' ', coalesce(first(body('Get_Decliner_Status')?['value'])?['Nachname'], ''), '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;">', items('Apply_To_Each_Decliner')?['emailAddress']?['address'], '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;">', coalesce(first(body('Get_Decliner_Status')?['value'])?['Department'], '—'), '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;">', if(empty(first(body('Get_Decliner_Status')?['value'])?['RegistrationDate']), '—', formatDateTime(first(body('Get_Decliner_Status')?['value'])?['RegistrationDate'], 'dd.MM.yyyy')), '</td></tr>')
  ```
- Run after `Increment_Count` Succeeded.

`If no`-Branch von `Condition_Still_Registered` bleibt leer.

#### 7. `Compose_DeclineList_Table` (Data Operation — Compose, **außerhalb** der Loop)

- Inputs (fx):
  ```
  concat('<table style="border-collapse:collapse;width:100%;font-size:13px;margin:12px 0;"><thead><tr style="background:#f5f5f5;"><th style="padding:8px 10px;text-align:left;">Name</th><th style="padding:8px 10px;text-align:left;">E-Mail</th><th style="padding:8px 10px;text-align:left;">Department</th><th style="padding:8px 10px;text-align:left;">Anmeldedatum</th></tr></thead><tbody>', variables('DeclineRowsHtml'), '</tbody></table>')
  ```
- Run after `Apply_To_Each_Decliner` Succeeded.

#### 8. `Has_Decliners` (Condition)

- Linke Seite (fx): `variables('StillRegisteredCount')`
- Operator: `is greater than`
- Rechte Seite: `0`
- Run after `Compose_DeclineList_Table` Succeeded.

`If no` bleibt leer (alle Decliner sind schon abgemeldet → kein Digest nötig).

**Im `If yes`-Branch:**

##### 8.a — `Get_Digest_Template` (SharePoint — Get items)

- Site Address: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- List GUID: `2c428d35-e6fb-42f9-8a20-580acd6d05f4` (DEX_EmailTemplates)
- Filter Query (fx):
  ```
  concat('TemplateType eq ''OutlookDeclineDigest'' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')
  ```
- Top Count: `1`

##### 8.b — `Create_Digest_Queue_Item` (SharePoint — Create item, in `DEX_Emails`)

- Site Address: gleich
- List GUID: `57aa0840-df98-41ae-a39b-323c0b80ae3b` (DEX_Emails)
- **Title (fx):**
  ```
  replace(replace(coalesce(first(outputs('Get_Digest_Template')?['body/value'])?['Subject'], 'FYI: attendees declined Outlook'), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title']), '{{DeclineCount}}', string(variables('StillRegisteredCount')))
  ```
- **Recipient (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['OrganizerEmail']`
- **RecipientName:** Klartext `Organizer`
- **EmailType Value:** `OutlookDeclineDigest`
- **EventTitle (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['Title']`
- **EventId (fx):** `first(outputs('Get_DEX_Event')?['body/value'])?['ID']`
- **Status Value:** `Pending`
- **Body (fx):**
  ```
  replace(replace(replace(coalesce(first(outputs('Get_Digest_Template')?['body/value'])?['BodyHtml'], ''), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title']), '{{DeclineCount}}', string(variables('StillRegisteredCount'))), '{{DeclineList}}', outputs('Compose_DeclineList_Table'))
  ```
- Run after `Get_Digest_Template` Succeeded.

`DEX_SEND_MAIL` braucht keine Änderung — der `OutlookDeclineDigest`-Eintrag verhält sich wie jeder andere Pending-Mail-Eintrag und wird gerendert + versendet.

### Test-Plan für den Digest

| Szenario | Erwartung |
|----------|-----------|
| User lehnt mit „Send response" ab | Reminder + Digest werden gequeued. Digest enthält den User in der Tabelle. |
| User lehnt mit „Don't send response" ab | **Reminder feuert nicht** (kein Mail-Trigger), aber **beim nächsten Decline** im selben Event ist er im Digest enthalten — weil der Digest direkt aus `attendees[].status.response` liest. |
| User lehnt ab und akzeptiert dann doch | Beim nächsten Decline im selben Event ist er **nicht mehr** im Digest — `status.response` zeigt jetzt `accepted`. |
| User lehnt ab und meldet sich offiziell ab | Beim nächsten Decline im selben Event ist er **nicht mehr** im Digest — `Get_Decliner_Status` filtert auf `Status ne 'Abgemeldet'`, abgemeldete fallen damit raus. |
| Event hat keinen Outlook-Termin (`CalendarLink` leer) | `Has_Outlook_Event` ist false → Digest-Branch wird übersprungen, nur Reminder wird gequeued. |
| Event hat 0 noch-angemeldete Decliner | `Has_Decliners` ist false → kein Digest gequeued. |

## App-seitige Unterstützung (bereits implementiert)

- `DexEventPlatform.tsx` parsed `?action=cancel&event=<n>` aus `window.location.search`
- Erkennt den Deep-Link schon beim ersten Render und zeigt einen Vollbild-
  Lade-Spinner (statt der LandingPage), solange die Events noch geladen werden.
  Dadurch sieht der User sofort, dass eine Aktion läuft und "hängt" nicht
  sekundenlang auf der Willkommensseite.
- Navigiert zu `my-events` mit `selectedEventId` + Intent `auto-cancel`
- `MyEventsPage.tsx` prüft den Intent, scrollt zur Event-Karte und **storniert
  die Registrierung direkt** (ohne zusätzlichen "Abmeldung bestätigen"-Klick).
  Der Klick auf den Action-Button in der Mail gilt als Bestätigung. Das ist
  sicher, weil der User eingeloggt sein muss und durch Item-Level Security
  ohnehin nur seine eigene Registrierung stornieren kann.
- Der OutlookDeclineReminder-Template-Body wird als fertig gewrapptes
  Deloitte-HTML (Logo, grüne Linie, Footer) in `DEX_EmailTemplates.BodyHtml`
  gespeichert — der `DEX_SEND_MAIL`-Flow ersetzt nur `{{LOGO_URL}}` und
  `{{ORB_URL}}` und muss keinen Template-Wrapper selbst erzeugen.
- Der Body endet schlicht mit "If you no longer want to attend, please also
  cancel your registration." / "Falls du nicht mehr teilnehmen möchtest,
  melde dich bitte auch offiziell ab." — ohne Waitlist-Bezug, damit die Mail
  auch für Events ohne Warteliste korrekt wirkt.

## Neues Template in DEX_EmailTemplates

Beim nächsten App-Start legt die App folgende Template-Einträge automatisch an (falls noch nicht vorhanden):

| TemplateType | Language | Subject |
|--------------|----------|---------|
| OutlookDeclineReminder | EN | Action Required: Do you also want to cancel your registration? {{EventTitle}} |
| OutlookDeclineReminder | DE | Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}} |

Der Body enthält einen großen roten "Anmeldung stornieren" / "Cancel my registration"-Action-Button, der auf `{{CancelUrl}}` zeigt.

## Änderung in DEX_SEND_MAIL (bei der Template-Auswahl)

- Wenn `EmailType == 'OutlookDeclineReminder'`:
  - Template aus `DEX_EmailTemplates` mit `TemplateType eq 'OutlookDeclineReminder'` und passender Sprache holen
  - `{{CancelUrl}}` ersetzen mit Expression:
    ```
    concat(
      'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView&action=cancel&event=',
      first(body('Get_Event_from_DEX_Events')?['value'])?['EventNumber']
    )
    ```
    (EventNumber, nicht ID!)

## Finaler Flow-JSON v2 (Stand 2026-05-05, mit Decline-Digest via Outlook-Connector)

**Wichtig:** Die `Recipient`-Spalte in `DEX_Emails` ist ein **Note-Feld** (Multi-line text, weil es auch `;`-separierte Mehrfach-Empfänger enthalten kann). SP-OData `$filter` erlaubt keinen `eq` auf Note-Feldern. Deshalb filtert `Get_Existing_Reminder` nur nach `EmailType + EventId`, und ein nachgeschalteter `Filter array`-Schritt (`Filter_By_Recipient`) pickt den aktuellen Sender heraus — mit Semikolon-Wrapping um Teil-Matches (z.B. `alice@x.de` in `alicebackup@x.de`) zu vermeiden.

**Neu in v2 (2026-05-05):** Direkt nach `Create_Reminder_Queue_Item` ein paralleler Digest-Pfad — `Get_Outlook_EventId` (liest **`CalendarLink`** aus `DEX_Events`, das ist die iCalUId) → `Has_Outlook_Event?` → `Get_Outlook_Event` (Office-365-Outlook-Connector `V3CalendarGetItem`) → `Filter_Declined_Attendees` (`status.response == 'declined'`) → pro Decliner einen SP-Lookup auf der Teilnehmerliste (`Status='Angemeldet'`-Filter) → HTML-Tabelle bauen → `OutlookDeclineDigest`-Item in `DEX_Emails` queuen mit Recipient = Organizer-Mail.

### Trigger

```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "mailboxAddress": "no_reply.events@deloitte.de",
      "importance": "Any",
      "hasAttachments": false,
      "includeAttachments": false,
      "folderId": "Inbox"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "SharedMailboxOnNewEmailV2"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "splitOn": "@triggerOutputs()?['body/value']"
}
```

### Top-Level Variables (vor dem Is_Decline_Mail-If)

```json
{
  "Init_DeclineRows_HTML": {
    "type": "InitializeVariable",
    "inputs": { "variables": [ { "name": "DeclineRowsHtml", "type": "string" } ] },
    "runAfter": {}
  },
  "Init_StillRegistered_Count": {
    "type": "InitializeVariable",
    "inputs": { "variables": [ { "name": "StillRegisteredCount", "type": "integer" } ] },
    "runAfter": { "Init_DeclineRows_HTML": [ "Succeeded" ] }
  }
}
```

### Is_Decline_Mail (Top-Level If, runAfter Init_StillRegistered_Count)

Die Condition ist unverändert gegenüber v1:

```json
{
  "type": "If",
  "expression": {
    "and": [
      {
        "equals": [
          "@or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'declined '), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'abgelehnt '), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'refusé'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'rifiutato:'), and(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg:')), or(contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'declined:'), contains(toLower(coalesce(triggerOutputs()?['body/bodyPreview'], '')), 'abgelehnt:'))))",
          "@true"
        ]
      }
    ]
  },
  "runAfter": { "Init_StillRegistered_Count": [ "Succeeded" ] }
}
```

### Sub-Actions (Reihenfolge im If-yes-Branch)

Die folgenden Actions sind **unverändert gegenüber v1** — Inputs siehe v1-Beschreibung weiter oben in diesem Dokument oder im Archiv-Block:

- `Cleaned_Subject` (Compose)
- `Real_Sender` (Compose)
- `Decliner_Lastname` (Compose)
- `Decliner_Firstname` (Compose)
- `Get_DEX_Event` (SP Get items, `28457815-1163-4e92-8b08-3ae43f477d9e`)
- `Event_Found` (If)
  - `Get_Teilnehmer_Entry` (SP HttpRequest auf Subsite/Teilnehmer)
  - `Final_Recipient_Email` (Compose)
  - `Still_Registered` (If)
    - `Get_Existing_Reminder` (SP Get items, EmailType + EventId)
    - `Filter_By_Recipient` (Query/Filter array)
    - `No_Reminder_Yet` (If)
      - `Get_Reminder_Template` (SP Get items)
      - `Assistant_Forward_Mailto` (Compose)
      - `Create_Reminder_Queue_Item` (SP Create item)

### Neu in v2: Digest-Branch nach `Create_Reminder_Queue_Item`

Alle folgenden Actions liegen im selben `No_Reminder_Yet`-If-yes-Branch wie `Create_Reminder_Queue_Item`, runAfter-Kette:
`Create_Reminder_Queue_Item → Get_Outlook_EventId → Has_Outlook_Event`.

```json
{
  "Get_Outlook_EventId": {
    "type": "Compose",
    "inputs": "@coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['CalendarLink'], '')",
    "runAfter": { "Create_Reminder_Queue_Item": [ "Succeeded" ] }
  },
  "Has_Outlook_Event": {
    "type": "If",
    "expression": { "and": [ { "greater": [ "@length(outputs('Get_Outlook_EventId'))", 0 ] } ] },
    "runAfter": { "Get_Outlook_EventId": [ "Succeeded" ] },
    "actions": {
      "Get_Outlook_Event": {
        "type": "OpenApiConnection",
        "inputs": {
          "parameters": {
            "table": "AAMkADU5YjlkMDBiLWU2MDktNGViMy1iNGIwLTI0YWFkNDkyN2VjMABGAAAAAABjJcNB5xJWS7D2nCeePixeBwAbtMj6YVUGQJroN6O--ImBAAAAAAEGAAAbtMj6YVUGQJroN6O--ImBAAKF4fCpAAA=",
            "id": "@outputs('Get_Outlook_EventId')"
          },
          "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
            "connection": "shared_office365",
            "operationId": "V3CalendarGetItem"
          },
          "retryPolicy": {
            "type": "exponential",
            "count": 4,
            "interval": "PT5S"
          }
        }
      },
      "Set_variable": {
        "type": "SetVariable",
        "inputs": { "name": "StillRegisteredCount", "value": 0 },
        "runAfter": { "Get_Outlook_Event": [ "Succeeded" ] }
      },
      "Filter_Declined_Attendees": {
        "type": "Query",
        "inputs": {
          "from": "@coalesce(first(body('Get_Outlook_Event')?['value'])?['attendees'], json('[]'))",
          "where": "@equals(toLower(coalesce(item()?['status']?['response'], '')), 'declined')"
        },
        "runAfter": { "Set_variable": [ "Succeeded" ] }
      },
      "Apply_To_Each_Decliner": {
        "type": "Foreach",
        "foreach": "@body('Filter_Declined_Attendees')",
        "actions": {
          "Get_Decliner_Status": {
            "type": "OpenApiConnection",
            "inputs": {
              "parameters": {
                "dataset": "@concat('https://deudeloitte.sharepoint.com', replace(first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl'], 'https://deudeloitte.sharepoint.com', ''))",
                "table": "Teilnehmer",
                "$filter": "@concat('ParticipantEmail eq ''', items('Apply_To_Each_Decliner')?['emailAddress']?['address'], ''' and Status ne ''Abgemeldet''')",
                "$top": 1
              },
              "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                "connection": "shared_sharepointonline",
                "operationId": "GetItems"
              }
            }
          },
          "Condition_Still_Registered": {
            "type": "If",
            "expression": { "and": [ { "greater": [ "@length(body('Get_Decliner_Status')?['value'])", 0 ] } ] },
            "runAfter": { "Get_Decliner_Status": [ "Succeeded" ] },
            "actions": {
              "Increment_Count": {
                "type": "IncrementVariable",
                "inputs": { "name": "StillRegisteredCount", "value": 1 }
              },
              "Append_Decliner_Row": {
                "type": "AppendToStringVariable",
                "inputs": {
                  "name": "DeclineRowsHtml",
                  "value": "@concat('<tr><td style=\"padding:6px 10px;border-bottom:1px solid #eee;\">', coalesce(first(body('Get_Decliner_Status')?['value'])?['Vorname'], ''), ' ', coalesce(first(body('Get_Decliner_Status')?['value'])?['Nachname'], ''), '</td><td style=\"padding:6px 10px;border-bottom:1px solid #eee;\">', items('Apply_To_Each_Decliner')?['emailAddress']?['address'], '</td><td style=\"padding:6px 10px;border-bottom:1px solid #eee;\">', coalesce(first(body('Get_Decliner_Status')?['value'])?['Department'], '—'), '</td><td style=\"padding:6px 10px;border-bottom:1px solid #eee;\">', if(empty(first(body('Get_Decliner_Status')?['value'])?['RegistrationDate']), '—', formatDateTime(first(body('Get_Decliner_Status')?['value'])?['RegistrationDate'], 'dd.MM.yyyy')), '</td></tr>')"
                },
                "runAfter": { "Increment_Count": [ "Succeeded" ] }
              }
            },
            "else": { "actions": {} }
          }
        },
        "runAfter": { "Filter_Declined_Attendees": [ "Succeeded" ] },
        "runtimeConfiguration": {
          "concurrency": { "repetitions": 1 }
        }
      },
      "Compose_DeclineList_Table": {
        "type": "Compose",
        "inputs": "@concat('<table style=\"border-collapse:collapse;width:100%;font-size:13px;margin:12px 0;\"><thead><tr style=\"background:#f5f5f5;\"><th style=\"padding:8px 10px;text-align:left;\">Name</th><th style=\"padding:8px 10px;text-align:left;\">E-Mail</th><th style=\"padding:8px 10px;text-align:left;\">Department</th><th style=\"padding:8px 10px;text-align:left;\">Anmeldedatum</th></tr></thead><tbody>', variables('DeclineRowsHtml'), '</tbody></table>')",
        "runAfter": { "Apply_To_Each_Decliner": [ "Succeeded" ] }
      },
      "Has_Decliners": {
        "type": "If",
        "expression": { "and": [ { "greater": [ "@variables('StillRegisteredCount')", 0 ] } ] },
        "runAfter": { "Compose_DeclineList_Table": [ "Succeeded" ] },
        "actions": {
          "Get_Digest_Template": {
            "type": "OpenApiConnection",
            "inputs": {
              "parameters": {
                "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                "table": "2c428d35-e6fb-42f9-8a20-580acd6d05f4",
                "$filter": "@concat('TemplateType eq ''OutlookDeclineDigest'' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')",
                "$top": 1
              },
              "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                "connection": "shared_sharepointonline",
                "operationId": "GetItems"
              }
            }
          },
          "Create_Digest_Queue_Item": {
            "type": "OpenApiConnection",
            "inputs": {
              "parameters": {
                "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                "item/Title": "@replace(replace(coalesce(first(outputs('Get_Digest_Template')?['body/value'])?['Subject'], 'FYI: attendees declined Outlook'), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title']), '{{DeclineCount}}', string(variables('StillRegisteredCount')))",
                "item/Recipient": "@first(outputs('Get_DEX_Event')?['body/value'])?['OrganizerEmail']",
                "item/RecipientName": "Organizer",
                "item/EmailType/Value": "OutlookDeclineDigest",
                "item/EventTitle": "@first(outputs('Get_DEX_Event')?['body/value'])?['Title']",
                "item/Status/Value": "Pending",
                "item/Body": "@replace(replace(replace(coalesce(first(outputs('Get_Digest_Template')?['body/value'])?['BodyHtml'], ''), '{{EventTitle}}', first(outputs('Get_DEX_Event')?['body/value'])?['Title']), '{{DeclineCount}}', string(variables('StillRegisteredCount'))), '{{DeclineList}}', outputs('Compose_DeclineList_Table'))",
                "item/EventId": "@first(outputs('Get_DEX_Event')?['body/value'])?['ID']"
              },
              "host": {
                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                "connection": "shared_sharepointonline",
                "operationId": "PostItem"
              }
            },
            "runAfter": { "Get_Digest_Template": [ "Succeeded" ] }
          }
        },
        "else": { "actions": {} }
      }
    },
    "else": { "actions": {} }
  }
}
```

### Hinweise zur Calendar-Lookup-Logik

- **`Get_Outlook_Event` table = Default-Calendar-ID des Postfachs `no_reply.events@deloitte.de`.** Aktuell hardcoded als die lange Base64-ID. Bei Tenant-Migration muss der Wert über die Office-365-Outlook-Connector-Dropdowns (Calendar Id) neu picked werden — die Connector-UI tauscht ihn dann automatisch im Code-View.
- **`Get_Outlook_Event` id = `outputs('Get_Outlook_EventId')`** = `iCalUId` aus `DEX_Events.CalendarLink`. `V3CalendarGetItem` akzeptiert die iCalUId als Lookup-Schlüssel und sucht damit über alle Mailbox-Kalender hinweg — anders als die kalenderspezifische `EntryId`, die mailbox-gebunden wäre.
- **Get_Decliner_Status table = `Teilnehmer`** (nicht leer lassen — leerer table-Wert führte in einer Vorab-Version zu HTTP 400 auf der Subsite-Liste).
- **Get_Decliner_Status Status-Filter `ne 'Abgemeldet'`, NICHT `eq 'Angemeldet'`** (Bug-Fix 2026-05-05): die Teilnehmer-Liste hat drei aktive Status — `Angemeldet` (frisch registriert), `QR versendet` (QR-Code gesendet) und `Eingecheckt` (am Eventtag eingecheckt). Alle drei zählen als „noch in der Liste" und müssen vom Decline-Digest erfasst werden. Der ursprüngliche Filter `eq 'Angemeldet'` schloss QR-versendete und eingecheckte Teilnehmer fälschlich aus → Decliner mit bereits versandtem QR-Code landeten nicht im Digest und der Organizer kriegte trotz aktivem Decline keine FYI-Mail. Der negative Filter `ne 'Abgemeldet'` ist robuster — er schließt ausschließlich offiziell Abgemeldete aus, was genau die Decliner sind die KEIN Catering / Hotel mehr blockieren.
- **Outlook-Event-Lookup nutzt `Get events (V4)` (Plural) mit `$filter=iCalUId eq '<id>'`, NICHT `Get event (V4)` (Singular).** Die Singular-Action (`V3CalendarGetItem`) erwartet eine **EntryId** — wenn man ihr eine **iCalUId** (aus `DEX_Events.CalendarLink`) übergibt, antwortet der Connector mit `ErrorInvalidIdMalformed`. Mit Plural + Filter umgeht man das, weil iCalUId eine Standard-Property der Event-Resource in Microsoft Graph ist und filterbar.
- **`Filter_Declined_Attendees.from = json('[]')` als Fallback, NICHT `createArray()`.** PA-`createArray()` braucht mindestens 1 Parameter; ohne Args wirft es einen Compile-Error („createArray expects a comma separated list of parameters"). `json('[]')` parst den String als JSON und liefert ein leeres Array — das saubere PA-Idiom für „leeres Array, falls null".
- **Plural-Result-Access:** seit der Umstellung auf `Get events (V4)` ist `body('Get_Outlook_Event')?['value']` ein Array. Zugriff auf den einzelnen Treffer immer mit `first(body('Get_Outlook_Event')?['value'])?['attendees']` — NICHT mehr mit `outputs('Get_Outlook_Event')?['body/attendees']`.
- **`Apply_To_Each_Decliner.runtimeConfiguration.concurrency.repetitions = 1`** ist Pflicht. Power Automate führt `Foreach` per Default mit bis zu **20 parallelen Iterationen** aus. Innerhalb der Loop wird an `DeclineRowsHtml` (string) appended und `StillRegisteredCount` (int) incrementiert — beides shared Variables. Bei paralleler Ausführung gehen Schreibvorgänge verloren (lost update) und die Digest-Tabelle hat verstümmelte Zeilen / falschen Counter. Sequentiell (`repetitions: 1`) löst das — Performance-Verlust ist minimal, weil pro Decliner nur ein SP-Read und zwei Variable-Writes laufen.
- **`Get_Outlook_Event.retryPolicy = exponential, count: 4, interval: PT5S`** fängt transiente Outlook-Connector-Fehler (`429 Too Many Requests`, `503`, kurzfristige Timeouts) ab. Ohne Retry würde der ganze Flow-Run kippen und der Digest für diese eine Decline-Mail wäre verloren — der Reminder ist ja bereits gequeued, also stünde der Decliner danach im nächsten Digest mit einer Run-Zeit Verspätung. Mit Retry ist beides resilient.

### Test-Plan Decline-Digest (Approach B)

| Szenario | Erwartung |
|----------|-----------|
| User lehnt mit „Send response" ab | Reminder + Digest werden gequeued. Digest enthält den User in der Tabelle. |
| User lehnt mit „Don't send response" ab | **Reminder feuert nicht** (kein Mail-Trigger), aber **beim nächsten Decline** im selben Event ist er im Digest enthalten — weil der Digest direkt aus `attendees[].status.response` liest. |
| User lehnt ab und akzeptiert dann doch | Beim nächsten Decline im selben Event ist er **nicht mehr** im Digest — `status.response` zeigt jetzt `accepted`. |
| User lehnt ab und meldet sich offiziell ab | Beim nächsten Decline im selben Event ist er **nicht mehr** im Digest — `Get_Decliner_Status` filtert auf `Status ne 'Abgemeldet'`, abgemeldete fallen damit raus. |
| Event hat keinen Outlook-Termin (`CalendarLink` leer) | `Has_Outlook_Event` ist false → Digest-Branch wird übersprungen, nur Reminder wird gequeued. |
| Event hat 0 noch-angemeldete Decliner | `Has_Decliners` ist false → kein Digest gequeued. |

## Archiv: DEX_OutlookDeclineHandler v1 (event-modified, deprecated 2026-04-14)

Die erste Version nutzte den `When an event is modified (V3)`-Trigger und filterte `body/attendees[*].status.response == 'declined'`. In der Praxis gibt der Shared-Mailbox-Trigger dieses pro-Attendee-Feld jedoch **nicht** zuverlässig zurück (nur `requiredAttendees`/`optionalAttendees` als Semicolon-String). Der Flow hat daher nie eine Reminder-Mail erzeugt. Die Mail-basierte Variante oben ersetzt diesen Ansatz.

Der v1-Flow ist deaktiviert aber nicht gelöscht (`DEX_OutlookDeclineHandler_v1_old`), kann nach 1 Woche stabiler Mail-Variante endgültig entfernt werden.

---

## 6. DEX_OutlookForwardHandler

**Trigger:** Neue Mail in `no_reply.events@deloitte.de` mit Subject startend mit `Meeting Forward Notification:` (EN) oder `Terminweiterleitungsbenachrichtigung:` (DE)
**Zweck:** Wenn ein Teilnehmer einen Outlook-Termin an Dritte weiterleitet, bekommt `no_reply.events@deloitte.de` eine Info-Mail. Der Flow prüft, ob die weitergeleitete Person bereits in der SharePoint-Teilnehmerliste eingetragen ist. Wenn **nein**, geht eine FYI-Mail an den Organizer raus (Template `OutlookForwardNotification` aus `DEX_EmailTemplates`, vom Flow gerendert, Body landet in `DEX_Emails`-Queue → DEX_SEND_MAIL versendet).
**Letztes Update:** 2026-04-15 (FW:/WG:-Varianten abgedeckt, Cleaned_Subject auf `last(split)` umgestellt)
**Listen-GUIDs:** DEX_Events `28457815-1163-4e92-8b08-3ae43f477d9e`, DEX_EmailTemplates `2c428d35-e6fb-42f9-8a20-580acd6d05f4`, DEX_Emails `57aa0840-df98-41ae-a39b-323c0b80ae3b`

### Hintergrund

Beispiel-Mail, die das auslöst:

```
From: Microsoft Outlook on behalf of von Rueden, Dr. Michael
Subject: Meeting Forward Notification: E2E M&A Activation Session
Body:
  Your meeting was forwarded
  von Rueden, Dr. Michael has forwarded your meeting request to additional recipients.

  Meeting:       E2E M&A Activation Session
  Meeting Time:  Thursday, 23 April 2026, 19:00 to Friday, 24 April 2026, 15:30.
  Recipients:    Mauß, Anna Kristina
```

Problem: Anna Kristina Mauß ist im Outlook-Termin drin — aber **nicht** in der SharePoint-Teilnehmerliste der Event-Subsite. Sie hat keine TeilnehmerID, keinen QR-Code, und der Organizer sieht sie nicht in der App. Dazu kommt: sie könnte ausserhalb der Audience-Filter (Location/JobTitle) liegen und eigentlich gar nicht teilnehmen dürfen.

### Ablauf

1. Trigger (neue Mail im Inbox von `no_reply.events@deloitte.de`)
2. **Is_ForwardNotification?** → Subject startet mit `Meeting Forward Notification:` ODER `Terminweiterleitungsbenachrichtigung:`
3. **Cleaned_Subject** → Event-Titel aus Subject extrahieren (alles nach dem `:`)
4. **Get_DEX_Event** → SharePoint Get items `DEX_Events` mit `Title eq '<Cleaned_Subject>'`
5. **Event_Found?** → weiter nur wenn Event existiert
6. **Parse_Recipients** → Recipient-Namen aus `body/body` extrahieren (alle Namen nach `Recipients` / `Empfänger` bis `All times listed`)
7. **Resolve_Recipient_Email** → Graph API User-Search per DisplayName (`Nachname, Vorname`) → Email
8. **Get_Teilnehmer_Entry** → SharePoint HTTP request auf Subsite-Teilnehmerliste: gibt es einen Eintrag mit `ParticipantEmail eq '<resolved_email>'` der nicht `Abgemeldet` ist?
9. **Already_Registered?** → wenn **ja**: Flow endet (Log-Eintrag "OK, schon eingeladen"). Wenn **nein**: weiter zu 10.
10. **Get_ForwardTemplate** → Template `OutlookForwardNotification` (DE/EN je nach Event) aus `DEX_EmailTemplates` laden. Das Template wird **von der App automatisch angelegt** (siehe Abschnitt "Template-Seeding durch die App") — inklusive Deloitte-Outlook-Wrapper.
11. **Rendered_Subject / Rendered_Body** → Platzhalter `{{EventTitle}}`, `{{Forwarder}}`, `{{Recipient}}`, `{{RecipientEmail}}`, `{{OrganizerFirstName}}`, `{{AppUrl}}` per `replace()` ersetzen.
12. **Create_FYI_Email** → DEX_Emails-Queue-Item anlegen mit den echten Spalten:
    - `Title` = Rendered Subject
    - `Recipient` = Organizer-Email (Plain-Text)
    - `Body` = Rendered HTML-Body (fertig gerendert, kein weiteres Template-Lookup im DEX_SEND_MAIL nötig)
    - `EmailType` = `Info`, `Status` = `Pending`, `EventId` = Event-ID, `EventTitle` = Cleaned_Subject

### SharePoint-Liste DEX_Outlook (unverändert)

Dieser Flow nutzt **nicht** die DEX_Outlook-Queue (da keine Outlook-Änderung getriggert wird — der Outlook-Termin hat den neuen Empfänger ja schon). Er nutzt die bestehende `DEX_Emails`-Queue für die FYI-Mail.

### UI-Anleitung zum Anlegen (Schritt für Schritt)

**1. Neuer Cloud-Flow anlegen**
1. https://make.powerautomate.com öffnen
2. Links **+ Create** → **Automated cloud flow**
3. **Flow name:** `DEX_OutlookForwardHandler`
4. Trigger: `When a new email arrives V3` → **Office 365 Outlook — When a new email arrives (V3)** auswählen
5. **Create**

**2. Trigger konfigurieren**
- **Folder:** `Inbox`
- **Show advanced options:**
  - **Original Mailbox Address:** `no_reply.events@deloitte.de`
  - **Include Attachments:** `No`
  - **Subject Filter:** leer lassen (Language-Filter via Condition in Schritt 3)
- Settings → Concurrency: **1** (sequentiell, maximumWaitingRuns 100)

**3. Condition `Is_ForwardNotification`**
- Linke Seite: **Expression (fx)** →
  ```
  or(
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'meeting forward notification:'),
    startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'terminweiterleitungsbenachrichtigung:')
  )
  ```
- Operator: `is equal to`
- Rechte Seite: **Expression (fx)** → `true`
- Wenn **no** → Terminate (Succeeded). Wenn **yes** → weiter.

**4. `Cleaned_Subject` (Compose)**
- Expression:
  ```
  trim(
    last(
      split(triggerOutputs()?['body/subject'], ':')
    )
  )
  ```
- Achtung: `last(split(...,':'))` kann bei Event-Titeln mit Doppelpunkt falsch sein. Alternativ robuster:
  ```
  trim(
    substring(
      triggerOutputs()?['body/subject'],
      add(indexOf(triggerOutputs()?['body/subject'], ':'), 1)
    )
  )
  ```
- Rename → `Cleaned_Subject`

**5. `Get_DEX_Event` (SharePoint Get items)**
- Site Address: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- List Name: `DEX_Events`
- Filter Query: **Expression (fx)** →
  ```
  concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')
  ```
- Top Count: `1`

**6. Condition `Event_Found`**
- Expression: `length(outputs('Get_DEX_Event')?['body/value'])` `is greater than` `0`
- Wenn **no** → Terminate (Succeeded, Message `Event not found in DEX_Events`)
- Wenn **yes** → weiter

**7. `Parse_Recipients` (Compose)**

Die Mail enthält die Recipient-Namen im HTML-Body. Der `bodyPreview` (Plaintext) hat das Format:
```
... Recipients   Mauß, Anna Kristina   All times listed ...
```
Expression zum Extrahieren:
```
trim(
  first(
    split(
      last(
        split(
          coalesce(triggerOutputs()?['body/bodyPreview'], ''),
          'Recipients'
        )
      ),
      'All times listed'
    )
  )
)
```
Für DE-Mails zusätzlich auch `Empfänger` als Split-Token. Alternative (robuster):
```
if(
  contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'Recipients'),
  trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Recipients')), 'All times listed'))),
  trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Empfänger')), 'Alle aufgeführten Zeiten')))
)
```
Rename → `Recipient_DisplayName`.

**8. `Resolve_Recipient_Email` (Office 365 Users — Search for users (V2))**
- Search Term: `@outputs('Recipient_DisplayName')`
- Top: `5`
- Nach der Action: `Filter_Matching_User` (Data Operation — Filter array) über `body('Resolve_Recipient_Email')?['value']` mit `item()?['displayName'] is equal to @outputs('Recipient_DisplayName')` (exakte Match).
- Compose `Recipient_Email`: `first(body('Filter_Matching_User'))?['mail']`

**9. Condition `Email_Resolved`**
- Expression: `outputs('Recipient_Email')` `is not equal to` `null`
- Wenn **no** → FYI-Mail an Organizer mit Hinweis "Recipient-Email konnte nicht aufgelöst werden: <DisplayName>" (siehe Schritt 12, aber mit anderer Nachricht)
- Wenn **yes** → weiter

**10. `Get_Teilnehmer_Entry` (Send HTTP request to SharePoint)**

Die Teilnehmer-Liste liegt auf der Event-Subsite. `SubsiteUrl` aus `Get_DEX_Event` holen:
- Site Address: `@{first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl']}`
- Method: `GET`
- Uri:
  ```
  _api/web/lists/getbytitle('Teilnehmer')/items?$filter=ParticipantEmail eq '@{outputs('Recipient_Email')}' and Status ne 'Abgemeldet'&$top=1
  ```
- Headers: `Accept: application/json;odata=nometadata`

**11. Condition `Already_Registered`**
- Expression: `length(body('Get_Teilnehmer_Entry')?['value'])` `is greater than` `0`
- Wenn **yes** → Terminate (Succeeded, Message `Recipient already registered`)
- Wenn **no** → weiter zu 12

**12. `Get_ForwardTemplate` (SharePoint Get items auf DEX_EmailTemplates)**

Das Template `OutlookForwardNotification` wird **von der App automatisch angelegt** (siehe unten, Abschnitt "Template-Seeding durch die App"). Der Flow holt es per Filter. Sprache wird vom Event übernommen (`EmailLanguage`-Feld), mit Fallback auf `EN`.

- Site Address: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- List Name: `DEX_EmailTemplates`
- Filter Query: **Expression (fx)** →
  ```
  concat('TemplateType eq ''OutlookForwardNotification'' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')
  ```
- Top Count: `1`

**12a. `Rendered_Subject` (Compose)**

Platzhalter im Subject ersetzen. Expression (fx):
```
replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['Subject'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' ')))
```

**12b. `Rendered_Body` (Compose)**

Die gleiche Replace-Kaskade auf dem `BodyHtml`-Feld, zusätzlich noch `{{AppUrl}}`. Expression (fx):
```
replace(replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['BodyHtml'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' '))), '{{AppUrl}}', 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView')
```

**13. `Create_FYI_Email` (SharePoint Create item, Liste DEX_Emails)**

Schreibt den fertigen Mail-Body in die Queue. `DEX_SEND_MAIL` verschickt ihn, ohne selbst Templates laden zu müssen — der komplette, schon gerenderte HTML-Body steht direkt im `Body`-Feld.

Echte Spalten von `DEX_Emails` (siehe `EventService.ts:196`):

| Feld | Typ | Wert für diesen Flow |
|------|-----|----------------------|
| `Title` | Text | = Subject (`outputs('Rendered_Subject')`) |
| `Recipient` | Note, Plain-Text | Organizer-Email (erste aus `;`-Liste) |
| `Cc` | Note, Plain-Text | leer |
| `RecipientName` | Text | = `first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';'))` |
| `Body` | Note (HTML) | = `outputs('Rendered_Body')` |
| `EmailType` | Choice (Anmeldung, Abmeldung, Warteliste, Nachruecken, **Info**) | `Info` |
| `EventTitle` | Text | = `outputs('Cleaned_Subject')` |
| `EventId` | Text | = `string(first(outputs('Get_DEX_Event')?['body/value'])?['Id'])` |
| `Status` | Choice (**Pending**, Sent, Failed) | `Pending` |
| `SentDate` | Date | leer (Flow setzt es beim Versand) |

Der `DEX_SEND_MAIL`-Flow pickt das Item auf (`Status eq 'Pending'`) und verschickt es per Office-365-Mail-Action — **keine Anpassung an DEX_SEND_MAIL nötig**, weil der Body schon fertig gerendert ist.

### Template-Seeding durch die App

Das Template `OutlookForwardNotification` (DE + EN) wird **automatisch beim App-Start** in die Liste `DEX_EmailTemplates` geschrieben. Zuständig sind drei Funktionen in `EventService.ts`:

- `ensureEmailTemplatesList()` — erstmaliges Anlegen der Liste mit allen Default-Templates
- `ensureMissingEmailTemplates()` — Nachlegen auf Tenants, wo die Liste schon existiert, aber `OutlookForwardNotification` noch fehlt
- `upgradeStandardEmailTemplates()` — Updated den `BodyHtml` wenn das Template in der App-Version verändert wurde (überschreibt User-Customizing)

Der HTML-Body wird über `wrapTemplateForStorage()` aus `services/EmailTemplates.ts` gerendert — damit sieht die FYI-Mail exakt wie die anderen DEX-Mails aus (schwarzer Deloitte-Header mit Logo, grüne Trennlinie, DEX-Orb, Content-Block, Footer mit Legal-Text).

**Platzhalter im Body (werden vom Flow per `replace()` ersetzt, siehe 12a/12b):**
- `{{OrganizerFirstName}}` — Vorname des Organizers
- `{{Forwarder}}` — Name/Email der Person, die den Termin weitergeleitet hat (aus `body/from`)
- `{{Recipient}}` — Name der hinzugefügten Person
- `{{RecipientEmail}}` — Resolved Email oder `nicht aufgelöst`
- `{{EventTitle}}` — Event-Titel
- `{{AppUrl}}` — DEX-App-URL (hardcoded im Flow)

**Code-Konstanten:** `OUTLOOK_FORWARD_BODY_DE` / `OUTLOOK_FORWARD_BODY_EN` in `EventService.ts`.

### Änderung in DEX_SEND_MAIL

**Keine Änderung nötig.** Der Flow verschickt einfach das fertige `Body`-Feld aus DEX_Emails per Office-365-Mail-Action — das neue Template wird für DEX_SEND_MAIL transparent behandelt (wie alle `Info`-Mails).

### Bekannte Einschränkungen

| Fall | Verhalten |
|------|-----------|
| Forwarder hat Outlook auf DE → `Terminweiterleitungsbenachrichtigung:` | ✅ (Condition `Is_ForwardNotification` deckt beide Sprachen ab) |
| Forwarder hat Outlook auf EN → `Meeting Forward Notification:` | ✅ |
| Notification wurde nochmal weitergeleitet (`FW:` / `WG:` davor) | ✅ Seit 2026-04-15: Condition kennt alle 6 Varianten (DE/EN × direkt/FW:/WG:). `body/from` ist dann allerdings der Weiterleiter, nicht der originale Forwarder — `{{Forwarder}}` in der FYI-Mail kann deshalb ungenau sein. |
| Forwarder hat andere Sprache (FR/IT/…) | ❌ Subject wird nicht erkannt — erweitern bei Bedarf |
| Recipient ist ein externer User (kein Azure AD Account) | ❌ `Resolve_Recipient_Email` findet keinen Treffer → `Email_Resolved`-else-Zweig ist leer, Flow terminiert ohne Mail |
| Recipient-Name steht in uneindeutiger Form (nur Vorname, Firmenkürzel) | ⚠️ Graph-Search kann 0 oder mehrere Treffer liefern → `Filter_Matching_User` per exakter `displayName`-Gleichheit |
| Mehrere Recipients in einer Forward-Notification | ⚠️ Flow behandelt nur den ersten — für Multi-Recipient Schleife über gesplittete Namen nötig |
| Event-Titel enthält Doppelpunkt | ⚠️ `Cleaned_Subject` per `last(split(..., ':'))` nimmt nur das letzte Segment → bei Events wie `DEX: Sommer-Event` landet nur `Sommer-Event` im Cleaned_Subject und `Get_DEX_Event` findet nichts |

### Teststrategie

1. Manuell in Outlook einen Testtermin aus `no_reply.events@deloitte.de` erstellen, einladen.
2. Als eingeladener User: **Forward** auf eine andere Person (Deloitte-interne) → Flow sollte die FYI-Mail an den Organizer queuen.
3. Als eingeladener User: **Forward** auf eine Person, die bereits registriert ist → Flow sollte **keine** FYI schicken.
4. Als eingeladener User: **Forward** auf einen Externen (keine Azure AD-Identität) → Email_Resolved-Condition sollte greifen, Flow terminiert sauber (aktuell keine FYI bei nicht-auflösbarer Mail — Verbesserungspotenzial).
5. Flow-Runs kontrollieren: jeder Schritt sollte `Succeeded` sein, Terminate-Zweige dokumentieren warum keine Mail verschickt wurde.

### Finaler Flow-JSON (Stand 2026-04-15, HTML-Parser für Name + Email, ohne Graph-Search)

**Key Changes vs. initiale Version:**
- `Recipient_DisplayName` extrahiert jetzt den Namen **direkt aus dem HTML-Body** (via `mailto:...">Name</a>`-Muster), nicht mehr über Graph-API-Search
- `Recipient_Email` ebenfalls direkt aus dem HTML — keine Office-365-User-Search mehr nötig
- Dadurch: `Resolve_Recipient_Email` (Office 365 Users) und `Filter_Matching_User` (Filter array) **entfernt** — Flow ist schneller und robuster bei Externen
- `Cleaned_Subject` auf `last(split(subject, ':'))` umgestellt (robust gegen FW:/WG:-Prefixes)
- `Get_DEX_Event` auf `substringof(subject, Title)` statt `eq` (matcht Events mit längerem Titel wie `E2E M&A Activation Session Munich`)
- `Email_Resolved` nutzt `empty()` statt `null`-Vergleich
- `Rendered_Subject`/`Rendered_Body`: 3 Platzhalter-Fixes (OrganizerFirstName aus `Nachname, Vorname`, Forwarder aus `bodyPreview`, Recipient-Fallback auf Email)



TRIGGER:
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": { "mailboxAddress": "no_reply.events@deloitte.de" },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connection": "shared_office365",
      "operationId": "SharedMailboxOnNewEmailV2"
    }
  },
  "recurrence": { "frequency": "Minute", "interval": 1 },
  "runtimeConfiguration": { "concurrency": { "runs": 1, "maximumWaitingRuns": 100 } },
  "splitOn": "@triggerOutputs()?['body/value']"
}
```

IS_FORWARDNOTIFICATION (If):
```json
{
  "type": "If",
  "expression": {
    "and": [
      {
        "equals": [
          "@or(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'meeting forward notification:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'terminweiterleitungsbenachrichtigung:')), or(or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw: meeting forward notification:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg: meeting forward notification:')), or(startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'fw: terminweiterleitungsbenachrichtigung:'), startsWith(toLower(coalesce(triggerOutputs()?['body/subject'], '')), 'wg: terminweiterleitungsbenachrichtigung:'))))",
          true
        ]
      }
    ]
  },
  "actions": {
    "Cleaned_Subject": {
      "type": "Compose",
      "inputs": "@trim(last(split(coalesce(triggerOutputs()?['body/subject'], ''), ':')))"
    },
    "Get_DEX_Event": {
      "type": "OpenApiConnection",
      "inputs": {
        "parameters": {
          "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
          "table": "28457815-1163-4e92-8b08-3ae43f477d9e",
          "$filter": "@concat('Title eq ''', replace(outputs('Cleaned_Subject'), '''', ''''''), '''')",
          "$top": 1
        },
        "host": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
          "connection": "shared_sharepointonline",
          "operationId": "GetItems"
        }
      },
      "runAfter": { "Cleaned_Subject": ["Succeeded"] }
    },
    "Event_Found": {
      "type": "If",
      "expression": {
        "and": [ { "greater": ["@length(outputs('Get_DEX_Event')?['body/value'])", 0] } ]
      },
      "actions": {
        "Recipient_DisplayName": {
          "type": "Compose",
          "inputs": "@if(contains(coalesce(triggerOutputs()?['body/bodyPreview'], ''), 'Recipients'), trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Recipients')), 'All times listed'))), trim(first(split(last(split(triggerOutputs()?['body/bodyPreview'], 'Empfänger')), 'Alle aufgeführten Zeiten'))))"
        },
        "Resolve_Recipient_Email": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "top": 5,
              "isSearchTermRequired": true,
              "searchTerm": "@outputs('Recipient_DisplayName')"
            },
            "host": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365users",
              "connection": "shared_office365users",
              "operationId": "SearchUserV2"
            }
          },
          "runAfter": { "Recipient_DisplayName": ["Succeeded"] }
        },
        "Filter_Matching_User": {
          "type": "Query",
          "inputs": {
            "from": "@body('Resolve_Recipient_Email')?['value']",
            "where": "@equals(item()?['displayName'],outputs('Recipient_DisplayName'))"
          },
          "runAfter": { "Resolve_Recipient_Email": ["Succeeded"] }
        },
        "Recipient_Email": {
          "type": "Compose",
          "inputs": "@first(body('Filter_Matching_User'))?['mail']",
          "runAfter": { "Filter_Matching_User": ["Succeeded"] }
        },
        "Email_Resolved": {
          "type": "If",
          "expression": {
            "and": [ { "equals": ["@empty(outputs('Recipient_Email'))", false] } ]
          },
          "actions": {
            "Get_Teilnehmer_Entry": {
              "type": "OpenApiConnection",
              "inputs": {
                "parameters": {
                  "dataset": "@first(outputs('Get_DEX_Event')?['body/value'])?['SubsiteUrl']",
                  "parameters/method": "GET",
                  "parameters/uri": "_api/web/lists/getbytitle('Teilnehmer')/items?$filter=ParticipantEmail eq '@{outputs('Recipient_Email')}' and Status ne 'Abgemeldet'&$top=1",
                  "parameters/headers": { "Accept": "application/json;odata=nometadata" }
                },
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                  "connection": "shared_sharepointonline",
                  "operationId": "HttpRequest"
                }
              }
            },
            "Already_Registered": {
              "type": "If",
              "expression": {
                "and": [ { "greater": ["@length(body('Get_Teilnehmer_Entry')?['value'])", 0] } ]
              },
              "actions": {
                "Terminate_2": { "type": "Terminate", "inputs": { "runStatus": "Succeeded" } }
              },
              "else": {
                "actions": {
                  "Get_ForwardTemplate": {
                    "type": "OpenApiConnection",
                    "inputs": {
                      "parameters": {
                        "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                        "table": "2c428d35-e6fb-42f9-8a20-580acd6d05f4",
                        "$filter": "@concat('TemplateType eq ''OutlookForwardNotification'' and Language eq ''', coalesce(first(outputs('Get_DEX_Event')?['body/value'])?['EmailLanguage'], 'EN'), '''')",
                        "$top": 1
                      },
                      "host": {
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                        "connection": "shared_sharepointonline",
                        "operationId": "GetItems"
                      }
                    }
                  },
                  "Rendered_Subject": {
                    "type": "Compose",
                    "inputs": "@replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['Subject'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' ')))",
                    "runAfter": { "Get_ForwardTemplate": ["Succeeded"] }
                  },
                  "Rendered_Body": {
                    "type": "Compose",
                    "inputs": "@replace(replace(replace(replace(replace(replace(first(outputs('Get_ForwardTemplate')?['body/value'])?['BodyHtml'], '{{EventTitle}}', outputs('Cleaned_Subject')), '{{Forwarder}}', triggerOutputs()?['body/from']), '{{Recipient}}', outputs('Recipient_DisplayName')), '{{RecipientEmail}}', coalesce(outputs('Recipient_Email'), 'nicht aufgelöst')), '{{OrganizerFirstName}}', first(split(first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';')), ' '))), '{{AppUrl}}', 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView')",
                    "runAfter": { "Rendered_Subject": ["Succeeded"] }
                  },
                  "Create_FYI_Email": {
                    "type": "OpenApiConnection",
                    "inputs": {
                      "parameters": {
                        "dataset": "https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform",
                        "table": "57aa0840-df98-41ae-a39b-323c0b80ae3b",
                        "item/Title": "@outputs('Rendered_Subject')",
                        "item/Recipient": "@first(split(first(outputs('Get_DEX_Event')?['body/value'])?['OrganizerEmail'], ';'))",
                        "item/RecipientName": "@first(split(first(outputs('Get_DEX_Event')?['body/value'])?['Organizer'], ';'))",
                        "item/Body": "@outputs('Rendered_Body')",
                        "item/EmailType/Value": "Info",
                        "item/EventTitle": "@outputs('Cleaned_Subject')",
                        "item/EventId": "@string(first(outputs('Get_DEX_Event')?['body/value'])?['Id'])",
                        "item/Status/Value": "Pending"
                      },
                      "host": {
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                        "connection": "shared_sharepointonline",
                        "operationId": "PostItem"
                      }
                    },
                    "runAfter": { "Rendered_Body": ["Succeeded"] }
                  }
                }
              },
              "runAfter": { "Get_Teilnehmer_Entry": ["Succeeded"] }
            }
          },
          "else": { "actions": {} },
          "runAfter": { "Recipient_Email": ["Succeeded"] }
        }
      },
      "else": {
        "actions": { "Terminate_1": { "type": "Terminate", "inputs": { "runStatus": "Succeeded" } } }
      },
      "runAfter": { "Get_DEX_Event": ["Succeeded"] }
    }
  },
  "else": {
    "actions": { "Terminate": { "type": "Terminate", "inputs": { "runStatus": "Succeeded" } } }
  },
  "runAfter": {}
}
```

### Offenes Verbesserungspotenzial

- **`Email_Resolved`-else-Zweig ist leer:** Wenn die weitergeleitete Person nicht per Graph-Search aufgelöst werden kann (z.B. Externe, Gastaccount mit anderem DisplayName-Format), verliert der Flow sie stumm. Besser wäre ein zweiter `Create_FYI_Email`-Block in diesem Zweig, der eine FYI-Mail mit `RecipientEmail = "nicht aufgelöst — bitte manuell prüfen: <DisplayName>"` an den Organizer schickt.
- **Mehrere Recipients:** Die `bodyPreview`-Parsing-Logik nimmt aktuell nur den ersten Namen. Wenn jemand den Termin an mehrere Personen gleichzeitig forwarded (z.B. "Mauß, Anna Kristina; Müller, Max"), sollte eine Apply-to-each-Schleife über `split(outputs('Recipient_DisplayName'), ';')` iterieren.
- **Subject-Parsing bei `:` im Event-Titel:** `Cleaned_Subject` per `substring(..., indexOf(':'))` schneidet nach dem **ersten** `:`. Bei Events wie `DEX: Sommer-Event` landet nur `Sommer-Event` im Cleaned_Subject und `Get_DEX_Event` findet nichts. Robustere Variante: per `add(indexOf('notification:'), 13)` bzw. `add(indexOf('benachrichtigung:'), 17)`.


---

## UI-Anleitung 2026-07-XX (v12.10) — Warteliste-Sortierung im DEX_IDReorder-Flow auf TeilnehmerID umstellen

Seit App-v12.10 sortiert die App-seitige Promote-Logik die Warteliste
nach `TeilnehmerID asc` statt `RegistrationDate asc`. Damit auch der
**User-Self-Cancel-Pfad** (= Power-Automate-Flow
`DEX_IDReorder_TeilnehmerIDs`) konsistent nachrückt, muss in jedem
Nachrück-Branch die Order-By-Klausel der Get-Items-Action umgestellt
werden.

**Hintergrund:** Nach jedem Lauf des Flows sind die TeilnehmerIDs
durchlaufend (1..N aktiv, N+1..M Warteliste). Wenn jetzt Platz 100
frei wird, soll TID 101 nachrücken — auch wenn TID 103 zeitlich
gesehen früher registriert war (z.B. nach Group-Switch oder
Reaktivierung). Mit `RegistrationDate`-Sortierung würde TID 103
fälschlich vorrücken.

**Betroffene Branches im Flow:**

1. **`Check_Nachruecken`** (Standard-Single-Capacity-Pfad)
   → `Get_Erster_Warteliste` (oder analog benannte Get-Items-Action)
2. **`Check_Nachruecken_Durchstarter`** (Split-Capacity, Durchstarter)
   → `Get_Erster_Warteliste_Durchstarter`
3. **`Check_Nachruecken_Funstarter`** (Split-Capacity, Funstarter)
   → `Get_Erster_Warteliste_Funstarter`
4. **`Check_Nachruecken_Shared`** (Shared-Waitlist-Modus, v10.20)
   → `Get_Erster_Warteliste_Shared`

**UI-Schritte pro Branch (Power-Automate-Editor):**

1. Im linken Aktionsbaum den betroffenen Branch öffnen
2. Auf die Get-Items-Action (z.B. `Get_Erster_Warteliste`) klicken
3. Im **Show advanced options**-Block den Wert von **Order by** finden
4. Aktueller Wert: `RegistrationDate asc`
5. Diesen Wert **leeren** und neu eintippen: `TeilnehmerID asc`
6. Filter-Query (Status eq 'Warteliste' etc.) **NICHT ändern**
7. Auf Speichern klicken (oben rechts)

**Wichtig:** `TeilnehmerID` ist ein Number-Feld in der Subsite-
`Teilnehmer`-Liste. SP-OData-Sortierung auf Number-Feldern funktioniert
direkt — keine Konvertierung nötig.

**Status 2026-07-XX (Tenant durchgeklickt):**
- ✅ `Get_Waitlist_First` (Standard / Single-Capacity) — `$orderby=TeilnehmerID asc` (siehe JSON-Snippet weiter oben in dieser Datei).
- ✅ `Get_Waitlist_First_Durchstarter` (Split-A) — siehe Snapshot unten.
- ✅ `Get_Waitlist_First_Funstarter` (Split-B) — siehe Snapshot unten.

### Snapshot Get_Waitlist_First_Durchstarter (v12.10)

```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "@outputs('Settings')?['siteAddress']",
      "parameters/method": "GET",
      "parameters/uri": "@concat('_api/web/lists/getbytitle(''', outputs('Settings')?['listName'], ''')/items?$select=Id,Vorname,Nachname,ParticipantName,ParticipantEmail,PreferredStarterType&$filter=Status eq ''Warteliste'' and PreferredStarterType eq ''Durchstarter''&$orderby=TeilnehmerID asc&$top=1')"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "HttpRequest"
    }
  }
}
```

### Snapshot Get_Waitlist_First_Funstarter (v12.10)

```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "@outputs('Settings')?['siteAddress']",
      "parameters/method": "GET",
      "parameters/uri": "@concat('_api/web/lists/getbytitle(''', outputs('Settings')?['listName'], ''')/items?$select=Id,Vorname,Nachname,ParticipantName,ParticipantEmail,PreferredStarterType&$filter=Status eq ''Warteliste'' and PreferredStarterType eq ''Funstarter''&$orderby=TeilnehmerID asc&$top=1')"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "HttpRequest"
    }
  }
}
```

**Unterschied zur Standard-Action:** zusätzliches `PreferredStarterType` im `$select` und im `$filter`. Sortierung in allen drei Branches identisch: `$orderby=TeilnehmerID asc`.

## UI-Anleitung 2026-05-27 (v17.20) — Cancellation-Body in Outlook-Kalender ausgrauen (Nice-to-have)

**Hintergrund:** Beim Sub-Event-Cancel (Action `Ausladen` in `DEX_Outlook`)
wird im Flow `DEX_Outlook_Einladungen` der Teilnehmer per Graph aus der
Attendees-Liste des Outlook-Termins entfernt. Outlook schickt dem
abgemeldeten Teilnehmer dann automatisch eine „Meeting Cancelled"-Notification.
Die enthält den **aktuellen Body** des Termins — also weiterhin die ganze
Agenda etc. Der Hinweis im Subject („Cancelled: …") geht in der Vorschau
unter, der Body wirkt unverändert wie ein gültiger Termin.

**Ziel:** Der Body soll im Cancel-Pfad **nur für den abgemeldeten
Teilnehmer** als „storniert" erkennbar sein — z.B. ausgegrauter Text mit
einem auffälligen roten Banner oben.

**Konzept (App-seitig vorbereitet, Flow-seitig zu ergänzen):**

Die App ruft beim Sub-Event-Cancel `queueOutlookEvent(attendee, eventId,
title, 'Ausladen')` auf. Damit der Flow den Body nur für den
abgemeldeten Empfänger ausgrauen kann, ohne den Termin für die übrigen
Teilnehmer zu verändern, gibt es zwei saubere Pfade — entscheide dich
für **einen**:

**Variante A (empfohlen, einfach):** Schicke per `DEX_SEND_MAIL` eine
zusätzliche Mail mit Inline-iCal-Attachment vom Typ
`METHOD:CANCEL` an den abgemeldeten Teilnehmer. Outlook erkennt das
und markiert die Kalendereintragung visuell als storniert. Das ist seit
v17.20 bereits durch den verbesserten Cancellation-Mail-Body
(grosses rotes „Stornierung"-Banner) teilweise abgedeckt — die echte
iCal-Cancel-Methode wäre ein optionales Upgrade.

**Variante B (komplexer):** Im Flow vor dem Attendee-Remove einen
Graph-PATCH auf `/users/{mailbox}/events/{eventId}` mit einem grau
gefärbten Body schicken, der das Original umschließt, dann Attendee
entfernen, dann den Body wieder zurückpatchen. Risiko: zwischen den
zwei PATCH-Calls bekommen alle übrigen Teilnehmer eine „Event Updated"-
Benachrichtigung. Macht den User-Flow unsauber.

**Status:** Variante A ist app-seitig durch den visuell verbesserten
Cancellation-Mail-Body (rotes Banner + durchgestrichener Event-Titel)
seit v17.20 live. Bei Bedarf kann der Tenant-Admin Variante B im
`DEX_Outlook_Einladungen`-Flow nachziehen — bis dahin reicht der
verbesserte Mail-Body als visuelle Bestätigung der Cancelation.
