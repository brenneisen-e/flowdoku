# CLAUDE.md - Project Instructions

## DEX Event Experience Platform (SPFx)

### Build & Deploy Workflow

When deploying changes to the SPFx project:

1. **Clean stale bundles BEFORE building** (kritisch, sonst waechst das .sppkg bei jedem Build um ca. 2.8 MB):
   ```bash
   cd dex-event-app-spfx
   rm -rf dist release temp sharepoint/solution/debug
   ```
   Hintergrund: `gulp bundle --ship` schreibt jede neue Bundle-Version (`dex-event-platform-web-part_<hash>.js`) zusaetzlich in `dist/`, `release/assets/` UND `sharepoint/solution/debug/ClientSideAssets/`, entfernt die alten aber NICHT. `gulp clean` raeumt diese Ordner NICHT auf. Beim Packagen landen alle Bundles (aus `release/assets/`) im `.sppkg` → es waechst unkontrolliert (beobachtet: 956 KB → 4.5 MB nach 4 Builds). Ein sauberes .sppkg ist ca. 950 KB.
2. Run `npm run package` to build and bundle
   - This automatically bumps the patch version (e.g. 1.0.37 → 1.0.38)
   - If `gulp` is not found, run `npm install` first, then use `npx gulp bundle --ship && npx gulp package-solution --ship`
   - **IMPORTANT:** Always use `npm run package` (not raw gulp commands) so the version gets bumped automatically
3. **Always** copy the built package to `dist/`:
   ```bash
   cp dex-event-app-spfx/sharepoint/solution/dex-event-platform.sppkg dist/dex-event-platform.sppkg
   ```
4. **Size-Check:** Das resultierende `.sppkg` sollte ca. 950 KB - 1.1 MB gross sein. Wenn es deutlich groesser ist (>2 MB), wurden die alten Bundles nicht entfernt - Schritt 1 wiederholen und neu bauen.
5. Commit and push both the source changes and the updated `dist/dex-event-platform.sppkg`

The `dist/dex-event-platform.sppkg` must always reflect the latest build so it can be downloaded directly from GitHub.

### Kein lokales Testen — immer direkt bauen

**WICHTIG:** Der Maintainer testet **nicht lokal** (kein `gulp serve`, kein Workbench, kein Browser-Run-Through). Schlag das auch nicht vor.

Stattdessen: Sobald Code-Änderungen abgeschlossen sind, **direkt** bauen und deployen — ohne Zwischenschritt „lokal testen ob es klappt". Konkret heißt das pro Iteration:

1. Code-Änderung committen.
2. Version bumpen (Minor, siehe unten).
3. `rm -rf dist release temp sharepoint/solution/debug && npx gulp bundle --ship && npx gulp package-solution --ship`
4. `cp sharepoint/solution/dex-event-platform.sppkg ../dist/`
5. `dist/dex-event-platform.sppkg` zum Commit dazu, push.

Validierung erfolgt direkt im SharePoint-Tenant nach Upload, nicht lokal. Wenn ein Bug auftaucht: nächste Iteration mit Fix + Bump + Build, fertig. Kein „bitte erst lokal verifizieren" als Antwort — das blockiert den Maintainer-Flow.

### Versioning Strategy

**WICHTIG (Stand 2026-04-16): Jede Build-Iteration muss um +0.1 (Minor) hochgezaehlt werden.**
SharePoint erkennt Patch-Updates (z.B. 4.1.0 → 4.1.1) **nicht** zuverlaessig als neue
Version und ignoriert den Upload — die App laeuft dann weiterhin mit der alten Version.
Deshalb IMMER Minor-Bump (x.y.0 → x.(y+1).0).

**Vorgehen pro Build:**
- `package.json`: `version` manuell auf naechsten Minor-Wert setzen (z.B. `4.1.0` → `4.2.0`)
- `config/package-solution.json`: `solution.version` und `features[*].version` auf `x.y.0.0` (z.B. `4.2.0.0`)
- `src/webparts/dexEventPlatform/version.ts`: `APP_VERSION` auf `x.y.0`
- Dann: `rm -rf dist release temp sharepoint/solution/debug && npx gulp bundle --ship && npx gulp package-solution --ship`
  (direkt gulp aufrufen, nicht `npm run package` — das bump-Script zaehlt nur Patch und wuerde daraus `4.2.1` machen)
- `dist/dex-event-platform.sppkg` kopieren und committen

**Major-Bump (x.0.0)** nur bei echten Breaking-Changes oder offiziellem Release-Marker
(z.B. `4.x.0` → `5.0.0`). Nicht fuer normale Iterationen.

**Patch-Bump (x.y.z mit z > 0)** ist seit v4.1.x deprecated fuer neue Deployments — nur
noch fuer Hotfixes auf einer bereits deployten Version nutzen, wenn der Tenant
SharePoint-seitig bereit ist, Patch-Updates zu akzeptieren (was meistens NICHT der
Fall ist — im Zweifel immer Minor-Bump).

### SharePoint Site

- Site URL: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- Lists: DEX_Events, DEX_Roles, DEX_Emails, DEX_Outlook, DEX_IDReorder, DEX_Participants, DEX_EmailTemplates
- Per-Event: Subsite with "Teilnehmer" registration list
- Shared Mailbox: `no_reply.events@deloitte.de`

### Event-Schema (Stand v10.20): kein Type-Unterschied mehr

Seit **v10.20** gibt es kein hartes B2Run-vs-Deloitte-Schema mehr. Alle Events
laufen auf demselben Eventschema. Die Funktion **"Geteilte Kapazität"** (zwei
Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste) ist
generisch nutzbar — der Organizer vergibt die Bezeichnungen frei.

Neue SP-Spalten in `DEX_Events`:

- `SplitLabelA` (Single line text) — Bezeichnung Gruppe A (z.B. "Vormittag",
  "VIP", "Lauf 5 km"). Leer = Default-Fallback `'Durchstarter'`.
- `SplitLabelB` (Single line text) — Bezeichnung Gruppe B. Leer =
  Default-Fallback `'Funstarter'`.
- `SplitSharedWaitlist` (Boolean) — `true` = eine gemeinsame Warteliste über
  beide Gruppen (FIFO), `false`/leer = getrennte Wartelisten pro Gruppe
  (alter B2Run-Stil mit typ-bewusstem Nachrücken).

Die SP-Spalte `EventType` ist **seit v5.2 deprecated** und wird **mit v10.20
auch in der UI nicht mehr abgefragt**. Der Type wird beim Laden aus
`CustomFields` abgeleitet (Presence von `b2run_startblock` ⇒ `'B2Run'`,
sonst `'Other'`). Per Admin-Center-Button **"B2Run migrieren"** kann ein
einzelnes Legacy-Event auf das neue Standard-Schema umgestellt werden — das
entfernt die b2run_*-Custom-Fields und persistiert `'Durchstarter'` /
`'Funstarter'` explizit in `SplitLabelA` / `SplitLabelB`.

**Caveat Power-Automate-Flow `DEX_IDReorder_TeilnehmerIDs`:** ohne Flow-
Anpassung promotet der Flow weiterhin typ-bewusst, auch wenn
`SplitSharedWaitlist=true` gesetzt ist. Für die Admin-Center-Abmeldung
läuft die Promotion clientseitig und respektiert `splitSharedWaitlist`
korrekt (siehe `AdminPage.tsx`, `useTypeFilter`-Logik). Damit der
gemeinsame Wartelisten-Modus **auch beim User-Self-Cancel** greift, muss
in der `Is_B2RunSplit`-Bedingung des Flows eine dritte Zeile ergänzt
werden — die UI-Schritt-für-Schritt-Anleitung steht in
`docs/flow-jsons.md` unter "UI-Anleitung 2026-05-06 (v10.20) — Shared-
Waitlist-Modus respektieren". Sobald die Änderung im Tenant
durchgeklickt und gespeichert ist, bitte den neuen Flow-JSON in
`docs/flow-jsons.md` einpflegen.

### OutlookDirty / Update-Confirm (v11.57)

Beim Bearbeiten eines bestehenden Events wird der Wizard-Save um eine
explizite Entscheidung erweitert, ob die Teilnehmer eine
„Aktualisierter Termin"-Benachrichtigung von Outlook bekommen sollen.
Mechanik:

- Neue SP-Spalte `OutlookDirty` (Boolean/Yes-No, Default `false`) auf
  `DEX_Events`. Wird in `ensureEventsList()` mit angelegt — alte Tenants
  ohne diese Spalte interpretieren `undefined` als `false`.
- Beim Mount des Wizards wird ein Snapshot von Titel, Startzeit, Endzeit
  und Outlook-Body in einem React-Ref festgehalten. Beim Save vergleicht
  die App den aktuellen Stand mit dem Snapshot.
- Wenn mindestens ein Outlook-relevantes Feld geändert wurde, der Event
  einen Outlook-Termin hat (`OutlookEventId` gesetzt) und Outlook nicht
  deaktiviert wurde, erscheint vor dem Speichern das Modal
  **„Outlook-Termin der Teilnehmer aktualisieren?"** mit einer Checkbox
  (Default UNCHECKED) + ausführlichem Erklärtext.
- Checkbox an + Speichern: `updateEvent` läuft, danach
  `queueOutlookEvent(eventId, 'UpdateEvent')` + `OutlookDirty=false`.
- Checkbox aus + Speichern: `updateEvent` läuft, KEIN `queueOutlookEvent`,
  stattdessen `OutlookDirty=true`. Beim nächsten Wizard-Lauf zeigt
  Schritt 1 (Grundlagen) eine gelbe Hinweis-Box, dass ein Update aussteht.
- Abbrechen: nichts wird gespeichert.

Sub-Events werden mit demselben Modal abgehandelt — pro Sub-Event mit
Outlook-relevanter Änderung läuft ebenfalls `UpdateEvent` (oder es wird
`OutlookDirty=true` gesetzt). Das Modal trägt einen Zusatzhinweis
„Dies aktualisiert auch alle Sub-Event-Outlook-Termine".

**Power-Automate-TODO:** der `DEX_CreateOutlookEvent`-Flow soll nach
erfolgreichem Create + der `DEX_Outlook_Einladungen`-Flow nach
erfolgreichem UpdateEvent zusätzlich `OutlookDirty=false` zurück auf das
DEX_Events-Item schreiben. Das ist eine optionale Optimierung — solange
die SPFx-App den Flag korrekt setzt, bleibt der Mechanismus funktional.
Die UI-Schritt-für-Schritt-Anleitung steht in `docs/flow-jsons.md` unter
„UI-Anleitung 2026-05-21 (v11.57) — OutlookDirty=false nach
erfolgreichem Create/Update setzen".

**v11.63 — pro-Event-Checkbox.** Das frühere globale „Outlook-Update
senden ja/nein"-Häkchen wurde durch eine Liste pro betroffenem Event
ersetzt. `detectOutlookRelevantChanges()` liefert seit v11.63 nicht mehr
`{ topChanged, affectedSubEventIds }`, sondern
`{ items: [{ kind: 'top'|'sub', eventId, title, changedFields[] }] }` —
jedes Item beschreibt ein konkret geändertes Event und listet die
geänderten Felder (`title` / `startDate` / `endDate` / `outlookBody`).
Das Modal rendert pro Item eine eigene Checkbox plus den klein-grauen
„Geändert: …"-Subtext. Angehakte Events lösen ein `UpdateEvent` in
`DEX_Outlook` aus und bekommen `OutlookDirty=false`, nicht angehakte
Events werden gespeichert und mit `OutlookDirty=true` markiert — pro
Event-ID, nicht mehr global. Die Hinweisbox in Schritt 1 (Grundlagen)
erkennt jetzt auch Sub-Event-Dirty-Marker via `childEventsOf(editEvent.id)`
und nennt Hauptevent und/oder die Zahl der dirty Sub-Events explizit
(„Outlook-Synchronisation steht aus: für das Hauptevent UND X
Sub-Event(s)"). Submit-Pipeline: `pendingOutlookUpdateForTopRef`
(boolean) + `pendingOutlookUpdateForSubEventsRef` (string[]) entscheiden,
welche Events ein `queueOutlookEvent('UpdateEvent')` bekommen;
`pendingOutlookDirtyWriteRefs` (Record<eventId, boolean>) hält die
OutlookDirty-Schreibwerte pro Event-ID.

### Team-Anmeldung — Phase 2 Registration-Flow (v11.82)

Mit v11.82 wird aus der Konfiguration aus v11.80/v11.81 ein
funktionierender Multi-Person-Anmelde-Flow. Der Stand zur App- und
Schema-Seite:

- **Schema-Erweiterung Subsite-Teilnehmerliste:** Drei neue Spalten
  `TeamId` (Text), `TeamLead` (Boolean), `TeamName` (Text) — angelegt
  sowohl in `createRegistrationList()` (neue Events) als auch in
  `fixRegistrationListColumns()` (bestehende Events; Admin per
  „Spalten fixen" nachrüstbar). In der Default-View landen die drei
  Spalten **am Ende**, nach allen Custom-Fields, damit die View bei
  Nicht-Team-Events unauffällig bleibt.
- **Atomare N-Sitz-Reservierung:** `EventService.reserveSeat(...,
  count?: number)` akzeptiert seit v11.82 einen optionalen
  `count`-Parameter (Default 1). Bei `count > 1` wird der
  Sitzplatz-Counter um `count` per ETag-CAS inkrementiert — entweder
  alle N Plätze gehen gleichzeitig in die Gruppe, oder das gesamte
  Team landet auf der Warteliste (`current + N > cap` → 'full').
  Kein Teil-Anmelden bei Engpässen.
- **`EventService.registerTeamMember(...)`**: ein einzelner Insert
  mit Team-Feldern (`TeamId`, `TeamLead`, `TeamName`). Wird vom
  Team-Submit pro Mitglied einmal aufgerufen, ohne den schweren
  Permission-/Dedup-Pfad von `registerForEvent` zu durchlaufen.
- **`EventContext.registerTeam(eventId, leadData, members,
  teamName)`**: orchestriert den ganzen Submit — Doppel-Anmelde-
  Prüfung pro Member, `TeamId` via `crypto.randomUUID()` (Fallback
  `Date.now()+random`), Sitzplatz-Reservierung, parallele
  `Promise.all`-Inserts pro Person, dann pro erfolgreichem Eintrag
  Bestätigungs-Mail (mit „Du wurdest als Teil eines Teams
  angemeldet"-Hinweis-Box plus Klartext-Verweis auf Self-Cancel
  über „Meine Events" falls keine Zustimmung) und
  Outlook-Einladung, plus KPI-Bump um N.
- **UI im Anmeldeformular:** `RegistrationPage` zeigt unter
  „Persönliche Daten" einen Toggle „Ich melde mich + mein Team an"
  nur wenn `event.teamRegistrationEnabled && teamSize >= 2 &&
  !registerForOther`. Aktiver Toggle blendet eine eigene Card auf
  mit (1) auffällig orange Pflicht-Hinweis-Box zur Zustimmung jedes
  Mitglieds, (2) optionalem Team-Name-Pflichtfeld (nur wenn
  `event.askTeamName`), (3) N-1 People-Picker-Slots (Pflicht im
  Modus „Nur komplette Teams", optional bei „Auch Teil-Teams") und
  (4) einer Pflicht-Bestätigungs-Checkbox. Der Submit-Button heißt
  in dem Modus „Team anmelden (N Personen)" und bleibt deaktiviert
  bis die Validation grün ist (keine Duplikate, alle Pflicht-Slots
  gefüllt, Checkbox angehakt).
- **„Meine Events"-Team-Badge:** Jede Registrierung mit `TeamId`
  rendert unter dem Event-Titel ein grünes Badge „Team „<Name>" —
  N/M belegt" plus eine Komma-Liste der aktiven
  Mitglieder-Namen. Wenn der eingeloggte User Lead ist, kommt
  zusätzlich ein kleines „du bist Team-Lead"-Chip dazu. Die andere
  Team-Mitgliederliste wird lazy via
  `EventContext.getTeamMembers(eventId, teamId)` geladen und im
  React-State gecached.
- **Sub-Events:** Für v11.82 ist Team-Anmeldung **auf das Haupt-
  Event beschränkt**. Sub-Event-Subsites bekommen die Team-Spalten
  noch nicht über den Wizard-Pfad zugewiesen — Pro-Sub-Event-
  Team-Anmeldung + Beitritts-/Approve-Flows kommen in v11.83+.

### Team-Anmeldung — Phase 3 (v11.83): Cancel-Promote, Add-Member, Open-Teams, Approve

Aufbauend auf v11.82 ergänzt v11.83 den vollständigen Lebenszyklus
einer Team-Anmeldung:

- **Cancel-Logik für Team-Mitglieder.** `EventContext.cancelRegistration`
  snapshottet vor dem MERGE die Felder `TeamId`, `TeamLead`, `TeamName`
  der eigenen Registrierung. Nach dem Cancel führt
  `handleTeamCancelPostStep` (in `EventContext.tsx`) drei Schritte aus:
  (1) verbleibende aktive Team-Mitglieder via `getTeamMembers` laden,
  (2) wenn die abgemeldete Person Lead war UND ≥1 Member übrig ist, das
  früheste verbleibende Mitglied per `EventService.promoteToTeamLead`
  zum neuen Lead promoten (Kriterium: kleinste `TeilnehmerID`, sonst
  früheste `RegistrationDate`, sonst kleinste Item-Id), (3) pro Member
  eine Info-Mail im Deloitte-Wrap-Layout in `DEX_Emails` queuen mit
  Abgemeldete-Person, aktuelle Belegung `N/TeamSize` und drei
  Handlungs-Optionen. Der Sitzplatz-Counter wird dabei NICHT
  zurückgerollt — die App folgt dem Standard-Reconcile
  (`syncSeatsToActiveCount`), der frei werdende Slot ist neutral und
  darf später vom Lead per Add-Member oder von Dritten via Open-Teams-
  Box belegt werden.

- **Add-Member-Modal in MyEvents.** Der Team-Badge zeigt jedem Lead bei
  freien Slots einen Button `[+ Mitglied hinzufügen (N Slots frei)]`,
  der ein Modal öffnet (orange Pflicht-Hinweisbox, People-Picker via
  `useRoles().searchUsers`, Pflicht-Bestätigungs-Checkbox).
  `EventContext.addTeamMember(eventId, teamId, teamName, member)`
  prüft per `isUserAlreadyOnEvent` auf Doppel-Anmeldungen, reserviert
  atomar einen Sitzplatz (split-aware, mit `PreferredStarterType`-
  Vererbung aus vorhandenen Team-Mitgliedern), legt das Mitglied via
  `registerTeamMember` an und queued Bestätigungs-Mail + Outlook +
  Info-Mails an die anderen Mitglieder. Bei Doppel-Anmeldung gibt's
  einen klaren Fehler-Toast statt eines Inserts.

- **Open-Teams-Box auf der Register-Page.** Wenn
  `event.teamOpenSlotsVisible` aktiv ist und der User selbst noch nicht
  beim Event angemeldet ist, lädt `RegistrationPage` via
  `EventContext.listOpenTeamsForEvent` alle Teams mit Aktiv-Count <
  TeamSize und zeigt sie als Card oberhalb des Formulars. Pro Team:
  Team-Name (falls vorhanden), Belegungs-Anzahl, Beitritts-Button. **Die
  App rendert bewusst KEINE Mitglieder-Namen** (Privatsphäre — nur
  Anzahl). Klick-Verhalten je nach `event.teamJoinRequiresApproval`:
  Direkter Beitritt via `EventContext.joinTeam(eventId, teamId,
  teamName)` (gleicher Pfad wie Add-Member, nur mit dem eingeloggten
  User selbst) ODER `EventContext.createTeamJoinRequest(eventId,
  teamId)` legt eine Pending-Zeile in `DEX_TeamJoinRequests` an und
  queued eine Notification-Mail an den Lead.

- **Approve-Queue `DEX_TeamJoinRequests`.** Neue globale SP-Liste auf der
  Site Collection (nicht pro Subsite) mit Spalten `EventId`, `TeamId`,
  `RequesterEmail`, `RequesterDisplayName`, `Status` (Choice
  `Pending`/`Approved`/`Rejected`), `Created` (Default), `DecidedDate`,
  `DecidedByEmail`. Angelegt durch
  `EventService.ensureTeamJoinRequestsList` in `initEvents`.
  Schreibrechte über `setQueueListPermissions` (analog DEX_Emails). Die
  Notification-Mail an den Lead enthält Approve-/Reject-Buttons als
  HTML-Links auf die App mit Query-Parametern
  `?action=teamjoin&request=<id>&decision=approve|reject` — aktuell
  NICHT hart als URL-Handler verdrahtet, der Lead nutzt stattdessen den
  UI-Block in MyEvents (orange Box „Beitritts-Anfragen (N)").
  `EventContext.decideTeamJoinRequest` Approve-Pfad: lädt die Request-
  Zeile, ruft `addTeamMember`-Logik auf, setzt Request-Status='Approved'.
  Reject-Pfad: Status='Rejected' + kurze Absage-Mail.

- **Doppel-Anmelde-Prävention.** Neuer Helper
  `EventService.isUserAlreadyOnEvent(subsiteUrl, email)` filtert auf
  Status in `Angemeldet | QR versendet | Eingecheckt | Warteliste`.
  Wird in `registerTeam` (v11.82 konsolidiert auf den Helper),
  `addTeamMember` (v11.83) und `createTeamJoinRequest` (v11.83) genutzt.

### Team-Anmeldung — Phase 5 (v11.86): Lead bearbeitet sein Team aus MyEvents

Mit v11.86 bekommt der Team-Lead in „Meine Events" einen
vollständigen Edit-Pfad direkt auf seiner eigenen Event-Karte —
ohne Umweg über das Admin Center.

- **Mitglieder-Karten im Team-Badge.** Statt der Inline-Komma-Liste
  „Mitglieder: …" rendert der Team-Badge jetzt pro Mitglied eine
  eigene Karte mit Profilfoto (40×40, Hover-Zoom scale 2.4× analog
  SettingsPage), Name, E-Mail und Standort (mit „ · "-Trenner).
  Aktive Mitglieder mit `TeamLead=true` bekommen einen grünen
  Lead-Pill rechts, abgemeldete Mitglieder werden ausgegraut mit
  einem grauen „abgemeldet"-Badge gerendert. Sortierung: Lead
  zuerst, danach nach `TeilnehmerID`, sonst nach Id.
- **Button „Team bearbeiten".** Sichtbar nur für Leads, sitzt
  neben dem bestehenden „+ Mitglied hinzufügen"-Button. Klick
  öffnet das Modal **„Team verwalten"**: Headline mit Team-Namen,
  Sub-Headline mit Belegung `N/TeamSize`, alle Mitglieder als
  Karten (gleicher Stil wie im Badge). Pro aktivem Mitglied außer
  dem Lead selbst rendert ein roter Trash-Button mit
  Title-Tooltip. Klick öffnet eine zweite Confirm-Modal-Ebene mit
  dem Namen der Person und einem Hinweis, dass die Abmeldung
  stellvertretend passiert und die Person eine eigene
  Bestätigungs-Mail bekommt.
- **`EventContext.cancelTeamMember(eventId, memberRegistration)`**:
  ruft `EventService.cancelRegistration` mit dem eingeloggten Lead
  als Audit-Akteur auf, schreibt den ChangeLog mit `asActor:
  'teamLead'`, queued die Standard-Abmelde-Mail an die abgemeldete
  Person, queued den `Ausladen`-Outlook-Event, den IDReorder,
  synct den Sitzplatz-Counter und ruft anschließend
  `handleTeamCancelPostStep(..., wasTeamLead=false, ...)` für die
  Info-Mails an die übrigen Mitglieder. Auto-Promote läuft hier
  nicht, weil der Lead sich nicht selbst löscht — Self-Schutz im
  Context wirft einen Warn-Log, falls jemand das versucht. Der
  Lead-Self-Cancel läuft weiterhin über den normalen Abmelden-Button
  via `cancelRegistration` und nutzt die bestehende Auto-Promote-
  Logik aus v11.83.
- **Schema.** Die SP-Spalte `Location` ist bereits in der
  Teilnehmerliste vorhanden (seit jeher), wurde nur als optionales
  Feld in `SPRegistration` nachgezogen, damit `getTeamMembers`
  sie typesafe ausliefert. Keine `fixRegistrationListColumns`-
  Änderung nötig.

### Team-Anmeldung — Phase 4 (v11.84): Admin-Center-Team-Management

Mit v11.84 bekommen Admins und Organizer eigener Events einen
direkten Eingriff in bestehende Teams aus dem Admin Center heraus —
ohne den Umweg über „Meine Events" als Team-Lead.

- **Teams-Sektion in `AdminPage.tsx`:** Bei Events mit
  `teamRegistrationEnabled === true` rendert die Admin-Detail-Seite
  oberhalb der Teilnehmer-Tabelle (nach Statistiken und
  Überbuchungs-Box) eine eigene collapsible Card „Teams (N)".
  Teams werden live aus dem bereits geladenen `registrations`-State
  per groupBy(`TeamId`) gebildet — kein zusätzlicher Roundtrip.
  Abgemeldete Mitglieder werden ausgeblendet, die Teams sortieren
  nach Lead-RegistrationDate (älteste zuerst), innerhalb eines
  Teams steht der Lead oben und danach die Mitglieder nach
  TeilnehmerID aufsteigend. Pro Mitglied: Profilfoto via
  `userphoto.aspx?accountname=<email>&size=L` (mit Hover-Zoom
  scale 2.4×), Name + Email, Status-Badge (außer „Angemeldet"),
  und ein grüner „Lead"-Pill für den TeamLead.
- **„Person hinzufügen"-Button pro Team:** Nur sichtbar wenn das
  Team noch freie Slots hat (`activeCount < teamSize`). Öffnet ein
  Modal mit dem gleichen orangen Pflicht-Hinweisbox-Pattern wie der
  Lead-Add in MyEvents — People-Picker via `searchUsers`,
  Pflicht-Bestätigungs-Checkbox. Submit ruft `addTeamMember` aus
  EventContext (existiert seit v11.83) auf — dedup-Schutz,
  Sitzplatz-Reservierung, Insert, Bestätigungs-Mail + Outlook +
  Info-Mails an die anderen Mitglieder. Nach Erfolg ein grüner
  Toast „X wurde zum Team hinzugefügt — Mail + Outlook werden
  versendet." und die Teilnehmer-Tabelle wird mit
  `getAllRegistrations()` neu geladen.
- **„Lead-Rolle übergeben"-Dropdown:** Nur sichtbar wenn das Team
  ≥ 2 aktive Mitglieder hat. Klick öffnet ein Inline-Dropdown mit
  allen anderen aktiven Mitgliedern (mit Foto + Name + Email).
  Auswahl ruft `EventContext.transferTeamLead(eventId, teamId,
  newLeadEmail)`. Die Context-Funktion lädt die aktuellen Members
  via `getTeamMembers`, identifiziert alten Lead + Ziel-Member und
  delegiert an `EventService.transferTeamLead(subsiteUrl,
  fromLeadItemId, toNewLeadItemId)` — zwei MERGE-Patches in Folge
  (best-effort, kein echtes Transactional weil SP keine Multi-Item-
  Transaktionen kennt). Bei Erfolg gehen Info-Mails an alle
  aktiven Mitglieder raus (im Deloitte-Layout via `wrapTemplate`),
  der neue Lead bekommt einen Extra-Hinweis auf seine erweiterten
  Rechte. Audit-Eintrag im ChangeLog mit
  `action='TeamLeadTransferred'` plus alter und neuer Lead-Email.
- **Berechtigungen:** Sichtbar für Admin (alle Events) oder
  Organizer (eigene Events; via `isOrganizerFor(selectedEvent)`).
  Die Buttons sind sonst komplett ausgeblendet. Die Teams-Sektion
  selbst bleibt für nicht-berechtigte Rollen verborgen, weil die
  Admin-Detail-Seite ohnehin nur für sie gerendert wird.

### Self-Check-in per QR-Code (v18.33)

Teilnehmer checken sich am Veranstaltungstag selbst ein, indem sie einen
event-spezifischen QR-Code mit der **nativen Handy-Kamera** scannen. Das
umgeht bewusst den In-App-Scanner (`CheckInPage`/`qr-scanner`), weil die
SharePoint-Mobile-App `getUserMedia` im WebView blockiert — der native
Kamera-Scan öffnet stattdessen einen Deep-Link, der die App startet und den
eingeloggten User automatisch eincheckt.

**Zwei QR-Modi, ein geheimer Token pro Event:**

- **Statisches PDF** (druckbar, bequem): URL `?action=selfcheckin&token=<Token>`.
  Der Token ist der Lookup-Key. Per Foto teilbar → mit Zeitfenster kombinieren.
- **Rotierende Live-Anzeige** (foto-sicher): URL
  `?action=selfcheckin&event=<Nr>&code=<HMAC>&t=<Fenster>`. Der Code wechselt
  alle `SELF_CHECKIN_STEP_SECONDS` (45s) und wird per HMAC-SHA256(Token,
  Fenster-Index) über **Web Crypto** clientseitig validiert (aktuelles + 1
  vorheriges Fenster gelten als frisch). Kein Server / keine Power-Automate-
  Änderung nötig.

**Neue SP-Spalten auf `DEX_Events`** (in `getEventsFieldDefinitions()`, daher
via `ensureMissingFields()` auch auf Bestands-Tenants nachgezogen):

- `SelfCheckInEnabled` (Boolean) — Feature an/aus.
- `SelfCheckInToken` (Single line text) — geheimer Token (Lookup-Key +
  HMAC-Schlüssel). Wird beim Aktivieren im Wizard einmalig per
  `generateSelfCheckInToken()` erzeugt und bleibt stabil.
- `SelfCheckInFrom` / `SelfCheckInTo` (DateTime, optional) — Check-in-
  Zeitfenster. Leer = Default „nur am Event-Tag" (Start- bis End-Datum,
  `isWithinCheckInWindow()`).

**Architektur:**

- `utils/selfCheckIn.ts` — HMAC/Fenster-Logik, Token-Gen, URL-Builder,
  Fenster-/Frische-Prüfung (alles rein client, Web Crypto).
- `utils/selfCheckInPdf.ts` — `downloadSelfCheckInPdf()` baut das A4-PDF per
  **jsPDF** (neue Dependency `jspdf`) mit QR (`qrcode`) + Anleitung.
- `EventContext.selfCheckIn(params)` — orchestriert: Event per Token ODER
  Event-Nr auflösen (`EventService.getEventBySelfCheckInToken` /
  `getEventByEventNumber`), Enabled/Fenster/Frische prüfen, eigene
  Registrierung via `getMyRegistration` finden, `checkInParticipant`. Gibt
  `SelfCheckInResult` (`success`/`already`/`not-registered`/`on-waitlist`/
  `not-found`/`disabled`/`closed`/`expired`/`error`).
- `SelfCheckInPage.tsx` — vollflächige Ergebnis-UI für den User
  (Deep-Link `?action=selfcheckin`, in `DexEventPlatform.tsx` vor dem
  Boot-Loader abgefangen).
- `SelfCheckInDisplayPage.tsx` — rotierende Live-Anzeige für den Organizer
  (Page `self-checkin-display`, Deep-Link `?action=selfcheckin-display&event=<id>`).
- Wizard: Toggle + Zeitfenster + Erklär-Modal in Schritt **Kapazität &
  Sichtbarkeit** (`EventCreationPage.tsx`); Persistenz im create-Objekt UND
  im Edit-`updates`-Payload.
- Admin Center: zwei `ActionTile`s „Self-Check-in: QR-PDF" und „… Live-
  Anzeige" (sichtbar bei `selfCheckInEnabled` + Admin/Organizer + Token).

**Sicherheit:** Jeder checkt nur **sich selbst** ein (Login-gebunden,
Item-Level-Security). Gegen „von zu Hause einchecken" wirken: Zeitfenster
(Default Event-Tag) + im Live-Modus die Code-Rotation (abfotografierter Code
verfällt). Restrisiko des statischen PDFs ist bewusst akzeptiert (nur
Anwesenheits-Schummeln, keine Fremd-Anmeldung).

### People-Picker-Feld → Person auf CC der An-/Abmelde-Mail (v18.41)

Pro People-Picker-Custom-Field (`type === 'user'` / `'roommate'`) kann der
Organizer im Wizard (Schritt **Felder**) das Häkchen **„Ausgewählte Person bei
An-/Abmelde-Mail auf CC setzen"** aktivieren (Property `ccOnEmails`). Use-Case:
ein „Assistenz"-Feld — die angegebene Person bekommt Anmelde- und Abmelde-Mail
automatisch in Kopie.

- **NUR E-Mails, NICHT Outlook:** Die CC wirkt ausschließlich auf
  `queueEmail` (Anmeldung / Warteliste / Abmeldung). Der Outlook-Termin
  (`queueOutlookEvent`) bleibt unangetastet — die CC-Person wird NICHT zum
  Kalendereintrag eingeladen.
- **Extraktion:** `collectCcEmailsFromFields(fields, customData, excludeEmail)`
  in `EventContext.tsx` liest aus den `ccOnEmails`-Feldern die E-Mail
  (Wertformat „Anzeigename <email>", Regex `/<([^>]+@[^>]+)>/`), dedupliziert
  und schließt den Teilnehmer selbst aus. Ergebnis geht als `cc`-Parameter in
  `queueEmail` (8. Argument).
- **Eingehängt in:** `registerForEvent` (Bestätigung/Warteliste, aus dem
  live `customData`), `cancelRegistration` (Self-Cancel, aus
  `myReg.CustomData`), `cancelTeamMember` (aus `memberRegistration.CustomData`).
- **Persistenz:** `ccOnEmails` ist eine Custom-Field-Property — daher in
  `serializeCustomFields` **und** im `cfForFix`-Mapping ergänzt (zweiter
  CustomFields-Write, siehe Abschnitt oben) sowie im SP→Event-Parse
  (`eventSpecificFields.map`) durchgereicht.

### Dokument-Custom-Feldtyp (v19.0)

Neuer Custom-Field-Typ **`document`** — der Organizer bittet Teilnehmer um einen
**PDF-/Bild-Upload**. Die Datei wird als **SP-Item-Attachment an die Teilnehmer-
Zeile** gehängt (kein Spaltenwert) und ist im **Admin Center** pro Teilnehmer
abrufbar. Nutzt die bestehende Attachment-Infrastruktur (`allowAttendeeUpload`)
weiter.

- **Type-Union** an 4 Stellen ergänzt: `types/index.ts` (`EventSpecificField`),
  `EventService.ts` (`CustomField`), `EventCreationPage.tsx` (`CustomFieldInput`
  + Wizard-Dropdown „Dokument (PDF/Bild-Upload)"). Persistenz läuft über den
  bestehenden `type`-Durchgriff (kein neues Property → keine `cfForFix`-Falle).
- **Keine SP-Spalte:** Dokument-Felder werden in `createRegistrationList` UND
  `fixRegistrationListColumns` per `continue` übersprungen; die „Antwort" ist die
  angehängte Datei, kein Spaltenwert. Daher auch nicht in `customData`.
- **Feld-↔Datei-Zuordnung über Dateinamen-Präfix:** ein Event kann mehrere
  Dokument-Felder haben. `addRegistrationAttachment(..., fieldPrefix)` schreibt
  `dxf-<sanitizedFieldId>--<ts>_<name>`; gelistet wird per `startsWith`-Filter.
  Helper `docFieldPrefix` / `stripDocPrefix` in `EventContext.tsx`.
- **Neue Context-Methoden:** `uploadFieldDocument`, `listFieldDocuments`,
  `deleteFieldDocument` (jeweils mit optionalem `participantEmail` für die
  stellvertretende Anmeldung; Default = eingeloggter User). Sie lösen das
  Listen-Item per `getMyRegistration` auf und rufen die EventService-Attachment-
  Helper.
- **Upload-Timing:** Attachments brauchen die Item-Id, die erst nach dem Insert
  existiert. `RegistrationPage` hält die gewählte Datei in `pendingDocFiles`
  (NICHT in `customData`) und lädt sie in `performRegistration` NACH erfolgreicher
  Anmeldung hoch. Pflicht-Dokumentfelder erzwingen bei NEU-Anmeldung eine Datei
  (bei bereits angemeldeter Person läuft das über „Meine Events").
- **„Meine Events":** pro Dokument-Feld ein `MyEventDocField`-Block
  (`MyEventsPage.tsx`) zum nachträglichen Ergänzen/Ersetzen/Löschen.
- **Admin Center:** der „Datei"-Button + Attachment-Modal (vorher nur bei
  `allowAttendeeUpload`) erscheint jetzt auch bei Events mit Dokument-Feld; im
  Modal wird der `dxf-…`-Präfix gestrippt und das Feld-Label als Badge angezeigt.
- **Akzeptierte Typen:** PDF + JPG/PNG (Upload-Limit 10 MB). Kein Power-Automate-
  Change (reine SP-REST-Attachments).

### Grüne Hervorhebung ausgefüllter Felder (v19.0)

Auf der Anmeldeseite bekommen **ausgefüllte** Custom-Felder (Dropdown gewählt,
Text/Zahl eingegeben, Person ausgewählt, Mehrfachauswahl getroffen, Dokument
hochgeladen) denselben **grünen Rand + zarten grünen Hintergrund** wie eine
ausgewählte Event-Section. Umgesetzt in `renderRegField` (`inputStyleGreen`) +
`MultiSelectDropdown` + `UserFieldPicker` (Chip). Checkboxen waren schon grün.

### Teams im Organizer-Center als 3-Spalten-Raster (v19.0)

Die Teams-Sektion im Admin Center (`AdminPage.tsx`) rendert die Team-Karten jetzt
in einem responsiven Raster (`repeat(auto-fill, minmax(300px, 1fr))`) statt
vollbreit gestapelt — und nummeriert die Teams durch (`1.`, `2.`, …). Spart
vertikalen Platz bei vielen Teams.

### Sicherheitshinweis vor dem Absenden der Anmeldung (v18.75)

Pro Event kann der Organizer einen **Bestätigungs-Dialog** aktivieren, der nach
dem Klick auf „Anmelden" und VOR der eigentlichen Anmeldung erscheint.
Konfiguration: Schritt 5 (Felder), eigene Section **ganz unten**.

- **Neue SP-Spalten auf `DEX_Events`** (via `getEventsFieldDefinitions()` →
  `ensureMissingFields`): `ConfirmDialogEnabled` (Boolean, Default false),
  `ConfirmDialogMode` (Single line text: `'summary'` | `'freetext'`),
  `ConfirmDialogText` (Note). Durchgereicht über `SPEvent` (EventService),
  `EVENT_SELECT`, `createEvent`-Payload, `DeloitteEvent` (types/index.ts),
  `CreateEventInput` (EventContext) und die Wizard-Create-/Update-Payloads.
- **Zwei Modi:** `'summary'` = Auswahl-Übersicht (Haupt-Event + **alle**
  Sub-Events als Checkbox-Liste **mit Datum/Uhrzeit**; der Teilnehmer kann vor
  dem Absenden einzelne Punkte ab- UND zuwählen — auch noch nicht gewählte
  Sub-Events; bei stellvertretender Anmeldung ist das Haupt-Event fixiert. Wird
  ein Sub-Event mit eigenen Pflichtfeldern zugewählt, öffnet erst das
  Sub-Event-Modal, danach erscheint der Dialog erneut). `'freetext'` = frei
  formulierter Hinweis (`ConfirmDialogText`) mit Pflicht-Bestätigungs-Checkbox.
- **Feld-Ausrichtung (v18.76/v18.77):** Im 2-Spalten-Custom-Field-Grid
  (`dex-reg-fields-grid`) bekommt die Inline-Hilfe eine **`minHeight` von ~2
  Zeilen** (`2.9em`), sodass die Eingaben benachbarter Felder auf gleicher Höhe
  stehen, wenn sich die Beschreibungen um eine Zeile unterscheiden. (v18.76
  nutzte stattdessen `flexGrow:1` + Flex-Spalte — das zog bei Feldern mit Inhalt
  UNTER der Eingabe, z.B. People-Picker mit „international suchen"-Schalter, die
  Nachbar-Eingabe bis ganz nach unten und erzeugte große Lücken; daher in v18.77
  ersetzt.)
- **Externe-Person-Felder (v18.76):** Vorname/Nachname/E-Mail sind NUR im
  Extern-Modus (`externalPerson`) frei editierbar; im normalen
  Stellvertreter-Modus werden sie read-only vom People-Picker befüllt.
- **RegistrationPage:** Gate in `handleSubmit` VOR `performRegistration` (Team-
  und Beitritts-Pfade sind ausgenommen — die haben eigene Bestätigungen). Der
  Dialog setzt `confirmDialogConfirmedRef`; beim Bestätigen wird die (ggf.
  angepasste) Auswahl in `registerForParent` / `selectedSessions` übernommen und
  `handleSubmit` erneut angestoßen (überspringt dann den Dialog). Reiner
  Client-Flow, **kein Power-Automate-Change**.

### Zustimmungs-Nachweis + externe Personen bei stellvertretender Anmeldung (v18.74)

Zwei zusammenhängende Erweiterungen am „Für andere registrieren"-Flow:

- **Zustimmungs-Nachweis (`ProxyConsent`):** Neue Note-Spalte `ProxyConsent` auf
  der Teilnehmerliste (`createRegistrationList` + `fixRegistrationListColumns` +
  Default-View). Bei jeder stellvertretenden Anmeldung (Teilnehmer-E-Mail ≠
  eingeloggter User) schreibt die App einen lesbaren Nachweis hinein, z.B.
  „Schriftliche Zustimmung der Person zur stellvertretenden Anmeldung bestätigt
  durch <Actor> (<Email>) am <Datum>". Der Organizer kann so in SharePoint
  nachweisen, dass die Zustimmung vorlag. Mechanik: `registerForEvent`/
  `reactivateRegistration` (EventService) nehmen einen `proxyConsent`-Parameter;
  `EventContext.registerForEvent` baut den String (intern = „Zustimmung",
  extern = „Schriftliche Zustimmung") und gibt ihn weiter, sobald
  `opts.proxyConsentConfirmed` gesetzt ist (RegistrationPage setzt das bei
  `registerForOther`).
- **Externe Personen (`externalPerson`):** Im „Für andere registrieren"-Modus
  gibt es für Organizer/Admins die Checkbox **„Person außerhalb Deloitte"**.
  Aktiv blendet sie den Deloitte-People-Picker aus und macht Vorname/Nachname/
  E-Mail **frei eintragbar**. Hinweis: Zustimmung **schriftlich** einholen.
  Versand: die **Bestätigungs-Mail geht direkt an die externe Person, mit den
  Organizern auf CC** (Nachweis) — **kein Outlook-Termin** (an externe Adressen
  nicht möglich). Das ändert das frühere Verhalten („Weiterleitung notwendig" an
  den Organizer umgeleitet): externe Empfänger bekommen die Mail jetzt direkt
  (`EventContext.registerForEvent` → `externalCcExtra` = Organizer-E-Mails in den
  CC gemischt; `skipOutlookForExternal` unterdrückt den Termin wie bisher).

### Header-Bild in Mail/Outlook frei einstellbar (v18.73)

Pro Event lassen sich **Breite** und **Innenabstand** (seitlich + oben/unten) des
Hero-Bildes (`{{ORB_URL}}` = Event-Bild) im Mail- UND Outlook-Termin-Kopf
einstellen — Use-Case: Bild größer und fast randlos statt klein mit viel
Weißraum.

- **Render:** `EmailTemplates.ts` → `WrapHeadingOpts` um `imageWidth`,
  `imagePaddingV`, `imagePaddingH` erweitert; gemeinsamer Helper `buildHeroRow()`
  baut die Hero-Zeile in `wrapTemplate()` + `wrapTemplateForStorage()`
  (`width:100%;max-width:Wpx` = responsiv, `width`-Attribut als Outlook-Fallback).
  Defaults (180 / 30 / 30) lassen Alt-Aufrufe unverändert.
- **Mails:** `buildEmailFromTemplate()` reicht die Felder an `wrapTemplate` durch;
  `applyEventTemplateOverride()` liest das **globale** Layout aus dem reservierten
  Piggyback-Key `_headerImageLayout` und merged es in JEDEN Template-Typ.
- **Outlook:** `buildOutlookBody(..., imgOpts?)` — die 3 Wizard-Call-Sites
  (Edit / Create / Sub-Event) geben das Layout mit.
- **Persistenz:** Piggyback-Key `_headerImageLayout` (`{ width, paddingV,
  paddingH }`) in `EmailTemplateOverrides`-JSON — beim Load gestrippt (eigener
  State `headerImageLayout`), beim Save in Create-/Edit-/Sub-Event-JSON gemerged.
  **Kein Flow-Change** nötig (der Flow ersetzt nur `{{ORB_URL}}`-Src, Breite +
  Padding stehen im gewrappten Body).
- **UI:** Block „Header-Bild" im `HtmlEditorModal` (Schritt 6), sichtbar in der
  Mail- UND Outlook-Vorschau, live in der Vorschau.

### Beschreibung-Hinweis: Name/Datum/Ort redundant (v18.73)

Schritt 1 (Grundlagen) zeigt unter dem Beschreibungs-Editor eine **orange
Hinweis-Box**, wenn die Beschreibung den **Event-Namen**, das **Datum** oder den
**Ort** enthält (diese werden bereits separat auf der Anmelde-Seite angezeigt).
Die Box enthält einen **klickbaren Beispieltext**, der per Klick die Beschreibung
mit einer einladenden Standard-Formulierung ersetzt (`setDescription`).
Erkennung clientseitig im Wizard (Plain-Text-Vergleich, Datum in mehreren
Formaten inkl. Monatsname).

### Event-spezifische Karte ausblenden wenn leer (v18.73)

Die Karte „Event-spezifische Informationen" auf der Anmelde-Seite
(`RegistrationPage.tsx`) wird **komplett ausgeblendet**, wenn es dort nichts
auszufüllen/auszuwählen gibt — Bedingung
`eventSpecificFields.length > 0 || isSplitGroup || childEvents.length > 0`.
Vorher erschien eine leere Karte mit „Keine zusätzlichen Informationen
erforderlich".

### Team-Beitritt: erst vormerken, dann anmelden (v18.73)

Bisher trat ein User über die „Offene Teams"-Box **sofort** beim Klick bei —
ohne seine event-spezifischen Infos anzugeben. Neu:

- Die „Offene Teams"-Box sitzt jetzt **unter** der „Ich melde mich + mein Team
  an"-Karte (vorher über „Persönliche Daten"). Reihenfolge: persönliche Daten →
  Team-Karte → offene Teams → event-spezifische Infos.
- Klick auf **„Vormerken"** wählt ein Team nur vor (`pendingJoinTeam`-State,
  gegenseitig exklusiv zum Team-Anlege-Modus). Die eigentliche Anmeldung
  (inkl. Pflicht-Custom-Felder) passiert erst über den **„Anmelden"**-Button
  (`performJoinSelectedTeam`). Der Button heißt dann „Team beitreten & anmelden"
  bzw. „Beitritt anfragen".
- **`customData` durchgereicht:** `joinTeam()` / `addTeamMember()` nehmen einen
  optionalen `customData`-Parameter (→ `registerTeamMember`, vorher hart `{}`).
  Bei Approval-Teams speichert `createTeamJoinRequest()` die Antworten als JSON
  in der neuen Spalte **`CustomData`** (Note) auf `DEX_TeamJoinRequests`; der
  Approve-Pfad (`decideTeamJoinRequest`) wendet sie auf den neuen Member an.
- Eigener Erfolgs-Screen für Beitritt (direkt angemeldet / Warteliste / Anfrage
  gesendet).

### Anmeldesprache vorgeben (v18.35)

Pro Event kann der Organizer die **Sprache der Anmeldeseite** fest vorgeben —
unabhängig von der App-Sprache des Teilnehmers. Use-Case: ein englisch-
sprachiges Event soll die Anmeldung (inkl. Datenschutz-Disclaimer) **immer auf
Englisch** zeigen, auch wenn der User die App auf Deutsch nutzt.

- **Neue SP-Spalte `RegistrationLanguage`** (Single line text, `''`/`'de'`/`'en'`)
  auf `DEX_Events`, via `getEventsFieldDefinitions()` (→ `ensureMissingFields`
  auch auf Bestands-Tenants). `''`/undefined = Default „folgt App-Sprache".
- **Wizard:** Dropdown „Sprache des Anmeldeformulars" in Schritt **Felder**
  (`EventCreationPage.tsx`): Automatisch / Immer Deutsch / Immer Englisch.
  Persistiert im create-Objekt UND im Edit-`updates`-Payload.
- **`RegistrationPage.tsx`:** überschreibt lokal `locale` (und `t`, sowie
  `eventLocale` für das Form-Chrome) mit der erzwungenen Sprache, sobald
  `event.registrationLanguage` gesetzt ist. Dadurch greifen **alle**
  bestehenden `t(...)`/`locale`-Verwendungen automatisch — inkl. Disclaimer
  (`t('reg.privacy')`). Custom-Field-EN-Varianten greifen über das bestehende
  `bilingualFields`-Verhalten.
- **Header-Hinweis (`Header.tsx`):** auf der Registrierungsseite eines Events
  mit fester Sprache zeigt der Header einen kleinen Chip **in der erzwungenen
  Sprache** („Anmeldung auf Deutsch" / „Registration in English").

### Sub-Event-Tabs in Schritt 6 (v11.57, mit v11.80 renumbered von 5)

Schritt 6 (Kommunikation) zeigt eine Tab-Leiste, sobald das Event
Sub-Events hat. Tabs: erster Tab „Haupt-Event: <title>", danach pro
Sub-Event ein Tab mit dessen Titel. Beim Tab-Wechsel wird der aktuelle
UI-State zwischen Top-Level-State und der `SubEventDraft`-Slice
gespiegelt. Persistiert werden pro Sub-Event mindestens: Mail-Sprache,
Outlook-Body (gewrappt), Mail-/Outlook-Logos (Piggyback in
EmailTemplateOverrides analog zum Hauptevent), DisableEmails,
DisableOutlook. Aktive Tab-UI ist konsistent mit dem AdminPage-Tab-Look
(grüne Unterstreichung).

### Custom-Field-Properties: zweiter CustomFields-Write beim Edit-Save (WICHTIG, wiederkehrender Bug)

**Symptom:** Eine neue Property an einem Custom-Field (z.B. `helpText`,
`showIf`, `confirmLabel`, `externalLinks`, `onlyForGroup`, `multi`,
`helpTextStyle`) wird im Wizard gesetzt, ist nach dem **Speichern eines
bestehenden Events** (Edit) aber sofort wieder weg / auf Default.

**Ursache:** `handleSubmit` in `EventCreationPage.tsx` schreibt beim
**Edit-Save** das `CustomFields`-JSON **zweimal**:

1. Erster Write (≈ Zeile 2769): `JSON.stringify(serializeCustomFields(...))`
   — der zentrale Serializer, der **alle** Properties korrekt persistiert.
2. **Zweiter Write** (≈ Zeile 3082): nach `fixRegistrationListColumns()`
   wird `CustomFields` **nochmal** überschrieben, um die `spInternalName`-
   Zuordnung der Teilnehmerlisten-Spalten nachzutragen. Dieser zweite
   Payload wird aus dem **separaten `cfForFix`-Mapping** (≈ Zeile 3030)
   gebaut — und dieses Mapping listet die Properties **einzeln und
   manuell auf**. Jede Property, die dort fehlt, wird vom zweiten Write
   wieder vom SP-Item **entfernt**.

**Regel:** Wenn du eine neue Custom-Field-Property einführst, musst du sie
an **beiden** Stellen ergänzen:
- in `serializeCustomFields()` (Helper, ≈ Zeile 196) **und**
- im `cfForFix`-Mapping in `handleSubmit` (≈ Zeile 3030).

Bug-Historie (immer dasselbe Muster): v11.21 (`helpText`/`showIf`),
v11.94 (`confirmLabel`), v11.15 (`externalLinks`), v18.20 (`helpTextStyle`),
v18.41 (`ccOnEmails`).
Der zweite Write feuert nur, wenn `fixResult.customFieldMap` Einträge
liefert (also Spalten gemappt/angelegt wurden) — deshalb fällt der Bug
oft erst bei Events mit echten Teilnehmerlisten-Spalten auf, nicht in der
Wizard-Live-Vorschau (die `serializeCustomFields` gar nicht durchläuft).

**Hinweis (Tech-Debt):** Der zweite Write baut die Feldliste komplett neu,
statt nur `spInternalName` in die bereits serialisierten Felder zu mergen.
Besserer Fix beim nächsten Touch: das Ergebnis von `serializeCustomFields`
nehmen und nur `spInternalName` pro Feld ergänzen — dann kann diese
Property-Liste nie wieder veralten. Caveat: die EN-Varianten (`labelEn`,
`helpTextEn`, `optionsEn`, `confirmLabelEn`) werden vom zweiten Write
aktuell **ebenfalls** gedroppt — bei Bilingual-Events also dieselbe Falle.

### SharePoint REST API — MERGE-Requests (WICHTIG)

Bei allen SharePoint-List-Item-Updates per `this._merge(url, body)`:

- `_merge` sendet mit Header `Content-Type: application/json;odata=nometadata`.
- **Bei `odata=nometadata` darf KEIN `__metadata: { type: '...' }` im Body stehen.**
- SharePoint lehnt das sonst mit HTTP 400 ab:
  `The property '__metadata' does not exist on type 'SP.Data.XxxListItem'.`
- Der Body muss **nur die echten Feld-Werte** enthalten, z.B. `{ Vorname: 'Max', Status: 'Angemeldet' }`.

**Gegenbeispiel (falsch):**
```ts
const body = { '__metadata': { 'type': 'SP.Data.TeilnehmerListItem' }, ...patch };
await this._merge(url, body);  // → HTTP 400
```

**Korrekt:**
```ts
await this._merge(url, patch);  // nur Felder
```

Referenz-Pattern siehe `updateRegistrationData()` in `EventService.ts` — verlaesslich seit Jahren im Produktivbetrieb.

**Ausnahme: `odata=verbose`** — andere Funktionen wie `updateEvent()` nutzen direkt `this.context.spHttpClient.post()` mit `Content-Type: application/json;odata=verbose`. Dort ist `__metadata` **Pflicht**. Diese Funktionen nutzen NICHT `_merge` sondern bauen Header manuell.

Merksatz:
- `_merge(...)` = `nometadata` → KEIN `__metadata`
- Manueller POST mit `odata=verbose` → `__metadata` ist Pflicht

### SharePoint Item-Type-Encoding bei Listennamen mit Sonderzeichen (WICHTIG)

Bei `odata=verbose`-Inserts erwartet SharePoint den exakten Type-Namen `SP.Data.<Listenname>ListItem`. **Sonderzeichen im Listennamen werden encodet** und MÜSSEN im Type-Namen genauso encodet werden, sonst schlägt der POST stillschweigend mit HTTP 400 fehl ("Cannot find resource for the request") oder der Request gilt als invalid:

| Zeichen | Encoding   | Beispiel                                                       |
|---------|------------|----------------------------------------------------------------|
| `_`     | `_x005f_`  | `DEX_Events` → `SP.Data.DEX_x005f_EventsListItem`              |
| `-`     | `_x002d_`  | `My-List` → `SP.Data.My_x002d_ListListItem`                    |
| ` `     | `_x0020_`  | `My List` → `SP.Data.My_x0020_ListListItem`                    |

Beispiele aus dem Codebase die korrekt funktionieren:
- `DEX_Events` → `SP.Data.DEX_x005f_EventsListItem` (siehe `updateEvent()`)
- `DEX_Roles`, `DEX_Emails`, `DEX_Outlook` etc. genauso encodet
- `Teilnehmer` (ohne Sonderzeichen) → `SP.Data.TeilnehmerListItem` (siehe `REG_LIST_ITEM_TYPE`)
- `DEX_TeilnehmerCounter` (v7.28+) → `SP.Data.DEX_x005f_TeilnehmerCounterListItem`

**Bug-Story v7.28 → v7.29**: Beim Anlegen der Counter-Liste war der Type fälschlich als `SP.Data.DEX_TeilnehmerCounterListItem` (ohne `_x005f_`) angegeben. Der Item-Insert ging mit HTTP 400 stillschweigend kaputt — die Liste wurde zwar angelegt, das Seed-Item aber nicht. Ergebnis: Counter leer → Fallback auf `max+1` → Race-Conditions bleiben.

Wenn du also eine neue SP-Liste mit Sonderzeichen im Namen anlegst und Items per `odata=verbose` hineinposten willst: **ImmerType-Namen mit `_x005f_`/`_x002d_` etc. encoden**. Falls unsicher: stattdessen `odata=nometadata` benutzen — dann ist gar kein `__metadata` nötig und SP leitet den Typ aus der URL ab.

### Key Architecture

- SPFx WebPart with React (no browser routing, context-based navigation)
- User authentication via SPFx WebPartContext (no manual login)
- UserContext provides current user data from SharePoint profile
- EventContext manages event data via SharePoint REST API
- Event creation creates a SharePoint subsite per event with a "Teilnehmer" list
- Item-Level Security on registration lists (users see only their own entries)
- DEX_Emails queue list for Power Automate email sending
  - v18.30: Spalte `Importance` (Single line text, leer/`Normal` = normal,
    `High` = Outlook hohe Wichtigkeit / rotes „!"). `queueEmail(..., importance?)`
    schreibt sie; idempotent via `ensureImportanceFieldExists` in
    `ensureEmailsList`. Der `DEX_SEND_MAIL`-Flow muss den Wert per Expression
    auslesen (siehe `docs/flow-jsons.md`, UI-Anleitung v18.30) — bis dahin
    gehen alle Mails normal raus. Genutzt von der DEX-Anfrage-Mail
    (`sendAdminInquiry`).
- 3 roles: User, Organizer, Admin (stored in DEX_Roles list)

### Roles & Permissions

**Event-Verwaltung:**

| Feature | User | Organizer | Admin |
|---------|------|-----------|-------|
| Events ansehen (eigener Standort) | ✅ | ✅ | ✅ |
| Alle Events ansehen | ❌ | ✅ | ✅ |
| Events erstellen | ❌ | ✅ | ✅ |
| Events bearbeiten | ❌ | ✅ (eigene) | ✅ (alle) |
| Events löschen | ❌ | ✅ (eigene) | ✅ (alle) |
| Event-Bild hochladen (als Item-Attachment) | ❌ | ✅ (eigene) | ✅ (alle) |
| Event-Dokumente hochladen | ❌ | ✅ (eigene) | ✅ (alle) |
| Agenda / Transferzeiten / Quiz pflegen | ❌ | ✅ (eigene) | ✅ (alle) |
| E-Mail-Templates pro Event anpassen | ❌ | ✅ (eigene) | ✅ (alle) |
| Benachrichtigungen pro Event deaktivieren (E-Mail / Outlook) | ❌ | ✅ (eigene) | ✅ (alle) |

**Registrierung:**

| Feature | User | Organizer | Admin |
|---------|------|-----------|-------|
| Selbst registrieren | ✅ | ✅ | ✅ |
| Für andere registrieren | ❌ | ✅ | ✅ |
| Eigene Angaben bearbeiten | ✅ | ✅ | ✅ |
| Eigene Registrierung stornieren | ✅ | ✅ | ✅ |

**Teilnehmerverwaltung (Admin Center):**

| Feature | User | Organizer | Admin |
|---------|------|-----------|-------|
| Teilnehmerliste sehen | ❌ | ✅ (eigene Events) | ✅ (alle) |
| Teilnehmer suchen / sortieren | ❌ | ✅ | ✅ |
| Teilnehmer ein- / auschecken | ❌ | ✅ (eigene) | ✅ (alle) |
| Teilnehmer abmelden | ❌ | ✅ (eigene) | ✅ (alle) |
| QR-Codes versenden | ❌ | ✅ (eigene) | ✅ (alle) |
| E-Mail-Adressen kopieren | ❌ | ✅ (eigene) | ✅ |
| Massenmail an Teilnehmer (RichText-Editor + Deloitte-Template) | ❌ | ✅ (eigene) | ✅ (alle) |
| Einladungsmail mit Anmelde-Link (an mich oder an Mailverteiler — Standort-/All-Verteiler hart geblockt, siehe v11.41) | ❌ | ✅ (eigene, ab v11.40) | ✅ (alle) |
| **IDs neu vergeben** (sequentielle Renummerierung) | ❌ | ✅ (eigene, ab v11.36) | ✅ |
| **Spalten fixen** (fehlende Felder + View-Reihenfolge) | ❌ | ❌ | ✅ |

**Administration:**

| Feature | User | Organizer | Admin |
|---------|------|-----------|-------|
| Rollen verwalten | ❌ | ❌ | ✅ |
| Rollen-Matrix einsehen | ❌ | ❌ | ✅ |
| Globale Email-Templates bearbeiten | ❌ | ❌ | ✅ |
| Logos / Default-Bilder verwalten | ❌ | ❌ | ✅ |

**SharePoint-Listen / Berechtigungs-Ebene:**

| Liste | User | Organizer | Admin |
|-------|------|-----------|-------|
| DEX_Events: Schreiben | ❌ | ✅ | ✅ |
| DEX_Events: Item-Attachments (Bild + Dokumente) | ❌ | ✅ (eigene) | ✅ |
| DEX_Roles: Schreiben | ❌ | ❌ | ✅ |
| DEX_Emails: Schreiben (Queue) | Nur eigene | Nur eigene | ✅ |
| DEX_Outlook: Schreiben (Queue) | Nur eigene | Nur eigene | ✅ |
| DEX_IDReorder: Schreiben | ❌ | ❌ | ✅ |
| DEX_EmailTemplates: Schreiben | ❌ | ❌ | ✅ |
| Event-Subsites: Full Control | ❌ | Eigene Events | ✅ |
| Teilnehmerliste (Subsite): Lesen | Nur eigener Eintrag | ✅ (eigene Events) | ✅ |
| Teilnehmerliste (Subsite): Schreiben | Nur eigener Eintrag | ✅ (eigene Events) | ✅ |

### Teilnehmerlisten-Schema (Subsites)

Spaltenreihenfolge in der Default View: **TeilnehmerID > Anrede > Vorname > Nachname > ParticipantEmail > Department > Location > JobTitle > Phone > Status > RegistrationDate > CancellationDate > [Custom Fields]**

Wenn Aenderungen am Teilnehmerlisten-Schema noetig sind (neue Spalten, Reihenfolge aendern):
1. `createRegistrationList()` in EventService.ts anpassen (fuer neue Events)
2. `fixRegistrationListColumns()` in EventService.ts anpassen (fuer bestehende Events)
3. Der Admin kann im Admin Center per Button **"Spalten fixen"** bestehende Events nachtraeglich aktualisieren
4. Dieser Button legt fehlende Spalten an und setzt die View-Reihenfolge korrekt

**v11.36 — Spalte `OverbookReview`** (Single line text): Marker für die
Überbuchungs-Bereinigung. `''` = normal, `'Pending'` = vom Button
**"Überbuchung prüfen"** als über Kapazität erkannt (wird NICHT in die
Default-View aufgenommen — die UI surfacet die Markierung über die Box
"Überbuchung – zu prüfen" oben in der Teilnehmerliste). Wird in
`createRegistrationList()` + `fixRegistrationListColumns()` angelegt.

### v11.36 — Überbuchungs-Schutz (atomar) + Bereinigungs-Tool

**Ursache Überbuchung:** Die alte Anmelde-Kapazitätsprüfung war
read-then-write (TOCTOU-Race) **und** fail-open (`catch` ließ den Status
auf `Angemeldet` bei SP-Throttling während Anmeldewellen) — Massen-
Überbuchung bei vielen zeitgleichen Anmeldungen. `MaxParticipants=0` bei
Gruppen-Events machte den Fallback zahnlos.

**Atomare Prävention:** `reserveSeat()` in EventService inkrementiert pro
Gruppe einen Sitzplatz-Zähler auf `DEX_TeilnehmerCounter` per **ETag-CAS
(IF-MATCH, 412-Retry)** — exakt das Muster von `getNextTeilnehmerId`. Neue
Counter-Felder: `SeatsTaken`, `SeatsTakenDurch`, `SeatsTakenFun` (Number,
in `ensureCounterListField()` angelegt). `EventContext.registerForEvent`
ruft `reserveSeat` → `'reserved'` = Angemeldet, `'full'`/`'error'` =
**fail-closed Warteliste** (nie optimistisch Angemeldet). `cancelRegistration`
+ Reorder + Detect rufen `syncSeatsToActiveCount()` (Reconcile aus echtem
Bestand; errt safe Richtung „voll", weil der Power-Automate-Nachrück-Flow
den Counter nicht anfasst). Rest-Edgecases (≤1 unter Extremlast) fängt das
Tool ab.

**Bereinigungs-Tool (Admin ODER Organizer eigener Events):** Button
**"Überbuchung prüfen"** → `detectOverbooking()` markiert pro Gruppe die
zuletzt über Kapazität Angemeldeten (`OverbookReview='Pending'`), **ohne
Status-Änderung**. Box oben in der Teilnehmerliste zeigt sie mit Buttons
pro Person + Sammel-„Alle bestätigen". Pro Person:
- **Bestätigen** → `resolveOverbookToWaitlist()` (Status=Warteliste,
  gruppentreu, ChangeLog-Audit „war fälschlich angemeldet, Original-
  Registrierung"). Modal: mit/ohne Mail (Deloitte-Wrap-Vorschlagstext,
  editierbar, **in die Queue**, nicht direkt) + in **beiden** Pfaden die
  Kalender-Abmelde-Frage (`queueOutlookEvent(...,'Ausladen')`).
- **Platz behalten** → Variante (a) `resolveOverbookKeepActive` (bleibt
  Angemeldet, Gruppe +1, Flow rückt einmal nicht nach — `Check_<Typ>_Free`
  strikt `<` absorbiert das) ODER (b) `resolveOverbookKeepAsFirstWaitlist`
  (Erste(r) auf Gruppen-Warteliste via RegistrationDate-Backdate, ChangeLog).
Nach jeder Aktion automatisch `reorderParticipantIDs()` (Aktive 1..N,
Warteliste N+1..) + `syncSeatsToActiveCount()` + Liste neu laden. Der
Reorder meldet Fortschritt via `onProgress`-Callback → %-Overlay in der UI.

**Kürzlich-abgemeldet-Hinweis (v11.36):** Die TeilnehmerIDs sind durch den
`DEX_IDReorder`-Flow **immer durchlaufend** — es gibt keinen „kaputt"-
Zustand zu erkennen. Beim Öffnen eines Events prüft `recentCancellation()`
nur, ob die jüngste `CancellationDate` (Status='Abgemeldet') **< 10 Min**
her ist. Falls ja, erscheint einmalig pro Event-Öffnung ein Hinweis-Modal:
die automatische Batch-Korrektur (Nachrücken + ID-Neuvergabe per Flow) läuft
evtl. noch → **ein paar Minuten warten**, NICHT parallel manuell „IDs neu
vergeben". Das Modal hat „Verstanden" (primär) + „Trotzdem jetzt
korrigieren" (sekundär, ruft `runIdReorder`). „IDs neu vergeben" ist seit
v11.36 für **Admin ODER Organizer eigener Events** freigeschaltet
(Teilnehmerverwaltung), mit %-Overlay.

### Icons / Design

**IMPORTANT:** KEINE Emojis im UI verwenden. Stattdessen ausschließlich **Fluent UI Icons** (modern, einfarbig, SVG):

```tsx
import { Icon } from '@fluentui/react/lib/Icon';
<Icon iconName="Calendar" style={{ fontSize: 16, color: 'var(--dex-gray-500)' }} />
```

- Nutze `@fluentui/react/lib/Icon` mit `iconName` (z.B. `Calendar`, `MapPin`, `People`, `Document`)
- Für Icon-Auswahl in der UI: `IconPicker` aus `@pnp/spfx-controls-react`
- Alle Icons einfarbig, skalierbar, einheitliches Design
- Keine Emoji-Symbole (❌ 📍📅🚌📄), stattdessen Fluent UI Icons (✅ MapPin, Calendar, Bus, Page)

### Wizard-Schritt-Nummerierung (1-basiert in UI / 0-basiert in Logik)

**WICHTIG (ab v9.32, erweitert v11.80):** Der Event-Erstellungs-Wizard hat **8 Schritte** (vorher 7 — mit v11.80 ist „Team-Anmeldung" als neuer Schritt 4 eingefügt worden). In der UI und in jeglicher Kommunikation mit dem User (Tooltips, Hilfetexte, Handbuch, Mail-Texte, Commit-Messages) sprechen wir **immer 1-basiert**:

- Schritt 1 = **Grundlagen** (Entwurf, Title, Datum, Beschreibung, Bild, Organizer, Test-Team, Check-In Team)
- Schritt 2 = **Ort & Programm** (Veranstaltungsort, Adresse, Agenda, Transferzeiten)
- Schritt 3 = **Kapazität & Sichtbarkeit** (Standortfilter, Mailverteiler, Filterverknüpfung, Deadlines, Teilnehmerzahl & Warteliste)
- Schritt 4 = **Team-Anmeldung** (v11.80 + erweitert v11.81 — Basis-Settings: Toggle „Team-Anmeldung erlauben", Team-Größe, Toggle „Team-Namen abfragen". Beitritts-Modus (v11.81, Sub-Box): Radio „Nur komplette Teams" vs. „Auch Teil-Teams erlaubt" (Default: komplette Teams), Checkbox „Unvollständige Teams öffentlich für Beitritt sichtbar" (Default aus — wenn aktiv sehen andere Teilnehmer offene Slots als „Team mit X freien Plätzen" ohne Namen der bereits angemeldeten Mitglieder), Checkbox „Beitritt erfordert Bestätigung durch Team-Kapitän" (Default aus, nur aktivierbar wenn vorige Option an — wenn aktiv landen Beitritte in einer Approve-Queue beim Team-Lead). Konfiguration wird persistiert; die tatsächliche Multi-Person-Anmelde-Logik folgt mit v11.82+. Persistierte SP-Spalten: `TeamRegistrationEnabled`, `TeamSize`, `AskTeamName`, `TeamPartialAllowed`, `TeamOpenSlotsVisible`, `TeamJoinRequiresApproval`.)
- Schritt 5 = **Felder** (Template, eigene Abfragen, ab v11.80 Toggle „Anrede abfragen?")
- Schritt 6 = **Kommunikation** (Mail-Sprache, Versand-Schalter, Organizer-BCC, Logos, Templates, Sub-Events)
- Schritt 7 = **Dokumente** (PDFs für Teilnehmer)
- Schritt 8 = **Fun-Zone** (Quiz)

In der React-Logik bleibt `currentStep` weiterhin **0-basiert** (`currentStep === 0` ist Grundlagen) — das ist ein Implementierungs-Detail. Wenn du in einer UI-Erklärung oder einem Tooltip „Schritt X" schreibst, immer **1-basiert** angeben (also „Schritt 6 (Kommunikation)" statt „Schritt 5 (Kommunikation)").

Jeder Step rendert oben eine eigene **Überschrift in Dunkelgrün** (`var(--dex-green-dark, #4a7c1f)`) im Format `Schritt N — Name` mit einem kurzen ein-Satz-Lead darunter, was in diesem Schritt eingestellt wird.

### Demo-Daten laden im Wizard (v11.88)

In Schritt 1 (Grundlagen) sitzt oben rechts ein „Demo"-Button. Vor v11.88 hat
er das Formular direkt mit einer fixen Test-Vorlage gefüllt — seit v11.88
öffnet er stattdessen ein Auswahl-Modal mit vier Varianten-Karten. Klick auf
eine Karte schließt das Modal sofort, füllt das Formular vollständig (inkl.
Reset von Team-, Split- und Sub-Event-Feldern) und springt auf Schritt 1
zurück. Keine Submit-Buttons — die Karte selbst ist der Submit.

Die vier Varianten in `EventCreationPage.tsx` (Map `DEMO_VARIANTS`):

| Key            | Vorlage                                | Kerneigenschaften                                                                                                  |
|----------------|----------------------------------------|--------------------------------------------------------------------------------------------------------------------|
| `standard`     | Einfaches Meeting / Lunch              | 50 Plätze, keine Gruppen, keine Sub-Events, Custom-Field „Essenspräferenz" (select, Pflicht)                       |
| `groups`       | Workshop mit zwei Gruppen              | Split Capacity 25 + 25, Gruppen „Vormittag" + „Nachmittag", gemeinsame Warteliste, keine Sub-Events                |
| `subevent`     | Conference mit Networking-Dinner       | 100 Plätze, ein Sub-Event („Networking-Dinner", 60 Plätze), Custom-Field „Hotel-Buchung" (select, optional)        |
| `subeventTeam` | Quizabend mit Team-Anmeldung           | 80 Plätze (= 20 Teams à 4), Team-Anmeldung aktiv (TeamSize 4, Team-Name abfragen, Teil-Teams + offene Slots), ein Sub-Event („Vorbereitungs-Briefing", 10 Plätze), Custom-Field „Essenspräferenz" |

Der Reset-Helper `resetDemoVariantBaseState()` setzt vor jedem Variant-Laden
alle Variant-spezifischen Felder (Split, Team, Sub-Events, Custom-Fields,
Agenda, Transferzeiten, Audience, Bild, Ansprechpartner) auf neutrale
Defaults zurück, damit Reste vorheriger Varianten nicht im Formular hängen
bleiben. Die alte `fillDemo`-Funktion mit der `Test_<datum>`-Vorlage bleibt
als interner Helfer im Code, wird aber von keinem Button mehr direkt
angesprochen.

### Tooltip-Stil (ab v9.32)

InfoTooltips erklären jedem Admin/Organizer **ausführlich und in Klartext**:

1. **Was du hier einstellst** — die Bedeutung des Felds in plain language
2. **Anzeige in der App** — wo erscheint der Wert, wer sieht ihn (vorher „Auswirkung in der App" — zu negativ formuliert)
3. **Automatismen** *(nur falls es welche gibt)* — welche Mails / Outlook-Termine reagieren automatisch (vorher „Auswirkung in Automatik")
4. **Auswirkung für Teilnehmer** — was bedeutet das konkret für reguläre User, mit Beispielen wie „können sich ab dem Datum nicht mehr selbst anmelden, aber Organizer und Co-Organizer schon"

Dabei werden **fette Schlagwörter** (`<strong>`) für die wichtigsten Begriffe und Zustände gesetzt, sodass das Auge beim Überfliegen die Kerninfo sofort findet. Tooltips dürfen lang sein — die Komponente ist bereits auf max. 480px Breite ausgelegt und scrollt nicht.

**Kein Tech-Jargon in Tooltips:** keine Erwähnungen von „Power Automate", „Flow", „DEX_IDReorder_TeilnehmerIDs", „SharePoint-List", „REST-API" o.ä. — Organizer und Admins sind keine Engineers und müssen die Implementierung nicht kennen. Stattdessen den Effekt beschreiben („wird automatisch nachgerückt", „bekommt eine Bestätigungs-Mail"), nicht den Mechanismus. Begriffe wie „FIFO" entweder ausschreiben („first-in, first-out") oder ganz weglassen.

Die `InfoTooltip`-Komponente nimmt ab v9.32 `text: React.ReactNode` an (vorher nur `string`) — bei JSX-Tooltips bilingual als `{isDe ? <>...</> : <>...</>}` rendern. Bei einfachen Texten ist weiterhin ein String erlaubt.

### German Text / Sonderzeichen — ALWAYS

**IMPORTANT (strikt ab v6.3.0, verschärft v10.25):** Alle deutschen Texte —
**ohne Ausnahme** — MÜSSEN echte Sonderzeichen verwenden (ä, ö, ü, ß, Ä, Ö,
Ü, „doppelte Anführungszeichen" `„…"`). **Keine ASCII-Substitutionen** (ae,
oe, ue, ss). Alle Dateien sind UTF-8, Encoding-Probleme sind daher kein
Argument.

Das gilt für **jeden Kontext**, in dem deutscher Text steht — diese Liste ist
nicht erschöpfend, sondern Beispiele:

- Sichtbare UI-Strings (JSX-Text, Buttons, Labels, Tooltips, **Hinweisboxen**,
  Modals, Banner, Toasts).
- Tooltip- und Hilfe-Texte in `<InfoTooltip text={…}>` — auch wenn sie als
  React-Fragment `<>…</>` strukturiert sind.
- Strings in Translation-Maps (`LanguageContext.tsx`).
- HTML-Strings in Mail-Bodies, Outlook-Termin-Bodies (`EmailTemplates.ts`).
- Alert-/Console-/Window-Confirm-Texten.
- Sub-Event-Placeholder, Default-Optionen in Catalogs etc.
- Code-Kommentare, JSDoc, Inline-`/* … */`-Kommentare.
- Commit-Messages.
- Handbuch-Einträge (`src/webparts/dexEventPlatform/components/manual/sections/*`).
- Markdown-Dokumentation (`docs/*.md`, `CLAUDE.md`).
- HTML-entitiy-escaped Anführungszeichen (`&bdquo;…&ldquo;`) sind OK, aber
  Umlaute selber **nicht** als `&auml;`/`&ouml;` etc. — direkt `ä`/`ö`
  schreiben.

**Examples:**
- ✅ `Löschen`, `öffnen`, `Übersicht`, `ausfüllen`, `hinzufügen`, `Zurück`,
  `für`, `Grüße`, `Bestätigung`, `zurückziehen`, `nächste`, `zusätzlich`,
  `Änderung`, `aufräumen`, `nötig`, `möchtest`, `Läufer`
- ❌ `Loeschen`, `oeffnen`, `Uebersicht`, `ausfuellen`, `hinzufuegen`,
  `Zurueck`, `fuer`, `Gruesse`, `Bestaetigung`, `zurueckziehen`, `naechste`,
  `zusaetzlich`, `Aenderung`, `aufraeumen`, `noetig`, `moechtest`, `Laeufer`

**Selbst-Check vor jedem Save / Edit:** wenn du gerade einen deutschen Text
geschrieben hast, suche im fertigen String nach `ae`, `oe`, `ue`, `ss` —
prüfe jede Stelle, ob das Wort eigentlich Umlaute haben sollte. Ja? Dann
direkt umstellen, **nicht** auf den Build warten.

Bei neuen Commits gilt das als blockierendes Review-Kriterium. Bestehende
ASCII-Substitutionen aus der Zeit vor v6.3.0 werden nicht flächendeckend
refactored (Scope-Eingrenzung), aber **jede berührte Datei** wird im Rahmen
der Änderung mit sauberem Deutsch versehen.

### Power Automate Flow Anleitungen

**WICHTIG:** Der User kann in Power Automate **KEINEN Code View** öffnen oder JSON direkt einfügen. Alle Anleitungen für Power Automate Flows müssen **ausschließlich über die UI** erklärt werden:

- **Expressions:** Immer über den **Expression-Tab (fx)** eingeben, nie als Text
- **Actions konfigurieren:** Immer über die **Parameter-Ansicht** (Dropdowns, Eingabefelder)
- **Conditions:** Werte über den **Expression-Tab (fx)** eingeben, nicht als String tippen
- **Run After:** Über die **Settings** Tab → Run after konfigurieren
- **Rename:** Über die **drei Punkte (⋮)** → Rename

Nie sagen "füge dieses JSON ein" oder "öffne den Code View" — stattdessen jeden Klick in der UI beschreiben.

### Power Automate Flow-Änderungen Workflow

**WICHTIG:** Wenn eine Änderung in einem Power Automate Flow nötig ist:

1. **Beschreibe die Änderung** als UI-Anleitung (siehe oben)
2. **Warte auf Bestätigung** vom User, dass die Änderung durchgeführt wurde
3. **Fordere den aktuellen Flow-JSON** vom User an (Code View → kopieren)
4. **Aktualisiere `docs/flow-jsons.md`** mit dem neuen JSON

Die Datei `docs/flow-jsons.md` enthält die vollständigen Flow-Definitionen aller 6 DEX-Flows:
- DEX_IDReorder_TeilnehmerIDs — TeilnehmerIDs renummerieren + Warteliste nachrücken
- DEX_SEND_MAIL — Mail-Versand aus DEX_Emails-Queue
- DEX_CreateOutlookEvent — Outlook-Termin initial anlegen
  - **Betreff im Termin (v18.42):** Neue Spalte `OutlookSubject` (Single line
    text) auf `DEX_Events` — eigener Betreff des Outlook-Termins, im
    Outlook-Body-Editor (Schritt 6) bearbeitbar (leer = Event-Titel).
    Per-Tab gespiegelt wie `outlookHeading` (Top + Sub-Events), persistiert in
    Create-/Update-/Sub-Payload. Der Flow mappt `item/subject` ←
    `coalesce(OutlookSubject, Title)` (Create + UpdateEvent-`Build_Update_Body`,
    siehe `docs/flow-jsons.md`, UI-Anleitung v18.42). Eine reine Betreff-
    Änderung gilt als Outlook-relevant (öffnet das Update-Modal). Der Editor
    zeigt zusätzlich **Termin + Ort read-only** („wird übernommen").
  - **Abweichendes Datum (v18.44):** Neue Spalten `OutlookStart` / `OutlookEnd`
    (DateTime) auf `DEX_Events` — optional abweichende Termin-Zeit (leer =
    Event-`StartDate`/`EndDate`). Im Outlook-Body-Editor (Schritt 6) mit
    **denselben DatePickern wie der Wizard** bearbeitbar (Parent rendert den
    Picker-Node, Modal nimmt `outlookDateEditor`-Prop). Per-Tab (Top +
    Sub-Events), als ISO gespeichert; Override-Werte werden direkt im aktiven
    Tab gelesen/geschrieben (kein Comm-Mirror nötig). Flow: `item/start` ←
    `coalesce(OutlookStart, StartDate)`, `item/end` analog (Create + Update,
    siehe `docs/flow-jsons.md` v18.44). Override-Änderung gilt als Termin-
    Änderung (öffnet Update-Modal). Der **Ort** (`OutlookLocation`) ist seit
    v18.44 ebenfalls pro Sub-Event im Reiter „Ort & Programm" UND im
    Outlook-Editor überschreibbar (leer = Auto aus Veranstaltungsort + Adresse).
  - **Ort im Termin (v18.34):** Neue Spalte `OutlookLocation` (Single line text)
    auf `DEX_Events` — lesbarer Ort (Veranstaltungsort + Adresse), gebaut von
    `utils/eventFormat.ts → buildOutlookLocation()`. Geschrieben beim Anlegen
    (`createEvent`-Payload) UND beim Bearbeiten (Wizard-`updates`), plus
    Backfill in `EventService.queueOutlookEvent` (einmal pro Event/Session) für
    Bestands-Events. Der Flow mappt `item/location` ←
    `triggerBody()?['OutlookLocation']` (siehe `docs/flow-jsons.md`, UI-Anleitung
    v18.34); der `DEX_Outlook_Einladungen`-UpdateEvent-Zweig patcht den Ort beim
    Aktualisieren mit. Eine reine Ort-Änderung im Wizard gilt seit v18.34 als
    Outlook-relevant (öffnet das Update-Modal).
  - **Trigger-Konsequenz (wichtig):** Der Flow lauscht ausschließlich auf
    `GetOnNewItems` in DEX_Events — er feuert **nur** beim Anlegen eines
    neuen DEX_Events-Items, **nicht** bei MERGE/PATCH-Updates. Wenn der
    User auf einem bestehenden Sub-Event (oder Top-Level-Event) nachträglich
    `DisableOutlook` von `true` auf `false` umstellt, würde der Flow den
    Trigger nie sehen — es entstünde nie ein Outlook-Termin.
  - **Lösung in der SPFx-App (v11.69, non-destructive):** Sub-Event in
    diesem Fall delete-and-recreate, ABER ohne die Teilnehmer-Subsite
    mitzulöschen. `EventService.deleteEventItemOnly(eventId)` entfernt
    ausschließlich das DEX_Events-Item via REST DELETE — KEIN Cascade auf
    Subsite, KEIN Outlook-DeleteEvent-Queue, KEIN EventImage-Recycle.
    Direkt anschließend wird `EventService.createEvent({ ...,
    existingSubsiteUrl, existingRegistrationListName })` aufgerufen — wenn
    BEIDE Felder gesetzt sind, überspringt `createEvent` das Anlegen
    einer neuen Subsite und koppelt stattdessen die mitgegebene Subsite +
    Teilnehmerliste an die neue DEX_Events-Zeile. Resultat: die alten
    Teilnehmer-Anmeldungen, TeilnehmerIDs und Custom-Field-Antworten
    bleiben unangetastet erhalten, das neue DEX_Events-Item triggert den
    DEX_CreateOutlookEvent-Flow (GetOnNewItems) → Outlook-Termin entsteht.
    Pfad in `persistSubEventsForParent` (`EventCreationPage.tsx`);
    Trigger sowohl per Outlook-Update-Modal
    (`pendingOutlookRecreateForSubEventsRef`, noOutlookYet-Checkbox) als
    auch beim Legacy-DisableOutlook-Toggle (true → false ohne
    OutlookEventId).
  - **Garantien (v11.69):** `deleteEventItemOnly` MUSS strikt
    non-cascading bleiben (kein `recycle()` auf der Subsite, kein
    DeleteEvent-Queue-Eintrag, kein DEX_Participants-Cleanup), und
    `createEvent({ existingSubsiteUrl, existingRegistrationListName })`
    MUSS strikt nicht-kreativ bleiben (kein neuer
    `createEventSubsite`-Call, kein neuer `createRegistrationList`-Call).
    Wenn eine dieser Garantien bricht, gehen Teilnehmer-Daten verloren.
- DEX_Outlook_Einladungen — Teilnehmer hinzufügen/entfernen, Termin aktualisieren/löschen
- DEX_OutlookDeclineHandler — Decline-Mails abfangen → Reminder-Mail queuen (inkl. weitergeleitete Declines FW:/WG:)
- DEX_OutlookForwardHandler — Meeting-Forward-Notifications abfangen → FYI-Mail an Organizer wenn weitergeleiteter Empfänger nicht registriert

Diese Datei **MUSS immer aktuell** gehalten werden wenn Flows geändert werden. Sie dient als einzige Referenz für den aktuellen Stand der Flows.

### Outlook-Flow-Parallelität — Option B Pro-Event-Lock (v18.48)

Der `DEX_Outlook_Einladungen`-Flow lief mit **Concurrency 1** streng seriell.
Bei Grossevents patcht jeder Lauf die **komplette** Teilnehmerliste (bis 1500
Personen) an Graph → 13–40 Min pro Lauf. Anmeldungen für **andere** Events
stauten sich dahinter. Mit v18.48 wird die Trigger-Concurrency erhöht
(im Tenant auf **100** gesetzt; Parallelität über verschiedene Events),
abgesichert durch einen **Pro-Event-Lock**, damit zwei Läufe für **dasselbe**
Event nicht gleichzeitig die Attendee-Liste lesen-und-schreiben (Race →
verlorene Einträge). Stand 2026-06-02: im Tenant umgesetzt und verifiziert.

- **Neue SP-Liste `DEX_OutlookLocks`** (Site-Collection-Root, angelegt durch
  `EventService.ensureOutlookLocksList()` in `initEvents`). Spalten: `EventId`
  (Single line text, **indiziert + „Eindeutige Werte erzwingen"**) als
  Lock-Schlüssel, `LockedAt` (DateTime, informativ). Schreibrechte via
  `setQueueListPermissions` (analog DEX_Emails).
- **Lock-Erwerb (im Flow):** Create-Item mit `EventId` → gelingt = Lock
  erworben; schlägt fehl (Eindeutigkeit, weil ein anderer Lauf desselben Events
  schon ein Item hält) = kurz warten + retry (Do-until). **Lock-Release:**
  Delete-Item am Ende, mit Run-After **succeeded + failed + skipped + timed out**,
  damit der Lock auch bei Abbruch immer freigegeben wird.
- **Reiner Flow-Umbau** (keine App-Logik außer der Listen-Provisionierung) —
  die vollständige UI-Schritt-für-Schritt-Anleitung steht in
  `docs/flow-jsons.md` unter „UI-Anleitung 2026-06-02 (v18.48) — Option B:
  Pro-Event-Lock für parallele Outlook-Läufe".

### Pflicht-Mitlaufende Artefakte bei App-Updates

**WICHTIG (ab v6.9):** Bei jedem neuen Feature oder größerem Umbau müssen folgende Artefakte synchron aktualisiert werden, bevor ein Release gebaut wird. Das ist Review-Kriterium, nicht optional:

| Artefakt | Pfad | Pflicht-Inhalt |
|---------|------|---------------|
| **Handbuch-Sektion** | `src/webparts/dexEventPlatform/components/manual/sections/*.tsx` | Neue Sektion pro Feature ODER bestehende Sektion um Schritte erweitern. Immer mit Prosa-Beschreibungen "was macht der Nutzer & warum", zweisprachig DE/EN. Registrieren in `handbookContent.ts`. |
| **Rollenmatrix** | `src/webparts/dexEventPlatform/components/RoleMatrixPage.tsx` | Neue User-/Organizer-/Admin-Fähigkeit als zusätzliche Zeile in das Matrix-Array einfügen — inkl. Description, die erklärt wofür das Feature ist. Zeigt dem Admin auf einen Blick, wer was darf. |
| **Prozess-Übersicht (Flowcharts)** | `src/webparts/dexEventPlatform/components/FlowchartPage.tsx` | Betroffene Flowcharts (RegistrationFlow, CancellationFlow, IDReorderFlow, EventCreationFlow etc.) müssen den neuen Ablauf widerspiegeln. Seit v6.9 mit Prosa-`details`-Feld pro FlowNode — dort jedes wichtige Step erklären, was passiert und warum. Neue Flows brauchen eigene Render-Function. |
| **docs/flow-jsons.md** | `docs/flow-jsons.md` | Wenn sich ein Power-Automate-Flow ändert (neue Action, geänderte Expression, neue Condition): Section entsprechend aktualisieren mit Struktur-Diagramm und Begründung. |
| **CLAUDE.md** | `CLAUDE.md` | Neue Konventionen, neue SharePoint-Listen, neue Schema-Felder: hier ergänzen, damit die Projekt-Regeln aktuell bleiben. |

Beim Review eines Feature-Commits wird geprüft: **Handbuch ✓, Rollenmatrix ✓, Flowchart ✓, ggf. Flow-JSONs ✓.** Fehlt einer, blockiert das den Release.

### Demo-Impersonation für Admins (v12.7, v13.7)

Admins können testweise als generischer User aus einem gewählten
Standort durch die App navigieren. Eintrag im Header-User-Menü (Foto
oben rechts) unter „Rollenverwaltung": **„Demo: als User testen"**.

Mechanik:
- **v13.7:** kein People-Picker mehr — nur ein Standort-Dropdown
  (die feste Standortliste aus dem Event-Wizard, siehe
  `EventCreationPage.tsx` → `locationOptions`). Der Demo-User ist
  immer ein synthetischer „Demo User" mit `demo.user@deloitte.de` und
  dem gewählten Standort. Hintergrund: für UI- und Filter-Tests reicht
  es, die Standortfilter-Logik durchzuspielen — eine echte Identität
  („Mails/Outlook als XY") lässt sich darüber ohnehin nicht testen,
  weil das Mail-/Outlook-Routing serverseitig läuft.
- Aktivierung speichert das Tripel `{email, firstName, surname, location}`
  in `localStorage` unter dem Key `dex_demo_impersonation`. Die App
  re-loaded danach automatisch.
- Beim Boot lesen `RoleContext` und `UserContext` den localStorage-Key:
  - `RoleContext`: `effectiveRole = 'User'`, `isAdmin = false`,
    `canCreateEvents = false`. Aber `originalIsAdmin = true` bleibt
    erhalten — das nutzen `Header` und Admin-Pages, um zu erkennen
    dass es sich um einen echten Admin im Demo-Modus handelt
    (Demo-Menü-Eintrag bleibt sichtbar, kritische Pages bleiben
    zugänglich).
  - `UserContext`: `firstName / surname / email / location` werden mit
    den Impersonations-Werten überschrieben — das treibt sowohl die
    Anzeige als auch den Standortfilter in `EventListPage`.
- Oranger Sticky-Banner ganz oben (`DexEventPlatform.tsx →
  ImpersonationBanner`) zeigt den aktuellen Status und hat einen
  „Demo-Modus beenden"-Button. Klick löscht den localStorage-Key
  und reloaded.

**Permission-Guards die `originalIsAdmin` statt `isAdmin` prüfen
müssen** (sonst können Admins im Demo-Modus eigene Admin-Funktionen
nicht mehr testen):
- `SettingsPage`, `RoleMatrixPage`: frühe Navigation zurück zu „start"
  wenn `!originalIsAdmin` (v12.15).
- `Header`-Menü-Eintrag „Demo: als User testen": Sichtbarkeit
  via `originalIsAdmin` (vorhanden seit v12.7).

### Mail-Editor: Überschrift- + Unter-Überschrift-Formatierung (v18.19–v18.22)

Der `HtmlEditorModal` (E-Mail-Template-Editor im Wizard, Schritt 6) erlaubt
pro Event die freie Formatierung von Überschrift (h1) und Unter-Überschrift
(h2):

- **v18.19:** Überschrift Größe (`headingFontSize`) + Farbe (`headingColor`).
- **v18.20/v18.21:** Body-Schriftgröße — kontrolliertes Dropdown, zeigt die
  Größe der Auswahl (wie in Word) und setzt sie robust via
  `execCommand('fontSize','7')` → Umschreiben in `<span style="font-size:Npx">`
  (manuelles `range.extractContents()` warf bei Auswahlen über
  Element-/Zeilen-Grenzen eine verschluckte Exception → kein Span).
- **v18.22:** Unter-Überschrift frei formatierbar (`subheadingColor`,
  `subheadingFontSize`, `subheadingBold`, `subheadingItalic`), Überschrift
  zusätzlich fett/kursiv (`headingBold`, `headingItalic`), und **freie
  Farbwahl** (Swatches + nativer `<input type="color">` + Hex-Code-Eingabe)
  für Überschrift-, Unter-Überschrift- UND Body-Textfarbe (`ColorControl`-
  Komponente in `HtmlEditorModal.tsx`).

**Render-/Persistenz-Kette (wichtig bei neuen Heading-Properties):**

1. `EmailTemplates.ts` → `wrapTemplate()` / `wrapTemplateForStorage()` nehmen
   einen optionalen Trailing-Parameter `opts: WrapHeadingOpts` (NICHT
   weitere Positions-Parameter — sonst brechen die ~25 Bestands-Aufrufe).
   Die h1/h2-Zeilen baut der gemeinsame Helper `buildHeadingsHtml()`. Defaults
   bleiben unverändert (h1: 400/normal/headingColor; h2: 20px/700/normal/#000),
   damit alle Alt-Aufrufe gleich aussehen.
2. `buildEmailFromTemplate()` reicht die Felder als `opts` an `wrapTemplate`.
3. `EventContext.applyEventTemplateOverride()` löst Override > SP-Template auf
   und MUSS jedes neue Feld durchreichen (Resolver-Return-Type + Empty-Check
   + Spread-Block ergänzen).
4. `EventCreationPage.tsx`: Override-Form ist der zentrale Type-Alias
   `EmailOverrideEntry` (neue Felder NUR dort ergänzen). Alle Editor-Handler
   laufen über den `patchOverride()`-Helper (merged ein Teil-Update und
   bewahrt alle übrigen Override-Felder — vorher droppte z.B. ein
   Heading-Text-Edit die zuvor gesetzte Farbe). Persistenz passiert über die
   `EmailTemplateOverrides`-Spalte (JSON); Sub-Events tragen dieselben Felder
   im Piggyback-JSON.

Die Felder sind **Override-only** (kein globales SP-Template-Feld) — analog
zu `headingColor`/`headingFontSize`.

### Mail-Template-Architektur (v12.11–v13.0)

Alle automatischen App-Mails leben jetzt in `DEX_EmailTemplates` als
Tenant-anpassbare Vorlagen mit Platzhalter-Substitution. Vorher war
ein guter Teil davon ad-hoc inline in `EventContext.tsx` /
`EventService.ts` als HTML zusammengebaut.

**Vollständige Template-Liste (DE + EN je Template):**
- `Anmeldung`, `Warteliste`, `Abmeldung`, `Nachruecken`, `EventErstellt`
- `OutlookDeclineReminder`, `OutlookDeclineReminder_OnBehalfOf`,
  `OutlookForwardNotification`, `OutlookDeclineDigest`
- **Team (v12.13–v12.14):** `TeamMemberJoined`, `TeamJoinRequest`,
  `TeamJoinRejected`, `TeamLeadTransferred`, `TeamMemberCancelled`
- **v13.0:** `RoommateRequest`, `GroupSwitchConfirmed`,
  `GroupSwitchWaitlist`, `OverbookingApology`
- **v18.63:** `OrgNachruecker` — Organizer-Benachrichtigung bei Abmeldung mit
  Nachrücker. Pre-wrapped gespeichert (wie `Nachruecken`), wird vom
  `DEX_IDReorder`-Flow nach erfolgreichem Promote an die Organizer gequeued
  (nicht von der App). Platzhalter: `{{EventTitle}}`, `{{CancelledName}}`, `{{PromotedName}}`. Flow-
  UI-Anleitung in `docs/flow-jsons.md` (v18.63).

**Reseed-Button (v12.11):** Admin → Settings → Karte „Default-Mail-
Templates re-seed" → Confirm. Ruft
`eventService.reseedDefaultEmailTemplates()` auf, was alle Standard-
Einträge in DEX_EmailTemplates mit den im Code definierten Defaults
überschreibt. Notwendig nach App-Updates, die die Standard-Texte
verändert haben — z.B. v12.11 (Nachrücken-Mail), v12.13 (Team-Mails),
v13.0 (Roommate / GroupSwitch / OverbookingApology). Eigene Tenant-
Anpassungen an Subject / Heading / Body gehen dabei verloren.

**Pattern bei jedem `queueEmail`-Call:**
```ts
const tpl = await eventService.getEmailTemplate('TemplateType', lang);
if (tpl) {
  const { subject, body } = buildEmailFromTemplate(tpl, vars);
  await eventService.queueEmail(subject, recipient, name, body,
    'TemplateType', event.title, eventId);
} else {
  // Inline-Fallback für Tenants ohne geseedetes Template
  // ...
}
```

`templateType`-Parameter in `queueEmail` wird konsequent auf den
echten Template-Namen gesetzt (kein generisches „Info" mehr für
identifizierbare Mails) — erleichtert das Debugging in der
DEX_Emails-Queue.

### Warteliste-Nachrück-Sortierung (v12.10)

Vorher: erster Warteliste-Eintrag nach `RegistrationDate asc`.

Neu (v12.10): erster Warteliste-Eintrag nach `TeilnehmerID asc`.

Hintergrund: nach dem `DEX_IDReorder`-Flow sind die TIDs durchlaufend
(1..N aktiv, N+1..M Warteliste). Bei einer freien Position 100 soll
TID 101 nachrücken — unabhängig davon, ob z.B. TID 103 zeitlich
gesehen früher registriert war (Reaktivierung nach Cancel,
Group-Switch etc.).

Code: `EventService.promoteFirstWaitlistItem` nutzt jetzt
`$orderby=TeilnehmerID asc`. `AdminPage`-Warteliste-Anzeige sortiert
ebenfalls nach `TeilnehmerID`.

**Power-Automate-Flow `DEX_IDReorder_TeilnehmerIDs`:** die App-seitige
Promote-Logik läuft nur beim Admin-Cancel-Pfad. Der eigentliche
User-Self-Cancel-Nachrück-Pfad lebt im Power-Automate-Flow. Damit der
auch nach TeilnehmerID sortiert, muss in jedem Nachrück-Branch
(Standard, Durchstarter, Funstarter, Shared) die Order-By-Klausel auf
`TeilnehmerID asc` umgestellt werden. Anleitung folgt in
`docs/flow-jsons.md`.

### Shared `<Modal>`-Komponente (v13.1)

Vorher hatte jede Modal-Komponente (~17 Stück) das gleiche Backdrop-
+ Card-Layout selbst implementiert mit leicht unterschiedlichem
z-index, Padding, Backdrop-Opacity. Seit v13.1 lebt in
`components/Modal.tsx` ein wiederverwendbarer Wrapper:

```tsx
<Modal open={open} onClose={onClose} maxWidth={520}
       dismissable={!busy} ariaLabel="Demo-Modus">
  ...
</Modal>
```

Übernommen: `InquiryModal` und `ImpersonateModal`. Weitere Modals
können bei Touch auf den Wrapper umgestellt werden — der bestehende
Inline-Code funktioniert weiterhin, ist aber Migrationskandidat
beim nächsten gezielten Touch der Datei.

### People-Picker — Member-Firm-Filter (v13.6)

Alle People-Picker in der App filtern Standardmäßig auf die deutsche
Member-Firm (`@deloitte.de` — entspricht dem internen DEALL-Verteiler).
Hintergrund: die SP `ClientPeoplePickerSearchUser`-API matched
tenant-weit und liefert bei Suchen wie „Nils Felten" auch
False-Positives aus internationalen Member-Firms (z.B. „Agarwalla,
Nilesh" oder „Das, Niladri" mit @deloitte.com / @deloitte.in).

Mechanik:
- `SharePointService.searchUsers(query, includeInternational?)` —
  zweiter Parameter steuert den Filter. Default `false`. Bei `false`:
  nur `@deloitte.de`. Bei `true`: zusätzlich `@deloitte.com` (alle
  internationalen Member-Firms wie DEUS/DECH/DECEMEA mappen darauf).
  Andere Domains (Gast-Accounts, externe Tenants) bleiben in beiden
  Modi geblockt — die App ist ohnehin nur für DEALL freigegeben.
- `RoleContext.searchUsers(query, includeInternational?)` reicht den
  Parameter durch.
- `components/InternationalSearchToggle.tsx` ist ein kleines
  Checkbox-Component mit Label „Auch international suchen
  (@deloitte.com)" — wird in jedem Picker unter dem Such-Input
  gerendert. Jeder Picker hält seinen eigenen Toggle-State (kein
  globales Setting), damit z.B. der Admin im Impersonate-Modal
  international suchen kann, ohne dass der Register-for-Other-Picker
  auf der Detailseite mit umschaltet.
- Toggle-State-Wechsel triggert die laufende Query nochmal — siehe
  `React.useEffect([includeIntl])`-Pattern in `ImpersonateModal.tsx`.

Betroffen sind ausschließlich Name-Picker. `searchUsersByLocation`
(Graph-basierter Standort-Query für den Mailverteiler-Aufbau) bleibt
unangetastet, weil es ohnehin standortbasiert vorfiltert.

### EventCreationPage-Refactor — offene Arbeit

`EventCreationPage.tsx` ist mit ~10.700 Zeilen die mit Abstand
größte Datei. Sie umfasst 8 Wizard-Schritte plus ~30 Modals und
diverse Sub-Logiken. Refactor-Wunsch: jeden Wizard-Step in eine
eigene Datei `components/eventCreation/Step1Basics.tsx` etc.
aufteilen. **Aktuell ungeschnitten** — Risiko vs. Nutzen ist hoch,
weil viele States über Step-Grenzen geteilt werden. Wenn die Datei
in Zukunft erneut wachsen würde, hier ansetzen.

### Aktionen-Dropdown im Admin-Center (v12.7–v12.8)

Statt ~20 Action-Kacheln in einem 4-Spalten-Grid sammeln sich alle
Aktionen in einer Dropdown-Liste **direkt in der Event-Detail-Card**
unter „Aktuell registriert".

Mechanik:
- Neuer Context `ActionsRegistryContext` (in `AdminPage.tsx`) — jeder
  `<ActionTile>` registriert sich beim Mount mit `{key, title, desc,
  badge, onClick, href, disabled}` und rendert in diesem Modus
  `null`.
- `<ActionsDropdown>` liest die Registry, sortiert alphabetisch nach
  Titel und rendert eine Liste-Box. Hover-Tooltip links zeigt die
  ausführliche `desc` der gerade gehoverten Aktion.
- ActionTiles mit `children` (z.B. Excel-Export-Sub-Dropdown) werden
  weiterhin als Kachel sichtbar gerendert — sie brauchen die
  visible-UI für die Sub-Menüs.

Ergebnis: Detail-Ansicht aufgeräumter, alle Aktionen in einer
einzigen klar strukturierten Liste mit Hover-Erklärung. Layout
einspaltig (vorher 2-spaltig mit separater Aktionen-Card rechts).
