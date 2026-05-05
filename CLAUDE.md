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
| **IDs neu vergeben** (sequentielle Renummerierung) | ❌ | ❌ | ✅ |
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

**WICHTIG (ab v9.32):** Der Event-Erstellungs-Wizard hat 7 Schritte. In der UI und in jeglicher Kommunikation mit dem User (Tooltips, Hilfetexte, Handbuch, Mail-Texte, Commit-Messages) sprechen wir **immer 1-basiert**:

- Schritt 1 = **Grundlagen** (Entwurf, Title, Datum, Beschreibung, Bild, Organizer, Test-Team, Check-In Team)
- Schritt 2 = **Ort & Programm** (Veranstaltungsort, Adresse, Agenda, Transferzeiten)
- Schritt 3 = **Kapazität & Sichtbarkeit** (Standortfilter, Mailverteiler, Filterverknüpfung, Deadlines, Teilnehmerzahl & Warteliste)
- Schritt 4 = **Felder** (Template, eigene Abfragen)
- Schritt 5 = **Kommunikation** (Mail-Sprache, Versand-Schalter, Organizer-BCC, Logos, Templates, Sub-Events)
- Schritt 6 = **Dokumente** (PDFs für Teilnehmer)
- Schritt 7 = **Fun-Zone** (Quiz)

In der React-Logik bleibt `currentStep` weiterhin **0-basiert** (`currentStep === 0` ist Grundlagen) — das ist ein Implementierungs-Detail. Wenn du in einer UI-Erklärung oder einem Tooltip „Schritt X" schreibst, immer **1-basiert** angeben (also „Schritt 5 (Kommunikation)" statt „Schritt 4 (Kommunikation)").

Jeder Step rendert oben eine eigene **Überschrift in Dunkelgrün** (`var(--dex-green-dark, #4a7c1f)`) im Format `Schritt N — Name` mit einem kurzen ein-Satz-Lead darunter, was in diesem Schritt eingestellt wird.

### Tooltip-Stil (ab v9.32)

InfoTooltips erklären jedem Admin/Organizer **ausführlich und in Klartext**:

1. **Was du hier einstellst** — die Bedeutung des Felds in plain language
2. **Auswirkung in der App** — wo erscheint der Wert, wer sieht ihn
3. **Auswirkung in Automatik** — welche Mails / Outlook-Termine / Power-Automate-Flows reagieren
4. **Auswirkung für Teilnehmer** — wie wirkt sich die Einstellung auf den Teilnehmer-Flow aus

Dabei werden **fette Schlagwörter** (`<strong>`) für die wichtigsten Begriffe und Zustände gesetzt, sodass das Auge beim Überfliegen die Kerninfo sofort findet. Tooltips dürfen lang sein — die Komponente ist bereits auf max. 480px Breite ausgelegt und scrollt nicht.

**Kein Tech-Jargon in Tooltips:** keine Erwähnungen von „Power Automate", „Flow", „DEX_IDReorder_TeilnehmerIDs", „SharePoint-List", „REST-API" o.ä. — Organizer und Admins sind keine Engineers und müssen die Implementierung nicht kennen. Stattdessen den Effekt beschreiben („wird automatisch nachgerückt", „bekommt eine Bestätigungs-Mail"), nicht den Mechanismus. Begriffe wie „FIFO" entweder ausschreiben („first-in, first-out") oder ganz weglassen.

Die `InfoTooltip`-Komponente nimmt ab v9.32 `text: React.ReactNode` an (vorher nur `string`) — bei JSX-Tooltips bilingual als `{isDe ? <>...</> : <>...</>}` rendern. Bei einfachen Texten ist weiterhin ein String erlaubt.

### German Text / Sonderzeichen

**IMPORTANT (strikt ab v6.3.0):** Alle deutschen Texte — sowohl in der sichtbaren UI als auch in Code-Kommentaren, Mail-Bodies, Outlook-Termin-Bodies, Alert-/Console-Texten, Commit-Messages, Handbuch-Einträgen (`src/webparts/dexEventPlatform/components/manual/sections/*`) und `docs/*.md` — MÜSSEN echte Sonderzeichen verwenden (ä, ö, ü, ß, Ä, Ö, Ü). **Keine ASCII-Substitutionen** (ae, oe, ue, ss). Alle Dateien sind UTF-8, Encoding-Probleme sind daher kein Argument.

Examples:
- ✅ `Löschen`, `öffnen`, `Übersicht`, `ausfüllen`, `hinzufügen`, `Zurück`, `für`, `Grüße`, `Bestätigung`, `zurückziehen`, `nächste`, `zusätzlich`, `Änderung`, `aufräumen`, `nötig`
- ❌ `Loeschen`, `oeffnen`, `Uebersicht`, `ausfuellen`, `hinzufuegen`, `Zurueck`, `fuer`, `Gruesse`, `Bestaetigung`, `zurueckziehen`, `naechste`, `zusaetzlich`, `Aenderung`, `aufraeumen`, `noetig`

Bei neuen Commits gilt das als Review-Kriterium. Bestehende ASCII-Substitutionen aus der Zeit vor v6.3.0 werden nicht flächendeckend refactored (Scope-Eingrenzung), aber jede berührte Datei wird im Rahmen der Änderung mit sauberem Deutsch versehen.

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
- DEX_Outlook_Einladungen — Teilnehmer hinzufügen/entfernen, Termin aktualisieren/löschen
- DEX_OutlookDeclineHandler — Decline-Mails abfangen → Reminder-Mail queuen (inkl. weitergeleitete Declines FW:/WG:)
- DEX_OutlookForwardHandler — Meeting-Forward-Notifications abfangen → FYI-Mail an Organizer wenn weitergeleiteter Empfänger nicht registriert

Diese Datei **MUSS immer aktuell** gehalten werden wenn Flows geändert werden. Sie dient als einzige Referenz für den aktuellen Stand der Flows.

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
