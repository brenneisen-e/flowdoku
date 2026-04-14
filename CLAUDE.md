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

- **Patch (1.0.x):** Bug fixes, UI tweaks, small improvements. Bumped automatically by `npm run package`.
- **Minor (1.x.0):** New user-facing feature completed and working end-to-end (e.g. event creation with subsites, registration flow, role management). Bump manually in `package.json` and `config/package-solution.json`.
- **Major (x.0.0):** Breaking changes, major architectural shifts, or official production release. Bump manually.

When to bump minor/major (update `package.json` version + `config/package-solution.json` solution.version):
- **1.1.0** — First fully working version: event creation (subsites), registration, cancellation, admin page, role management all functional on SharePoint
- **1.2.0** — Power Automate integration (waitlist promotion, ID reassignment)
- **2.0.0** — Production release / go-live after pilot phase

### SharePoint Site

- Site URL: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform`
- Lists: DEX_Events, DEX_Roles, DEX_Emails, DEX_Outlook, DEX_IDReorder, DEX_Participants, DEX_EmailTemplates
- Per-Event: Subsite with "Teilnehmer" registration list
- Shared Mailbox: `no_reply.events@deloitte.de`

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

### German Text / Sonderzeichen

**IMPORTANT:** All user-facing text in components MUST use proper German characters (ä, ö, ü, ß). Never use ASCII substitutions (ae, oe, ue, ss) in visible UI text. Code comments may use ASCII substitutions if needed.

Examples:
- ✅ `Löschen`, `öffnen`, `Übersicht`, `ausfüllen`, `hinzufügen`, `Zurück`, `für`
- ❌ `Loeschen`, `oeffnen`, `Uebersicht`, `ausfuellen`, `hinzufuegen`, `Zurueck`, `fuer`

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

Die Datei `docs/flow-jsons.md` enthält die vollständigen Flow-Definitionen aller 4 DEX-Flows:
- DEX_IDReorder_TeilnehmerIDs
- DEX_SEND_MAIL
- DEX_CreateOutlookEvent
- DEX_Outlook_Einladungen

Diese Datei **MUSS immer aktuell** gehalten werden wenn Flows geändert werden. Sie dient als einzige Referenz für den aktuellen Stand der Flows.
