# CLAUDE.md - Project Instructions

## DEX Event Experience Platform (SPFx)

### Build & Deploy Workflow

When deploying changes to the SPFx project:

1. Run `cd dex-event-app-spfx && npm run package` to build and bundle
   - This automatically bumps the patch version (e.g. 1.0.37 → 1.0.38)
   - If `gulp` is not found, run `npm install` first, then use `npx gulp bundle --ship && npx gulp package-solution --ship`
   - **IMPORTANT:** Always use `npm run package` (not raw gulp commands) so the version gets bumped automatically
2. **Always** copy the built package to `dist/`:
   ```bash
   cp dex-event-app-spfx/sharepoint/solution/dex-event-platform.sppkg dist/dex-event-platform.sppkg
   ```
3. Commit and push both the source changes and the updated `dist/dex-event-platform.sppkg`

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

- Site URL: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-B2Run`
- Lists: DEX_Events, DEX_Roles
- Per-Event: Subsite with "Teilnehmer" registration list

### Key Architecture

- SPFx WebPart with React (no browser routing, context-based navigation)
- User authentication via SPFx WebPartContext (no manual login)
- UserContext provides current user data from SharePoint profile
- EventContext manages event data via SharePoint REST API
- Event creation creates a SharePoint subsite per event with a "Teilnehmer" list
- Item-Level Security on registration lists (users see only their own entries)
- 3 roles: User, EventAdmin, SuperAdmin (stored in DEX_Roles list)

### German Text / Sonderzeichen

**IMPORTANT:** All user-facing text in components MUST use proper German characters (ä, ö, ü, ß). Never use ASCII substitutions (ae, oe, ue, ss) in visible UI text. Code comments may use ASCII substitutions if needed.

Examples:
- ✅ `Löschen`, `öffnen`, `Übersicht`, `ausfüllen`, `hinzufügen`, `Zurück`, `für`
- ❌ `Loeschen`, `oeffnen`, `Uebersicht`, `ausfuellen`, `hinzufuegen`, `Zurueck`, `fuer`
