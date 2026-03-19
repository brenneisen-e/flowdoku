# CLAUDE.md - Project Instructions

## DEX Event Experience Platform (SPFx)

### Build & Deploy Workflow

When deploying changes to the SPFx project:

1. Run `cd dex-event-app-spfx && npm run package` to build and bundle
2. **Always** copy the built package to `dist/`:
   ```bash
   cp dex-event-app-spfx/sharepoint/solution/dex-event-platform.sppkg dist/dex-event-platform.sppkg
   ```
3. Commit and push both the source changes and the updated `dist/dex-event-platform.sppkg`

The `dist/dex-event-platform.sppkg` must always reflect the latest build so it can be downloaded directly from GitHub.

### SharePoint Site

- Site URL: `https://deudeloitte.sharepoint.com/sites/DOL-c-DE-B2Run`
- Lists: Events, Anmeldungen, Abmeldungen, Check-In

### Key Architecture

- SPFx WebPart with React (no browser routing, context-based navigation)
- User authentication via SPFx WebPartContext (no manual login)
- UserContext provides current user data from SharePoint profile
- EventContext manages event data (currently mock, to be connected to SP lists)
