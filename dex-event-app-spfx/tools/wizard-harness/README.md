# Wizard-Harness — den Event-Wizard ohne SharePoint ansehen

Seit v31.2. Nutzer-Ansage 07.09.2026: „am Ende bitte auch alle Schritte im
Wizard als PNG selber angucken." SPFx hat keinen lokalen Workbench mehr; der
Harness bündelt `EventCreationPage` mit gemockten Providern und macht mit
Chromium je Schritt ein Bild.

```bash
cd dex-event-app-spfx/tools/wizard-harness
npm i --no-audit --no-fund esbuild@0.21.5 playwright@1.55.0   # einmalig, landet in tools/wizard-harness/node_modules
node build.js            # Bundle + flaches CSS → out/
node shot.js edit        # out/shots/edit-step-1.png … edit-step-10.png (Beispiel-Event mit drei Sub-Events)
node shot.js create      # out/shots/create-initial.png (Anlege-Modus: Nutzungsbedingungen)
```

Chromium: `shot.js` sucht unter `/opt/pw-browsers/…` (Remote-Umgebung); lokal
`npx playwright install chromium` oder den Pfad in `shot.js` anpassen.

Was der Harness ersetzt (siehe `entry.tsx`): `EventContext`, `RoleContext`,
`NavigationContext`, `UserContext` bekommen Proxy-Mocks (alle Funktionen sind
leere `async`-Stubs, `events` enthält ein Beispiel-Event mit Programm und
drei Sub-Events); `LanguageProvider` und `DialogProvider` laufen echt;
`@microsoft/*` wird durch einen rekursiven Proxy ersetzt; das SCSS-Modul
liefert Klassennamen unverändert. Der Dienst (`EventService`) wird nie
gebaut — `window.__dexSpfxContext` fehlt absichtlich.

Was der Harness NICHT zeigt: Speichern, Personensuche, Mail-Vorschau mit
echten Vorlagen, Bilder aus SharePoint. Es geht um Darstellung, Reihenfolge
und Texte — genau das, was der UI-Leitfaden (`docs/ui-leitfaden.md`) regelt.

`out/` und `node_modules/` sind per `.gitignore` ausgeschlossen.
