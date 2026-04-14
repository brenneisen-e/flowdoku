# Migration Data Folder

Hier kommen **Teilnehmerlisten-CSVs** rein, die per Seed-Migration in DEX importiert werden sollen.

## Konvention

**Eine CSV pro Event.** Dateiname = Event-Titel (genauer: lowercase + Bindestriche statt Sonderzeichen):

| CSV-Datei | Event-Titel in DEX_Events |
|-----------|---------------------------|
| `e2e-ma-activation-session-munich.csv` | E2E M&A Activation Session Munich |
| `munich-plogging-city-clean-up.csv` | Munich Plogging - City Clean up |
| `koeln-city-clean-up.csv` | Köln City Clean-up |
| `berlin-plogging-city-clean-up.csv` | Berlin Plogging - City Clean up |
| `better-futures-day-stuttgart.csv` | Better Futures Day Stuttgart |

## CSV-Format

Erwartete Spalten (genau wie aus dem alten Power App Export):

```
Title, Teilnehmername, HotelRequired, RoomType, PreferredRoommate, Status, Nr.,
Anrede, Vorname, Nachname, E-Mail, Allergies, FoodPreferences,
Freitextfeld1, Freitextfeld2, Dropdownfeld1, Dropdownfeld2,
LocationHotel, Department, CompanyName, JobTitle, Office
```

Spalten-Reihenfolge ist egal (Header wird per Name gelesen). Pflichtfelder: **Vorname, Nachname, E-Mail**. Andere optional.

## Migrations-Verhalten

**Alle Migrationen aus diesem Ordner laufen SILENT:**
- ✅ Insert in die Teilnehmer-Liste der Subsite (Status `Angemeldet`, sofern nicht im CSV explizit `Abgemeldet`)
- ✅ Dual-Write in `DEX_Participants`
- ❌ **Kein** Anmelde-Mail (kein DEX_Emails-Eintrag)
- ❌ **Kein** Outlook-Termin (kein DEX_Outlook-Eintrag)
- ✅ Idempotent: wenn die Teilnehmer-Liste der Subsite schon Items hat, wird die Migration übersprungen

Das ist sinnvoll für **Bestands-Migrationen aus der alten App** — die User haben dort schon einen Outlook-Termin und brauchen keine neue Mail.

## Workflow

1. CSV in diesen Ordner legen mit dem konventionellen Dateinamen
2. Mir Bescheid geben → ich
   - parse die CSV
   - generiere `dex-event-app-spfx/src/webparts/dexEventPlatform/data/seed-<slug>.ts`
   - registriere die Migration in `EventService.seedNewEventParticipants()`
   - bumpe die Version, builde das Package
3. Du deploysed die neue Version
4. Beim ersten Admin-App-Start läuft die Migration einmal (idempotent)
5. Nach erfolgreicher Migration: ich nehme den Code wieder raus (analog JP Morgan)

## Bestehende Migrations-Snapshots

| Datum | Datei | Event | Status |
|-------|-------|-------|--------|
| 2026-04-08 | (intern) | Assistenz Meeting 2026 (113 Pers) | ✅ migriert + Code entfernt |
| 2026-04-13 | (intern) | SR&T P_MD_D Meeting (326 Pers) | ✅ migriert + Code entfernt |
| 2026-04-14 | `JPMorgan_Listen.xlsx` | JP Morgan Corporate Challenge (485 Pers) | ✅ migriert + Code entfernt |
