# SPFx bei Deloitte — Lessons Learned aus der DEX-Entwicklung

Praxis-Leitfaden aus ~30 Releases der DEX Event Experience Platform (SPFx 1.22.2,
React 17, TypeScript 4.7). Gedacht als Checkliste für **jedes weitere SPFx-Projekt**
im Tenant (z.B. eine per iframe eingebettete Vite-App wie „Banking Horizon 2030").

---

## 1. Versionierung — der wichtigste Einzelpunkt

**SharePoint erkennt kleine Versionssprünge im App-Katalog oft NICHT.** Ein Upload
mit gleicher oder nur um die Patch-Stelle erhöhter Version wird kommentarlos
angenommen — aber die Seiten liefern weiter das alte Bundle aus.

Regeln, die sich bewährt haben:

- **Bei JEDEM Deploy mindestens die MINOR-Stelle erhöhen** (28.11.0 → 28.12.0).
  Niemals nur Patch. Im Zweifel Major (das hast du im anderen Projekt schon
  richtig entschieden: jede Änderung = neue Major/Minor).
- Die Version steht an **drei Stellen** und muss überall identisch gepflegt werden:
  1. `package.json` → `"version": "x.y.0"`
  2. `config/package-solution.json` → `solution.version` **und** `features[*].version`
     (vierstellig: `x.y.0.0`)
  3. eine eigene `version.ts` mit `export const APP_VERSION = 'x.y.0'`
- **APP_VERSION sichtbar in der UI rendern** (Badge im Footer/Header). Das ist das
  einzige zuverlässige Mittel zu prüfen, ob der Deploy wirklich ankam — nicht der
  App-Katalog, nicht der Browser-Cache. Erst wenn der Badge die neue Nummer zeigt,
  ist das neue Bundle aktiv.
- Bundle-Dateinamen enthalten einen Hash — im Browser-Netzwerk-Tab sieht man am
  Dateinamen (`…web-part_<hash>.js`), ob wirklich das neue Bundle geladen wird.

## 2. Build-Ritual (deterministisch, immer gleich)

```bash
# im SPFx-Projektordner
rm -rf dist release temp sharepoint/solution/debug   # IMMER clean bauen
npx gulp bundle --ship
npx gulp package-solution --ship
```

- **Keine npm-Wrapper-Skripte mit Auto-Bump** benutzen, wenn die nur die
  Patch-Stelle erhöhen (siehe Punkt 1).
- Nach dem Build prüfen: In `release/assets/` darf **genau EIN** WebPart-Bundle
  (`<name>-web-part_<hash>.js`) liegen. Mehrere = vergessenes Clean → das sppkg
  kann das falsche referenzieren.
- Das fertige Paket liegt unter `sharepoint/solution/<name>.sppkg`.

## 3. sppkg-Paket & App-Katalog

- `includeClientSideAssets: true` in `package-solution.json` packt die Assets ins
  sppkg — SharePoint hostet sie dann selbst (Office-365-CDN /
  `ClientSideAssets`-Bibliothek). Kein eigenes Hosting nötig, ABER: siehe Punkt 4
  (anderer Origin!).
- „**Valid app package: Yes**" beim Upload prüfen. Der OPC-Fehler („package is
  invalid") kommt fast immer von kaputten Zip-Strukturen — sppkg **nie manuell
  umpacken**; wenn doch: keine Ordner-Einträge im Zip, korrekte
  `[Content_Types].xml`, Pfade mit Vorwärts-Slashes.
- Nach dem Bereitstellen: harter Reload (Strg+F5) UND Versions-Badge prüfen.
  Der SPFx-Loader cached aggressiv.

## 4. Die Blackscreen-Falle: ES-Module + CDN = CORS

Das ist die wahrscheinlichste Ursache für „App lädt, aber bleibt schwarz" bei
einer eingebetteten Vite-/moderne-Bundler-App:

- `<script type="module">` wird vom Browser **immer mit CORS** geladen.
- Die ClientSideAssets kommen vom **Office-365-CDN = anderer Origin** als die
  SharePoint-Seite. Ohne CORS-Header blockt der Browser das Modul — still,
  ohne sichtbaren Fehler auf der Seite (nur in der Konsole).
- **Lösung: als IIFE/klassisches Skript bundeln** (Vite: `build.rollupOptions` /
  `format: 'iife'`, kein `type="module"`). Klassische Skripte laden ohne CORS.
- Alternative: Assets same-origin hosten (z.B. SiteAssets der Ziel-Site) — dann
  aber Berechtigungen/Upload-Pipeline selbst bauen. IIFE ist der einfachere Weg.

**Immer eine sichtbare Lade-/Fehleranzeige einbauen** (im iframe wie im WebPart):
ein `<div>Lade …</div>`, das erst das App-JS ersetzt, plus `window.onerror` →
Fehlertext ins DOM. Ein stummer Blackscreen ist nicht debugbar; mit Anzeige sieht
man sofort „JS nie geladen" vs. „JS crasht".

## 5. React-Crashes = weiße Seite (kein Error Boundary = Totalausfall)

Zwei reale Vorfälle aus DEX, beide enden als komplett weiße Seite, weil React
bei einem unbehandelten Fehler den **gesamten Baum unmountet**:

1. **Hook-Reihenfolge:** Ein `if (!data) return <Fallback/>` **mitten** im
   Komponentenkörper, mit weiteren Hooks danach. Mountet die Komponente erst ohne
   Daten (weniger Hooks) und rendert später mit Daten (mehr Hooks), wirft React
   „Rendered more hooks than during the previous render". Trat auf, sobald eine
   Seite VOR dem Daten-Load gemountet wurde (Refresh-Restore). **Regel:** Early
   Returns nur VOR dem ersten Hook — oder die Seite erst mounten, wenn die Daten
   da sind (Loader eine Ebene höher).
2. **Stub-/Preview-Contexts:** Rendert man echte Komponenten in einem
   Vorschau-/Testmodus mit gemocktem Context, muss der Mock **jede** Funktion
   liefern, die die Komponente aufruft — auch die in `useEffect`. Fehlt eine,
   gibt es `TypeError: xy is not a function` beim Mount → weiße Seite. Mocks bei
   jeder neuen Context-Funktion mitpflegen (oder Proxy mit No-op-Fallback).

Empfehlung generell: **Error Boundary um den App-Root** + Fehlertext rendern.

## 6. SharePoint REST — harte Grenzen & Muster

- **2-MB-Request-Limit:** REST lehnt Bodies > 2.097.152 Bytes mit HTTP 400
  „The request message is too big" ab. Klassischer Auslöser: Base64-Bilder in
  Feldern — bei DEX steckte dasselbe Logo dreimal im selben MERGE. **Regeln:**
  Bilder vor dem Einbetten hart verkleinern (Canvas, z.B. max 600px), Payload-
  Größe vor dem POST prüfen und eine verständliche Fehlermeldung liefern.
- **Bilder als Item-Attachments** speichern statt in SiteAssets: wer das Item
  editieren darf, darf auch Attachments anhängen — keine extra Berechtigungen.
  Mit Namens-Präfixen arbeiten (`__eventimage__…`), beim Ersetzen erst alte
  Präfix-Treffer löschen. **Achtung Präfix-Kollision:** ein zweites Präfix darf
  nicht mit dem ersten BEGINNEN, sonst löscht die Aufräum-Logik es mit.
- **MERGE-Updates:** POST mit Headern `IF-MATCH: *`, `X-HTTP-Method: MERGE`,
  `Content-Type: application/json;odata=verbose` + `__metadata.type` im Body.
- **Einzelne JSON-Keys patchen:** Wenn mehrere Features sich ein JSON-Feld teilen,
  nie blind überschreiben — GET → parse → Key setzen/löschen → MERGE
  (read-modify-write). Sonst löschen sich Features gegenseitig die Daten.
- **„Piggyback"-Muster:** Statt für jedes kleine Flag eine neue SP-Spalte zu
  provisionieren, Zusatzfelder als Keys in EINEM Note-Feld-JSON ablegen
  (`_meinFlag: true`). Dazu gehört Disziplin: beim Laden die Piggyback-Keys aus
  dem Roh-JSON strippen, beim Speichern alle Keys frisch zusammenbauen —
  sonst überschreiben stale Werte die neuen.
- **TypeScript-Falle TS2590** („union type too complex"): lange Ketten von
  Objekt-Spreads (`{...a, ...b, ...c, …}`) sprengen den Compiler. Fix: Configs in
  ein `Array<Record<string, unknown>>` sammeln und mit `Object.assign({}, ...)`
  mergen.
- **Metadaten-Fallbacks:** `$expand=Author` + `Created` liefern Ersteller/Datum
  auch für Zeilen, die nicht über die App entstanden sind.
- **Schema-Ensure nur einmal pro Version:** idempotente `ensure*`-Calls (Listen/
  Felder anlegen) hinter ein versionsgebundenes localStorage-Flag legen —
  spart beim Boot mehrere Sekunden.

## 7. Microsoft Graph aus SPFx

- Client holen: `context.msGraphClientFactory.getClient('3')`.
- Benötigte Scopes in `package-solution.json` → `webApiPermissionRequests`
  deklarieren (Admin muss sie im SharePoint-Admin-Center freigeben).
- **SP-Userprofil-Properties sind im Tenant unzuverlässig gefüllt** (z.B.
  „Company"). Verlässlicher: Graph — `/me?$select=companyName` bzw.
  `/users/{mail}?$select=companyName`. Muster: SP-Profil zuerst, Graph als
  Fallback, alles best-effort mit try/catch.
- Verteiler auflösen: `/groups` per Mail filtern → `transitiveMembers`.

## 8. Frontend-Muster, die sich bewährt haben

- **Kein Browser-Routing in SPFx** — Navigation als React-State; für
  Refresh-Festigkeit den Zustand (Seite + ausgewählte Entität) pro Tab in
  `sessionStorage` sichern und beim Boot restaurieren. Deep-Links (URL-Parameter)
  haben Vorrang vor dem Restore.
- **Code-Splitting:** schwere Sekundärseiten via `React.lazy` als eigene Chunks —
  verkleinert das Boot-Bundle massiv.
- **Bild-Caching:** Event-/Profilbilder über IndexedDB cachen (URL → Blob) —
  zweiter Aufruf rendert sofort ohne SP-Roundtrip.
- **Canvas & Bilder:** same-origin SP-Bilder sind Canvas-lesbar (Content-Analyse,
  Kompression); alles in try/catch mit Fallback, falls „tainted".
- **Fluent UI Icons** sind im SPFx-Kontext bereits registriert — nutzen statt
  eigener Assets.

## 9. Debugging-Workflow, der funktioniert

1. **Versions-Badge** checken — läuft überhaupt der neue Code?
2. **F12-Konsole** lesen: die eigene Bundle-Datei (`…web-part_<hash>.js`) in den
   Stacktraces suchen. `TypeError: xy is not a function` in einem Effekt-Flush =
   fast immer fehlende Context-/Prop-Funktion.
3. **Eigene Perf-/Debug-Logs** mit Präfix (`[DEX][perf] …`) einbauen — bei
   Nutzer-Screenshots der Konsole sieht man sofort Boot-Phasen und Versionen.
4. 404s auf gelöschte Subsites etc. sind oft **erwartetes Rauschen** — erst die
   letzte, unbehandelte Exception zählt.
5. Fehlertexte des Servers durchreichen statt schlucken: SharePoint-Fehler stecken
   in `error.message.value` (odata verbose) — extrahieren und dem User anzeigen.

## 10. Checkliste pro Release

- [ ] Version an allen 3 Stellen erhöht (MINOR+)
- [ ] `rm -rf dist release temp sharepoint/solution/debug`
- [ ] `gulp bundle --ship` fehlerfrei (tsc + eslint laufen darin mit)
- [ ] genau EIN WebPart-Bundle in `release/assets/`
- [ ] sppkg hochgeladen, „Valid app package: Yes", bereitgestellt
- [ ] Harter Reload, Versions-Badge zeigt neue Nummer
- [ ] Kernflows kurz durchklicken (inkl. Seiten-REFRESH — der findet
      Mount-Reihenfolge-Bugs, die normale Navigation nie triggert)
- [ ] Release-Notes gepflegt (nutzerverständlich + technisch)
