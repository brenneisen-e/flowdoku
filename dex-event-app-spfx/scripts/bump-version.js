/**
 * bump-version.js
 *
 * Zaehlt die Patch-Version automatisch hoch in:
 *   - package.json (npm version)
 *   - config/package-solution.json (SPFx solution + feature version)
 *
 * Wird vor jedem "npm run package" ausgefuehrt.
 * Format: Major.Minor.Patch (npm) -> Major.Minor.Patch.0 (SPFx)
 */

const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const pkgPath = path.join(rootDir, 'package.json');
const spfxPath = path.join(rootDir, 'config', 'package-solution.json');

// package.json lesen und Patch-Version erhoehen
const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
const parts = pkg.version.split('.').map(Number);
parts[2] = parts[2] + 1; // Patch hochzaehlen
const newVersion = parts.join('.');
const spfxVersion = `${parts[0]}.${parts[1]}.${parts[2]}.0`;

pkg.version = newVersion;
fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + '\n');

// package-solution.json aktualisieren
const spfx = JSON.parse(fs.readFileSync(spfxPath, 'utf8'));
spfx.solution.version = spfxVersion;
if (spfx.solution.features) {
  spfx.solution.features.forEach(f => { f.version = spfxVersion; });
}
fs.writeFileSync(spfxPath, JSON.stringify(spfx, null, 2) + '\n');

console.log(`Version bumped: ${newVersion} (SPFx: ${spfxVersion})`);
