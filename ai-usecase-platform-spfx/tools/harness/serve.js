/* Kleiner Datei-Server für out/ — ohne Abhängigkeit.
 *
 *   node serve.js [port]      → http://127.0.0.1:5173/?role=admin
 *
 * Wozu, obwohl `file://` genügen würde: Ein Mensch kann die Plattform damit in
 * einem echten Browser ansehen (Parameter: siehe entry.tsx), und die
 * Kachelbilder, die die App als Anhang „hochlädt", liegen danach unter der
 * Adresse, die die App sich merkt (`/sites/AIUC/Lists/…/Attachments/…`). Die
 * Attrappe legt sie per PUT auf `/__harness/anhang` ab; hier werden sie
 * ausgeliefert. Sie leben nur im Speicher dieses Prozesses.
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'out');
const TYPEN = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml' };

function start(port) {
  const anhaenge = new Map();
  const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    if (url.pathname === '/__harness/anhang' && req.method === 'PUT') {
      const teile = [];
      req.on('data', c => teile.push(c));
      req.on('end', () => {
        anhaenge.set(url.searchParams.get('pfad'), { typ: req.headers['content-type'] || 'application/octet-stream', daten: Buffer.concat(teile) });
        res.writeHead(204).end();
      });
      return;
    }
    if (anhaenge.has(decodeURIComponent(url.pathname))) {
      const a = anhaenge.get(decodeURIComponent(url.pathname));
      res.writeHead(200, { 'Content-Type': a.typ }).end(a.daten);
      return;
    }
    const rel = url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname).replace(/^\/+/, '');
    const datei = path.join(OUT, rel);
    // Nur innerhalb von out/ — kein `..` aus dem Ordner heraus.
    if (datei.indexOf(OUT + path.sep) !== 0 || !fs.existsSync(datei) || !fs.statSync(datei).isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain' }).end('nicht gefunden');
      return;
    }
    res.writeHead(200, { 'Content-Type': TYPEN[path.extname(datei)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
    fs.createReadStream(datei).pipe(res);
  });
  return new Promise(resolve => server.listen(port, '127.0.0.1', () => resolve(server)));
}

module.exports = { start };

if (require.main === module) {
  start(parseInt(process.argv[2] || '5173', 10)).then(s => {
    console.log(`Harness unter http://127.0.0.1:${s.address().port}/?role=admin`);
  });
}
