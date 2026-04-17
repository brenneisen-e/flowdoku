<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>DEX Check-in Scanner</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Aptos, Arial, sans-serif; background: #1a1a2e; color: #fff; min-height: 100vh; overflow-x: hidden; }
  .header { background: #000; padding: 16px 20px; border-bottom: 3px solid #86bc25; display: flex; justify-content: space-between; align-items: center; }
  .header h1 { font-size: 1.1rem; font-weight: 500; }
  .header .logo { color: #86bc25; font-weight: 700; font-size: 1.2rem; }
  .header .logo span { color: #86bc25; }
  .counter { background: #86bc25; color: #000; padding: 4px 14px; border-radius: 20px; font-weight: 700; font-size: 0.9rem; }
  .scanner-container { position: relative; max-width: 500px; margin: 0 auto; padding: 16px; }
  video { width: 100%; border-radius: 12px; border: 3px solid #86bc25; display: block; }
  .scan-overlay { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 200px; height: 200px; border: 3px solid rgba(134,188,37,0.6); border-radius: 12px; pointer-events: none; }
  .result-card { max-width: 500px; margin: 0 auto 16px; padding: 20px; border-radius: 12px; animation: slideUp 0.3s ease; }
  .result-success { background: #1b5e20; border: 2px solid #86bc25; }
  .result-error { background: #b71c1c; border: 2px solid #ef5350; }
  .result-info { background: #0d47a1; border: 2px solid #42a5f5; }
  .result-already { background: #e65100; border: 2px solid #ff9800; }
  .result-card h2 { font-size: 1.3rem; margin-bottom: 4px; }
  .result-card .email { color: rgba(255,255,255,0.7); font-size: 0.9rem; }
  .result-card .status { display: inline-block; padding: 4px 12px; border-radius: 12px; font-weight: 600; font-size: 0.85rem; margin-top: 8px; }
  .result-card .event-name { font-size: 0.85rem; color: rgba(255,255,255,0.8); margin-top: 4px; }
  .manual-input { max-width: 500px; margin: 16px auto; padding: 0 16px; display: flex; gap: 8px; }
  .manual-input input { flex: 1; padding: 12px 16px; border-radius: 8px; border: 2px solid #333; background: #16213e; color: #fff; font-size: 1rem; }
  .manual-input input:focus { border-color: #86bc25; outline: none; }
  .manual-input button { padding: 12px 20px; border-radius: 8px; border: none; background: #86bc25; color: #000; font-weight: 700; cursor: pointer; font-size: 1rem; }
  .btn-start { display: block; max-width: 500px; margin: 24px auto; padding: 16px; width: calc(100% - 32px); border-radius: 12px; border: none; background: #86bc25; color: #000; font-size: 1.1rem; font-weight: 700; cursor: pointer; }
  .btn-start:hover { background: #75a821; }
  .history { max-width: 500px; margin: 16px auto; padding: 0 16px; }
  .history h3 { font-size: 0.85rem; color: #666; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; }
  .history-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 16px; background: #16213e; border-radius: 8px; margin-bottom: 6px; font-size: 0.85rem; }
  .history-item .name { font-weight: 600; }
  .history-item .time { color: #666; }
  .history-item .badge { padding: 2px 10px; border-radius: 10px; font-size: 0.75rem; font-weight: 600; }
  .badge-ok { background: #1b5e20; color: #86bc25; }
  .badge-err { background: #b71c1c; color: #ef5350; }
  .info-text { text-align: center; color: #666; font-size: 0.85rem; padding: 16px; }
  @keyframes slideUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
  @media (max-width: 480px) { .header h1 { font-size: 0.9rem; } .result-card h2 { font-size: 1.1rem; } }
</style>
</head>
<body>

<div class="header">
  <div>
    <span class="logo">Deloitte<span>.</span></span>
    <h1 id="eventTitle">DEX Check-in Scanner</h1>
  </div>
  <div class="counter" id="counterBadge">0 eingecheckt</div>
</div>

<div class="scanner-container" id="scannerContainer" style="display:none;">
  <video id="video" playsinline muted></video>
  <div class="scan-overlay"></div>
</div>

<div id="resultArea"></div>

<div class="manual-input">
  <input type="text" id="manualInput" placeholder="DEX|1|email@deloitte.de" />
  <button onclick="processManual()">Check</button>
</div>

<button class="btn-start" id="startBtn" onclick="startScanner()">
  Kamera starten
</button>

<div class="history" id="historySection" style="display:none;">
  <h3>Check-in Verlauf</h3>
  <div id="historyList"></div>
</div>

<p class="info-text">QR-Code Format: DEX|EventNumber|Email</p>

<script>
// Parameter aus URL lesen
const params = new URLSearchParams(window.location.search);
const siteUrl = params.get('siteUrl') || '';
const subsiteUrl = params.get('subsiteUrl') || '';
const eventTitle = params.get('eventTitle') || '';
const eventNumber = params.get('eventNumber') || '';

if (eventTitle) {
  document.getElementById('eventTitle').textContent = 'Check-in: ' + decodeURIComponent(eventTitle);
  document.title = 'Check-in: ' + decodeURIComponent(eventTitle);
}

let checkedInCount = 0;
let scanInterval = null;
let isProcessing = false;
let lastScannedCode = '';
const history = [];

// SharePoint Request Digest holen
async function getDigest() {
  const resp = await fetch(siteUrl + '/_api/contextinfo', {
    method: 'POST',
    headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
    credentials: 'include',
  });
  const data = await resp.json();
  return data.FormDigestValue;
}

// Kamera starten
async function startScanner() {
  const container = document.getElementById('scannerContainer');
  const video = document.getElementById('video');
  const btn = document.getElementById('startBtn');

  try {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: 'environment' }, width: { ideal: 640 }, height: { ideal: 480 } }
    });
    video.srcObject = stream;
    await video.play();
    container.style.display = 'block';
    btn.textContent = 'Kamera stoppen';
    btn.onclick = stopScanner;
    startDetection();
  } catch (err) {
    alert('Kamera konnte nicht gestartet werden: ' + (err.message || err.name) + '\n\nBitte Kamera-Berechtigung in den Browser-Einstellungen erlauben.');
  }
}

function stopScanner() {
  const video = document.getElementById('video');
  const container = document.getElementById('scannerContainer');
  const btn = document.getElementById('startBtn');

  if (scanInterval) { clearInterval(scanInterval); scanInterval = null; }
  if (video.srcObject) {
    video.srcObject.getTracks().forEach(t => t.stop());
    video.srcObject = null;
  }
  container.style.display = 'none';
  btn.textContent = 'Kamera starten';
  btn.onclick = startScanner;
}

// BarcodeDetector Loop
function startDetection() {
  if (!window.BarcodeDetector) {
    alert('BarcodeDetector nicht verfuegbar. Bitte Chrome oder Edge verwenden.');
    return;
  }
  const detector = new BarcodeDetector({ formats: ['qr_code'] });
  const video = document.getElementById('video');

  scanInterval = setInterval(async () => {
    if (isProcessing || video.readyState !== 4) return;
    try {
      const barcodes = await detector.detect(video);
      if (barcodes.length > 0) {
        const code = barcodes[0].rawValue;
        if (code && code !== lastScannedCode) {
          lastScannedCode = code;
          await processCode(code);
          // Nach 3 Sekunden wieder scannen
          setTimeout(() => { lastScannedCode = ''; }, 3000);
        }
      }
    } catch {}
  }, 400);
}

// QR-Code verarbeiten
async function processCode(code) {
  isProcessing = true;
  const area = document.getElementById('resultArea');

  const parts = code.split('|');
  if (parts.length !== 3 || parts[0] !== 'DEX') {
    showResult('error', 'Ungültiger QR-Code', code, '');
    isProcessing = false;
    return;
  }

  const en = parts[1];
  const email = parts[2];

  // Subsite URL bestimmen: entweder aus URL-Parameter oder aus Events-Liste
  let targetSubsiteUrl = subsiteUrl;
  if (!targetSubsiteUrl && siteUrl) {
    // Event aus DEX_Events laden
    try {
      const evResp = await fetch(
        siteUrl + "/_api/web/lists/getbytitle('DEX_Events')/items?$filter=EventNumber eq " + en + "&$select=SubsiteUrl,Title&$top=1",
        { headers: { 'Accept': 'application/json;odata=nometadata' }, credentials: 'include' }
      );
      const evData = await evResp.json();
      if (evData.value && evData.value.length > 0) {
        targetSubsiteUrl = evData.value[0].SubsiteUrl;
      }
    } catch {}
  }

  if (!targetSubsiteUrl) {
    showResult('error', 'Event nicht gefunden', 'Event #' + en, email);
    isProcessing = false;
    return;
  }

  // Registrierung suchen
  try {
    const regResp = await fetch(
      targetSubsiteUrl + "/_api/web/lists/getbytitle('Teilnehmer')/items?$filter=ParticipantEmail eq '" + email.replace(/'/g, "''") + "'&$select=Id,Vorname,Nachname,ParticipantName,ParticipantEmail,Status&$top=1",
      { headers: { 'Accept': 'application/json;odata=nometadata' }, credentials: 'include' }
    );
    const regData = await regResp.json();

    if (!regData.value || regData.value.length === 0) {
      showResult('error', 'Nicht registriert', email, '');
      addHistory(email, false);
      isProcessing = false;
      return;
    }

    const reg = regData.value[0];
    const name = (reg.Vorname && reg.Nachname) ? reg.Vorname + ' ' + reg.Nachname : reg.ParticipantName;

    if (reg.Status === 'Eingecheckt') {
      showResult('already', name, email, 'Bereits eingecheckt');
      isProcessing = false;
      return;
    }

    if (reg.Status === 'Abgemeldet') {
      showResult('error', name, email, 'Abgemeldet - Check-in nicht moeglich');
      addHistory(name, false);
      isProcessing = false;
      return;
    }

    // Check-in durchfuehren
    const digest = await getDigest();
    const checkResp = await fetch(
      targetSubsiteUrl + "/_api/web/lists/getbytitle('Teilnehmer')/items(" + reg.Id + ")",
      {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
          'X-RequestDigest': digest,
        },
        credentials: 'include',
        body: JSON.stringify({ 'Status': 'Eingecheckt' }),
      }
    );

    if (checkResp.ok) {
      checkedInCount++;
      document.getElementById('counterBadge').textContent = checkedInCount + ' eingecheckt';
      showResult('success', name, email, 'Eingecheckt');
      addHistory(name, true);
    } else {
      showResult('error', name, email, 'Check-in fehlgeschlagen');
      addHistory(name, false);
    }
  } catch (err) {
    showResult('error', 'Fehler', err.message || '', '');
  }

  isProcessing = false;
}

function showResult(type, title, subtitle, status) {
  const area = document.getElementById('resultArea');
  const cls = type === 'success' ? 'result-success' : type === 'already' ? 'result-already' : type === 'info' ? 'result-info' : 'result-error';
  const statusColor = type === 'success' ? '#86bc25' : type === 'already' ? '#ff9800' : '#ef5350';
  area.innerHTML = '<div class="result-card ' + cls + '">' +
    '<h2>' + escHtml(title) + '</h2>' +
    (subtitle ? '<div class="email">' + escHtml(subtitle) + '</div>' : '') +
    (status ? '<div class="status" style="background:rgba(0,0,0,0.3);color:' + statusColor + '">' + escHtml(status) + '</div>' : '') +
    '</div>';
  // Auto-clear nach 4 Sekunden
  setTimeout(() => { area.innerHTML = ''; }, 4000);
}

function addHistory(name, ok) {
  const time = new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  history.unshift({ name, ok, time });
  if (history.length > 20) history.pop();
  renderHistory();
}

function renderHistory() {
  const section = document.getElementById('historySection');
  const list = document.getElementById('historyList');
  if (history.length === 0) { section.style.display = 'none'; return; }
  section.style.display = 'block';
  list.innerHTML = history.map(h =>
    '<div class="history-item">' +
    '<span class="name">' + escHtml(h.name) + '</span>' +
    '<span><span class="badge ' + (h.ok ? 'badge-ok' : 'badge-err') + '">' + (h.ok ? 'OK' : 'Fehler') + '</span> <span class="time">' + h.time + '</span></span>' +
    '</div>'
  ).join('');
}

function processManual() {
  const input = document.getElementById('manualInput');
  if (input.value.trim()) { processCode(input.value.trim()); input.value = ''; }
}

document.getElementById('manualInput').addEventListener('keydown', e => {
  if (e.key === 'Enter') processManual();
});

function escHtml(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
</script>
</body>
</html>
