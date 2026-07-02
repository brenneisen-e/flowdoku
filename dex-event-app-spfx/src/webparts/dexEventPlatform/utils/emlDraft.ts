/**
 * v26.47: Sendefertige Outlook-Entwürfe als .eml-Datei.
 *
 * Hintergrund: Der DEX-Mail-Flow (no_reply-Postfach) kann KEINE Mails an
 * externe Adressen zustellen. Externe Einladungen verschickt deshalb die
 * anmeldende Person aus dem EIGENEN Postfach. Damit sie nichts tippen muss,
 * erzeugt die App eine .eml-Datei mit dem Header `X-Unsent: 1` — Outlook
 * (Desktop) öffnet so eine Datei als UNGESENDETEN ENTWURF mit fertig
 * gefülltem An/Cc/Betreff/Text und aktivem „Senden"-Button.
 *
 * Hinweis: Im neuen Outlook/OWA kann die Datei stattdessen als empfangene
 * Mail angezeigt werden — dann per „Weiterleiten" nutzen. Der Text der
 * Instruktions-Mail weist darauf hin.
 */

/** UTF-8 → Base64 (browser-sicher, auch für Umlaute/Emoji). */
function utf8ToBase64(s: string): string {
  const bytes = new TextEncoder().encode(s);
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

/** Base64 in 76-Zeichen-Zeilen umbrechen (MIME-Konvention). */
function wrap76(s: string): string {
  const out: string[] = [];
  for (let i = 0; i < s.length; i += 76) out.push(s.slice(i, i + 76));
  return out.join('\r\n');
}

/** Betreff RFC-2047-kodieren (UTF-8, Base64) — Umlaute im Subject-Header. */
function encodeSubject(subject: string): string {
  return `=?utf-8?B?${utf8ToBase64(subject)}?=`;
}

export function buildUnsentEmlDraft(opts: {
  to: string[];
  cc?: string[];
  subject: string;
  html: string;
}): string {
  const to = (opts.to || []).map(s => (s || '').trim()).filter(Boolean);
  const cc = (opts.cc || []).map(s => (s || '').trim()).filter(Boolean);
  const headers: string[] = [`To: ${to.join(', ')}`];
  if (cc.length) headers.push(`Cc: ${cc.join(', ')}`);
  headers.push(
    `Subject: ${encodeSubject(opts.subject || '')}`,
    'X-Unsent: 1',
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset=utf-8',
    'Content-Transfer-Encoding: base64',
  );
  return `${headers.join('\r\n')}\r\n\r\n${wrap76(utf8ToBase64(opts.html || ''))}`;
}

/** .eml als Datei-Download auslösen. */
export function downloadEml(fileName: string, emlContent: string): void {
  const blob = new Blob([emlContent], { type: 'message/rfc822' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.endsWith('.eml') ? fileName : `${fileName}.eml`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
