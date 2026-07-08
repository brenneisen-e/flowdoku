/**
 * v27.10 (Refactor): unverändert aus RegistrationPage.tsx extrahiert
 * (reiner String→HTML-Helfer ohne Komponenten-Abhängigkeiten).
 *
 * v26.91: Feld-Beschreibungen dürfen ein kleines Markdown-Subset tragen:
 *   **fett**            → <strong>
 *   [Text](https://…)   → Link (nur http/https/mailto)
 *   nackte http(s)-URL  → automatisch verlinkt
 * v26.96: Neu erfasste Beschreibungen kommen aus einem echten Rich-Text-Editor
 * und sind bereits HTML. In diesem Fall wird das HTML (organizer-authored,
 * sicherer Origin) nur von gefährlichen Teilen befreit und direkt ausgegeben.
 * Alt-Beschreibungen (reiner Text / Markdown) gehen weiter durch das Subset.
 */
export function renderFieldDescHtml(raw: string): string {
  if (!raw) return '';
  // Enthält der Text echte HTML-Tags (Rich-Text-Editor), NICHT escapen —
  // nur script/style/Event-Handler/javascript: entfernen.
  if (/<[a-z][\s\S]*>/i.test(raw)) {
    return raw
      .replace(/<\s*script[\s\S]*?<\s*\/\s*script\s*>/gi, '')
      .replace(/<\s*style[\s\S]*?<\s*\/\s*style\s*>/gi, '')
      .replace(/\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
      .replace(/javascript:/gi, '');
  }
  const linkStyle = 'color:var(--dex-green,#86bc25);font-weight:600;text-decoration:underline;';
  let html = raw
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  // [Text](url) — nur sichere Schemata
  html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+|mailto:[^\s)]+)\)/g,
    (_m, label, url) => `<a href="${url}" target="_blank" rel="noopener noreferrer" style="${linkStyle}">${label}</a>`);
  // nackte http(s)-URLs (nicht die, die schon in einem href="…" stehen)
  html = html.replace(/(^|[\s(])(https?:\/\/[^\s<)]+)/g,
    (_m, pre, url) => `${pre}<a href="${url}" target="_blank" rel="noopener noreferrer" style="${linkStyle}">${url}</a>`);
  // **fett**
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  // Zeilenumbrüche erhalten
  html = html.replace(/\n/g, '<br />');
  return html;
}
