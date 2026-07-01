/**
 * v26.33: Deep-Link-Parameter robust auslesen — aus dem URL-HASH *und* dem
 * Query-String.
 *
 * Hintergrund (Bug „Sorry, something went wrong / No item exists"): Wird ein
 * Ticket-/Benachrichtigungs-Link aus MICROSOFT TEAMS heraus geklickt, hängt der
 * Teams-Client eine Reihe eigener Parameter an den Query-String an
 * (`xsdata`, `TeamsCID`, `OR=Teams-HL`, `CT`, `clickparams`, `startedResponseCatch`).
 * Zusammen mit `env=WebView` und unseren eigenen Query-Parametern
 * (`action=tickets&id=…`) bringt das den SharePoint-Seiten-Resolver zum
 * Absturz („No item exists at …DEX.aspx?…") — noch BEVOR die App lädt.
 *
 * Lösung: Unsere App-Parameter wandern in den URL-HASH (`#action=tickets&id=4`).
 * Der Hash wird NIE an den Server geschickt — SharePoint kann daran also nicht
 * scheitern und liefert die Seite normal aus; die App liest den Hash
 * clientseitig. Bestehende (Query-basierte) Links funktionieren weiterhin: diese
 * Funktion liest BEIDE Quellen, wobei der Hash Vorrang hat.
 */
export function deepLinkParams(): URLSearchParams {
  const out = new URLSearchParams();
  try {
    const q = new URLSearchParams(window.location.search || '');
    q.forEach((v, k) => out.set(k, v));
  } catch { /* */ }
  try {
    let h = window.location.hash || '';
    // Führende '#', '#/' oder '#?' abstreifen.
    h = h.replace(/^#/, '').replace(/^[?/]+/, '');
    if (h) {
      const hp = new URLSearchParams(h);
      hp.forEach((v, k) => out.set(k, v)); // Hash hat Vorrang.
    }
  } catch { /* */ }
  return out;
}

/**
 * Baut einen App-Deep-Link, dessen Parameter im HASH stehen (Teams-/SharePoint-
 * resistent). `base` ist die Seiten-URL inkl. `?env=WebView` (unser appBase).
 * `params` = die App-Parameter (z.B. { action: 'tickets', id: '4' }).
 */
export function buildHashDeepLink(base: string, params: Record<string, string | number | undefined>): string {
  const parts: string[] = [];
  for (const k of Object.keys(params)) {
    const v = params[k];
    if (v === undefined || v === null || v === '') continue;
    parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`);
  }
  const frag = parts.join('&');
  return frag ? `${base}#${frag}` : base;
}
