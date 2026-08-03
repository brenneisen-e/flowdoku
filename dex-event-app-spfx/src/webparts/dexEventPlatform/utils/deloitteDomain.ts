// v27.11: Einheitliche Intern/Extern-Klassifikation für E-Mail-Adressen.
//
// Hintergrund (Bug 2026-08): Die Personensuche wurde mit v26.57/v26.58 auf
// ALLE Deloitte-Member-Firm-Domains erweitert (@deloitte.com, @deloitte.at,
// @deloitte.nl, @deloitteCE.com, …) — die Intern/Extern-Klassifikation bei
// der Anmeldung Dritter prüfte aber weiterhin nur auf @deloitte.de. Interne
// Kolleg:innen außerhalb Deutschlands liefen dadurch fälschlich durch den
// kompletten Extern-Flow (ConsentReview='Pending', keine Bestätigungs-Mail,
// kein Outlook-Termin, QR-Umleitung an Organizer).
//
// Diese Datei ist die EINZIGE Quelle der Wahrheit für „ist diese Adresse ein
// Deloitte-Postfach?" — dieselbe Domain-Regel wie die International-Suche in
// SharePointService.searchUsers (v26.58): das Label „deloitte…" muss ein
// Domain-Label BEGINNEN (@deloitte.de, @deloitte.co.uk, @deloitteCE.com,
// @xy.deloitte.com), damit z.B. „notdeloitte.com" NICHT matcht.

/** true = Deloitte-Postfach (beliebige Member Firm), false = extern. */
export const isDeloitteInternalEmail = (email: string): boolean => {
  const v = (email || '').trim().toLowerCase();
  const at = v.lastIndexOf('@');
  if (at < 0) return false;
  const domain = v.slice(at + 1);
  return /(^|\.)deloitte[a-z0-9-]*\./.test(domain);
};

/** Bequemer Gegenpart: true = KEIN Deloitte-Postfach (externe Person). */
export const isExternalEmail = (email: string): boolean => {
  const v = (email || '').trim();
  return !!v && !isDeloitteInternalEmail(v);
};
