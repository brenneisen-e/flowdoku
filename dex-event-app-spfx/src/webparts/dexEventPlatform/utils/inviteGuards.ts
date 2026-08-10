/**
 * v28.94: Aus `AdminPage` herausgeloest. Prueft, ob eine Adresse ueberhaupt
 * eingeladen werden darf — Verteiler, Funktionspostfaecher und Standort-Token
 * sind keine Personen und wuerden im Verteiler nur Rueckläufer erzeugen.
 */
export // v11.41: Einladungsmail-Empfänger-Blocker. Die Einladungsmail darf NIE an
// komplette Standort-Verteiler ('de.duesseldorf@...', 'duesseldorf@...' etc.)
// oder an pauschale 'all'-Listen ('deall@...', 'all@...', 'alldeloitte@...')
// gehen. Hintergrund: solche Aussendungen sind ohne CMC-/Marketing-Freigabe
// nicht erlaubt — kleinere explizite Verteilergruppen (Team-Mailboxen,
// Funktions-Accounts) bleiben aber zulässig.
const DEX_LOCATION_TOKENS: string[] = [
  'berlin', 'dresden', 'duesseldorf', 'dusseldorf', 'düsseldorf',
  'frankfurt', 'goerlitz', 'görlitz', 'halle', 'hamburg', 'hannover',
  'koeln', 'köln', 'cologne', 'leipzig', 'magdeburg', 'mannheim',
  'muenchen', 'münchen', 'munich', 'nuernberg', 'nürnberg', 'nuremberg',
  'stuttgart', 'walldorf',
];
export /** Wenn die Adresse als unerlaubter Massen-Verteiler erkannt wird, gibt den
 *  Block-Grund zurück — sonst null. Heuristik bewusst konservativ: matched
 *  nur, wenn der Local-Part (bzw. der gesamte Token, falls kein '@' vorhanden)
 *  eindeutig ein Standort-/All-Verteiler ist. Team-Mailboxen wie
 *  'frankfurt-event-team@' bleiben erlaubt.
 *
 *  v11.44: Auch reine Tokens ohne '@' werden geprüft — der Mailverteiler
 *  kann Einträge wie 'All' oder 'Duesseldorf' enthalten, die direkt aus dem
 *  Standort-/Location-Picker stammen. Vorher wurden die durchgelassen, weil
 *  der Parser an `at <= 0` zurückkehrte. */
function getBlockedInviteReason(email: string): string | null {
  const lc = (email || '').trim().toLowerCase();
  if (!lc) return null;
  const at = lc.indexOf('@');
  // Mit '@': Local-Part vor dem '@' prüfen. Ohne '@': gesamten Token prüfen
  // (z.B. wenn die Sichtbarkeit per Location-Picker auf 'All' gesetzt war
  // und 'All' so im audienceFilter landet).
  const local = at > 0 ? lc.slice(0, at) : lc;
  // Token-Split: Local-Part nach .-_ tokenisieren.
  const tokens = local.split(/[._-]/).filter(Boolean);
  // (1) deall / de.all / de-all / alldeloitte etc. — globaler DE-Verteiler.
  if (local === 'deall' || local === 'alldeloitte' || tokens.includes('deall') || tokens.includes('alldeloitte')) {
    return 'globaler Deloitte-DE-Verteiler';
  }
  // (2) 'all' als eigenständiger Token oder Local-Part — pauschale Liste.
  if (local === 'all' || tokens.includes('all')) {
    return 'globaler "all"-Verteiler';
  }
  // (3) Standort-Verteiler: Local-Part ist exakt eine Stadt ODER beginnt /
  //     endet mit 'de.' / 'de-' und enthält eine Stadt als Token.
  for (const loc of DEX_LOCATION_TOKENS) {
    if (local === loc) return `Standort-Verteiler (${loc})`;
    // 'de.<loc>' / 'de-<loc>' / '<loc>.de' / '<loc>-de'
    if (tokens.length === 2 && tokens.includes('de') && tokens.includes(loc)) {
      return `Standort-Verteiler (${loc})`;
    }
  }
  return null;
}
export /** Liefert pro Empfänger die Block-Begründung — leeres Array = alles OK. */
function getBlockedInviteRecipients(emails: string[]): Array<{ email: string; reason: string }> {
  const out: Array<{ email: string; reason: string }> = [];
  for (const e of emails) {
    const reason = getBlockedInviteReason(e);
    if (reason) out.push({ email: e, reason });
  }
  return out;
}
