/**
 * v31.70: Eingefügte Verteiler lesen — Adresse UND Name.
 *
 * Bis v31.69 zog die Massenmail aus einem eingefügten Verteiler nur die
 * E-Mail-Adressen (Regex über den Rohtext). Für „Nachrücker" reichte das:
 * Angeschrieben wurden Teilnehmer aus der Liste, deren Namen die App kennt.
 * Für die Erinnerung (Nutzer-Ansage 17.09.2026: „Reminder an alle, die noch
 * nicht zurückgemeldet haben, des Verteilers") sind die Empfänger aber genau
 * die, die NICHT in der Teilnehmerliste stehen — ihr Name kommt nur aus dem
 * Verteiler selbst.
 *
 * Erkannt werden die üblichen Formen eines Outlook-An-Feldes oder
 * Verteiler-Exports:
 *   Max Mustermann <mmustermann@deloitte.de>
 *   "Mustermann, Max" <mmustermann@deloitte.de>
 *   Mustermann, Max (DE - Köln) <mmustermann@deloitte.de>
 *   mmustermann@deloitte.de
 * getrennt durch Semikolon, Zeilenumbruch oder Komma. Das Komma ist
 * zweideutig („Nachname, Vorname"): Ein Stück wird nur dann am Komma
 * geteilt, wenn es mehr als eine Adresse enthält.
 */

export interface PastedRecipient {
  /** kleingeschrieben, ohne Leerraum */
  email: string;
  /** wie im Verteiler geschrieben, ohne Klammern und Anführungszeichen; '' wenn nur die Adresse da stand */
  name: string;
}

const EMAIL_RE = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

function nameAus(chunk: string, email: string): string {
  return chunk
    .replace(email, ' ')
    .replace(/mailto:/gi, ' ')
    .replace(/[<>"'“”„‚‘’]/g, ' ')
    .replace(/\(\s*\)/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^[\s,;:-]+|[\s,;:-]+$/g, '')
    .trim();
}

function zerlege(chunk: string, out: PastedRecipient[], seen: Set<string>): void {
  const emails = chunk.match(EMAIL_RE) || [];
  if (emails.length === 0) return;
  if (emails.length > 1) {
    // Mehrere Adressen in einem Stück → Komma war hier ein Trenner.
    const teile = chunk.split(',');
    if (teile.length > 1) { for (const t of teile) zerlege(t, out, seen); return; }
    // Ohne Komma (z. B. nur Leerzeichen): jede Adresse einzeln, Namen unbekannt.
    for (const e of emails) {
      const lc = e.toLowerCase();
      if (!seen.has(lc)) { seen.add(lc); out.push({ email: lc, name: '' }); }
    }
    return;
  }
  const lc = emails[0].toLowerCase();
  if (seen.has(lc)) return;
  seen.add(lc);
  out.push({ email: lc, name: nameAus(chunk, emails[0]) });
}

/** Liest einen eingefügten Verteiler; jede Adresse genau einmal, in Reihenfolge. */
export function parsePastedRecipients(raw: string): PastedRecipient[] {
  const out: PastedRecipient[] = [];
  const seen = new Set<string>();
  if (!raw) return out;
  for (const chunk of raw.split(/[;\n\r]+/)) {
    if (chunk.trim()) zerlege(chunk, out, seen);
  }
  return out;
}

/** Nur die Adressen (kleingeschrieben) — der bisherige Vertrag der Nachrücker-Suche. */
export function parsePastedEmails(raw: string): string[] {
  return parsePastedRecipients(raw).map(r => r.email);
}

/** „Mustermann, Max" → { vorname: 'Max', nachname: 'Mustermann' }; „Max Mustermann" → umgekehrt gelesen. */
export function splitName(name: string): { vorname: string; nachname: string } {
  const n = (name || '').replace(/\s*\([^)]*\)\s*$/, '').trim();
  if (!n) return { vorname: '', nachname: '' };
  const komma = n.indexOf(',');
  if (komma > 0) return { nachname: n.slice(0, komma).trim(), vorname: n.slice(komma + 1).trim() };
  const teile = n.split(/\s+/);
  if (teile.length === 1) return { vorname: teile[0], nachname: '' };
  return { vorname: teile.slice(0, -1).join(' '), nachname: teile[teile.length - 1] };
}
