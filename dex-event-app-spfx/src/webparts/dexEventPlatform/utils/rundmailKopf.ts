/**
 * v32.58: Wer bei einer Rundmail (Einladung, Massenmail) in welchem Feld steht —
 * EINE Regel für beide Dialoge.
 *
 * Nutzer-Ansage 07.10.2026: „alle Organizer (auch der Absender des Klicks)
 * sollen immer in CC stehen und nicht im An-Feld", und der Absender entscheidet
 * selbst, ob die Empfänger sichtbar im An-Feld oder verdeckt im BCC stehen.
 *
 * Bis v32.57 hatte jeder Dialog seine eigene Antwort: Die Einladung setzte den
 * Absender ins An-Feld und die Organizer (auch ausgeblendete) ins CC, gegen die
 * Empfänger entdoppelt; die Massenmail setzte im BCC-Modus den ERSTEN Organizer
 * ins An-Feld. Beides widersprach der Regel oben, und im Kopf stand der
 * Absender doppelt (An und CC, Screenshot 30.09.2026).
 *
 * Die Organizer kommen über `visibleOrganizerEmails` — wer im Assistenten
 * ausgeblendet ist, steht nicht im Kopf einer Mail an die Teilnehmenden
 * (Nutzer-Ansage 10.09.2026, v31.10). Der Absender steht trotzdem immer drin:
 * Er hat den Knopf gedrückt und braucht seine Kopie.
 */

/** Das Gruppenpostfach, aus dem der Flow sendet. Es trägt das An-Feld, wenn
 *  die Empfänger verdeckt im BCC stehen — der Flow („Send an email from a
 *  shared mailbox") braucht ein An-Feld, und ein Organizer darf es nicht sein. */
export const NO_REPLY_MAILBOX = 'no_reply.events@deloitte.de';

const lc = (e: string): string => (e || '').trim().toLowerCase();

/**
 * Das CC einer Rundmail: Organizer, dann der Absender, dann das von Hand
 * ergänzte CC — jede Adresse einmal, Schreibweise der ersten Nennung.
 */
export function rundmailCc(organizer: string[], absender: string, zusatz: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const raw of [...organizer, absender, ...zusatz]) {
    const e = (raw || '').trim();
    const k = lc(e);
    if (!e || k.indexOf('@') <= 0 || seen.has(k)) continue;
    seen.add(k);
    out.push(e);
  }
  return out;
}

/**
 * Die Empfänger ohne alle, die schon im CC stehen. Wer im CC steht, steht
 * nicht zusätzlich im An-Feld oder im BCC — zugestellt wird ihm die Mail über
 * das CC. Ohne diesen Schritt stünde ein Organizer, der bei seinem eigenen
 * Event angemeldet ist, wieder im An-Feld.
 */
export function ohneCc(empfaenger: string[], cc: string[]): string[] {
  const ccLc = new Set(cc.map(lc));
  return empfaenger.filter(e => !ccLc.has(lc(e)));
}
