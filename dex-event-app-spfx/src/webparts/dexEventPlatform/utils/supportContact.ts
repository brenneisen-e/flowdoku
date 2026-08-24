/**
 * v29.43 — Postanschrift des DEX-Teams.
 *
 * Bis hierher standen in der App die persönlichen Konten der beiden Entwickler:
 * Organizer-Anfragen gingen an sie, die Onboarding-Mail hatte sie im Cc, und
 * die Kontakt-Hinweise nannten sie namentlich. Das ist zwei Personen, kein
 * Postfach — bei Urlaub, Wechsel oder Weggang landet die Anfrage nirgends.
 *
 * Deshalb geht alles, was an „das DEX-Team" adressiert ist, an dieses
 * Funktionspostfach. Die Konstante steht an EINER Stelle, damit die nächste
 * Adressänderung nicht wieder ein Dutzend Fundstellen bedeutet.
 *
 * NICHT betroffen:
 *  - Fragen aus dem Fragen-Dialog. Die gehen an die Organizer des Events bzw.
 *    an die Power-User aus der Rollenverwaltung (`TicketContext`), also an
 *    Personen, die in SharePoint gepflegt werden — dort ändert man das, nicht
 *    hier im Code.
 *  - Die Autoren-Nennung („Entwickelt von …") auf der Startseite. Das ist ein
 *    Credit, keine Postanschrift.
 */

/** Funktionspostfach des DEX-Teams — Anfragen, Onboarding-Kopien, Kontakt. */
export const DEX_TEAM_EMAIL = 'dex.event@deloitte.de';

/** Empfängerliste für Team-Mails (Semikolon-getrennt, wie queueEmail sie erwartet). */
export const DEX_TEAM_RECIPIENTS = DEX_TEAM_EMAIL;
