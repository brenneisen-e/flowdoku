/**
 * v28.94: Aus `AdminPage` herausgeloest. Die Kategorien des Aktionen-Menues
 * im Organizer Center — reine Beschriftungen ohne Logik.
 */
// v7.6: Wiederverwendbare Action-Kachel für den Aktionen-Bereich.
// Default in Grau, beim Hover/Focus kippt Border + Icon + Hintergrund auf
// Deloitte-Grün. Unterstützt Button (onClick), Link (href, öffnet in neuem
// Tab) und passive Wrapper (children-Mode für Spezialfälle wie das
// Excel-Dropdown). Badge zeigt zwingend "Organizer" (grüner Tint) oder
// "Nur Admin" (oranger Tint), damit auf einen Blick klar ist, für welche
// Rolle die Aktion gedacht ist.
// v20.3: Kategorien für das Aktionen-Dropdown — die Aktionen werden nicht
// mehr als flache Alphabet-Liste gerendert, sondern als aufklappbare
// Kategorien (mehrzeilige Einträge: Titel fett, Beschreibung darunter).
export type ActionCategoryKey = 'event' | 'participants' | 'mails' | 'checkin' | 'maintenance';
export const ACTION_CATEGORY_ORDER: ActionCategoryKey[] = ['event', 'participants', 'mails', 'checkin', 'maintenance'];
// v20.4: pro Kategorie zusätzlich eine Kurzbeschreibung, was darin steckt —
// sichtbar direkt im zugeklappten Kategorie-Kopf.
export const ACTION_CATEGORY_LABELS: Record<ActionCategoryKey, { de: string; en: string; descDe: string; descEn: string }> = {
  event: {
    de: 'Event', en: 'Event',
    // v27.13: „in SharePoint öffnen" entfernt — alle Aktionen laufen über die App.
    descDe: 'Event bearbeiten, Link teilen, Änderungsprotokoll ansehen.',
    descEn: 'Edit the event, share the link, view the change history.',
  },
  participants: {
    de: 'Teilnehmer', en: 'Participants',
    descDe: 'Teilnehmerliste exportieren, Warteliste nachrücken, Nummern neu vergeben, Überbuchung prüfen.',
    descEn: 'Export the participant list, promote from the waitlist, renumber participants, check overbooking.',
  },
  mails: {
    de: 'E-Mails', en: 'Emails',
    descDe: 'Mails an Teilnehmer schreiben, Einladungsmail verschicken, E-Mail-Adressen kopieren.',
    descEn: 'Write emails to participants, send the invitation email, copy email addresses.',
  },
  checkin: {
    de: 'Check-in', en: 'Check-in',
    descDe: 'Check-in am Event-Tag starten, QR-Codes versenden, Self-Check-in (QR aushängen/anzeigen) einrichten.',
    descEn: 'Start check-in on event day, send QR codes, set up self check-in (post/show the QR).',
  },
  maintenance: {
    de: 'Wartung & Reparatur', en: 'Maintenance & repair',
    descDe: 'Werkzeuge für Sonderfälle: Daten reparieren, Zähler korrigieren, alte Events umstellen.',
    descEn: 'Tools for edge cases: repair data, correct counters, migrate old events.',
  },
};

