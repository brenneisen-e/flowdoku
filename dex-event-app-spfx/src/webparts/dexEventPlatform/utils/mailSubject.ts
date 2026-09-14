/**
 * v30.18 — Mail-Betreff für Kalender-Tage.
 *
 * Kalender-Tage heißen wie ihr Datum („Fr. 04.09.2026"). Als Termin-Name in
 * Outlook ist das seit v30.7 gefixt (OutlookSubject = Hauptevent-Titel) —
 * als MAIL-Betreff stand das Datum aber weiterhin allein da und war in der
 * Inbox keinem Event zuzuordnen. Der Betreff führt jetzt den Hauptevent-
 * Namen an: „Soft Opening … — Fr. 04.09.2026".
 *
 * Bewusst NUR der Betreff: `{{EventTitle}}` im Mail-BODY bleibt der
 * Tages-Titel, weil Organizer-Texte ihn als Datum verwenden („… für den
 * Office-Tag am {{EventTitle}}") — ein globaler Tausch würde diese Sätze
 * zerreißen. Greift nur bei Kindern eines Kalender-Events; überall sonst
 * ist die Funktion ein No-op, ebenso wenn der Hauptevent-Titel schon im
 * Betreff steckt.
 */
export function withParentTitleSubject(
  subject: string,
  calendarParent: { title?: string; subEventCalendar?: boolean } | undefined | null,
): string {
  if (!calendarParent || !calendarParent.subEventCalendar) return subject;
  const pt = (calendarParent.title || '').trim();
  if (!pt) return subject;
  if ((subject || '').toLowerCase().indexOf(pt.toLowerCase()) >= 0) return subject;
  return `${pt} — ${subject}`;
}

/**
 * v31.31: Der Betreff einer Nachrück-Mail sagt, dass es eine ist.
 *
 * Nutzer-Wunsch 14.09.2026: „wäre es möglich, dass Personen, die nachrücken,
 * keine ganz normale Anmeldemail bekommen, sondern zumindest im Betreff steht
 * ‚Anmeldebestätigung Event xx (von Warteliste nachgerückt)‘?"
 *
 * Warum das hier steht und nicht nur in der Vorlage: Die Vorlage `Nachruecken`
 * hat zwar einen eigenen Betreff („Du hast einen Platz: …"), aber sie ist je
 * Event überschreibbar (`applyEventTemplateOverride`) — und genau das ist der
 * Fall, in dem sie zur Kopie der normalen Anmeldebestätigung wird. Dann steht
 * im Postfach der nachgerückten Person eine Mail, die von einer regulären
 * Anmeldung nicht zu unterscheiden ist, obwohl sich dahinter eine Entscheidung
 * verbirgt, die sie kennen muss.
 *
 * Der Zusatz wird deshalb am VERSANDORT angehängt, nach allen Vorlagen und
 * Überschreibungen — und nur, wenn er nicht ohnehin schon dasteht (sonst
 * doppelt er sich bei der gepflegten Vorlage).
 */
export function withPromotionSubject(subject: string, isDe: boolean): string {
  const s = (subject || '').trim();
  const lc = s.toLowerCase();
  if (lc.indexOf('nachgerückt') >= 0 || lc.indexOf('nachgerueckt') >= 0
    || lc.indexOf('waiting list') >= 0 || lc.indexOf('waitlist') >= 0) return s;
  return isDe
    ? `${s} (von Warteliste nachgerückt)`
    : `${s} (moved up from the waiting list)`;
}
