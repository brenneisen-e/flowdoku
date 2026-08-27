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
