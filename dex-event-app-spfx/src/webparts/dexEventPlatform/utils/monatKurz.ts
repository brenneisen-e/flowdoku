/**
 * v31.73: Monatsbeschriftung des Zeitstrahls („Okt. 2026") — EINE Stelle für
 * die Listen-Ansicht, die Monats-Trenner der Karten-Ansicht und die
 * Organizer-Eventübersicht. Leer bei fehlendem oder unlesbarem Datum (dann
 * kein Trenner und keine Beschriftung auf der Schiene).
 */
export function monatKurz(iso: string | undefined, locale: string): string {
  const d = iso ? new Date(iso) : null;
  if (!d || isNaN(d.getTime())) return '';
  return d.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { month: 'short', year: 'numeric' });
}
