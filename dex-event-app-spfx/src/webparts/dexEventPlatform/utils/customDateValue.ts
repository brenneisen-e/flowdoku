/**
 * v32.53: Speicherformat von Datums-Antworten — aus RegistrationPage
 * herausgelöst, weil „Meine Events → Angaben bearbeiten" dasselbe Format
 * schreiben muss. Bis v32.52 zeigte der Bearbeiten-Modus ein Datumsfeld als
 * Freitext (Nutzer-Befund 30.09.2026); wer dort etwas eintippte, schrieb ein
 * Format, das Organizer Center und Export nicht als Datum lesen.
 *
 * v31.9: Der gespeicherte String bleibt ZEICHENGLEICH zum früheren nativen
 * Feld — `YYYY-MM-DD` bzw. `YYYY-MM-DDTHH:mm`, beides in lokaler Zeit. Nur die
 * Anzeige ist deutsch, nie der Speicherwert (dieselbe Roundtrip-Falle wie
 * `{{Organizer}}`, v30.74).
 */
const pad2 = (n: number): string => String(n).padStart(2, '0');

export const parseCustomDateValue = (raw: string): Date | null => {
  const m = (raw || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/);
  if (!m) return null;
  const d = new Date(
    parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10),
    m[4] ? parseInt(m[4], 10) : 0, m[5] ? parseInt(m[5], 10) : 0, 0, 0,
  );
  return isNaN(d.getTime()) ? null : d;
};

export const formatCustomDateValue = (d: Date | null, withTime: boolean): string => {
  if (!d || isNaN(d.getTime())) return '';
  const day = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  return withTime ? `${day}T${pad2(d.getHours())}:${pad2(d.getMinutes())}` : day;
};
