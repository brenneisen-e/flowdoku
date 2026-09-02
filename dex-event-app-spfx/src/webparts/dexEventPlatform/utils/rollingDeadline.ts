// v30.67: Rollierende Fristen (`_subDeadlineRule`) — EINMAL ausgerechnet.
//
// Vorher stand die Rechnung als `amount * 86400000` im Wizard
// (EventCreationPage.rollingDeadlineIso) und noch einmal als Fallback der
// Anmeldeseite (utils/eventFormat.ts, subEventRegDeadline). Beide verschoben
// einen absoluten UTC-Zeitpunkt um exakt 24 h je Tag. Über eine Sommerzeit-
// Grenze hinweg ist das kein Kalendertag mehr: Der Berliner Offset wechselt
// zwischen +1 und +2, das Ergebnis liegt eine Stunde daneben — und weil die
// erzeugten Kalender-Termine bewusst auf 00:00 stehen, rutscht diese Stunde
// über Mitternacht auf den VORTAG. Termin 30.03. (Montag nach der
// Umstellung), Regel „1 Tag vorher": Frist 28.03. 23:00 statt 29.03. 00:00 —
// ein Buchungsfenster, das 25 Stunden früher zuging als bei allen Geschwistern.
//
// Deshalb: Tage im Berliner KALENDER verschieben (Datumsteil ± n, Uhrzeit
// bleibt), dann wieder nach UTC. Stunden bleiben absolut — „bis 2 Stunden
// vorher" meint wirklich 2 Stunden.
import { berlinLocalToUtcIso, isoToLocal } from './berlinTime';

const pad2 = (n: number): string => (n < 10 ? '0' : '') + n;

/**
 * Frist relativ zum Termin-Start als UTC-ISO. `after=true` rechnet VORWÄRTS
 * (Abmelden nach Beginn). Leerer String bei ungültigem Start oder amount <= 0.
 */
export const rollingDeadlineIso = (startIso: string, amount: number, unit: 'days' | 'hours', after?: boolean): string => {
  const t = new Date(startIso || '').getTime();
  if (!isFinite(t) || !(amount > 0)) return '';
  const sign = after ? 1 : -1;
  if (unit === 'hours') return new Date(t + sign * amount * 3600000).toISOString();
  const local = isoToLocal(startIso); // „YYYY-MM-DDTHH:MM" in Europe/Berlin
  if (!local) return '';
  const [datePart, timePart] = local.split('T');
  const [y, m, d] = datePart.split('-').map(Number);
  // Date.UTC verträgt Tag-Überläufe (0, 32, …) und liefert den Kalendertag
  // ohne jede Zeitzonen-Arithmetik — genau das ist hier gewollt.
  const shifted = new Date(Date.UTC(y, m - 1, d + sign * amount));
  const shiftedLocal = `${shifted.getUTCFullYear()}-${pad2(shifted.getUTCMonth() + 1)}-${pad2(shifted.getUTCDate())}T${timePart}`;
  return berlinLocalToUtcIso(shiftedLocal);
};
