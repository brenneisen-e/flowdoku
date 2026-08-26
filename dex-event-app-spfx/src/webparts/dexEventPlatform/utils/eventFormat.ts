// v18.34 — Hilfsfunktion: lesbaren Outlook-Termin-Ort aus Veranstaltungsort +
// Adresse bauen. `LocationAddress` liegt in DEX_Events als JSON-String
// ({ street, houseNo, zip, city }) vor — der Power-Automate-Flow kann das nicht
// sinnvoll parsen, deshalb baut die App den fertigen Orts-String und legt ihn
// in der Spalte `OutlookLocation` ab. Der Flow mappt diese Spalte 1:1 in das
// Location-Feld des Outlook-Termins.

/**
 * v22.22 — Liegt das Event vollständig in der Vergangenheit?
 *
 * Ende (Fallback: Ende des Start-Tages, 23:59:59) < jetzt. Ohne verwertbare
 * Datumsangaben immer `false` (fail-open: lieber eine Abmeldung zu viel
 * erlauben als eine legitime blockieren). Genutzt für die Abmelde-Sperre
 * vergangener Events (Meine Events / Auto-Cancel-Deep-Link) und die stille
 * Organizer-Abmeldung im Admin Center (keine Mail / kein Outlook / kein
 * Nachrücken mehr).
 */
export function isEventOver(ev: { startDate?: string; endDate?: string }): boolean {
  const endRaw = (ev.endDate || '').trim();
  let end = endRaw ? new Date(endRaw).getTime() : NaN;
  if (!Number.isFinite(end)) {
    const startRaw = (ev.startDate || '').trim();
    if (startRaw) {
      const s = new Date(startRaw);
      if (Number.isFinite(s.getTime())) {
        end = new Date(s.getFullYear(), s.getMonth(), s.getDate(), 23, 59, 59).getTime();
      }
    }
  }
  return Number.isFinite(end) && end < Date.now();
}

/**
 * v22.54 — Ist die Anmeldung für ein einzelnes Event/Sub-Event noch offen?
 *
 * Offen = Anmelde-Frist nicht abgelaufen UND das Event ist nicht vorbei.
 * Ohne gesetzte Frist gilt es als offen (bis das Event vorbei ist).
 */
export function isRegistrationOpen(ev: { registrationDeadline?: string; startDate?: string; endDate?: string }): boolean {
  const dl = (ev.registrationDeadline || '').trim();
  const deadlinePassed = !!dl && new Date(dl).getTime() < Date.now();
  return !deadlinePassed && !isEventOver(ev);
}

/**
 * v22.54 — Ist die Anmeldung für das GESAMTE Event geschlossen?
 *
 * Leitlinie: Solange das Hauptevent ODER mindestens ein Sub-Event noch offen
 * ist, bleibt die Anmeldung möglich. Erst wenn das Hauptevent zu ist UND kein
 * Sub-Event mehr offen ist, gilt das Event als komplett geschlossen.
 *
 * Behebt den Bug, dass eine abgelaufene Klammer-/Hauptevent-Frist die ganze
 * Anmeldung sperrte, obwohl die Sub-Events noch offen waren.
 */
export function isRegistrationFullyClosed(
  event: { registrationDeadline?: string; startDate?: string; endDate?: string; klammerDeadline?: string; subDeadlineRule?: { reg?: unknown } },
  childEvents: Array<{ registrationDeadline?: string; startDate?: string; endDate?: string }>,
): boolean {
  // v28.20: EXPLIZITE Klammer-Frist (Piggyback _klammerDeadline) — vom
  // Organizer bewusst gesetzt. Abgelaufen = Gesamt-Event zu, offene
  // Sub-Events ändern daran nichts. (Bewusst getrennt von der Spalte
  // RegistrationDeadline, deren Alt-Werte bei Klammern wirkungslos sind.)
  // v30.6: AUSSER die rollierende „Anmeldung bis"-Regel ist aktiv — dann
  // gelten die Fristen je Termin, und eine (aus der Zeit VOR der Regel
  // stehengebliebene) Klammer-Frist sperrte sonst das gesamte Event: Beim
  // Soft Opening war Tag 1 zu (25.08.), alle weiteren Tage offen — und die
  // Anmeldeseite zeigte trotzdem „Anmeldefrist abgelaufen" für alles.
  const rollingReg = !!(event.subDeadlineRule && event.subDeadlineRule.reg);
  const kdl = (event.klammerDeadline || '').trim();
  if (kdl && !rollingReg) {
    const kt = new Date(kdl).getTime();
    if (Number.isFinite(kt) && kt < Date.now()) return true;
  }
  // Hauptevent-Gate bleibt rein frist-basiert (wie bisher) — kein
  // unerwartetes Schließen über die Event-Vorbei-Logik.
  const dl = (event.registrationDeadline || '').trim();
  const mainDeadlinePassed = !!dl && new Date(dl).getTime() < Date.now();
  if (!mainDeadlinePassed) return false;
  // Hauptevent-Frist abgelaufen: nur geschlossen, wenn auch kein Sub-Event
  // mehr offen ist.
  return !(childEvents || []).some(isRegistrationOpen);
}

interface AddressParts {
  street?: string;
  houseNo?: string;
  zip?: string;
  city?: string;
}

/**
 * Baut einen lesbaren Ort-String, z.B.
 * „Rheinterrasse Düsseldorf, Schwannstraße 6, 40476 Düsseldorf".
 *
 * @param location     Veranstaltungsort (Freitext, z.B. Gebäude-/Raumname)
 * @param address      Adresse als JSON-String ODER Objekt ({street,houseNo,zip,city})
 */
export function buildOutlookLocation(
  location?: string,
  address?: string | AddressParts | null
): string {
  let addr: AddressParts = {};
  if (typeof address === 'string') {
    try { addr = address ? JSON.parse(address) : {}; } catch { addr = {}; }
  } else if (address && typeof address === 'object') {
    addr = address;
  }
  const streetLine = [addr.street, addr.houseNo].map(s => (s || '').trim()).filter(Boolean).join(' ');
  const cityLine = [addr.zip, addr.city].map(s => (s || '').trim()).filter(Boolean).join(' ');
  return [location, streetLine, cityLine].map(s => (s || '').trim()).filter(Boolean).join(', ');
}

/**
 * v29.61 — Zeitraum eines GANZTÄGIGEN Events als Text.
 *
 * DEX speichert ganztägig als 00:00–23:59 (der Outlook-Flow rechnet daraus die
 * echten Ganztags-Grenzen, siehe docs/flow-jsons.md v29.52). Diese Uhrzeiten
 * anzuzeigen ist irreführend: Der Organizer hat „ganztägig" gewählt und liest
 * dann „00:00 - 23:59" — das sieht aus, als hätte die Einstellung nicht
 * gegriffen. Genau so gemeldet worden.
 *
 * Liefert deshalb nur die Datumsangaben plus den Zusatz. Ein eintägiger Termin
 * bekommt nur EIN Datum, nicht zweimal dasselbe.
 */
export function formatAllDayPeriod(startIso: string, endIso: string, isDe: boolean): string {
  if (!startIso) return '';
  const start = new Date(startIso);
  if (isNaN(start.getTime())) return '';
  const end = endIso ? new Date(endIso) : null;
  const fmt = (d: Date): string => d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', {
    day: '2-digit', month: '2-digit', year: 'numeric',
  });
  const suffix = isDe ? ' (ganztägig)' : ' (all day)';
  if (!end || isNaN(end.getTime()) || start.toDateString() === end.toDateString()) {
    return fmt(start) + suffix;
  }
  return `${fmt(start)} – ${fmt(end)}${suffix}`;
}
