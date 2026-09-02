/* parentRegs — v30.67: Die EINE Auflösung „welche Klammer-Zeile gehört zu
 * dieser Person" und die EINE Zählung „wie viele Personen sind im
 * Klammer-Modus angemeldet".
 *
 * Bisher suchten fünf Stellen (ConsolidatedView: hasParentReg,
 * parentFieldFilled, zwei Tabellenzellen; createKlammerActions:
 * openMainFieldsEdit, submitAssignAssistant) mit `registrations.find(...)`
 * ohne Status-Filter und ohne definierte Reihenfolge — `getAllRegistrations`
 * liefert `$orderby=Id asc`, also immer die ÄLTESTE Zeile, und die kann
 * abgemeldet sein. Folgen: eine abgemeldete Klammer-Zeile schluckte den roten
 * Kasten „Fehlende Klammer-Anmeldung", blendete „Zur Klammer hinzufügen" aus
 * und fing die Felder-Bearbeitung ab (Hotelwunsch landete auf einer Zeile,
 * die das HotelPlanningPanel nie liest).
 */
import { SPRegistration } from '../../../services/EventService';
import { DUP_ACTIVE_STATI } from '../../admin/adminConstants';

/**
 * Die AKTIVE Klammer-Zeile einer Person (Status in `DUP_ACTIVE_STATI`),
 * bei mehreren die neueste (höchste Id) — denn genau dort steht die
 * aktuelle Antwort, wenn eine Person zweimal auf der Klammer liegt.
 * `undefined`, wenn es nur abgemeldete oder gar keine Zeilen gibt.
 */
export function activeParentRegOf(registrations: SPRegistration[], emailKey: string): SPRegistration | undefined {
  if (!emailKey) return undefined;
  let best: SPRegistration | undefined;
  for (const r of registrations) {
    if ((r.ParticipantEmail || '').toLowerCase().trim() !== emailKey) continue;
    if (DUP_ACTIVE_STATI.indexOf(r.Status || '') < 0) continue;
    if (!best || (r.Id || 0) > (best.Id || 0)) best = r;
  }
  return best;
}

/**
 * Anzahl der im Klammer-Modus ANGEMELDETEN Personen — dieselbe Definition
 * wie die KPI-Kachel „Angemeldet" (`Angemeldet | QR versendet | Eingecheckt`,
 * entdoppelt über die E-Mail; emaillose Zeile zählt als eigener Kopf, s.
 * KpiTiles v23.3). Die Detail-Karte zählte bis v30.67 `consolidatedFiltered`,
 * und das enthält bewusst auch `Warteliste` — zwei Zahlen für dieselbe
 * Aussage, 20 auseinander.
 */
export function countConsolidatedActive(subEventRegsByEventId: Record<string, SPRegistration[]>): number {
  const keys = new Set<string>();
  for (const id of Object.keys(subEventRegsByEventId)) {
    for (const r of subEventRegsByEventId[id] || []) {
      if (r.Status !== 'Angemeldet' && r.Status !== 'QR versendet' && r.Status !== 'Eingecheckt') continue;
      keys.add((r.ParticipantEmail || '').toLowerCase().trim() || `__noemail#${r.Id}`);
    }
  }
  return keys.size;
}
