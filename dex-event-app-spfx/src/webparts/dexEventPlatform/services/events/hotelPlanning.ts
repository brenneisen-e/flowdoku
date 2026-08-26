/**
 * v30.6 — Modularisierung Stufe 2: Thema „Hotel-Planung" (v28.38).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import type { EventService } from '../EventService';
import { REG_LIST_NAME, HOTEL_COLS_READY } from '../EventService';

/**
 * v28.38: Spalten für die Hotel-Zuordnung auf einer Teilnehmerliste anlegen.
 * Die Zuordnung gehört bewusst an die TEILNEHMERZEILE und nicht ans Event:
 * so steht sie in der Teilnehmertabelle, läuft in jeden bestehenden Export
 * mit und blaeht den Event-Datensatz nicht auf (2-MB-Grenze, s. v28.31).
 * Idempotent — vorhandene Spalten liefern 500/400 und werden ignoriert.
 */
export async function ensureHotelColumns(svc: EventService, subsiteUrl: string): Promise<void> {
  if (!subsiteUrl) return;
  // v28.61: Nur EINMAL je Liste und Sitzung. Vorher lief das vor jeder
  // einzelnen Zuordnung — drei POSTs, die alle mit „Feld existiert bereits"
  // scheiterten, bevor überhaupt geschrieben wurde. Beim Umstellen eines
  // Hotels in der Personenliste war genau das die spuerbare Verzoegerung.
  if (HOTEL_COLS_READY.has(subsiteUrl)) return;
  const base = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`;
  const fields: Array<{ title: string; type: number }> = [
    { title: 'Hotel', type: 2 },       // Text: Name des Hotels
    { title: 'HotelFrom', type: 4 },   // DateTime: Anreise
    { title: 'HotelTo', type: 4 },     // DateTime: Abreise
  ];
  for (const f of fields) {
    try {
      await svc._post(base, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
      });
    } catch { /* existiert bereits oder keine Rechte — beides unkritisch */ }
  }
  HOTEL_COLS_READY.add(subsiteUrl);
}

/**
 * v28.38: Hotel-Zuordnung einer einzelnen Teilnehmerzeile setzen oder löschen
 * (leeres Hotel = Zuordnung aufheben). Liefert true bei Erfolg.
 */
export async function setHotelAssignment(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  hotel: string,
  fromIso: string,
  toIso: string,
): Promise<boolean> {
  if (!subsiteUrl || !itemId) return false;
  try {
    const resp = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      {
        'Hotel': hotel || '',
        'HotelFrom': hotel ? (fromIso || null) : null,
        'HotelTo': hotel ? (toIso || null) : null,
      },
    );
    return resp.ok || resp.status === 406;
  } catch { return false; }
}
