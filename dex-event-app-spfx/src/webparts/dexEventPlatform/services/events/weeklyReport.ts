/**
 * v30.11 — Modularisierung Stufe 2: Thema „Wochenbericht" (v23.8).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 *
 * DEX_WeeklyReports ist die Tracking-Liste des wöchentlichen Admin-Berichts:
 * pro Versand eine Zeile (Created = Versandzeit, PeriodFrom/PeriodTo =
 * abgedeckter Zeitraum, DraftEventIds = bereits gemeldete Entwürfe) — die
 * Quelle der Wahrheit für „wann lief der letzte Bericht".
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

// ==================== Wochenbericht (v23.8) ====================

/** Tracking-Liste für den wöchentlichen Admin-Bericht. Pro versendetem
 *  Bericht eine Zeile (Created = Versandzeit; PeriodFrom/PeriodTo = der
 *  abgedeckte Zeitraum). Quelle der Wahrheit für „wann lief der letzte
 *  Bericht". */
export async function ensureWeeklyReportsList(svc: EventService): Promise<void> {
  const listName = 'DEX_WeeklyReports';
  const exists = await svc.listExists(listName);
  if (!exists) {
    const createResp = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Versand-Protokoll des wöchentlichen Admin-Berichts (v23.8).',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    if (!createResp.ok) {
      console.warn('[DEX] DEX_WeeklyReports konnte nicht angelegt werden — vermutlich fehlen Owner-Rechte.');
      return;
    }
    for (const f of [{ title: 'PeriodFrom', type: 4 }, { title: 'PeriodTo', type: 4 }]) {
      try {
        await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
        });
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
    try { await svc.configureDefaultView(listName, ['PeriodFrom', 'PeriodTo']); } catch { /* */ }
  }
  // v23.36: DraftEventIds-Snapshot (JSON-Array der Event-IDs, die beim letzten
  // Bericht noch Entwürfe waren) — idempotent nachziehen, auch auf Bestands-
  // Listen. So erkennt der nächste Bericht „Entwurf ist live gegangen".
  try {
    const fieldsResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$filter=InternalName eq 'DraftEventIds'&$top=1`,
      SPHttpClient.configurations.v1
    );
    const fieldsData = fieldsResp.ok ? await fieldsResp.json() : null;
    const has = fieldsData && (fieldsData.value || fieldsData.d?.results || []).length > 0;
    if (!has) {
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' }, 'Title': 'DraftEventIds', 'FieldTypeKind': 3, 'Required': false,
      });
    }
  } catch { /* best-effort */ }
}

/** Letzter Bericht: Created (Versandzeit) + PeriodTo + Entwurfs-Snapshot. */
export async function getLastWeeklyReport(svc: EventService): Promise<{ created: string; periodTo: string; draftEventIds: string[] } | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_WeeklyReports')/items?$select=Created,PeriodTo,DraftEventIds&$orderby=Created desc&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return null;
    let draftEventIds: string[] = [];
    try { const arr = JSON.parse(items[0].DraftEventIds || '[]'); if (Array.isArray(arr)) draftEventIds = arr.map((x: unknown) => String(x)); } catch { /* */ }
    return { created: items[0].Created || '', periodTo: items[0].PeriodTo || items[0].Created || '', draftEventIds };
  } catch { return null; }
}

/** Versand des Berichts protokollieren (eine Zeile) + Entwurfs-Snapshot. */
export async function recordWeeklyReport(svc: EventService, fromIso: string, toIso: string, draftEventIds?: string[]): Promise<void> {
  try {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_WeeklyReports')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_WeeklyReportsListItem' },
      'Title': `Weekly ${toIso.slice(0, 10)}`,
      'PeriodFrom': fromIso,
      'PeriodTo': toIso,
      'DraftEventIds': JSON.stringify(draftEventIds || []),
    });
  } catch (e) { console.warn('[DEX] recordWeeklyReport failed:', e); }
}

