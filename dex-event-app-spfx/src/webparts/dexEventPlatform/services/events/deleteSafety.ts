/**
 * v31.62 — Löschen, das man zurückholen kann.
 *
 * Anlass 15.09.2026: Der Recreate-Pfad des Wizards entfernte eine
 * DEX_Events-Zeile per REST-DELETE. Das landet NICHT im Papierkorb — die
 * Zeile war weg, und mit ihr der einzige Verweis (`CalendarLink`) auf einen
 * Outlook-Termin mit 64 Eingeladenen. Die Reparatur-Aktion „Outlook-Termin
 * wiederfinden" (v31.49) suchte im Papierkorb und fand nichts.
 *
 * Das Audit danach (Agent, 15.09.2026): Nur `deleteEvent` und
 * `deleteParticipantData` recycelten; rund zwanzig weitere Stellen löschten
 * hart — Teilnehmerzeilen samt Antworten, Subsites, Dateien, Register.
 *
 * Zwei Regeln, die hier zentral stehen:
 *  1. **Recyceln statt löschen.** SharePoint bietet `/recycle` auf Listen-
 *     zeilen, Webs, Ordnern und Dateien und `/recycleObject` auf Anhängen —
 *     93 Tage Papierkorb, ein Admin holt es zurück. Kostet nichts.
 *  2. **Vor dem unumkehrbaren Schritt einen Schnappschuss.** Auch der
 *     Papierkorb ist endlich (93 Tage, Speicherlimit); die Zeile als JSON im
 *     Änderungsprotokoll (`DEX_ChangeLog.Details`, 30 000 Zeichen) lässt
 *     sich von Hand wiederherstellen. Große Felder (Base64, HTML-Bodies)
 *     werden gekürzt — die Schlüssel (CalendarLink, SubsiteUrl, EventNumber,
 *     ParentEventId, CustomData) bleiben immer drin.
 *
 * Bewusst NICHT recycelt: flüchtige Zeilen (Bearbeitungs-Präsenz, Locks),
 * Zeilen, deren Löschen der Zweck ist (Organizer-Archiv), Rollbacks.
 */
import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

/** Felder, die einen Schnappschuss sprengen würden — werden auf eine Länge gekappt. */
const GROSSE_FELDER = ['OutlookBody', 'EmailTemplateOverrides', 'EmailImageBase64', 'Description', 'BodyHtml', 'Attachments', 'Agenda', 'Documents', 'FunZone', 'Quiz'];
const KAPPUNG = 400;
const SNAPSHOT_MAX = 28000;

/** Zeile für den Schnappschuss eindampfen: Metadaten raus, Riesenfelder kappen. */
export function kompakteZeile(row: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const k of Object.keys(row)) {
    if (k.indexOf('odata') >= 0 || k === '__metadata' || k.indexOf('OData__') === 0) continue;
    const v = row[k];
    if (v === null || v === undefined || v === '') continue;
    if (typeof v === 'string' && (GROSSE_FELDER.indexOf(k) >= 0 || v.length > 2000)) {
      out[k] = v.length > KAPPUNG ? `${v.slice(0, KAPPUNG)}… [gekürzt, ${v.length} Zeichen]` : v;
    } else {
      out[k] = v;
    }
  }
  let json = JSON.stringify(out);
  if (json.length > SNAPSHOT_MAX) {
    // Noch zu groß: alles Lange raus, nur Skalare behalten.
    for (const k of Object.keys(out)) {
      if (typeof out[k] === 'string' && (out[k] as string).length > 120) delete out[k];
    }
    json = JSON.stringify(out);
  }
  return out;
}

/** Eine Listenzeile roh lesen (nometadata) — null, wenn nicht lesbar. */
export async function leseZeile(svc: EventService, itemUrl: string): Promise<Record<string, unknown> | null> {
  try {
    const r = await svc._sp.get(itemUrl, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
    if (!r.ok) return null;
    return await r.json() as Record<string, unknown>;
  } catch {
    return null;
  }
}

export interface SchnappschussInput {
  action: string;
  targetType: 'Event' | 'Participant' | 'Subsite' | 'Other';
  targetId: string;
  targetName?: string;
  eventId?: string;
  eventTitle?: string;
  /** Wo die Zeile lag — für die Wiederherstellung von Hand. */
  quelle: string;
  row: Record<string, unknown> | null;
  grund?: string;
}

/**
 * Schnappschuss ins Änderungsprotokoll — best-effort, wirft nie. Ohne
 * lesbare Zeile wird das protokolliert (`snapshot: null`), damit man später
 * weiß, dass es keinen gab, statt es zu vermuten.
 */
export async function schnappschussVorLoeschen(svc: EventService, s: SchnappschussInput): Promise<void> {
  try {
    await svc.writeChangeLog({
      action: s.action,
      targetType: s.targetType,
      targetId: s.targetId,
      targetName: s.targetName || (s.row && typeof s.row.Title === 'string' ? s.row.Title : ''),
      eventId: s.eventId,
      eventTitle: s.eventTitle,
      details: {
        quelle: s.quelle,
        grund: s.grund || '',
        wiederherstellung: 'Zeile liegt 93 Tage im Papierkorb der Site (Site contents → Recycle bin); danach nur noch aus diesem Schnappschuss von Hand.',
        snapshot: s.row ? kompakteZeile(s.row) : null,
      },
    });
  } catch { /* Protokoll ist best-effort — das Löschen darf daran nicht scheitern */ }
}

/** Listenzeile in den Papierkorb (statt DELETE). `itemUrl` = …/items(id). */
export async function recycleZeile(svc: EventService, itemUrl: string): Promise<boolean> {
  try {
    const r = await svc._post(`${itemUrl}/recycle`, {});
    return r.ok || r.status === 200 || r.status === 204;
  } catch {
    return false;
  }
}

/** Anhang in den Papierkorb (statt DELETE). `attachmentUrl` = …/AttachmentFiles/getByFileName('x'). */
export async function recycleAnhang(svc: EventService, attachmentUrl: string): Promise<boolean> {
  try {
    const r = await svc._post(`${attachmentUrl}/recycleObject`, {});
    return r.ok || r.status === 200 || r.status === 204;
  } catch {
    return false;
  }
}
