/**
 * v31.2 — Thema „Wer bearbeitet dieses Event gerade?" (Liste DEX_EditPresence).
 *
 * Nutzer-Auftrag 07.09.2026: „Im Event-Wizard soll angezeigt werden, wenn ein
 * anderer Organizer aktuell im Edit-Modus ist — das soll pulsieren und die
 * Person mit Foto angezeigt werden."
 *
 * Warum eine eigene Liste und kein Feld auf DEX_Events: Ein Herzschlag alle
 * 20 Sekunden auf die Event-Zeile würde Versionen erzeugen, die
 * `Modified`-Anzeige verfälschen und mit dem Speichern des Wizards
 * kollidieren (IF-MATCH). Eine kleine Nebenliste kostet nichts, hat keine
 * Zeilen-Sicherheit (jeder Organizer muss die Zeilen der anderen lesen) und
 * braucht keinen Flow.
 *
 * Spalten: EventId (Text, DEX_Events-Item-Id), UserEmail (Text), UserName
 * (Text), LastSeen (DateTime, ISO). Eine Zeile je Person je Event; der
 * Herzschlag setzt nur LastSeen. „Anwesend" heißt: LastSeen jünger als
 * `PRESENCE_FRESH_MS`. Beim Verlassen löscht der Wizard die eigene Zeile
 * best-effort; bleibt sie stehen (Tab geschlossen), fällt sie nach der Frist
 * aus der Anzeige und wird beim nächsten Lesen aufgeräumt.
 *
 * Herausgelöst in ein eigenes Modul wie alle Listen-Themen; im EventService
 * stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

export const PRESENCE_LIST_NAME = 'DEX_EditPresence';
const ITEM_TYPE = 'SP.Data.DEX_x005f_EditPresenceListItem';
/** Jünger als das gilt als „gerade im Wizard" (Herzschlag alle 20 s). */
export const PRESENCE_FRESH_MS = 75 * 1000;
/** Älter als das wird beim Lesen aufgeräumt (Tab zu, Rechner zu). */
const PRESENCE_STALE_MS = 6 * 60 * 60 * 1000;

export interface EditPresenceEntry {
  itemId: number;
  email: string;
  name: string;
  lastSeen: string;
}

export async function ensureEditPresenceList(svc: EventService): Promise<void> {
  const exists = await svc.listExists(PRESENCE_LIST_NAME);
  if (exists) return;
  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': PRESENCE_LIST_NAME,
    'Description': 'Wer bearbeitet gerade welches Event im Wizard (Herzschlag, v31.2). Kein Flow — die App liest und schreibt selbst.',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  const fields: Array<{ title: string; type: number }> = [
    { title: 'EventId', type: 2 },
    { title: 'UserEmail', type: 2 },
    { title: 'UserName', type: 2 },
    { title: 'LastSeen', type: 4 },
  ];
  for (const f of fields) {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${PRESENCE_LIST_NAME}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    });
  }
  await svc.configureDefaultView(PRESENCE_LIST_NAME, ['EventId', 'UserEmail', 'UserName', 'LastSeen']);
  // Jeder Organizer schreibt seine eigene Zeile und liest die der anderen —
  // dieselbe Rechtelage wie bei den Queue-Listen, ohne Zeilen-Sicherheit.
  try { await svc.setQueueListPermissions(PRESENCE_LIST_NAME); } catch { /* best-effort */ }
}

function esc(v: string): string { return (v || '').replace(/'/g, "''"); }

/**
 * Herzschlag: eigene Zeile anlegen oder LastSeen setzen. Liefert die Item-Id
 * (für den Aufräum-DELETE beim Verlassen) oder null bei Fehler — der Aufrufer
 * versucht es beim nächsten Takt einfach wieder.
 */
export async function heartbeatEditPresence(
  svc: EventService,
  eventId: string,
  email: string,
  name: string,
  knownItemId: number | null
): Promise<number | null> {
  const base = `${svc.siteUrl}/_api/web/lists/getbytitle('${PRESENCE_LIST_NAME}')/items`;
  const now = new Date().toISOString();
  try {
    let id = knownItemId;
    if (id === null) {
      const q = `${base}?$select=Id&$filter=EventId eq '${esc(eventId)}' and UserEmail eq '${esc(email.toLowerCase())}'&$top=1`;
      const r = await svc._sp.get(q, SPHttpClient.configurations.v1);
      if (r.ok) {
        const d = await r.json();
        const rows = (d.value || d.d?.results || []) as Array<{ Id: number }>;
        if (rows.length) id = rows[0].Id;
      } else if (r.status === 404) {
        // Liste fehlt noch (erster Aufruf im Tenant) — anlegen und weiter.
        await ensureEditPresenceList(svc);
      } else {
        return null;
      }
    }
    if (id !== null) {
      const m = await svc._merge(`${base}(${id})`, { 'LastSeen': now, 'UserName': name });
      if (m.ok) return id;
      // Zeile weg (jemand hat aufgeräumt) → neu anlegen.
      if (m.status !== 404) return null;
    }
    const c = await svc._post(base, {
      '__metadata': { 'type': ITEM_TYPE },
      'Title': `${name} → ${eventId}`,
      'EventId': eventId,
      'UserEmail': email.toLowerCase(),
      'UserName': name,
      'LastSeen': now,
    });
    if (!c.ok) return null;
    const cd = await c.json();
    const newId = (cd.d?.Id ?? cd.Id) as number | undefined;
    return typeof newId === 'number' ? newId : null;
  } catch {
    return null;
  }
}

/**
 * Wer ist gerade (frisch) im Wizard dieses Events — ohne mich selbst?
 * Liefert null, wenn die Liste nicht lesbar war (kein Aussage = keine
 * Anzeige, aber auch kein „niemand da"). Sehr alte Zeilen werden dabei
 * best-effort gelöscht, höchstens fünf je Aufruf.
 */
export async function readEditPresence(
  svc: EventService,
  eventId: string,
  excludeEmail: string
): Promise<EditPresenceEntry[] | null> {
  const base = `${svc.siteUrl}/_api/web/lists/getbytitle('${PRESENCE_LIST_NAME}')/items`;
  try {
    const q = `${base}?$select=Id,UserEmail,UserName,LastSeen&$filter=EventId eq '${esc(eventId)}'&$top=100`;
    const r = await svc._sp.get(q, SPHttpClient.configurations.v1);
    if (!r.ok) return null;
    const d = await r.json();
    const rows = (d.value || d.d?.results || []) as Array<{ Id: number; UserEmail: string; UserName: string; LastSeen: string }>;
    const nowMs = Date.now();
    const me = (excludeEmail || '').toLowerCase();
    const fresh: EditPresenceEntry[] = [];
    let cleaned = 0;
    for (const row of rows) {
      const seen = row.LastSeen ? new Date(row.LastSeen).getTime() : 0;
      const age = nowMs - seen;
      if (!isFinite(age) || age > PRESENCE_STALE_MS) {
        if (cleaned < 5) { cleaned++; svc._delete(`${base}(${row.Id})`).catch(() => { /* best-effort */ }); }
        continue;
      }
      if (age > PRESENCE_FRESH_MS) continue;
      if ((row.UserEmail || '').toLowerCase() === me) continue;
      fresh.push({ itemId: row.Id, email: row.UserEmail || '', name: row.UserName || row.UserEmail || '', lastSeen: row.LastSeen });
    }
    // Eine Person mit zwei Tabs nur einmal.
    const seenMail = new Set<string>();
    return fresh.filter(e => { const k = e.email.toLowerCase(); if (seenMail.has(k)) return false; seenMail.add(k); return true; });
  } catch {
    return null;
  }
}

/** Eigene Zeile beim Verlassen entfernen — best-effort, Ergebnis egal. */
export async function clearEditPresence(svc: EventService, itemId: number): Promise<void> {
  try {
    await svc._delete(`${svc.siteUrl}/_api/web/lists/getbytitle('${PRESENCE_LIST_NAME}')/items(${itemId})`);
  } catch { /* best-effort */ }
}
