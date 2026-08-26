/**
 * v30.12 — Modularisierung Stufe 2: Thema „Organizer-Verwaltung".
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 *
 * Drei zusammengehörige Stücke: das Organizer-Archiv (DEX_OrganizerArchived,
 * reiner Anzeige-Filter „Event aus MEINER Übersicht ausblenden", v24.6), die
 * Organizer-Anträge (DEX_OrganizerRequests, v23.37 — Admins bestätigen in
 * der App, inkl. Site-Leserecht-Vergabe) und die Rollen-Abfragen auf
 * DEX_Roles (roleFilter kennt die Legacy-Schreibweisen) samt der
 * Wochenbericht-Zähler (Events/Anmeldungen seit Datum).
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

export async function ensureOrganizerRequestsList(svc: EventService): Promise<void> {
  const listName = 'DEX_OrganizerRequests';
  const exists = await svc.listExists(listName);
  if (exists) return;
  const cr = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Anträge „Organizer werden" (v23.37) — Admins bestätigen in der App.',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  if (!cr.ok) { console.warn('[DEX] DEX_OrganizerRequests konnte nicht angelegt werden.'); return; }
  const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
    { title: 'RequesterEmail', type: 2 },
    { title: 'RequesterName', type: 2 },
    { title: 'RequesterLocation', type: 2 },
    { title: 'Message', type: 3 },
    { title: 'Status', type: 6, choices: ['Pending', 'Approved', 'Rejected'], metaType: 'SP.FieldChoice' },
    { title: 'DecidedDate', type: 4 },
    { title: 'DecidedByEmail', type: 2 },
  ];
  for (const f of fields) {
    const payload: Record<string, unknown> = { '__metadata': { 'type': f.metaType || 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false };
    if (f.choices) payload['Choices'] = { 'results': f.choices };
    try { await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload); } catch { /* */ }
  }
  try { await svc.configureDefaultView(listName, ['RequesterName', 'RequesterEmail', 'RequesterLocation', 'Status', 'Created', 'DecidedByEmail']); } catch { /* */ }
  try { await svc.setQueueListPermissions(listName); } catch { /* */ }
}

// ==================== v24.6: Organizer-Archiv (pro Person ausblenden) ====================
// Reiner Anzeige-Filter: ein abgelaufenes Event kann der Organizer aus SEINER
// Übersicht ausblenden (eine Zeile pro Event+Person). Das Event selbst bleibt
// mit allen Daten erhalten und für andere sichtbar — KEINE Datenlöschung.
export async function ensureOrganizerArchivedList(svc: EventService): Promise<void> {
  const listName = 'DEX_OrganizerArchived';
  const exists = await svc.listExists(listName);
  if (exists) return;
  const cr = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Pro Organizer ausgeblendete (archivierte) Events (v24.6) — reiner Anzeige-Filter, keine Datenlöschung.',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  if (!cr.ok) { console.warn('[DEX] DEX_OrganizerArchived konnte nicht angelegt werden.'); return; }
  for (const f of [{ title: 'EventId', type: 2 }, { title: 'OrganizerEmail', type: 2 }]) {
    try { await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, { '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false }); } catch { /* */ }
  }
  try { await svc.configureDefaultView(listName, ['EventId', 'OrganizerEmail', 'Created']); } catch { /* */ }
  try { await svc.setQueueListPermissions(listName); } catch { /* */ }
}

export async function getOrganizerArchivedEventIds(svc: EventService, email: string): Promise<Set<string>> {
  const out = new Set<string>();
  try {
    const e = (email || '').replace(/'/g, "''");
    if (!e) return out;
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items?$select=Id,EventId,OrganizerEmail&$filter=OrganizerEmail eq '${e}'&$top=2000`,
      SPHttpClient.configurations.v1);
    if (resp.ok) {
      const d = await resp.json();
      const rows = d.value || d.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const r of rows) { const id = String((r as any).EventId || ''); if (id) out.add(id); }
    }
  } catch { /* best-effort */ }
  return out;
}

export async function archiveEventForOrganizer(svc: EventService, eventId: string, email: string): Promise<boolean> {
  try {
    const resp = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_OrganizerArchivedListItem' },
      'Title': String(eventId).slice(0, 250),
      'EventId': String(eventId),
      'OrganizerEmail': email,
    });
    return resp.ok;
  } catch { return false; }
}

export async function unarchiveEventForOrganizer(svc: EventService, eventId: string, email: string): Promise<boolean> {
  try {
    const e = (email || '').replace(/'/g, "''");
    const idEsc = String(eventId).replace(/'/g, "''");
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items?$select=Id&$filter=OrganizerEmail eq '${e}' and EventId eq '${idEsc}'&$top=50`,
      SPHttpClient.configurations.v1);
    if (!resp.ok) return false;
    const d = await resp.json();
    const rows = d.value || d.d?.results || [];
    let okAll = true;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const r of rows) { const del = await svc._delete(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items(${(r as any).Id})`); if (!del.ok) okAll = false; }
    return okAll;
  } catch { return false; }
}

export async function createOrganizerRequest(svc: EventService, email: string, name: string, location: string, message: string): Promise<{ ok: boolean; itemId?: number }> {
  try {
    const resp = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerRequests')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_OrganizerRequestsListItem' },
      'Title': (name || email || 'Antrag').slice(0, 250),
      'RequesterEmail': email,
      'RequesterName': name,
      'RequesterLocation': location || '',
      'Message': message || '',
      'Status': 'Pending',
    });
    if (!resp.ok) return { ok: false };
    try { const j = await resp.json(); return { ok: true, itemId: j?.d?.Id || j?.Id || 0 }; } catch { return { ok: true }; }
  } catch { return { ok: false }; }
}

export async function getOrganizerRequests(svc: EventService, onlyPending: boolean = true): Promise<Array<{ id: number; email: string; name: string; location: string; message: string; status: string; created: string }>> {
  try {
    const filter = onlyPending ? `&$filter=Status eq 'Pending'` : '';
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerRequests')/items?$select=Id,RequesterEmail,RequesterName,RequesterLocation,Message,Status,Created&$orderby=Created desc&$top=200${filter}`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return items.map((i: any) => ({ id: i.Id, email: i.RequesterEmail || '', name: i.RequesterName || '', location: i.RequesterLocation || '', message: i.Message || '', status: i.Status || '', created: i.Created || '' }));
  } catch { return []; }
}

/** v26.58: Einzelnen Organizer-Antrag inkl. Entscheidungs-Metadaten laden —
 *  für den approveorg-Deep-Link, wenn der Antrag bereits entschieden wurde
 *  („bereits freigegeben durch X am Y" statt kommentarlos Landing Page). */
export async function getOrganizerRequestDetails(svc: EventService, id: number): Promise<{ id: number; email: string; name: string; status: string; decidedByEmail: string; decidedDate: string } | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerRequests')/items(${id})?$select=Id,RequesterEmail,RequesterName,Status,DecidedByEmail,DecidedDate`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return null;
    const d = await resp.json();
    const it = d.d || d;
    if (!it || !it.Id) return null;
    return {
      id: Number(it.Id),
      email: String(it.RequesterEmail || ''),
      name: String(it.RequesterName || ''),
      status: String(it.Status || ''),
      decidedByEmail: String(it.DecidedByEmail || ''),
      decidedDate: String(it.DecidedDate || ''),
    };
  } catch { return null; }
}

/** v26.59: Einer Person Leserechte auf die Site geben. Weg: User im Web
 *  sicherstellen (ensureuser), dann in die Standard-Besucher-Gruppe
 *  (associatedvisitorgroup, Permission Level „Lesen") aufnehmen — Gruppen-
 *  Mitgliedschaft ist sauberer als Einzel-Berechtigungen. Fallback: direkte
 *  Read-Rollenzuweisung (RoleTypeKind=2) aufs Web, falls es keine
 *  Besucher-Gruppe gibt. Erfordert Berechtigungs-Verwaltungsrechte des
 *  Aufrufers (Admins haben Full Control). Genutzt vom grantaccess-Deep-Link
 *  aus der „SharePoint-Zugriff benötigt"-Mail. */
export async function grantSiteReadAccess(svc: EventService, email: string): Promise<boolean> {
  const mail = (email || '').trim();
  if (!mail) return false;
  try {
    const ensure = await svc._post(`${svc.siteUrl}/_api/web/ensureuser`, { 'logonName': `i:0#.f|membership|${mail}` });
    if (!ensure.ok) return false;
    const ud = await ensure.json();
    const userId = Number(ud?.d?.Id ?? ud?.Id ?? 0);
    const loginName = String(ud?.d?.LoginName ?? ud?.LoginName ?? '') || `i:0#.f|membership|${mail}`;
    if (!userId) return false;
    try {
      const vg = await svc._sp.get(
        `${svc.siteUrl}/_api/web/associatedvisitorgroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (vg.ok) {
        const vgd = await vg.json();
        const gid = Number(vgd?.Id ?? vgd?.d?.Id ?? 0);
        if (gid > 0) {
          const add = await svc._post(`${svc.siteUrl}/_api/web/sitegroups(${gid})/users`, {
            '__metadata': { 'type': 'SP.User' },
            'LoginName': loginName,
          });
          if (add.ok) return true;
        }
      }
    } catch { /* Fallback unten */ }
    const rd = await svc._sp.get(
      `${svc.siteUrl}/_api/web/roledefinitions?$filter=RoleTypeKind eq 2&$select=Id&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!rd.ok) return false;
    const rdd = await rd.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const roleId = Number(((rdd.value || rdd.d?.results || [])[0] as any)?.Id || 0);
    if (!roleId) return false;
    const ra = await svc._post(`${svc.siteUrl}/_api/web/roleassignments/addroleassignment(principalid=${userId},roledefid=${roleId})`, {});
    return ra.ok;
  } catch { return false; }
}

export async function updateOrganizerRequestStatus(svc: EventService, id: number, status: 'Approved' | 'Rejected', decidedByEmail: string): Promise<boolean> {
  try {
    const r = await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerRequests')/items(${id})`, {
      'Status': status, 'DecidedDate': new Date().toISOString(), 'DecidedByEmail': decidedByEmail,
    });
    return r.ok;
  } catch { return false; }
}

/**
 * v28.89: OData-Filter auf die Rollen-Spalte, inklusive der ALTEN Werte.
 *
 * DEX_Roles kannte früher `SuperAdmin` (heute `Admin`) und `EventAdmin`
 * (heute `Organizer`). `RoleContext.migrateRole` bildet beide weiterhin ab
 * und schreibt sie im Hintergrund um — das passiert aber nur, wenn jemand
 * die Rollenliste öffnet, und `updateRole` scheitert still (fehlende
 * Schreibrechte). Es können also dauerhaft Legacy-Zeilen stehen bleiben.
 *
 * Für die Anzeige ist das egal, für Rechte-Prüfungen nicht: `Role eq
 * 'Organizer'` findet einen Legacy-Organizer nicht, er gilt dann als „ohne
 * Rolle" — und bekommt beim Speichern eines Events, in dem er als
 * Co-Organizer steht, einen Freigabe-Antrag, obwohl er längst freigegeben
 * ist (dasselbe beim Deep-Link der DEX-Anfrage).
 */
function roleFilter(role: string): string {
  const legacy: Record<string, string[]> = { Admin: ['SuperAdmin'], Organizer: ['EventAdmin'] };
  const values = [role].concat(legacy[role] || []);
  return values
    .map(v => `Role eq '${encodeURIComponent(v.replace(/'/g, "''"))}'`)
    .join(' or ');
}

/** E-Mail-Adressen (Title) aller DEX_Roles-Einträge mit der gegebenen Rolle. */
/** v23.38: Rollen-Empfänger mit E-Mail UND Anzeigename (für personalisierte
 *  Mails wie den Wochenbericht — „Hallo <Name>" statt generisch „Admin"). */
export async function getRoleRecipients(svc: EventService, role: string): Promise<Array<{ email: string; name: string }>> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=${roleFilter(role)}&$select=Title,UserName&$top=5000`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    const seen = new Set<string>();
    const out: Array<{ email: string; name: string }> = [];
    for (const i of items) {
      const email = (i.Title || '').trim();
      if (!email) continue;
      const key = email.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ email, name: (i.UserName || '').trim() || email });
    }
    return out;
  } catch { return []; }
}

export async function getRoleEmails(svc: EventService, role: string): Promise<string[]> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=${roleFilter(role)}&$select=Title&$top=5000`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    const set = new Set<string>();
    for (const i of items) { const e = (i.Title || '').trim().toLowerCase(); if (e) set.add(e); }
    return Array.from(set);
  } catch { return []; }
}

/** DEX_Roles-Einträge einer Rolle, die seit `fromIso` neu angelegt wurden. */
export async function getRoleItemsCreatedSince(svc: EventService, role: string, fromIso: string): Promise<Array<{ email: string; created: string }>> {
  try {
    const esc = role.replace(/'/g, "''");
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Role eq '${encodeURIComponent(esc)}' and Created ge '${fromIso}'&$select=Title,Created&$orderby=Created desc&$top=500`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    return items.map((i: { Title?: string; Created?: string }) => ({ email: (i.Title || '').trim(), created: i.Created || '' }));
  } catch { return []; }
}

/** DEX_Events-Items, die seit `fromIso` erstellt wurden — mit Ersteller
 *  (SP-Author) + Titel. */
export async function getEventsCreatedSince(svc: EventService, fromIso: string): Promise<Array<{ title: string; author: string; created: string; isDraft: boolean }>> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Title,Created,IsFictive,Author/Title&$expand=Author&$filter=Created ge '${fromIso}'&$orderby=Created desc&$top=500`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return items.map((i: any) => ({ title: i.Title || '(ohne Titel)', author: (i.Author && i.Author.Title) || '—', created: i.Created || '', isDraft: !!i.IsFictive }));
  } catch { return []; }
}

/** Aktive Anmeldungen einer Teilnehmer-Subsite zählen: total (alle aktiven)
 *  + since (RegistrationDate ≥ fromIso). Status-Filter = Angemeldet/QR
 *  versendet/Eingecheckt/Warteliste (keine Abgemeldeten). */
export async function countRegistrations(svc: EventService, subsiteUrl: string, fromIso: string): Promise<{ total: number; since: number }> {
  if (!subsiteUrl) return { total: 0, since: 0 };
  const active = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
  let total = 0; let since = 0;
  let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status,RegistrationDate&$top=5000`;
  let guard = 0;
  const fromTs = new Date(fromIso).getTime();
  while (url && guard < 50) {
    guard++;
    try {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      for (const i of items) {
        if (active.indexOf(i.Status) < 0) continue;
        total++;
        const ts = i.RegistrationDate ? new Date(i.RegistrationDate).getTime() : 0;
        if (ts && ts >= fromTs) since++;
      }
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    } catch { break; }
  }
  return { total, since };
}
