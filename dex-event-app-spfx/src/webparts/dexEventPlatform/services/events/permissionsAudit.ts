/**
 * v30.66 — Modularisierung Stufe 2: Thema „Berechtigungs-Audit" (Item-Level-
 * Security der Queue-Listen, Site-weiter Einzel-Freigaben-Scan, verwaiste
 * Subsites). Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 *
 * Die vier `_`-Helfer (Rollenzuweisungen lesen/löschen, Listen und Subsites
 * eines Webs auflisten) haben ausserhalb dieses Themas keinen Aufrufer und
 * bleiben deshalb modul-intern.
 */

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { EventService, OrphanScanResult, PermCleanupFinding, PermCleanupReport } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

/**
 * v21: Item-Level-Security („nur eigene Elemente", 2/2) auf den globalen
 * Queue-Listen DEX_Outlook + DEX_IDReorder nachziehen. DEX_Emails und
 * DEX_AccessFix haben sie bereits; DEX_TeamJoinRequests bekommt sie
 * bewusst NICHT (der Team-Lead muss fremde Beitritts-Anfragen lesen).
 * Idempotent; wird vom Admin-Reparatur-Button mit ausgeführt.
 */
export async function hardenQueueListsIls(svc: EventService): Promise<{ fixed: string[]; failed: string[] }> {
  const targets = ['DEX_Outlook', 'DEX_IDReorder'];
  const fixed: string[] = [];
  const failed: string[] = [];
  for (const listName of targets) {
    try {
      // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400).
      const st = await svc._setListSecurity(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 2, WriteSecurity: 2 });
      if (st >= 200 && st < 300) fixed.push(listName); else failed.push(listName);
    } catch { failed.push(listName); }
  }
  return { fixed, failed };
}

// ==================== v26.79: Berechtigungen aufräumen ====================
// Scannt die GESAMTE Site-Collection (Haupt-Web + alle Listen/Bibliotheken +
// alle Subsites + deren Listen) und findet EINZEL-Freigaben (direkte
// Nutzer-Berechtigungen), die über das Rollen-Konzept hinaus SCHREIB-/Vollzugriff
// geben. Soll-Konzept: Schreiben nur über die Gruppen (Owners, Members,
// Visitors/DEALL) plus die im Rollen-Konzept vorgesehenen Einzelpersonen
// (Admins global; Organizer auf DEX_Events/Web-Root + ihren eigenen
// Event-Subsites). Alle anderen direkten Nutzer-Schreibrechte gelten als
// „manuelle Über-Freigabe" und werden entfernt — LESERECHTE bleiben erhalten
// (die betroffene Person liest weiterhin über die Visitors-/DEALL-Gruppe;
// internationale Leser sind bewusst OK). Gruppen werden NIE angefasst.
//
// Zusätzlich wird die Element-Sicherheit (ReadSecurity/WriteSecurity=2 „nur
// eigene Elemente") auf den sensiblen Listen (Teilnehmer, DEX_Emails,
// DEX_Outlook, DEX_IDReorder) geprüft und korrigiert.
//
// apply=false → reiner Prüf-/Dry-Run-Bericht (ändert NICHTS).
// apply=true  → entfernt die Über-Freigaben und korrigiert die Element-Sicherheit.
export async function auditOrCleanupPermissions(
  svc: EventService,
  apply: boolean,
  ctx: { adminEmails: string[]; organizerEmails: string[]; subsiteOrganizers: Record<string, string>; selfEmail: string },
  onProgress?: (msg: string, done: number, total: number) => void
): Promise<PermCleanupReport> {
  const report: PermCleanupReport = {
    apply, websScanned: 0, listsScanned: 0, strayWriteFound: 0, strayWriteRemoved: 0,
    ilsIssues: 0, ilsFixed: 0, errors: 0, findings: [],
  };
  const MAX_FINDINGS = 800;
  const addFinding = (f: PermCleanupFinding): void => { if (report.findings.length < MAX_FINDINGS) report.findings.push(f); };
  const norm = (e: string): string => (e || '').trim().toLowerCase();
  const globalAllowed = new Set([...(ctx.adminEmails || []), ...(ctx.organizerEmails || []), ctx.selfEmail].map(norm).filter(Boolean));
  // 1073741825 = Limited Access (System-verwaltet), 1073741826 = Read → beide
  // gelten NICHT als „Schreibrecht". Alles andere (Contribute/Edit/Design/Full
  // Control + Custom-Level) zählt als elevated.
  const isElevated = (ids: number[]): boolean => ids.some(id => id !== 1073741825 && id !== 1073741826);
  const extractEmail = (login: string): string => {
    const l = login || '';
    const m = /\|membership\|([^|]+)$/i.exec(l) || /\|([^|]+@[^|]+)$/.exec(l);
    return m ? m[1] : '';
  };
  const isSystemPrincipal = (email: string, login: string): boolean => {
    const l = (login || '').toLowerCase();
    if (!email) return true; // ohne E-Mail nicht sicher klassifizierbar → nicht anfassen
    if (email.indexOf('@sharepoint') >= 0) return true;
    if (l.indexOf('app@sharepoint') >= 0 || l.indexOf('|spo-grid') >= 0 || l.indexOf('c:0(.s|true') >= 0) return true;
    return false;
  };

  const processSecurable = async (scopeBase: string, label: string, allowed: Set<string>): Promise<void> => {
    let assigns: Array<{ pid: number; type: number; title: string; login: string; email: string; roleIds: number[]; roleNames: string[] }>;
    try {
      assigns = await _readRoleAssignments(svc, scopeBase);
    } catch {
      report.errors++;
      addFinding({ scope: label, kind: 'error', detail: 'Berechtigungen konnten nicht gelesen werden.', fixed: false });
      return;
    }
    for (const a of assigns) {
      if (a.type !== 1) continue; // nur einzelne User (Gruppen NIE anfassen)
      const email = norm(a.email || extractEmail(a.login));
      if (isSystemPrincipal(email, a.login)) continue;
      if (allowed.has(email)) continue; // Admin/Organizer/Ich → legitim
      if (!isElevated(a.roleIds)) continue; // reine Leseberechtigung → OK (int. Leser)
      report.strayWriteFound++;
      let fixed = false;
      if (apply) {
        try {
          const r = await _deletePrincipalAssignment(svc, scopeBase, a.pid);
          fixed = r.ok || r.status === 200 || r.status === 204;
          if (fixed) report.strayWriteRemoved++; else report.errors++;
        } catch { report.errors++; }
      }
      addFinding({
        scope: label, kind: 'stray-write', principal: email || a.title,
        detail: `${(a.roleNames.filter(Boolean).join(', ') || 'Schreibzugriff')} — ${apply ? (fixed ? 'Schreibrecht entfernt (Lesen bleibt, sofern Gruppenmitglied)' : 'Entfernen fehlgeschlagen') : 'würde entfernt (Lesen bleibt über Gruppe)'}`,
        fixed,
      });
    }
  };

  // Liest ReadSecurity/WriteSecurity einer Liste (roh + geparst). Loggt bei
  // Bedarf die exakte Server-Antwort — Diagnose, warum eine Korrektur ggf.
  // nicht greift (Format/Stale/Rechte).
  const readIls = async (listBase: string, label: string): Promise<{ rs: number; ws: number; raw: string; status: number } | null> => {
    try {
      const resp = await svc._sp.get(
        `${listBase}?$select=ReadSecurity,WriteSecurity`, SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (!resp.ok) { console.warn('[DEX PermFix] ILS-Read HTTP', resp.status, label); return null; }
      const d = await resp.json();
      const raw = JSON.stringify(d).slice(0, 200);
      return { rs: Number(d.ReadSecurity ?? d.d?.ReadSecurity), ws: Number(d.WriteSecurity ?? d.d?.WriteSecurity), raw, status: resp.status };
    } catch (e) { console.warn('[DEX PermFix] ILS-Read ERROR', label, e); return null; }
  };
  const checkIls = async (listBase: string, label: string): Promise<void> => {
    const before = await readIls(listBase, label);
    if (!before || !Number.isFinite(before.rs) || !Number.isFinite(before.ws)) {
      console.warn('[DEX PermFix] ILS unlesbar/format', label, 'raw=', before?.raw);
      return;
    }
    if (before.rs === 2 && before.ws === 2) return;
    report.ilsIssues++;
    console.warn(`[DEX PermFix] ILS FALSCH ${before.rs}/${before.ws} (soll 2/2)`, label, 'raw=', before.raw);
    let fixed = false;
    if (apply) {
      // v26.87: zuverlässiger nometadata-MERGE (Digest + kein __metadata).
      // Der bisherige verbose+__metadata-MERGE gab unter SPFx odata-version
      // 3.0 flächendeckend HTTP 400 zurück → nichts wurde je korrigiert.
      const mergeStatus = await svc._setListSecurity(listBase, { ReadSecurity: 2, WriteSecurity: 2 });
      // ENTSCHEIDEND: Read-back — nur wenn der Server danach WIRKLICH 2/2
      // meldet, gilt es als korrigiert (nicht blind dem MERGE-Status trauen).
      const after = await readIls(listBase, label);
      fixed = !!after && after.rs === 2 && after.ws === 2;
      console.warn(`[DEX PermFix] ILS-FIX ${label} | MERGE-Status=${mergeStatus} | nachher=${after ? `${after.rs}/${after.ws}` : 'null'} | raw=${after?.raw} | => ${fixed ? 'OK' : 'WEITER FALSCH'}`);
      if (fixed) report.ilsFixed++; else report.errors++;
    }
    addFinding({ scope: label, kind: 'ils', detail: `Element-Sicherheit ${before.rs}/${before.ws} statt 2/2 („nur eigene Elemente") — ${apply ? (fixed ? 'korrigiert' : 'Korrektur fehlgeschlagen (siehe Konsole)') : 'würde korrigiert'}`, fixed });
  };
  const ILS_LISTS = new Set(['DEX_Emails', 'DEX_Outlook', 'DEX_IDReorder']);

  // ---- 1. Haupt-Web (Site-Root) ----
  onProgress?.('Hauptseite …', 0, 1);
  await processSecurable(`${svc.siteUrl}/_api/web`, 'Hauptseite (Web-Root)', globalAllowed);
  report.websScanned++;

  // ---- 2. Listen/Bibliotheken des Haupt-Webs ----
  const rootLists = await _listSecurables(svc, svc.siteUrl);
  // Alle Subsites der Collection einsammeln (BFS, max. 3 Ebenen, gedeckelt) —
  // Event-Subsites hängen direkt unter dem Root, tiefere Verschachtelung wird
  // vorsorglich mitgenommen, damit wirklich der GANZE SharePoint erfasst ist.
  const subwebs: Array<{ url: string; serverRel: string; title: string; unique: boolean }> = [];
  const seenWebs = new Set<string>([svc.siteUrl.toLowerCase().replace(/\/+$/, '')]);
  let frontier = [svc.siteUrl];
  for (let depth = 0; depth < 3 && frontier.length > 0 && subwebs.length < 500; depth++) {
    const next: string[] = [];
    for (const wurl of frontier) {
      const kids = await _childWebs(svc, wurl);
      for (const k of kids) {
        const kkey = k.url.toLowerCase().replace(/\/+$/, '');
        if (!k.url || seenWebs.has(kkey)) continue;
        seenWebs.add(kkey);
        subwebs.push(k);
        next.push(k.url);
        if (subwebs.length >= 500) break;
      }
      if (subwebs.length >= 500) break;
    }
    frontier = next;
  }
  const total = 1 + rootLists.length + subwebs.length;
  let done = 1;
  for (const l of rootLists) {
    onProgress?.(`Liste ${l.title} …`, done, total);
    if (l.unique) { await processSecurable(l.base, `Liste ${l.title}`, globalAllowed); report.listsScanned++; }
    if (ILS_LISTS.has(l.title)) await checkIls(l.base, `Liste ${l.title}`);
    done++;
  }

  // ---- 3. Subsites (Event-Subsites) + deren Listen ----
  for (const w of subwebs) {
    const wlabel = w.title || w.serverRel || w.url;
    onProgress?.(`Subsite ${wlabel} …`, done, total);
    const key = (s: string): string => norm(s).replace(/\/+$/, '');
    const orgStr = ctx.subsiteOrganizers[key(w.serverRel)] || ctx.subsiteOrganizers[key(w.url)] || '';
    const allowed = new Set(globalAllowed);
    orgStr.split(/[;,]/).map(norm).filter(Boolean).forEach(e => allowed.add(e));
    if (w.unique) { await processSecurable(`${w.url}/_api/web`, `Subsite ${wlabel} – Web`, allowed); }
    report.websScanned++;
    try {
      const subLists = await _listSecurables(svc, w.url);
      for (const sl of subLists) {
        if (sl.unique) { await processSecurable(sl.base, `Subsite ${wlabel} – Liste ${sl.title}`, allowed); report.listsScanned++; }
        if (sl.title === REG_LIST_NAME) await checkIls(sl.base, `Subsite ${wlabel} – Teilnehmerliste`);
      }
    } catch { /* Subsite-Listen nicht lesbar */ }
    done++;
  }
  onProgress?.('Fertig', total, total);
  return report;
}

/** Liest die Rollenzuweisungen eines Securables (Web oder Liste). scopeBase =
 *  voll-qualifizierte API-URL bis zum Securable (…/_api/web bzw.
 *  …/_api/web/lists/getbytitle('X')). */
async function _readRoleAssignments(svc: EventService, scopeBase: string): Promise<Array<{ pid: number; type: number; title: string; login: string; email: string; roleIds: number[]; roleNames: string[] }>> {
  // $top hoch setzen: ein über-freigegebenes Securable (genau der Fall, den
  // wir suchen) kann viele Einzel-Zuweisungen haben — ohne $top würde OData
  // bei 100 abschneiden und Über-Freigaben stillschweigend übersehen.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const rows: any[] = [];
  let url: string | null = `${scopeBase}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/LoginName,Member/PrincipalType,Member/Email,RoleDefinitionBindings/Id,RoleDefinitionBindings/Name&$top=2000`;
  let guard = 0;
  while (url && guard < 20) {
    guard++;
    const resp: SPHttpClientResponse = await svc._sp.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
    if (!resp.ok) { if (guard === 1) throw new Error(`roleassignments ${resp.status}`); break; }
    const d = await resp.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const page: any[] = d.value || d.d?.results || [];
    for (const p of page) rows.push(p);
    // Falls SharePoint doch paginiert: nextLink verfolgen (nometadata-Feldname).
    url = d['odata.nextLink'] || d['@odata.nextLink'] || null;
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return rows.map((ra: any) => {
    const m = ra.Member || {};
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const binds: any[] = ra.RoleDefinitionBindings || ra.RoleDefinitionBindings?.results || [];
    return {
      pid: Number(ra.PrincipalId),
      type: Number(m.PrincipalType) || 0,
      title: m.Title || '',
      login: m.LoginName || '',
      email: m.Email || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      roleIds: binds.map((b: any) => Number(b.Id)),
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      roleNames: binds.map((b: any) => b.Name || ''),
    };
  });
}

/** Entfernt ALLE Rollenzuweisungen eines Principals auf einem Securable
 *  (Downgrade auf „kein direktes Recht" — Leserecht über Gruppen bleibt). */
async function _deletePrincipalAssignment(svc: EventService, scopeBase: string, principalId: number): Promise<SPHttpClientResponse> {
  // v26.81: Digest des Ziel-Webs mitschicken (Cross-Web-Schreibzugriff auf
  // Subsites würde sonst mit 403 abgelehnt).
  const digest = await svc._webDigest(svc._webOf(scopeBase));
  return svc._sp.post(
    `${scopeBase}/roleassignments/getbyprincipalid(${principalId})`,
    SPHttpClient.configurations.v1,
    { headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'odata-version': '', 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', ...(digest ? { 'X-RequestDigest': digest } : {}) } }
  );
}

/** Alle Listen/Bibliotheken eines Webs mit Unique-Flag. base = API-URL zum
 *  Securable der Liste. Titel mit Sonderzeichen werden für getbytitle escaped. */
async function _listSecurables(svc: EventService, webUrl: string): Promise<Array<{ title: string; hidden: boolean; unique: boolean; base: string }>> {
  try {
    const resp = await svc._sp.get(
      `${webUrl}/_api/web/lists?$select=Title,Hidden,HasUniqueRoleAssignments&$top=1000`,
      SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );
    if (!resp.ok) return [];
    const d = await resp.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = d.value || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return rows.map((l: any) => ({
      title: l.Title || '',
      hidden: !!l.Hidden,
      unique: !!l.HasUniqueRoleAssignments,
      base: `${webUrl}/_api/web/lists/getbytitle('${String(l.Title || '').replace(/'/g, "''")}')`,
    })).filter(l => l.title);
  } catch { return []; }
}

/** Direkte Kind-Webs eines Webs (eine Ebene — Event-Subsites hängen direkt
 *  unter dem Root). */
async function _childWebs(svc: EventService, webUrl: string): Promise<Array<{ url: string; serverRel: string; title: string; unique: boolean }>> {
  try {
    const resp = await svc._sp.get(
      `${webUrl}/_api/web/webs?$select=Url,ServerRelativeUrl,Title,HasUniqueRoleAssignments&$top=1000`,
      SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );
    if (!resp.ok) return [];
    const d = await resp.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = d.value || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return rows.map((w: any) => ({ url: w.Url || '', serverRel: w.ServerRelativeUrl || '', title: w.Title || '', unique: !!w.HasUniqueRoleAssignments })).filter(w => w.url);
  } catch { return []; }
}

/** v30.67: SubsiteUrls ALLER DEX_Events-Zeilen, strikt gelesen: jeder
 *  HTTP-Fehler, Netzwerkfehler oder ein abgebrochenes Paging WIRFT, statt
 *  eine halbe Liste zurückzugeben. `total` zählt alle Zeilen (auch ohne
 *  SubsiteUrl), damit der Aufrufer „Liste leer" von „nichts referenziert"
 *  unterscheiden kann. */
async function _referencedSubsiteUrlsStrict(svc: EventService): Promise<{ total: number; subsiteUrls: string[] }> {
  const out: string[] = [];
  let total = 0;
  let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Id,SubsiteUrl&$top=5000`;
  let guard = 0;
  while (url) {
    guard++;
    if (guard > 20) throw new Error('DEX_Events: mehr als 20 Seiten — Lesen abgebrochen, Prüfung nicht möglich.');
    let resp: SPHttpClientResponse;
    try {
      resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    } catch (e) {
      throw new Error(`DEX_Events konnte nicht gelesen werden (${e instanceof Error ? e.message : String(e)}) — Prüfung abgebrochen, sonst gälte jede Subsite als verwaist.`);
    }
    if (!resp.ok) {
      throw new Error(`DEX_Events konnte nicht gelesen werden (HTTP ${resp.status}) — Prüfung abgebrochen, sonst gälte jede Subsite als verwaist.`);
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: any;
    try { data = await resp.json(); } catch { throw new Error('DEX_Events: Antwort nicht lesbar — Prüfung abgebrochen.'); }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: any[] = data.value || data.d?.results || [];
    for (const it of items) {
      total++;
      const u = String(it.SubsiteUrl || '').trim();
      if (u) out.push(u);
    }
    url = data['odata.nextLink'] || data['@odata.nextLink'] || (data.d && data.d.__next) || null;
  }
  return { total, subsiteUrls: out };
}

// ==================== v26.81: Verwaiste Subsites finden ====================
// Vergleicht alle real existierenden Subsites (Webs) mit den von DEX_Events
// referenzierten SubsiteUrls. Ein Web, das von KEINEM Event (Haupt- oder
// Sub-Event) referenziert wird, ist ein „Rest" — z.B. eine Test-Subsite,
// deren Event bereits gelöscht wurde. Reine Analyse, ändert NICHTS.
export async function findOrphanSubsites(
  svc: EventService,
  onProgress?: (msg: string, done: number, total: number) => void
): Promise<OrphanScanResult> {
  const result: OrphanScanResult = { websScanned: 0, eventSubsites: 0, orphans: [] };
  const norm = (s: string): string => (s || '').trim().toLowerCase().replace(/\/+$/, '');

  // 1. Referenzierte Subsites aus DEX_Events (alle Events inkl. Sub-Events).
  //
  // v30.67: STRIKT lesen. `getAllEventsForKpi` bricht bei `!resp.ok` die
  // Schleife ab und liefert die bis dahin gelesenen Zeilen — bei einem 429
  // (GETs laufen seit v29.50 bewusst ohne Retry) oder 403 auf der ersten
  // Seite also `[]`. `referenced` blieb leer, und damit galt JEDE der bis zu
  // 500 Subsites als verwaist — jede Live-Teilnehmerliste mit „Endgültig
  // löschen"-Knopf daneben. Die beiden anderen Aufrufer (`getKpiTotals`,
  // `recomputeEventKpiOnly`) schützen sich mit `if (all.length === 0)`; hier
  // fehlte der Wächter. Jetzt: Lesefehler → Abbruch mit Fehler (die UI zeigt
  // ihn), und eine wirklich leere Event-Liste ist ein Sonderfall, der
  // ebenfalls KEINE Löschliste erzeugt.
  onProgress?.('Events werden gelesen …', 0, 1);
  const eventRows = await _referencedSubsiteUrlsStrict(svc);
  if (eventRows.total === 0) {
    throw new Error('DEX_Events enthält keine Events — die Prüfung wird nicht ausgeführt, weil dann jede Subsite als verwaist gälte. Bitte zuerst prüfen, ob die Liste wirklich leer ist.');
  }
  const referenced = new Set<string>();
  for (const u of eventRows.subsiteUrls) {
    referenced.add(norm(u));
    try { referenced.add(norm(new URL(u).pathname)); } catch { /* */ }
  }
  result.eventSubsites = referenced.size;

  // 2. Alle Subsites einsammeln (BFS, max. 3 Ebenen, cap 500).
  const allWebs: Array<{ url: string; serverRel: string; title: string }> = [];
  const seen = new Set<string>([norm(svc.siteUrl)]);
  let frontier = [svc.siteUrl];
  for (let depth = 0; depth < 3 && frontier.length > 0 && allWebs.length < 500; depth++) {
    const next: string[] = [];
    for (const wurl of frontier) {
      const kids = await _childWebs(svc, wurl);
      for (const k of kids) {
        const kk = norm(k.url);
        if (!k.url || seen.has(kk)) continue;
        seen.add(kk);
        allWebs.push({ url: k.url, serverRel: k.serverRel, title: k.title });
        next.push(k.url);
        if (allWebs.length >= 500) break;
      }
      if (allWebs.length >= 500) break;
    }
    frontier = next;
  }
  result.websScanned = allWebs.length;

  // 3. Nicht referenzierte Webs = Rest-Kandidaten; Metadaten nachladen.
  const candidates = allWebs.filter(w => !referenced.has(norm(w.url)) && !referenced.has(norm(w.serverRel)));
  let i = 0;
  for (const w of candidates) {
    i++;
    onProgress?.(`Prüfe „${w.title || w.serverRel}" …`, i, candidates.length);
    let created = '';
    try {
      const wr = await svc._sp.get(`${w.url}/_api/web?$select=Created,Title`, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
      if (wr.ok) { const wd = await wr.json(); created = wd.Created || wd.d?.Created || ''; }
    } catch { /* */ }
    let hasParticipantList = false;
    let participantCount = 0;
    try {
      const lr = await svc._sp.get(`${w.url}/_api/web/lists/getbytitle('${REG_LIST_NAME}')?$select=ItemCount`, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
      if (lr.ok) { const ld = await lr.json(); hasParticipantList = true; participantCount = Number(ld.ItemCount ?? ld.d?.ItemCount) || 0; }
    } catch { /* Liste fehlt → kein Event-Rest oder anders strukturiert */ }
    result.orphans.push({ url: w.url, serverRel: w.serverRel, title: w.title, created, hasParticipantList, participantCount });
  }
  onProgress?.('Fertig', candidates.length, candidates.length);
  return result;
}

/** Löscht eine (verwaiste) Subsite endgültig — inkl. aller Listen darin.
 *  Nur für Admins (Owner-Rechte nötig). SharePoint verlangt, dass das Web
 *  keine eigenen Unter-Webs mehr hat. */
export async function deleteSubsiteWeb(svc: EventService, webUrl: string): Promise<boolean> {
  try {
    const digest = await svc._webDigest(webUrl);
    const resp = await svc._sp.post(
      `${webUrl}/_api/web`, SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'odata-version': '', 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', ...(digest ? { 'X-RequestDigest': digest } : {}) } }
    );
    return resp.ok || resp.status === 200 || resp.status === 204;
  } catch { return false; }
}
