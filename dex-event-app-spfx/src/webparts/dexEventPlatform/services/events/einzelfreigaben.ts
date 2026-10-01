/**
 * v32.56 — Einzelfreigaben auf die App normalisieren (täglich, Admin-Start).
 *
 * Anlass (Nutzer-Ansage 01.10.2026): Personen außerhalb von DEALL schicken
 * über SharePoint eine Zugriffsanfrage auf `DEX.aspx`. Wer sie mit „Edit"
 * genehmigt, gibt der Person ein PERSÖNLICHES Schreibrecht auf die App-Seite
 * — und trotzdem sieht sie kein Event, weil DEX jedes Teilnehmer-Recht an die
 * Besucher-Gruppe hängt (v31.6, `isInVisitorsGroup`). Gewünscht: „ein
 * Automatismus wie das Archivieren, der regelmäßig die Rechte nachzieht bzw.
 * reduziert bei Leuten, die dazugekommen sind".
 *
 * Was der Lauf tut, je Person mit direkter Freigabe:
 *   1. NACHZIEHEN — in die Besucher-Gruppe aufnehmen und per Nachlesen
 *      bestätigen (ohne den Direkt-Lese-Rückfall von `grantSiteReadAccess`:
 *      der würde genau die Einzelfreigabe erzeugen, die hier verschwinden soll).
 *   2. REDUZIEREN — erst danach die direkte Freigabe entfernen und nachlesen,
 *      ob der Principal wirklich weg ist (CLAUDE.md: ein DELETE-Status ist bei
 *      Rollenzuweisungen nicht belastbar). Klappt Schritt 1 nicht, bleibt die
 *      Freigabe stehen — sonst stünde die Person ganz ohne Zugang da.
 *
 * Bewusst eng geschnitten: nur die Stellen, an denen Zugriffsanfragen landen —
 * das Haupt-Web (nur Schreibrechte; ein direktes Leserecht dort ist harmlos)
 * sowie die Bibliothek „Websiteseiten“ und ihre Seiten mit eigenen Rechten
 * (dort JEDE direkte Freigabe außer „Beschränkter Zugriff“ — die App-Seite
 * soll erben). Die übrigen Listen tragen Einzelrechte, die DEX selbst vergibt
 * (Listen-Ersteller, Umfragen, Check-in-Team auf den Subsites); die bleiben
 * beim großen Aufräumen im Admin Hub (`auditOrCleanupPermissions`), das ein
 * Mensch anstößt. Admins, Organizer und die angemeldete Person selbst werden
 * nie angefasst, Gruppen auch nicht. Die Mitglieder-Gruppe wird nur
 * GEMELDET: Wer dort ohne Rolle steht, hat Schreibrecht auf alles Erbende —
 * ob das gewollt ist, entscheidet ein Mensch.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { _readRoleAssignments, _deletePrincipalAssignment } from './permissionsAudit';
import { isInVisitorsGroup } from './organizer';

export type EinzelfreigabeErgebnis =
  | 'reduziert'          // in Besucher-Gruppe bestätigt, Einzelfreigabe entfernt
  | 'gefunden'           // Prüflauf (apply=false)
  | 'besucher-fehlt'     // Aufnahme in die Besucher-Gruppe nicht bestätigt → nichts entfernt
  | 'entfernen-fehlgeschlagen'
  | 'nur-gemeldet';      // Mitglieder-Gruppe

export interface EinzelfreigabeBefund {
  ort: string;
  email: string;
  name: string;
  rechte: string;
  ergebnis: EinzelfreigabeErgebnis;
}

export interface EinzelfreigabeBericht {
  apply: boolean;
  geprueft: number;
  befunde: EinzelfreigabeBefund[];
  /** Stellen, deren Rechte nicht lesbar waren — „nichts gefunden" gilt dort nicht. */
  leseFehler: string[];
}

const LIMITED_ACCESS = 1073741825;
const READ = 1073741826;

const norm = (e: string): string => (e || '').trim().toLowerCase();
const emailAusLogin = (login: string): string => {
  const m = /\|membership\|([^|]+)$/i.exec(login || '') || /\|([^|]+@[^|]+)$/.exec(login || '');
  return m ? m[1] : '';
};
const istSystem = (email: string, login: string): boolean => {
  const l = (login || '').toLowerCase();
  if (!email) return true; // ohne Adresse nicht sicher zuzuordnen → nicht anfassen
  if (email.indexOf('@sharepoint') >= 0) return true;
  return l.indexOf('app@sharepoint') >= 0 || l.indexOf('|spo-grid') >= 0 || l.indexOf('c:0(.s|true') >= 0;
};

/** Nimmt die Person in die Besucher-Gruppe auf und liest nach. Kein Rückfall
 *  auf ein direktes Leserecht (s. Kopfkommentar). */
async function inBesucherGruppe(svc: EventService, email: string): Promise<boolean> {
  if ((await isInVisitorsGroup(svc, email)) === true) return true;
  const gid = await svc.getVisitorsGroupId();
  if (!gid) return false;
  try {
    const ensure = await svc._post(`${svc.siteUrl}/_api/web/ensureuser`, { 'logonName': `i:0#.f|membership|${email}` });
    if (!ensure.ok) return false;
    const ud = await ensure.json();
    const loginName = String(ud?.d?.LoginName ?? ud?.LoginName ?? '') || `i:0#.f|membership|${email}`;
    await svc._post(`${svc.siteUrl}/_api/web/sitegroups(${gid})/users`, {
      '__metadata': { 'type': 'SP.User' },
      'LoginName': loginName,
    });
  } catch { return false; }
  return (await isInVisitorsGroup(svc, email)) === true;
}

export async function einzelfreigabenNormalisieren(
  svc: EventService,
  apply: boolean,
  ctx: { erlaubt: string[]; selfEmail: string },
): Promise<EinzelfreigabeBericht> {
  const bericht: EinzelfreigabeBericht = { apply, geprueft: 0, befunde: [], leseFehler: [] };
  const erlaubt = new Set([...ctx.erlaubt, ctx.selfEmail].map(norm).filter(Boolean));

  // Securables einsammeln: Haupt-Web + Websiteseiten-Bibliothek + ihre Seiten
  // mit eigenen Rechten. `nurSchreiben` = direktes Leserecht ist dort in Ordnung.
  const ziele: Array<{ base: string; ort: string; nurSchreiben: boolean }> = [
    { base: `${svc.siteUrl}/_api/web`, ort: 'Hauptseite (Web)', nurSchreiben: true },
  ];
  try {
    const lr = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists?$filter=BaseTemplate eq 119&$select=Id,Title,HasUniqueRoleAssignments`,
      SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    if (!lr.ok) throw new Error(String(lr.status));
    const ld = await lr.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const lib of (ld.value || []) as any[]) {
      const libBase = `${svc.siteUrl}/_api/web/lists(guid'${lib.Id}')`;
      const libTitel = String(lib.Title || 'Websiteseiten');
      if (lib.HasUniqueRoleAssignments) ziele.push({ base: libBase, ort: `Bibliothek ${libTitel}`, nurSchreiben: false });
      const ir = await svc._sp.get(
        `${libBase}/items?$select=Id,FileLeafRef,HasUniqueRoleAssignments&$top=2000`,
        SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      if (!ir.ok) { bericht.leseFehler.push(`Bibliothek ${libTitel} (HTTP ${ir.status})`); continue; }
      const id = await ir.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const it of (id.value || []) as any[]) {
        if (it.HasUniqueRoleAssignments) ziele.push({ base: `${libBase}/items(${it.Id})`, ort: `Seite ${it.FileLeafRef || it.Id}`, nurSchreiben: false });
      }
    }
  } catch (e) {
    bericht.leseFehler.push(`Websiteseiten (${e instanceof Error ? e.message : String(e)})`);
  }

  for (const z of ziele) {
    let zuw: Awaited<ReturnType<typeof _readRoleAssignments>>;
    try { zuw = await _readRoleAssignments(svc, z.base); } catch { bericht.leseFehler.push(z.ort); continue; }
    bericht.geprueft++;
    for (const a of zuw) {
      if (a.type !== 1) continue; // nur Einzelpersonen, Gruppen nie
      const email = norm(a.email || emailAusLogin(a.login));
      if (istSystem(email, a.login) || erlaubt.has(email)) continue;
      const echte = a.roleIds.filter(r => r !== LIMITED_ACCESS);
      if (echte.length === 0) continue;
      if (z.nurSchreiben && echte.every(r => r === READ)) continue;
      const befund: EinzelfreigabeBefund = {
        ort: z.ort, email, name: a.title, rechte: a.roleNames.filter(Boolean).join(', ') || 'direkte Freigabe',
        ergebnis: 'gefunden',
      };
      if (apply) {
        if (!(await inBesucherGruppe(svc, email))) {
          befund.ergebnis = 'besucher-fehlt';
        } else {
          try { await _deletePrincipalAssignment(svc, z.base, a.pid); } catch { /* Nachlesen entscheidet */ }
          let weg = false;
          try { weg = !(await _readRoleAssignments(svc, z.base)).some(x => x.pid === a.pid && x.roleIds.some(r => r !== LIMITED_ACCESS)); } catch { weg = false; }
          befund.ergebnis = weg ? 'reduziert' : 'entfernen-fehlgeschlagen';
        }
      }
      bericht.befunde.push(befund);
    }
  }

  // Mitglieder-Gruppe: nur melden.
  try {
    const mr = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedmembergroup/users?$select=Title,Email,LoginName,PrincipalType&$top=2000`,
      SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    if (mr.ok) {
      const md = await mr.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const u of (md.value || []) as any[]) {
        if (Number(u.PrincipalType) !== 1) continue;
        const email = norm(u.Email || emailAusLogin(u.LoginName));
        if (istSystem(email, u.LoginName) || erlaubt.has(email)) continue;
        bericht.befunde.push({ ort: 'Mitglieder-Gruppe', email, name: u.Title || '', rechte: 'Bearbeiten (über Gruppe)', ergebnis: 'nur-gemeldet' });
      }
    } else if (mr.status !== 404) {
      bericht.leseFehler.push(`Mitglieder-Gruppe (HTTP ${mr.status})`);
    }
  } catch { bericht.leseFehler.push('Mitglieder-Gruppe'); }

  return bericht;
}
