/**
 * v31.91: Person umbenennen — alte gegen neue E-Mail-Adresse (und Namen) an
 * ALLEN Stellen tauschen, an denen die App eine Person über die Adresse
 * schlüsselt.
 *
 * Anlass (24.09.2026): Inga Fuhr hat geheiratet und heißt jetzt Guenther;
 * das Konto trägt iguenther@deloitte.de, DEX_Roles noch infuhr@deloitte.de.
 * Das Rechte-Audit meldete „gleiche Rechte, andere Schreibweise". Nutzer-
 * Ansage: „bitte an allen entsprechenden Stellen anpassen."
 *
 * CLAUDE.md: Die E-Mail-Adresse ist der EINZIGE Schlüssel — und genau
 * deshalb reicht es nicht, eine Zeile zu ändern. Stellen:
 *  1. DEX_Roles — Title (= E-Mail), UserName.
 *  2. DEX_Events — Organizer/OrganizerEmail (parallele `;`-Listen),
 *     ContactEmail, ContactOrganizerEmail, und die Piggybacks
 *     `_coOrganizers`, `_qrScanners`, `_testTeam` in EmailTemplateOverrides
 *     (Klammer UND Termine — jede Zeile trägt ihre eigene Kopie).
 *  3. DEX_Participants — Email (+ Title, Vorname, Nachname).
 *  4. Teilnehmerlisten aller Subsites — ParticipantEmail, Vorname, Nachname,
 *     ParticipantName; dazu RegisteredByEmail/-Name und
 *     CancelledByEmail/-Name, wo die Person Anmeldende war.
 *
 * Nicht angefasst: SharePoint-Berechtigungen (das Konto ist dasselbe
 * Principal, nur die Adresse hat sich geändert), Mail-Queue-Historie,
 * ChangeLog-Einträge (Historie bleibt, wie sie war).
 *
 * `dryRun` zählt nur. Der echte Lauf schreibt sequentiell, zählt Fehler und
 * bricht NICHT ab — jede Stelle ist für sich reparierbar, und eine halbe
 * Umbenennung ist besser als keine, solange das Ergebnis sagt, was fehlt.
 */
import { SPHttpClient } from '@microsoft/sp-http';
import { EventService, REG_LIST_NAME } from '../EventService';

export interface PersonRenameArgs {
  oldEmail: string;
  newEmail: string;
  /** Anzeigename, wie ihn Graph liefert („Guenther, Inga") — für Organizer-Listen, Piggybacks, DEX_Roles. */
  newDisplayName: string;
  newFirstName: string;
  newLastName: string;
}

export interface PersonRenameBereich {
  /** Gefundene Stellen. */
  hits: number;
  /** Geschrieben (0 im Trockenlauf). */
  written: number;
  failed: number;
  /** Nicht lesbar (Liste verweigert) — Zahl der Listen/Abfragen. */
  unreadable: number;
}

export interface PersonRenameResult {
  dryRun: boolean;
  roles: PersonRenameBereich;
  eventFields: PersonRenameBereich;
  eventPiggybacks: PersonRenameBereich;
  registry: PersonRenameBereich;
  registrations: PersonRenameBereich;
  registeredBy: PersonRenameBereich;
  /** Titel der Events mit Treffern (für die Vorschau). */
  eventTitles: string[];
  /** Subsites mit Treffern in der Teilnehmerliste (Site-Name). */
  listSites: string[];
}

const leer = (): PersonRenameBereich => ({ hits: 0, written: 0, failed: 0, unreadable: 0 });
const lc = (s: string | null | undefined): string => (s || '').trim().toLowerCase();
const esc = (s: string): string => s.replace(/'/g, "''");

/** HTML eines Note-Felds strippen und an `;`/`,`/Zeilenumbruch splitten — wie registration.ts. */
const splitList = (raw: string | null | undefined): string[] =>
  (raw || '')
    .replace(/<br\s*\/?>/gi, ';')
    .replace(/<\/div>\s*<div[^>]*>/gi, ';')
    .replace(/<[^>]+>/g, '')
    .split(/[;\n\r]+/)
    .map(s => s.trim())
    .filter(Boolean);

export async function renamePerson(
  svc: EventService,
  args: PersonRenameArgs,
  opts: { dryRun: boolean },
  onProgress?: (label: string) => void,
): Promise<PersonRenameResult> {
  const oldLc = lc(args.oldEmail);
  const newEmail = (args.newEmail || '').trim();
  const newName = (args.newDisplayName || '').trim();
  const first = (args.newFirstName || '').trim();
  const last = (args.newLastName || '').trim();
  const dry = !!opts.dryRun;
  const out: PersonRenameResult = {
    dryRun: dry, roles: leer(), eventFields: leer(), eventPiggybacks: leer(), registry: leer(),
    registrations: leer(), registeredBy: leer(), eventTitles: [], listSites: [],
  };
  if (!oldLc || !newEmail) return out;
  const hdr = { headers: { 'Accept': 'application/json;odata=nometadata' } };
  const merge = async (url: string, body: Record<string, unknown>, b: PersonRenameBereich): Promise<void> => {
    if (dry) return;
    try {
      const r = await svc._merge(url, body);
      if (r.ok || r.status === 204) b.written++; else b.failed++;
    } catch { b.failed++; }
  };

  // 1) DEX_Roles
  if (onProgress) onProgress('DEX_Roles');
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$select=Id,Title,UserName&$top=5000`,
      SPHttpClient.configurations.v1, hdr);
    if (!r.ok) out.roles.unreadable++;
    else {
      const d = await r.json();
      const rows: Array<{ Id: number; Title?: string; UserName?: string }> = d.value || d.d?.results || [];
      for (const row of rows) {
        if (lc(row.Title) !== oldLc) continue;
        out.roles.hits++;
        await merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items(${row.Id})`,
          { Title: newEmail, ...(newName ? { UserName: newName } : {}) }, out.roles);
      }
    }
  } catch { out.roles.unreadable++; }

  // 2) DEX_Events — Felder und Piggybacks
  if (onProgress) onProgress('DEX_Events');
  let events: Awaited<ReturnType<EventService['getEvents']>> = [];
  try { events = await svc.getEvents(); } catch { out.eventFields.unreadable++; }
  for (const ev of events) {
    const patch: Record<string, unknown> = {};
    // Organizer / OrganizerEmail sind parallele Listen — der Name an der
    // Position der alten Adresse wird mit ersetzt.
    const emails = splitList(ev.OrganizerEmail);
    const names = splitList(ev.Organizer);
    const idx = emails.findIndex(e => lc(e) === oldLc);
    if (idx >= 0) {
      emails[idx] = newEmail;
      if (newName && names.length === emails.length) names[idx] = newName;
      patch['OrganizerEmail'] = emails.join('; ');
      if (newName && names.length === emails.length) patch['Organizer'] = names.join('; ');
    }
    if (lc(ev.ContactEmail) === oldLc) patch['ContactEmail'] = newEmail;
    if (lc(ev.ContactOrganizerEmail) === oldLc) patch['ContactOrganizerEmail'] = newEmail;
    if (Object.keys(patch).length > 0) {
      out.eventFields.hits++;
      if (out.eventTitles.indexOf(ev.Title || String(ev.Id)) < 0) out.eventTitles.push(ev.Title || String(ev.Id));
      await merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${ev.Id})`, patch, out.eventFields);
    }
    // Piggybacks
    let ov: Record<string, unknown> | null = null;
    try { ov = JSON.parse(EventService.stripNoteWrapper(ev.EmailTemplateOverrides) || '{}'); } catch { ov = null; }
    if (ov && typeof ov === 'object') {
      for (const key of ['_coOrganizers', '_qrScanners', '_testTeam']) {
        const list = ov[key];
        if (!Array.isArray(list)) continue;
        let touched = false;
        const next = (list as Array<Record<string, unknown>>).map(x => {
          if (x && lc(String(x.email || '')) === oldLc) {
            touched = true;
            return { ...x, email: newEmail, ...(newName ? { name: newName } : {}) };
          }
          return x;
        });
        if (!touched) continue;
        out.eventPiggybacks.hits++;
        if (out.eventTitles.indexOf(ev.Title || String(ev.Id)) < 0) out.eventTitles.push(ev.Title || String(ev.Id));
        if (!dry) {
          try {
            const ok = await svc.patchEventOverridesValue(Number(ev.Id), key, next);
            if (ok) out.eventPiggybacks.written++; else out.eventPiggybacks.failed++;
          } catch { out.eventPiggybacks.failed++; }
        }
      }
    }
  }

  // 3) DEX_Participants
  if (onProgress) onProgress('DEX_Participants');
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$filter=Email eq '${esc(args.oldEmail.trim())}'&$select=Id,Title,Email,Vorname,Nachname&$top=50`,
      SPHttpClient.configurations.v1, hdr);
    if (!r.ok) out.registry.unreadable++;
    else {
      const d = await r.json();
      const rows: Array<{ Id: number; Title?: string; Email?: string }> = d.value || d.d?.results || [];
      for (const row of rows) {
        out.registry.hits++;
        const body: Record<string, unknown> = { Email: newEmail };
        if (lc(row.Title) === oldLc) body['Title'] = newEmail;
        if (first) body['Vorname'] = first;
        if (last) body['Nachname'] = last;
        await merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${row.Id})`, body, out.registry);
      }
    }
  } catch { out.registry.unreadable++; }

  // 4) Teilnehmerlisten aller Subsites
  const sites = Array.from(new Set(events.map(e => (e.SubsiteUrl || '').trim()).filter(Boolean)));
  const siteName = (s: string): string => s.replace(/\/+$/, '').split('/').pop() || s;
  let i = 0;
  for (const site of sites) {
    i++;
    if (onProgress) onProgress(`Teilnehmerlisten ${i}/${sites.length}`);
    const base = `${site}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`;
    // 4a) eigene Zeilen
    try {
      const r = await svc._sp.get(
        `${base}?$filter=ParticipantEmail eq '${esc(args.oldEmail.trim())}'&$select=Id,ParticipantEmail,Vorname,Nachname,ParticipantName&$top=50`,
        SPHttpClient.configurations.v1, hdr);
      if (!r.ok) {
        // 404 = Liste/Subsite weg (recycelt) — kein Lesefehler im Sinne der Zählung.
        if (r.status !== 404) out.registrations.unreadable++;
      } else {
        const d = await r.json();
        const rows: Array<{ Id: number }> = d.value || d.d?.results || [];
        if (rows.length > 0 && out.listSites.indexOf(siteName(site)) < 0) out.listSites.push(siteName(site));
        for (const row of rows) {
          out.registrations.hits++;
          const body: Record<string, unknown> = { ParticipantEmail: newEmail };
          if (first) body['Vorname'] = first;
          if (last) body['Nachname'] = last;
          if (first || last) body['ParticipantName'] = `${first} ${last}`.trim();
          await merge(`${base}(${row.Id})`, body, out.registrations);
        }
      }
    } catch { out.registrations.unreadable++; }
    // 4b) Zeilen, die die Person für andere angelegt oder abgemeldet hat
    try {
      const r = await svc._sp.get(
        `${base}?$filter=RegisteredByEmail eq '${esc(args.oldEmail.trim())}' or CancelledByEmail eq '${esc(args.oldEmail.trim())}'&$select=Id,RegisteredByEmail,CancelledByEmail&$top=500`,
        SPHttpClient.configurations.v1, hdr);
      if (r.ok) {
        const d = await r.json();
        const rows: Array<{ Id: number; RegisteredByEmail?: string; CancelledByEmail?: string }> = d.value || d.d?.results || [];
        for (const row of rows) {
          out.registeredBy.hits++;
          const body: Record<string, unknown> = {};
          if (lc(row.RegisteredByEmail) === oldLc) { body['RegisteredByEmail'] = newEmail; if (newName) body['RegisteredByName'] = newName; }
          if (lc(row.CancelledByEmail) === oldLc) { body['CancelledByEmail'] = newEmail; if (newName) body['CancelledByName'] = newName; }
          if (Object.keys(body).length === 0) continue;
          await merge(`${base}(${row.Id})`, body, out.registeredBy);
        }
      }
      // Spalten fehlen auf Alt-Listen (400) oder Liste weg (404): still übergehen.
    } catch { /* still */ }
  }

  if (!dry) {
    try {
      await svc.writeChangeLog({
        action: 'PersonRenamed',
        targetType: 'Participant',
        targetId: `${oldLc}→${lc(newEmail)}`,
        targetName: newName || newEmail,
        details: {
          note: `Person umbenannt (v31.91): ${args.oldEmail} → ${newEmail}${newName ? ` (${newName})` : ''}. Rollen ${out.roles.written}/${out.roles.hits}, Event-Felder ${out.eventFields.written}/${out.eventFields.hits}, Teams ${out.eventPiggybacks.written}/${out.eventPiggybacks.hits}, Register ${out.registry.written}/${out.registry.hits}, Anmeldungen ${out.registrations.written}/${out.registrations.hits}, Registriert-von ${out.registeredBy.written}/${out.registeredBy.hits}.`,
        },
      });
    } catch { /* best-effort */ }
  }
  return out;
}
