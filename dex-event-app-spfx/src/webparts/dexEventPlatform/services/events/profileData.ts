/**
 * v30.7 — Modularisierung Stufe 2, Tranche 3: Thema „Profil-Daten"
 * (SP-UserProfile/Graph-Lookups, Teilnehmer-Profil-Backfill,
 * canRegisterForOthers). Herausgelöst aus EventService; dort stehen
 * Delegations-Stubs mit unveränderter Signatur.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { SPRegistration } from '../EventService';
// EventService als WERT-Import: nur fuer den deferred Zugriff auf die
// statische stripNoteWrapper — der Zyklus ist unkritisch, weil der Zugriff
// erst zur Laufzeit in Funktionskoerpern passiert (wie REG_LIST_NAME).
import { REG_LIST_NAME, EventService } from '../EventService';


/**
 * Permission-Check: Darf der aktuell eingeloggte User einen anderen Teilnehmer
 * registrieren? Wird in registerForEvent() und reactivateRegistration() aufgerufen,
 * wenn ParticipantEmail !== session-Email.
 *
 * Erlaubt wenn (OR):
 *   - DEX_Roles enthält den User als 'Admin'
 *   - Der User ist in event.OrganizerEmail für das Event auf der zugehörigen
 *     Subsite eingetragen (Event-scope Organizer)
 *   - Der User ist Assistant (JobTitle enthält 'assistant') UND der Target
 *     ist Partner oder Director (JobTitle enthält 'partner' oder 'director')
 *
 * Bei Fehlern lieber konservativ `false` zurückgeben statt durchlassen.
 */
export async function canRegisterForOthers(svc: EventService, subsiteUrl: string, targetParticipantEmail: string): Promise<boolean> {
  // v19.6: Mehrere mögliche Identitäten des eingeloggten Users sammeln.
  // `pageContext.user.email` ist im SharePoint-Mobile-WebView nicht immer
  // gesetzt bzw. weicht vom in OrganizerEmail gespeicherten SMTP-Wert ab —
  // deshalb zusätzlich die E-Mail aus dem `loginName` (Claims-Format
  // `i:0#.f|membership|user@domain`) als Fallback heranziehen.
  const sessionIdentities = new Set<string>();
  const rawEmail = (svc.context.pageContext.user.email || '').toLowerCase().trim();
  if (rawEmail) sessionIdentities.add(rawEmail);
  const loginName = (svc.context.pageContext.user.loginName || '').toLowerCase();
  const loginMatch = loginName.match(/[^|]+@[^|\s]+$/);
  if (loginMatch) sessionIdentities.add(loginMatch[0].trim());
  if (sessionIdentities.size === 0) return false;
  const sessionEmail = rawEmail || Array.from(sessionIdentities)[0];
  const matchesSession = (emails: string[]): boolean =>
    emails.some(e => sessionIdentities.has((e || '').toLowerCase().trim()));

  // 1. DEX_Roles prüfen: Admin- ODER Organizer-Rolle haben?
  //    v19.6 BUG-FIX: Vorher liess dieser Check NUR die Admin-Rolle durch.
  //    Die Rollenmatrix sieht „Für andere registrieren" generell für
  //    Organizer vor — deshalb hier Admin UND Organizer akzeptieren.
  try {
    const esc = sessionEmail.replace(/'/g, "''");
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
      SPHttpClient.configurations.v1
    );
    if (resp.ok) {
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length > 0 && (items[0].Role === 'Admin' || items[0].Role === 'IT-Admin' || items[0].Role === 'Organizer')) return true;
    }
  } catch { /* ignore - fallback auf weitere Checks */ }

  // 2. Event-Organizer ODER Co-Organizer dieses Events?
  //    v19.6 BUG-FIX: Vorher wurde NUR `OrganizerEmail` (Haupt-Organizer)
  //    geprüft — Co-Organizer stehen aber in `EmailTemplateOverrides._coOrganizers`
  //    und wurden so fälschlich abgelehnt. Zudem war der Note-Strip auf EINE
  //    `<div>`-Ebene begrenzt und der Split nur auf `;` — bei mehreren
  //    Organizern (mehrere `<div>`/`<br>` oder Komma-Trennung) schlug der
  //    Match fehl. Jetzt: HTML robust strippen, an `;`/`,`/Zeilenumbruch
  //    splitten und Haupt- + Co-Organizer kombiniert gegen ALLE
  //    Session-Identitäten matchen.
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${encodeURIComponent(subsiteUrl.replace(/'/g, "''"))}'&$top=1&$select=OrganizerEmail,EmailTemplateOverrides`,
      SPHttpClient.configurations.v1
    );
    if (resp.ok) {
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length > 0) {
        const splitEmails = (raw: string | null | undefined): string[] =>
          (raw || '')
            .replace(/<br\s*\/?>/gi, ';')
            .replace(/<\/div>\s*<div[^>]*>/gi, ';')
            .replace(/<[^>]+>/g, '')
            .split(/[;,\n\r]+/)
            .map(s => s.trim().toLowerCase())
            .filter(Boolean);
        const mainOrgEmails = splitEmails(items[0].OrganizerEmail);
        // Co-Organizer aus dem EmailTemplateOverrides-Piggyback `_coOrganizers`.
        let coOrgEmails: string[] = [];
        try {
          const ovRaw = EventService.stripNoteWrapper(items[0].EmailTemplateOverrides) || '{}';
          const ov = JSON.parse(ovRaw);
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const list = (ov as any)._coOrganizers;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          if (Array.isArray(list)) coOrgEmails = list.map((x: any) => String(x?.email || '').toLowerCase().trim()).filter(Boolean);
        } catch { /* kein/ungültiges Override-JSON → keine Co-Organizer */ }
        if (matchesSession([...mainOrgEmails, ...coOrgEmails])) return true;
      }
    }
  } catch { /* ignore */ }

  // 3. Assistant-Ausnahme: User-JobTitle ist eine Assistenz UND Target ist
  //    Partner/Director.
  //    v23.9 BUG-FIX: Vorher nur `indexOf('assistant')` — das matcht das
  //    ENGLISCHE „Assistant", aber NICHT das deutsche „Assistenz" (und auch
  //    nicht „Assistentin"/„Teamassistenz"). Eine Assistenz mit dem Job-Title
  //    „Assistenz" fiel deshalb durch und durfte NICHT stellvertretend
  //    anmelden — die Anmeldung schlug mit der generischen (irreführenden)
  //    „bereits registriert"-Meldung fehl. Jetzt beide Schreibweisen matchen
  //    (gleiche Logik wie isEventVisibleForUser).
  try {
    const sessionProfile = await svc.getCurrentUserProfile();
    const sessionJt = (sessionProfile.jobTitle || '').toLowerCase();
    if (sessionJt.indexOf('assisten') >= 0 || sessionJt.indexOf('assistan') >= 0) {
      const targetProfile = await svc.getUserProfileByEmail(targetParticipantEmail);
      const targetJt = (targetProfile.jobTitle || '').toLowerCase();
      if (targetJt.indexOf('partner') >= 0 || targetJt.indexOf('director') >= 0) {
        return true;
      }
      // Assistant darf nicht für Non-Partner/Director registrieren
      return false;
    }
  } catch { /* ignore */ }

  return false;
}

/**
 * Profildaten des aktuellen Users laden für die Teilnehmerliste.
 */
export async function getCurrentUserProfile(svc: EventService): Promise<{ department: string; location: string; jobTitle: string; phone: string; firstName: string; lastName: string; displayName: string; company: string }> {
  const empty = { department: '', location: '', jobTitle: '', phone: '', firstName: '', lastName: '', displayName: '', company: '' };
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return empty;

    const data = await response.json();
    const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];
    const get = (key: string): string => {
      const p = props.find(x => x.Key === key);
      return p && p.Value ? p.Value : '';
    };

    // v24.33: Company kommt im Tenant nicht aus der SP-UserProfile-Property —
    // wenn leer, via Graph `/me?$select=companyName` nachladen.
    let company = get('Company') || get('SPS-Company') || get('CompanyName');
    if (!company) { company = await svc.getMyCompanyViaGraph(); }
    return {
      department: get('Department'),
      location: get('Office'),
      jobTitle: get('Title'),
      phone: get('WorkPhone') || get('CellPhone'),
      firstName: get('FirstName'),
      lastName: get('LastName'),
      displayName: get('PreferredName'),
      company,
    };
  } catch {
    return empty;
  }
}

/**
 * Cleanup: bei den N jüngsten Teilnehmer-Einträgen jedes Events JobTitle, Department,
 * Location und Phone aus dem aktuellen Benutzerprofil neu laden und überschreiben.
 * Notwendig weil bis v3.0.x diese Felder versehentlich vom EINGELOGGTEN User (statt
 * vom registrierten Teilnehmer) gezogen wurden, wenn jemand für eine andere Person
 * registriert hat.
 *
 * Idempotent: wenn die Daten bereits stimmen (Profil-Lookup liefert dasselbe), passiert
 * nichts. Wird typisch einmalig per LocalStorage-Flag in EventContext getriggert.
 *
 * Liefert die Anzahl tatsächlich aktualisierter Items.
 */
/**
 * Cleanup nur für EIN Event: lae alle Teilnehmer-Profile per Email nachladen
 * und JobTitle/Department/Location/Phone updaten falls abweichend.
 * Wird per Admin-Button im Admin Center pro Event getriggert.
 */
export async function fixEventParticipantsProfileData(svc: EventService, subsiteUrl: string, n: number = 1000): Promise<{ scanned: number; updated: number; failedLookups: number }> {
  let scanned = 0;
  let updated = 0;
  let failedLookups = 0;
  const sleep = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
  if (!subsiteUrl) return { scanned, updated, failedLookups };
  try {
    const listResp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items?$select=Id,ParticipantEmail,Vorname,Nachname,ParticipantName,JobTitle,Department,Location,Phone&$orderby=RegistrationDate desc&$top=${n}`,
      SPHttpClient.configurations.v1
    );
    if (!listResp.ok) return { scanned, updated, failedLookups };
    const listData = await listResp.json();
    const items = listData.value || listData.d?.results || [];
    // v22.58: erkennt kaputte Namen (SharePoint-Claims-Token), die durch den
    // sauberen Profil-Namen ersetzt werden müssen.
    const looksLikeClaim = (s: string): boolean => /\|membership\||0#\.f\||^i:0#/i.test((s || '').trim());
    for (const it of items) {
      scanned += 1;
      const email: string = (it.ParticipantEmail || '').trim();
      if (!email) continue;
      let profile = { department: '', location: '', jobTitle: '', phone: '', firstName: '', lastName: '', displayName: '' };
      let success = false;
      for (let attempt = 0; attempt < 3; attempt += 1) {
        try {
          const p = await svc.getUserProfileByEmail(email);
          if (p && (p.jobTitle || p.department || p.location || p.firstName || p.lastName)) {
            profile = p; success = true; break;
          }
        } catch { /* */ }
        await sleep(500 * (attempt + 1));
      }
      if (!success) { failedLookups += 1; continue; }
      try {
        // Name-Reparatur: nur wenn der aktuelle Name kaputt (Claims) oder leer
        // ist UND das Profil einen sauberen Namen liefert.
        const vornameBroken = looksLikeClaim(it.Vorname) || !(it.Vorname || '').trim();
        const nameFix = vornameBroken && (profile.firstName || profile.lastName)
          ? {
              Vorname: profile.firstName || '',
              Nachname: profile.lastName || '',
              ParticipantName: (profile.displayName && !looksLikeClaim(profile.displayName) ? profile.displayName : `${profile.firstName} ${profile.lastName}`.trim()),
            }
          : null;
        const profileUpdate =
          (profile.jobTitle && profile.jobTitle !== (it.JobTitle || '')) ||
          (profile.department && profile.department !== (it.Department || '')) ||
          (profile.location && profile.location !== (it.Location || '')) ||
          (profile.phone && profile.phone !== (it.Phone || ''));
        if (nameFix || profileUpdate) {
          const ok = await svc._merge(
            `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${it.Id})`,
            {
              'JobTitle': profile.jobTitle || it.JobTitle || '',
              'Department': profile.department || it.Department || '',
              'Location': profile.location || it.Location || '',
              'Phone': profile.phone || it.Phone || '',
              ...(nameFix || {}),
            }
          );
          if (ok && (ok as { ok: boolean }).ok) updated += 1;
        }
      } catch { /* */ }
      await sleep(200);
    }
  } catch { /* */ }
  return { scanned, updated, failedLookups };
}

export async function fixRecentParticipantsProfileData(svc: EventService, n: number): Promise<{ scanned: number; updated: number; failedLookups: number }> {
  let scanned = 0;
  let updated = 0;
  let failedLookups = 0;
  const sleep = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
  try {
    const events = await svc.getEvents();
    for (const evt of events) {
      if (!evt.SubsiteUrl) continue;
      try {
        const listResp = await svc._sp.get(
          `${evt.SubsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items?$select=Id,ParticipantEmail,JobTitle,Department,Location,Phone&$orderby=RegistrationDate desc&$top=${n}`,
          SPHttpClient.configurations.v1
        );
        if (!listResp.ok) continue;
        const listData = await listResp.json();
        const items = listData.value || listData.d?.results || [];
        for (const it of items) {
          scanned += 1;
          const email: string = (it.ParticipantEmail || '').trim();
          if (!email) continue;
          // Profil-Lookup mit Retry on Failure (max 3 Versuche, exponential backoff)
          let profile = { department: '', location: '', jobTitle: '', phone: '' };
          let success = false;
          for (let attempt = 0; attempt < 3; attempt += 1) {
            try {
              const p = await svc.getUserProfileByEmail(email);
              if (p && (p.jobTitle || p.department || p.location)) {
                profile = p;
                success = true;
                break;
              }
            } catch { /* */ }
            await sleep(500 * (attempt + 1));
          }
          if (!success) { failedLookups += 1; continue; }
          try {
            const needsUpdate =
              (profile.jobTitle && profile.jobTitle !== (it.JobTitle || '')) ||
              (profile.department && profile.department !== (it.Department || '')) ||
              (profile.location && profile.location !== (it.Location || '')) ||
              (profile.phone && profile.phone !== (it.Phone || ''));
            if (!needsUpdate) {
              await sleep(200); continue;
            }
            const ok = await svc._merge(
              `${evt.SubsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${it.Id})`,
              {
                'JobTitle': profile.jobTitle || it.JobTitle || '',
                'Department': profile.department || it.Department || '',
                'Location': profile.location || it.Location || '',
                'Phone': profile.phone || it.Phone || '',
              }
            );
            if (ok && (ok as { ok: boolean }).ok) updated += 1;
          } catch { /* einzelnen überspringen */ }
          // Throttle gegen Rate-Limit der UserProfile-API
          await sleep(200);
        }
      } catch { /* */ }
    }
  } catch { /* */ }
  return { scanned, updated, failedLookups };
}

/**
 * Profildaten eines bestimmten Users via Email laden (für "Register for someone else"
 * und "Profile neu laden"). Robust gegen UPN != SMTP-Mismatches.
 *
 * Strategie:
 *   1. Direkter Lookup mit Claim `i:0#.f|membership|<email>` (funktioniert wenn UPN==SMTP).
 *   2. Wenn leer: per `siteusers/getbyemail` den echten LoginName auflösen
 *      (deckt UPN != SMTP und Guest-Accounts ab) und GetPropertiesFor mit
 *      diesem LoginName erneut aufrufen.
 *
 * Rückgabe ist gefüllt sobald einer der Wege Properties liefert, sonst leer.
 */
export async function getUserProfileByEmail(svc: EventService, email: string): Promise<{ department: string; location: string; jobTitle: string; phone: string; firstName: string; lastName: string; displayName: string; company: string }> {
  const empty = { department: '', location: '', jobTitle: '', phone: '', firstName: '', lastName: '', displayName: '', company: '' };
  if (!email) return empty;

  const extractProfile = (props: Array<{ Key: string; Value: string }>): { department: string; location: string; jobTitle: string; phone: string; firstName: string; lastName: string; displayName: string; company: string } => {
    const get = (key: string): string => {
      const p = props.find(x => x.Key === key);
      return p && p.Value ? p.Value : '';
    };
    return {
      department: get('Department'),
      location: get('Office') || get('SPS-Location'),
      jobTitle: get('Title') || get('SPS-JobTitle'),
      phone: get('WorkPhone') || get('CellPhone'),
      // v22.57: Namen mitliefern (für die Absage-Zeile, damit nie ein
      // Claims-Token wie „0#.f|membership|…" als Vorname landet).
      firstName: get('FirstName'),
      lastName: get('LastName'),
      displayName: get('PreferredName'),
      // v24.29: Unternehmenszugehörigkeit / Rechtsträger.
      company: get('Company') || get('SPS-Company') || get('CompanyName'),
    };
  };

  // 1) Direkter Claim per SMTP-Email (schnell, funktioniert für Standard-Tenants)
  try {
    const directUrl = `${svc.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${encodeURIComponent(email)}'`;
    const response = await svc._sp.get(directUrl, SPHttpClient.configurations.v1);
    if (response.ok) {
      const data = await response.json();
      const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];
      const profile = extractProfile(props);
      if (profile.jobTitle || profile.department || profile.location || profile.phone) {
        return profile;
      }
    }
  } catch { /* weiter zu Fallback */ }

  // 2) Fallback: echten LoginName (UPN-Claim) über siteusers/getbyemail auflösen
  // Deckt UPN != SMTP, Guest-Accounts und Alias-SMTP-Adressen ab.
  try {
    const siteUserUrl = `${svc.siteUrl}/_api/web/siteusers/getbyemail('${email.replace(/'/g, "''")}')?$select=LoginName`;
    const siteUserResp = await svc._sp.get(siteUserUrl, SPHttpClient.configurations.v1);
    if (!siteUserResp.ok) return empty;
    const siteUserData = await siteUserResp.json();
    const loginName: string = siteUserData.LoginName || siteUserData.d?.LoginName || '';
    if (!loginName) return empty;

    const profileUrl = `${svc.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(loginName)}'`;
    const profileResp = await svc._sp.get(profileUrl, SPHttpClient.configurations.v1);
    if (!profileResp.ok) return empty;
    const profileData = await profileResp.json();
    const props: Array<{ Key: string; Value: string }> = profileData.UserProfileProperties || [];
    return extractProfile(props);
  } catch {
    return empty;
  }
}

/**
 * v28.65: Claims-Login-Tokens in einer Teilnehmerliste reparieren.
 *
 * Hintergrund siehe `utils/displayName.ts`: Bei Personen, deren Eintrag in
 * der versteckten „User Information List" ohne Anzeigename gestempelt wurde,
 * lieferte `pageContext.user.displayName` das Login-Token
 * („0#.f|membership|user@deloitte.de"). Bis v28.64 landete das 1:1 in der
 * Teilnehmerzeile. Diese Methode zieht die betroffenen Namen aus dem
 * Benutzerprofil nach.
 *
 * Geprüft werden `ParticipantName`, `Vorname`, `Nachname` (Quelle:
 * `ParticipantEmail`) sowie die Audit-Felder `RegisteredByName` und
 * `CancelledByName` (Quelle: die jeweilige Audit-E-Mail). Zeilen ohne Token
 * bleiben unangetastet; ist die Person im Profil nicht auflösbar, wird
 * wenigstens die E-Mail statt des Tokens gesetzt — lesbar und eindeutig.
 */
export async function repairClaimNamesInRegistrations(
  svc: EventService,
  subsiteUrl: string,
  onProgress?: (done: number, total: number) => void,
): Promise<{ scanned: number; hits: number; fixed: number; failed: number }> {
  const out = { scanned: 0, hits: 0, fixed: 0, failed: 0 };
  if (!subsiteUrl) return out;
  const looksLikeClaim = (s: string): boolean => /\|membership\b|^i:0[#|]|^c:0|0#\.[a-z]\||^\d+#\./i.test((s || '').trim());
  const mailFromClaim = (s: string): string => {
    const m = (s || '').match(/\|([^|]+@[^|\s]+)\s*$/);
    return m ? m[1].trim().toLowerCase() : '';
  };
  let rows: SPRegistration[] = [];
  try { rows = await svc.getAllRegistrations(subsiteUrl); } catch { return out; }
  out.scanned = rows.length;

  const affected = rows.filter(r =>
    looksLikeClaim(r.ParticipantName || '') || looksLikeClaim(r.Vorname || '') || looksLikeClaim(r.Nachname || '')
    || looksLikeClaim(r.RegisteredByName || '') || looksLikeClaim(r.CancelledByName || ''));
  out.hits = affected.length;
  if (affected.length === 0) return out;

  // Profile je E-Mail nur einmal holen — dieselbe Person taucht oft mehrfach
  // auf (Klammer-Schattenzeile plus Sub-Events).
  const cache: Record<string, { firstName: string; lastName: string; displayName: string }> = {};
  const nameFor = async (email: string): Promise<{ firstName: string; lastName: string; displayName: string }> => {
    const key = (email || '').toLowerCase();
    if (!key) return { firstName: '', lastName: '', displayName: '' };
    if (cache[key]) return cache[key];
    let p = { firstName: '', lastName: '', displayName: '' };
    try {
      const prof = await svc.getUserProfileByEmail(key);
      p = {
        firstName: looksLikeClaim(prof.firstName) ? '' : (prof.firstName || '').trim(),
        lastName: looksLikeClaim(prof.lastName) ? '' : (prof.lastName || '').trim(),
        displayName: looksLikeClaim(prof.displayName) ? '' : (prof.displayName || '').trim(),
      };
    } catch { /* nicht auflösbar — E-Mail als Fallback */ }
    cache[key] = p;
    return p;
  };

  for (let i = 0; i < affected.length; i++) {
    const r = affected[i];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const patch: Record<string, any> = {};
    const ownEmail = (r.ParticipantEmail || r.Title || '').trim();
    if (looksLikeClaim(r.ParticipantName || '') || looksLikeClaim(r.Vorname || '') || looksLikeClaim(r.Nachname || '')) {
      const mail = ownEmail || mailFromClaim(r.ParticipantName || '');
      // eslint-disable-next-line no-await-in-loop
      const p = await nameFor(mail);
      const first = looksLikeClaim(r.Vorname || '') ? p.firstName : ((r.Vorname || '').trim() || p.firstName);
      const last = looksLikeClaim(r.Nachname || '') ? p.lastName : ((r.Nachname || '').trim() || p.lastName);
      const display = p.displayName || `${first} ${last}`.trim() || mail;
      if (looksLikeClaim(r.ParticipantName || '')) patch['ParticipantName'] = display;
      if (looksLikeClaim(r.Vorname || '')) patch['Vorname'] = first;
      if (looksLikeClaim(r.Nachname || '')) patch['Nachname'] = last;
    }
    if (looksLikeClaim(r.RegisteredByName || '')) {
      const mail = (r.RegisteredByEmail || '').trim() || mailFromClaim(r.RegisteredByName || '');
      // eslint-disable-next-line no-await-in-loop
      const p = await nameFor(mail);
      patch['RegisteredByName'] = p.displayName || `${p.firstName} ${p.lastName}`.trim() || mail;
    }
    if (looksLikeClaim(r.CancelledByName || '')) {
      const mail = (r.CancelledByEmail || '').trim() || mailFromClaim(r.CancelledByName || '');
      // eslint-disable-next-line no-await-in-loop
      const p = await nameFor(mail);
      patch['CancelledByName'] = p.displayName || `${p.firstName} ${p.lastName}`.trim() || mail;
    }
    if (Object.keys(patch).length === 0) { if (onProgress) onProgress(i + 1, affected.length); continue; }
    try {
      // eslint-disable-next-line no-await-in-loop
      const resp = await svc._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${r.Id})`, patch,
      );
      if (resp.ok || resp.status === 406) out.fixed++; else out.failed++;
    } catch { out.failed++; }
    if (onProgress) onProgress(i + 1, affected.length);
  }
  return out;
}

/**
 * v28.65: Anzeigenamen zu einer E-Mail auflösen (für die Organizer-Reparatur
 * im Admin Center). Leer, wenn das Profil nichts hergibt.
 */
export async function displayNameForEmail(svc: EventService, email: string): Promise<string> {
  if (!email) return '';
  const looksLikeClaim = (s: string): boolean => /\|membership\b|^i:0[#|]|^c:0|0#\.[a-z]\|/i.test((s || '').trim());
  try {
    const p = await svc.getUserProfileByEmail(email);
    const cand = [p.displayName, [p.lastName, p.firstName].filter(Boolean).join(', ')];
    for (const c of cand) {
      const v = (c || '').trim();
      if (v && !looksLikeClaim(v)) return v;
    }
  } catch { /* nicht auflösbar */ }
  return '';
}

