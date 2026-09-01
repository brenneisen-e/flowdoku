/**
 * v30.66 — Modularisierung Stufe 2: Thema „Überbuchung" — erkennen, dass mehr
 * aktive Anmeldungen als Plätze da sind, und die drei Auflösungen anbieten
 * (auf die Warteliste, Platz behalten, als Erste auf die Warteliste) samt der
 * Entschuldigungs-Mail.
 *
 * Bei geteilten Kapazitäten ist `maxParticipants` 0 — die Grenze steht in
 * durchstarterCapacity/funstarterCapacity (siehe CLAUDE.md).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import { wrapTemplateForStorage, buildEmailFromTemplate } from '../EmailTemplates';
import { ACTIVE_STATI } from '../EventService';
import type { EventService, SPRegistration } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

/**
 * v11.36: Überbuchung erkennen + markieren (ändert KEINEN Status).
 * Pro Gruppe (bzw. gesamt bei Nicht-Split) werden die zuletzt angemeldeten
 * Einträge über Kapazität (höchste SP-Id = zuletzt registriert) mit
 * OverbookReview='Pending' markiert. First-come-first-served: wer zuerst
 * da war, behält den Platz.
 */
export async function detectOverbooking(
  svc: EventService,
  subsiteUrl: string,
  opts: { isSplit: boolean; maxParticipants?: number; durchstarterCapacity?: number; funstarterCapacity?: number }
): Promise<{ groups: Array<{ group: string; cap: number; activeBefore: number; marked: number }>; total: number; errors: number }> {
  const regs = await svc.getAllRegistrations(subsiteUrl); // Id asc
  const groups: Array<{ group: string; cap: number; activeBefore: number; marked: number }> = [];
  let total = 0;
  let errors = 0;
  const markExcess = async (items: SPRegistration[], cap: number, label: string): Promise<void> => {
    const before = items.length;
    let marked = 0;
    if (cap > 0 && before > cap) {
      const excess = items.slice(cap); // Id asc → ab Index cap = die neuesten
      for (const it of excess) {
        if (it.OverbookReview === 'Pending') { marked++; total++; continue; }
        try {
          const resp = await svc._merge(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${it.Id})`,
            { 'OverbookReview': 'Pending' }
          );
          if (resp.ok || resp.status === 406) { marked++; total++; } else { errors++; }
        } catch { errors++; }
      }
    }
    groups.push({ group: label, cap, activeBefore: before, marked });
  };
  const isActive = (r: SPRegistration): boolean => ACTIVE_STATI.indexOf(r.Status) >= 0;
  if (opts.isSplit) {
    await markExcess(regs.filter(r => isActive(r) && r.StarterType === 'Durchstarter'), opts.durchstarterCapacity || 0, 'Durchstarter');
    await markExcess(regs.filter(r => isActive(r) && r.StarterType === 'Funstarter'), opts.funstarterCapacity || 0, 'Funstarter');
  } else {
    await markExcess(regs.filter(isActive), opts.maxParticipants || 0, 'all');
  }
  return { groups, total, errors };
}

/** v11.36: Review-Marker einer Zeile entfernen (ohne Status-Änderung). */
export async function clearOverbookMark(svc: EventService, subsiteUrl: string, itemId: number): Promise<boolean> {
  try {
    const resp = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      { 'OverbookReview': '' }
    );
    return resp.ok || resp.status === 406;
  } catch { return false; }
}

/**
 * v11.36: Markierte Person auf die Warteliste setzen (= „Bestätigen").
 * Gruppentreu: PreferredStarterType bleibt die Gruppe, StarterType wird
 * geleert (wie bei switchSplitGroup → Warteliste).
 */
export async function resolveOverbookToWaitlist(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  group: string
): Promise<boolean> {
  try {
    // Audit: festhalten, dass die Person kurz einen bestätigten Platz hatte
    // und wegen der technischen Überbuchung auf Warteliste korrigiert wurde
    // (inkl. Original-Registrierung) — dauerhaft nachvollziehbar, unabhängig
    // von der späteren TeilnehmerID-Neuvergabe.
    let changeLog = '';
    let origDate = '';
    try {
      const ex = await svc._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog,RegistrationDate`,
        SPHttpClient.configurations.v1
      );
      if (ex.ok) {
        const d = await ex.json();
        changeLog = d.ChangeLog || d.d?.ChangeLog || '';
        origDate = d.RegistrationDate || d.d?.RegistrationDate || '';
      }
    } catch { /* ChangeLog optional */ }
    const stamp = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    const note = `[${stamp}] Überbuchung: war fälschlich angemeldet (technisches Problem bei zeitgleicher Anmeldung) → auf Warteliste korrigiert (Original-Registrierung: ${origDate || 'unbekannt'})`;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = group
      ? { 'Status': 'Warteliste', 'StarterType': '', 'PreferredStarterType': group, 'OverbookReview': '' }
      : { 'Status': 'Warteliste', 'OverbookReview': '' };
    body['ChangeLog'] = changeLog ? `${changeLog}\n${note}` : note;
    const resp = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      body
    );
    return resp.ok || resp.status === 406;
  } catch { return false; }
}

/**
 * v11.36: „Platz behalten" — Variante (a): bleibt angemeldet (Marker weg).
 * Die Gruppe bleibt damit ggf. +1 über Kapazität; der Power-Automate-Flow
 * (Check_<Typ>_Free bzw. Check_Nachrücken, strikt `<`) rückt beim nächsten
 * Frei-Werden in dieser Gruppe einmal NICHT nach — die Überzahl wird so
 * automatisch absorbiert. Identisch zu clearOverbookMark, eigener Name
 * nur fürs Audit/Lesbarkeit.
 */
export async function resolveOverbookKeepActive(svc: EventService, subsiteUrl: string, itemId: number): Promise<boolean> {
  return svc.clearOverbookMark(subsiteUrl, itemId);
}

/**
 * v11.36: „Platz behalten" — Variante (b): Person wird Erste(r) auf der
 * gruppeneigenen Warteliste. Der Nachrück-Flow sortiert Warteliste nach
 * RegistrationDate asc — daher setzen wir RegistrationDate knapp VOR den
 * frühesten aktuellen Wartelisten-Eintrag derselben Gruppe. Original-Datum
 * wird im ChangeLog vermerkt.
 */
export async function resolveOverbookKeepAsFirstWaitlist(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  group: string
): Promise<boolean> {
  try {
    const all = await svc.getAllRegistrations(subsiteUrl);
    const sameGroupWaitlist = all.filter(r =>
      r.Status === 'Warteliste' && (!group || r.PreferredStarterType === group)
    );
    let newDateMs = Date.now();
    for (const w of sameGroupWaitlist) {
      const t = new Date(w.RegistrationDate).getTime();
      if (!isNaN(t) && t < newDateMs) newDateMs = t;
    }
    newDateMs -= 1000; // 1s vor den/die bisherige(n) Erste(n)
    const self = all.find(r => r.Id === itemId);
    const origDate = self?.RegistrationDate || '';
    const stamp = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    let changeLog = '';
    try {
      const ex = await svc._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
        SPHttpClient.configurations.v1
      );
      if (ex.ok) { const d = await ex.json(); changeLog = d.ChangeLog || d.d?.ChangeLog || ''; }
    } catch { /* ChangeLog optional */ }
    const note = `[${stamp}] Überbuchung: Platz behalten → Erste(r) auf Warteliste (Original-Registrierung: ${origDate})`;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = {
      'Status': 'Warteliste',
      'StarterType': '',
      'PreferredStarterType': group || '',
      'RegistrationDate': new Date(newDateMs).toISOString(),
      'OverbookReview': '',
      'ChangeLog': changeLog ? `${changeLog}\n${note}` : note,
    };
    const resp = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      body
    );
    return resp.ok || resp.status === 406;
  } catch { return false; }
}

/**
 * v11.36: Vorgeschlagener Entschuldigungs-Mailtext (Deloitte-Wrap, DE/EN)
 * für „Bestätigen mit Mail". Der Admin kann den Text im Modal vor dem
 * Versand editieren.
 */
// v13.0: Lädt das OverbookingApology-Template aus DEX_EmailTemplates;
// wenn das Template existiert wird daraus die Mail gebaut (inkl.
// Reseed-Funktionalität für Admins). Fallback ist der alte Inline-Text
// damit ältere Tenants ohne Template-Update nicht ohne Mail dastehen.
export async function buildOverbookApologyEmail(
  svc: EventService,
  name: string,
  eventTitle: string,
  lang: string,
  waitlistPos?: number
): Promise<{ subject: string; body: string }> {
  const isDe = (lang || 'EN').toUpperCase() === 'DE';
  const first = (name || '').split(' ')[0] || name;
  const hasPos = typeof waitlistPos === 'number' && waitlistPos > 0;
  const posBlock = hasPos
    ? (isDe
      ? `<p>Du stehst jetzt auf <strong>Warteliste-Platz ${waitlistPos}</strong>.</p>`
      : `<p>You are now <strong>waitlist position ${waitlistPos}</strong>.</p>`)
    : '';
  const tpl = await svc.getEmailTemplate('OverbookingApology', lang).catch(() => null);
  const vars: Record<string, string> = {
    Name: first || name,
    EventTitle: eventTitle,
    WaitlistPositionBlock: posBlock,
    WaitlistPosition: hasPos ? String(waitlistPos) : '',
    AppUrl: `${svc.siteUrl}/SitePages/DEX.aspx?env=WebView`,
  };
  if (tpl) {
    return buildEmailFromTemplate(tpl, vars);
  }
  // Fallback-Inline (alte Pfade)
  const heading = isDe ? 'Anmeldung korrigiert' : 'Registration corrected';
  if (isDe) {
    const inner = `<p>Hallo ${first},</p>`
      + `<p>leider müssen wir uns für ein technisches Problem entschuldigen: durch sehr viele zeitgleiche Anmeldungen wurde dir für <strong>${eventTitle}</strong> versehentlich ein Platz bestätigt, obwohl die Kapazität bereits erschöpft war.</p>`
      + `<p>Wir mussten deine Anmeldung daher auf die <strong>Warteliste</strong> korrigieren. Das tut uns aufrichtig leid — es lag nicht an dir, sondern an einem Ansturm auf die Anmeldung.</p>`
      + posBlock
      + `<p>Sobald ein Platz frei wird, rückst du automatisch nach und bekommst sofort eine Bestätigung. Du musst nichts weiter tun.</p>`
      + `<p style="margin-top:24px;"><strong>Vielen Dank für dein Verständnis</strong><br><br><strong>Dein Event-Team</strong></p>`;
    return {
      subject: `Wichtig: Korrektur deiner Anmeldung — ${eventTitle}`,
      body: wrapTemplateForStorage('#ed8b00', heading, eventTitle, inner),
    };
  }
  const inner = `<p>Hi ${first},</p>`
    + `<p>we sincerely apologize for a technical problem: due to a large number of simultaneous registrations, you were mistakenly confirmed a spot for <strong>${eventTitle}</strong> although capacity was already full.</p>`
    + `<p>We therefore had to move your registration to the <strong>waitlist</strong>. We're truly sorry — this was not your fault but caused by a registration rush.</p>`
    + posBlock
    + `<p>As soon as a spot opens up you will be promoted automatically and notified right away. Nothing else is needed from your side.</p>`
    + `<p style="margin-top:24px;"><strong>Thank you for your understanding</strong><br><br><strong>Your Event Team</strong></p>`;
  return {
    subject: `Important: correction of your registration — ${eventTitle}`,
    body: wrapTemplateForStorage('#ed8b00', heading, eventTitle, inner),
  };
}
