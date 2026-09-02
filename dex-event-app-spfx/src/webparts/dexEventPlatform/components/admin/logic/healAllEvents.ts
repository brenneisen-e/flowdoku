/**
 * v30.70: Nachrücken & IDs für ALLE Events nachholen — Admin Center, Wartung.
 *
 * Entstanden am 02.09.2026: Der Flow `DEX_IDReorder_TeilnehmerIDs` brach einen
 * Tag lang jeden Lauf mit 502 ab (Gruppen-Zählung auf Events ohne
 * `StarterType`-Spalte). Jede Abmeldung in dieser Zeit hat einen Platz frei
 * gemacht, den niemand bekommen hat — bei Leuten daneben auf der Warteliste.
 * Pro Event nachklicken hieße 30 Events × „Freie Plätze füllen"; das hier macht
 * es in einem Durchgang und berichtet am Ende.
 *
 * Bewusst ein eigenständiges Modul ohne React-Hook: Der Aufrufer ist der
 * Admin-Hub, der kein ausgewähltes Event hat — `useWaitlistActions` verlangt
 * eines. Die beiden Helfer, die beide Seiten brauchen (`idSequenceCheck`,
 * `notifyPromotedFor`), liegen deshalb hier und werden vom Hook importiert;
 * zwei Kopien des Mail-Aufbaus wären beim nächsten Template-Wechsel
 * auseinandergelaufen.
 *
 * Zwei Phasen, bewusst getrennt:
 *  1. PLANEN — alle Listen lesen, je Event über `buildPromotionPlan`
 *     ausrechnen, wer nachrücken darf, und prüfen, ob die Nummern Lücken
 *     haben. Noch wird nichts geschrieben. Das Ergebnis steht im
 *     Bestätigungs-Dialog: Die Aktion verschickt Nachrück-Mails und
 *     Outlook-Einladungen, und ein Knopf, der ungefragt an 30 Leute mailt,
 *     ist der falsche Knopf.
 *  2. AUSFÜHREN — promoten + benachrichtigen, IDs neu vergeben, Zähler
 *     syncen. Sequentiell (SharePoint-Throttling), mit Fehlerzähler statt
 *     stillem catch — ein 429 hinterließe sonst genau die halben Zustände,
 *     die man später nicht mehr nachrechnen kann (v29.2).
 *
 * Beschnittene Listen werden ÜBERSPRUNGEN, nicht als leer gerechnet: Eine
 * Subsite ohne Vollzugriff meldet sich über `onHttpError`, und „0 Aktive" aus
 * einer 403-Sicht hieße sonst „alle Plätze frei" (v30.62). Ebenso
 * übersprungen: geteilte Kapazität mit GEMEINSAMER Warteliste — dort
 * entscheidet die Reihenfolge, nicht die Gruppe (v30.67), und das rechnet nur
 * `runManualPromote` richtig. Beide Gruppen stehen namentlich im Bericht.
 */
import { EventService, SPRegistration } from '../../../services/EventService';
import { applyEventTemplateOverride, formatOrganizerList } from '../../../context/EventContext';
import { buildEmailFromTemplate, promotionEmail } from '../../../services/EmailTemplates';
import { withParentTitleSubject } from '../../../utils/mailSubject';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { buildPromotionPlan, promotionPlanLines, isSplitCapacityOf, PromotionPlan } from '../../../utils/promotionPlan';
import { DeloitteEvent } from '../../../types';

/**
 * v11.70 / v11.71 / v22.12: Sind die TeilnehmerIDs der nicht-abgemeldeten
 * Zeilen lückenlos 1..N? `recent: true` heißt „Lücke, Duplikat oder Zeile ohne
 * Nummer" — typischer Auslöser: gerade erfolgte Abmeldung, der
 * DEX_IDReorder-Flow ist noch nicht fertig. `detail` benennt, WAS in den
 * geladenen Daten falsch ist, `whenIso` die jüngste Abmeldung (nur
 * `CancellationDate` — ein Gruppenwechsel setzt keines, siehe IdGapHintBox).
 */
export function idSequenceCheck(regs: SPRegistration[]): { recent: boolean; whenIso: string; detail: string } {
  const active = regs.filter(r => r.Status !== 'Abgemeldet');
  if (active.length === 0) return { recent: false, whenIso: '', detail: '' };
  const ids: number[] = [];
  let noId = 0;
  for (const r of active) {
    const id = Number(r.TeilnehmerID);
    if (!isFinite(id) || id <= 0) { noId++; continue; }
    ids.push(id);
  }
  ids.sort((a, b) => a - b);
  let dups = 0;
  let firstGapAt = 0;
  for (let i = 0; i < ids.length; i++) {
    if (i > 0 && ids[i] === ids[i - 1]) dups++;
    if (firstGapAt === 0 && ids[i] !== i + 1) firstGapAt = i + 1;
  }
  if (noId === 0 && dups === 0 && firstGapAt === 0) return { recent: false, whenIso: '', detail: '' };
  const parts: string[] = [];
  if (firstGapAt > 0) parts.push(`Nummern nicht durchgängig (erwartet Nr. ${firstGapAt})`);
  if (dups > 0) parts.push(`${dups} doppelte Nummer${dups === 1 ? '' : 'n'}`);
  if (noId > 0) parts.push(`${noId} Eintr${noId === 1 ? 'ag' : 'äge'} ohne Nummer`);
  let latest = 0;
  for (const r of regs) {
    if (r.Status !== 'Abgemeldet') continue;
    const t = new Date(r.CancellationDate || '').getTime();
    if (!isNaN(t) && t > latest) latest = t;
  }
  return {
    recent: true,
    whenIso: latest > 0 ? new Date(latest).toISOString() : '',
    detail: `${active.length} aktive Einträge — ${parts.join(', ')}`,
  };
}

/**
 * v29.16: Nachrück-Mail + Outlook-Einladung für EINE nachgerückte Person am
 * Event `ev`. `allEvents` nur für den Klammer-Titel im Betreff.
 */
export async function notifyPromotedFor(
  svc: EventService,
  allEvents: DeloitteEvent[],
  ev: DeloitteEvent,
  promoted: { email: string; name?: string },
): Promise<void> {
  if (!ev.disableEmails) {
    try {
      const lang = ev.emailLanguage || 'EN';
      const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
      const promoteVars = {
        Name: promotedFirstName,
        EventTitle: ev.title,
        Organizer: formatOrganizerList(ev.organizers, lang),
        AppUrl: `${svc.siteUrl}/SitePages/DEX.aspx?env=WebView`,
        WaitlistPosition: '',
      };
      let emailData: { subject: string; body: string };
      const spTplRaw = await svc.getEmailTemplate('Nachruecken', lang).catch(() => null);
      const spTpl = applyEventTemplateOverride(spTplRaw, ev.emailTemplateOverrides, 'Nachruecken');
      if (spTpl) {
        emailData = buildEmailFromTemplate(spTpl, promoteVars);
      } else {
        emailData = promotionEmail(promotedFirstName, ev.title);
      }
      await svc.queueEmail(
        withParentTitleSubject(emailData.subject, ev.parentEventId ? allEvents.find(e => e.id === ev.parentEventId) : undefined),
        promoted.email, promoted.name || '', emailData.body,
        'Nachruecken', ev.title, ev.id
      );
    } catch (err) { console.warn('[DEX] promote-email failed:', err); }
  }
  if (!ev.disableOutlook) {
    try {
      await svc.queueOutlookEvent(promoted.email, ev.id, ev.title, 'Einladen');
    } catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
  }
}

export interface HealAllDeps {
  svc: EventService;
  allEvents: DeloitteEvent[];
  isDe: boolean;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  confirmDialog: (message: React.ReactNode, opts?: { confirmLabel?: string }) => Promise<boolean>;
  /** Fortschrittszeile für die Kachel („Prüfe 3/12: …"). `null` = fertig. */
  onProgress: (label: string | null) => void;
}

/** Ergebnis in einer Zeile für die Kachel; `cancelled` = im Dialog abgebrochen. */
export async function healAllEvents(d: HealAllDeps): Promise<{ text: string; isError: boolean; cancelled: boolean }> {
  const { svc, allEvents, isDe, getAllRegistrations, confirmDialog, onProgress } = d;
  // Alle aktiven Events mit Subsite — auch Sub-Events, jeder Termin hat eigene
  // Liste, eigenen Zähler, eigene Warteliste. Klammern im subEventsOnlyMode
  // raus: dort ist niemand buchbar, die Zeilen sind Schatten (v15.25), und
  // `cap 0` würde als „unbegrenzt" gelesen.
  const seen = new Set<string>();
  const candidates = allEvents.filter(e => {
    const sub = (e.subsiteUrl || '').trim();
    if (!sub || e.status !== 'Active' || e.subEventsOnlyMode) return false;
    if (seen.has(sub)) return false;
    seen.add(sub);
    return true;
  });
  if (candidates.length === 0) {
    return { text: isDe ? 'Keine aktiven Events mit Teilnehmerliste.' : 'No active events with a participant list.', isError: false, cancelled: false };
  }

  // ---- Phase 1: planen ---------------------------------------------------
  type Planned = { ev: DeloitteEvent; plan: PromotionPlan; idGap: boolean; blocked: number | null };
  const planned: Planned[] = [];
  for (let i = 0; i < candidates.length; i++) {
    const ev = candidates[i];
    onProgress(isDe ? `Prüfe ${i + 1}/${candidates.length}: ${ev.title}` : `Checking ${i + 1}/${candidates.length}: ${ev.title}`);
    // Objekt statt `let`, weil TS eine Zuweisung im Rückruf nicht sieht und
    // die Variable sonst als `null` festschreibt.
    const err: { status: number | null } = { status: null };
    const regs = await getAllRegistrations(ev.id, s => { err.status = s; });
    planned.push({ ev, plan: buildPromotionPlan(ev, regs, isDe), idGap: idSequenceCheck(regs).recent, blocked: err.status });
  }
  const blocked = planned.filter(p => p.blocked !== null);
  const shared = planned.filter(p => p.blocked === null && p.plan.sharedWaitlist && p.plan.anyWaiting);
  const usable = planned.filter(p => p.blocked === null && !p.plan.sharedWaitlist);
  const withPromote = usable.filter(p => p.plan.total > 0);
  const withGap = usable.filter(p => p.idGap);
  const totalPromote = withPromote.reduce((n, p) => n + p.plan.total, 0);

  const evLabel = (ev: DeloitteEvent): string => {
    const parent = ev.parentEventId ? allEvents.find(e => e.id === ev.parentEventId) : undefined;
    return parent ? `${parent.title} › ${shortSubEventTitle(ev.title, parent.title)}` : ev.title;
  };
  const lines: string[] = [];
  lines.push(isDe ? `${usable.length} ${usable.length === 1 ? 'Event' : 'Events'} geprüft.` : `${usable.length} ${usable.length === 1 ? 'event' : 'events'} checked.`);
  if (totalPromote > 0) {
    lines.push('');
    lines.push(isDe
      ? `${totalPromote} ${totalPromote === 1 ? 'Person rückt' : 'Personen rücken'} nach:`
      : `${totalPromote} ${totalPromote === 1 ? 'person moves' : 'people move'} up:`);
    for (const p of withPromote) {
      lines.push(`• ${evLabel(p.ev)} — ${p.plan.total}`);
      for (const l of promotionPlanLines(p.plan)) lines.push(`    ${l}`);
    }
  } else {
    lines.push(isDe ? 'Niemand muss nachrücken.' : 'Nobody needs to move up.');
  }
  lines.push('');
  lines.push(isDe
    ? `${withGap.length} ${withGap.length === 1 ? 'Event' : 'Events'} mit Nummern-Lücken ${withGap.length === 1 ? 'wird' : 'werden'} neu nummeriert; die Platzzähler aller ${usable.length} Events werden abgeglichen.`
    : `${withGap.length} ${withGap.length === 1 ? 'event' : 'events'} with ID gaps will be renumbered; the seat counters of all ${usable.length} events will be reconciled.`);
  if (shared.length > 0) {
    lines.push('');
    lines.push(isDe
      ? `Nicht automatisch — gemeinsame Warteliste bei geteilten Gruppen, bitte je Event im Organizer Center über „Freie Plätze mit Warteliste füllen" (${shared.length}):`
      : `Not automatic — shared waitlist with split groups, please use “Fill free seats from waitlist” per event in the Organizer Center (${shared.length}):`);
    for (const p of shared) lines.push(`• ${evLabel(p.ev)}`);
  }
  if (blocked.length > 0) {
    lines.push('');
    lines.push(isDe
      ? `Übersprungen — kein Vollzugriff auf die Teilnehmerliste (${blocked.length}):`
      : `Skipped — no full access to the participant list (${blocked.length}):`);
    for (const p of blocked) lines.push(`• ${evLabel(p.ev)} (HTTP ${p.blocked})`);
  }
  if (totalPromote > 0) {
    lines.push('');
    lines.push(isDe
      ? 'Jede nachgerückte Person bekommt den Status „Angemeldet", eine Nachrück-Mail und eine Outlook-Einladung.'
      : 'Each promoted person gets status “Registered”, a promotion email and an Outlook invite.');
  }
  onProgress(null);
  const ok = await confirmDialog(lines.join('\n'), {
    confirmLabel: totalPromote > 0 ? (isDe ? 'Nachrücken & heilen' : 'Promote & heal') : (isDe ? 'Heilen' : 'Heal'),
  });
  if (!ok) return { text: '', isError: false, cancelled: true };

  // ---- Phase 2: ausführen -----------------------------------------------
  let promotedTotal = 0;
  let reorderedEvents = 0;
  let syncedEvents = 0;
  let errors = 0;
  const promotedNames: string[] = [];
  for (let i = 0; i < usable.length; i++) {
    const { ev, plan, idGap } = usable[i];
    const sub = (ev.subsiteUrl || '').trim();
    onProgress(isDe ? `Heile ${i + 1}/${usable.length}: ${ev.title}` : `Healing ${i + 1}/${usable.length}: ${ev.title}`);
    let promotedHere = 0;
    for (const g of plan.groups) {
      for (let k = 0; k < g.count; k++) {
        // Obergrenze bewusst NICHT mitgeben — dieselbe Begründung wie in
        // runManualPromote: Der Service zählt über die ganze Liste und kann
        // eine Gruppe nicht trennen; die Anzahl steht oben fest.
        const promoted = await svc.promoteFirstWaitlistItem(sub, undefined, undefined, g.key);
        if (promoted && promoted.success && promoted.email) {
          promotedHere++;
          promotedNames.push(promoted.name || promoted.email);
          await notifyPromotedFor(svc, allEvents, ev, { email: promoted.email, name: promoted.name });
        } else {
          // Warteliste unerwartet leer (jemand hat sich zwischendurch
          // abgemeldet) — kein Fehler, nur nichts mehr zu tun.
          break;
        }
      }
    }
    promotedTotal += promotedHere;
    // IDs nur anfassen, wo es nötig ist: Der Reorder schreibt jede geänderte
    // Zeile, und auf 30 sauberen Events wäre das reines Rauschen gegen das
    // SharePoint-Throttling.
    if (idGap || promotedHere > 0) {
      try {
        const r = await svc.reorderParticipantIDs(sub, () => { /* Fortschritt je Event nicht nötig */ });
        reorderedEvents++;
        errors += r.errors;
      } catch (e) {
        errors++;
        console.warn('[DEX] healAll: reorder failed', ev.title, e);
      }
    }
    try {
      await svc.syncSeatsToActiveCount(sub, { isSplit: isSplitCapacityOf(ev) });
      syncedEvents++;
    } catch (e) {
      errors++;
      console.warn('[DEX] healAll: seat sync failed', ev.title, e);
    }
  }
  onProgress(null);

  const parts: string[] = [];
  parts.push(isDe
    ? `${promotedTotal} ${promotedTotal === 1 ? 'Person' : 'Personen'} nachgerückt`
    : `${promotedTotal} ${promotedTotal === 1 ? 'person' : 'people'} promoted`);
  if (promotedNames.length > 0) parts[0] += ` (${promotedNames.slice(0, 6).join(', ')}${promotedNames.length > 6 ? ', …' : ''})`;
  parts.push(isDe ? `${reorderedEvents} Events neu nummeriert` : `${reorderedEvents} events renumbered`);
  parts.push(isDe ? `${syncedEvents} Zähler abgeglichen` : `${syncedEvents} counters reconciled`);
  if (shared.length > 0) parts.push(isDe ? `${shared.length} mit gemeinsamer Warteliste — bitte einzeln` : `${shared.length} with shared waitlist — please handle individually`);
  if (blocked.length > 0) parts.push(isDe ? `${blocked.length} übersprungen (kein Zugriff)` : `${blocked.length} skipped (no access)`);
  if (errors > 0) parts.push(isDe ? `${errors} Fehler — Konsole prüfen` : `${errors} errors — check console`);
  return { text: parts.join(' · '), isError: errors > 0, cancelled: false };
}
