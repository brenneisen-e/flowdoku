/* useMailComposers — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 3667-4096 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { AudiencePerson, MassmailAudience } from '../../admin/adminTypes';
import { DeloitteEvent } from '../../../types';
import { MailHeaderImage, applyHeroImage, hasOwnHeaderImage, mailHeaderOpts } from '../../../utils/mailHeaderImage';
import { eventHeaderImageLayout } from '../../admin/adminConstants';
import { getCachedImage } from '../../../utils/imageCache';
import { isB2RunKoelnTitle } from '../../../data/b2runKoeln';
import { replacePlaceholders, wrapTemplate } from '../../../services/EmailTemplates';
import { EventService, SPRegistration } from '../../../services/EventService';

export interface UseMailComposersCtx {
  currentUser: import("../../../types/index").User;
  emailBody: string;
  emailHeading: string;
  emailSubject: string;
  eventServiceRef: EventService;
  getGroupMembers: (groupEmail: string) => Promise<{ groupName: string; members: { email: string; displayName: string; firstName?: string; lastName?: string; jobTitle?: string; location?: string; }[]; }>;
  inviteBody: string;
  inviteEventPhotoB64: string;
  inviteHeaderImage: MailHeaderImage;
  inviteHeading: string;
  inviteHydratingRef: React.MutableRefObject<boolean>;
  inviteSubheading: string;
  inviteSubject: string;
  inviteTarget: "organizer" | "audience" | "pending" | "uninvited";
  isDe: boolean;
  massmailEventPhotoB64: string;
  massmailHeaderImage: MailHeaderImage;
  massmailHydratingRef: React.MutableRefObject<boolean>;
  massmailMode: "closed" | "pick" | "paste" | "editor";
  massmailSubheading: string;
  pendingCheckBusy: boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setEmailBody: React.Dispatch<React.SetStateAction<string>>;
  setEmailHeading: React.Dispatch<React.SetStateAction<string>>;
  setEmailSubject: React.Dispatch<React.SetStateAction<string>>;
  setInviteAddInput: React.Dispatch<React.SetStateAction<string>>;
  setInviteAudienceOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setInviteBody: React.Dispatch<React.SetStateAction<string>>;
  setInviteCustomEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setInviteDraftSaved: React.Dispatch<React.SetStateAction<boolean>>;
  setInviteEventPhotoB64: React.Dispatch<React.SetStateAction<string>>;
  setInviteHeaderImage: React.Dispatch<React.SetStateAction<MailHeaderImage>>;
  setInviteHeading: React.Dispatch<React.SetStateAction<string>>;
  setInviteSubheading: React.Dispatch<React.SetStateAction<string>>;
  setInviteSubject: React.Dispatch<React.SetStateAction<string>>;
  setInviteTarget: React.Dispatch<React.SetStateAction<"organizer" | "audience" | "pending" | "uninvited">>;
  setInvitedLc: React.Dispatch<React.SetStateAction<Set<string>>>;
  setMassmailAudience: React.Dispatch<React.SetStateAction<MassmailAudience>>;
  setMassmailCc: React.Dispatch<React.SetStateAction<string[]>>;
  setMassmailDraftSaved: React.Dispatch<React.SetStateAction<boolean>>;
  setMassmailEventPhotoB64: React.Dispatch<React.SetStateAction<string>>;
  setMassmailHeaderImage: React.Dispatch<React.SetStateAction<MailHeaderImage>>;
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setMassmailSubheading: React.Dispatch<React.SetStateAction<string>>;
  setMassmailTestMsg: React.Dispatch<React.SetStateAction<string>>;
  setMassmailTesting: React.Dispatch<React.SetStateAction<boolean>>;
  setPendingCheckBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setPendingPeople: React.Dispatch<React.SetStateAction<{ people: AudiencePerson[]; reachable: number; }>>;
  setShowInviteModal: React.Dispatch<React.SetStateAction<boolean>>;
  setVisibilityResolved: React.Dispatch<React.SetStateAction<AudiencePerson[]>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  showEmailModal: boolean;
  showInviteModal: boolean;
  siteUrl: string;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
  visibilityResolved: AudiencePerson[];
}

export interface UseMailComposersResult {
  applyInviteHero: (wrappedHtml: string) => string;
  applyMassmailHero: (wrappedHtml: string) => string;
  inviteHeaderOpts: { imageWidth: number; imagePaddingV: number; imagePaddingH: number; };
  massmailHeaderOpts: { imageWidth: number; imagePaddingV: number; imagePaddingH: number; };
  openInviteModal: () => void;
  openMassmailPicker: () => void;
  openPendingReminder: () => Promise<void>;
  resetInviteDraft: () => void;
  resetMassmailDraft: () => void;
  resolveAudienceEmails: (ev: DeloitteEvent) => Promise<AudiencePerson[]>;
  saveInviteDraft: () => void;
  saveMassmailDraft: () => void;
  sendMassmailTestToOrganizers: () => Promise<void>;
}

export function useMailComposers(ctx: UseMailComposersCtx): UseMailComposersResult {
  const {
    currentUser, emailBody, emailHeading, emailSubject, eventServiceRef, getGroupMembers,
    inviteBody, inviteEventPhotoB64, inviteHeaderImage, inviteHeading, inviteHydratingRef,
    inviteSubheading, inviteSubject, inviteTarget, isDe, massmailEventPhotoB64,
    massmailHeaderImage, massmailHydratingRef, massmailMode, massmailSubheading, pendingCheckBusy,
    registrations, selectedEvent, setEmailBody, setEmailHeading, setEmailSubject,
    setInviteAddInput, setInviteAudienceOpen, setInviteBody, setInviteCustomEmails,
    setInviteDraftSaved, setInviteEventPhotoB64, setInviteHeaderImage, setInviteHeading,
    setInviteSubheading, setInviteSubject, setInviteTarget, setInvitedLc, setMassmailAudience,
    setMassmailCc, setMassmailDraftSaved, setMassmailEventPhotoB64, setMassmailHeaderImage,
    setMassmailMode, setMassmailPasteRaw, setMassmailSubheading, setMassmailTestMsg,
    setMassmailTesting, setPendingCheckBusy, setPendingPeople, setShowInviteModal,
    setVisibilityResolved, showAlert, showEmailModal, showInviteModal, siteUrl,
    subEventRegsByEventId, visibilityResolved,
  } = ctx;
  // v22.5: Einladungsmail — Default-Texte bauen, Entwurf laden/speichern
  // (localStorage pro Event), Modal öffnen, zurücksetzen.
  const inviteDraftKey = (id: string): string => `dex_invite_draft_${id}`;
  // v26.89: B2Run-Köln-Events bekommen einen eigenen, dynamischen Einladungs-
  // text-Vorschlag (bilingual DE + EN) — mit Datum, Ort und Platzzahl aus dem
  // Event. Der Organizer kann ihn wie jeden anderen Entwurf frei überschreiben.
  const buildB2RunKoelnInviteDefaults = (ev: DeloitteEvent, appUrl: string, signatureNames: string): { subject: string; heading: string; subheading: string; body: string } => {
    const start = ev.startDate ? new Date(ev.startDate) : null;
    const validStart = start && !isNaN(start.getTime()) ? start : null;
    const dateDe = validStart ? validStart.toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }) : '';
    const dateEn = validStart ? validStart.toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }) : '';
    const venue = (ev.location || '').trim() || 'RheinEnergieStadion';
    const plaetze = (ev.maxParticipants && ev.maxParticipants > 0) ? ev.maxParticipants : 100;
    const dateLineDe = dateDe || 'Datum folgt';
    const dateLineEn = dateEn || 'date to follow';
    // „DEX App" als grün gestylter Link (statt der langen URL im Klartext).
    const appLinkDe = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">DEX App</a>`;
    const appLinkEn = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">DEX app</a>`;
    const b2runSite = '<a href="https://www.b2run.de" style="color:#86bc25;font-weight:600;">B2RUN Website</a>';
    const de = `
<p>Liebes Team,</p>
<p>es ist so weit: wir haben unseren Standort wieder für den <strong>B2Run Firmenlauf</strong> angemeldet. Es werden vorerst <strong>${plaetze} Startplätze</strong> zur Verfügung stehen. Diese werden nach dem Motto „First come, first run" vergeben.</p>
<p>Die Startgebühren inkl. einer Spende an Menschen für Menschen, die Kosten für ein Laufshirt sowie ein Teamzelt mit einer kleinen Verpflegung nach dem Lauf werden dabei von Deloitte übernommen.</p>
<p style="text-align:center;"><strong>Die Anmeldung ist ab sofort über die ${appLinkDe} möglich.</strong></p>
<p><strong>Wichtige Hinweise:</strong></p>
<ul>
  <li>Falls ihr es nicht unter die ersten ${plaetze} schaffen solltet, meldet euch bitte trotzdem über die App an. Ihr werdet automatisch der Reihe nach auf eine Warteliste gesetzt.</li>
  <li>Falls ihr aus wichtigen Gründen nicht am Lauf teilnehmen könnt, sagt eure Teilnahme bitte frühzeitig über die App („My Events") wieder ab. Die Plätze werden automatisch von der Warteliste – der Reihe nach – vergeben und ihr erhaltet eine automatische E-Mail.</li>
  <li>Bitte meldet euch nur an, wenn ihr auch wirklich am B2RUN teilnehmen könnt und möchtet. <strong>Wir zahlen für jede Anmeldung eine Teilnehmergebühr, die wir im Falle eines No-Shows nicht erstattet bekommen.</strong></li>
</ul>
<p><strong>Infos zum Lauf/Event:</strong></p>
<ul>
  <li>${dateLineDe} am <strong>${venue}</strong></li>
  <li>Beginn der Veranstaltung: 15:00 Uhr</li>
  <li>Teamtreff Deloitte, Startnummernübergabe und Aufwärmen: 16:00 Uhr</li>
  <li>Startzeit Deloitte: 17:00 Uhr</li>
  <li>Distanz: 5,3 km</li>
  <li>Anschließend: Get-Together, Teamfotos und Catering sowie Afterparty im Stadioninnenraum (ab 20:00 Uhr)</li>
</ul>
<p><em>Genauere Infos zu den Zeiten und Treffpunkten werden wir euch Mitte / Ende August mitteilen.</em></p>
<p><strong>Startfelder</strong></p>
<p>Beim B2Run gibt es zwei Startfelder: „Funstarter" und „Durchstarter".</p>
<ul>
  <li>Das <strong>Durchstarter</strong>-Feld ist für schnelle und ambitionierte Läufer:innen gedacht, die „freie Bahn" haben möchten (Richtwerte Männer &lt;4 Min/km, Frauen &lt;5 Min/km).</li>
  <li>Für das <strong>Funstarter</strong>-Feld gibt es keine Richtwerte – hier steht der Laufspaß im Vordergrund.</li>
  <li>Sofern jemand am Durchstarter-Lauf teilnehmen möchte, wählt dies entsprechend bei der Anmeldung aus. Die Laufstrecke für beide Startfelder beträgt 5,3 Kilometer.</li>
</ul>
<p><strong>Laufshirts</strong></p>
<p>Für jede:n Läufer:in gibt es ein Deloitte-Laufshirt. Wählt bitte bei der Anmeldung eure Größe aus.</p>
<p>Weitere Informationen sind auf der ${b2runSite} zu finden.</p>
<p>Bei Fragen wendet euch bitte direkt an unser Gruppenpostfach.</p>
<p>Auf die Plätze, fertig, los.</p>
<p>Mit sportlichen Grüßen<br />${signatureNames}</p>`.trim();
    const en = `
<p>Dear team,</p>
<p>The time has come: we have registered our location for the <strong>B2Run company run</strong> again. For now <strong>${plaetze} starting places</strong> will be available, allocated on a „first come, first run" basis.</p>
<p>The registration fees (incl. a donation to Menschen für Menschen), the cost of a running shirt and a team tent with light refreshments after the run are covered by Deloitte.</p>
<p style="text-align:center;"><strong>Registration is now open via the ${appLinkEn}.</strong></p>
<p><strong>Important information:</strong></p>
<ul>
  <li>If you don't make it into the first ${plaetze}, please register via the app anyway. You will automatically be placed on a waiting list in order of registration.</li>
  <li>If you are unable to take part for important reasons, please cancel your participation early via the app („My Events"). Places are allocated automatically from the waiting list – in order – and you will receive an automatic email.</li>
  <li>Please only register if you can really take part in the B2RUN. <strong>We pay a participation fee for each registration that is not refunded in the event of a no-show.</strong></li>
</ul>
<p><strong>Event details:</strong></p>
<ul>
  <li>${dateLineEn} at <strong>${venue}</strong></li>
  <li>Start of the event: 3:00 p.m.</li>
  <li>Deloitte team meeting, race-number handout and warm-up: 4:00 p.m.</li>
  <li>Deloitte start time: 5:00 p.m.</li>
  <li>Distance: 5.3 km</li>
  <li>Afterwards: get-together, team photos and catering plus after-party inside the stadium (from 8:00 p.m.)</li>
</ul>
<p><em>We will share more detailed information about times and meeting points in mid/late August.</em></p>
<p><strong>Starting fields</strong></p>
<p>B2Run has two starting fields: „Funstarter" and „Durchstarter".</p>
<ul>
  <li>The <strong>Durchstarter</strong> field is intended for fast and ambitious runners who want a „clear track" (guideline times: men &lt;4 min/km, women &lt;5 min/km).</li>
  <li>The <strong>Funstarter</strong> field has no guideline times – the focus here is on having fun.</li>
  <li>If you would like to take part in the Durchstarter run, please select this when registering. The distance for both fields is 5.3 kilometres.</li>
</ul>
<p><strong>Running shirts</strong></p>
<p>Every runner receives a Deloitte running shirt. Please select your size when registering.</p>
<p>Further information can be found on the ${b2runSite}.</p>
<p>If you have any questions, please contact our group mailbox directly.</p>
<p>On your marks, get set, go.</p>
<p>With sporting regards<br />${signatureNames}</p>`.trim();
    // v26.90: Zweisprachig — Hauptsprache oben, die jeweils andere Version unten
    // (per Trennlinie abgesetzt), wie es die B2Run-Kommunikation üblicherweise macht.
    const divider = '<p style="margin:28px 0 20px;border-top:1px solid #d0d0ce;"></p>';
    return {
      subject: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
      heading: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
      subheading: isDe ? `${dateLineDe} · ${plaetze} Startplätze` : `${dateLineEn} · ${plaetze} starting places`,
      body: isDe ? `${de}\n${divider}\n${en}` : `${en}\n${divider}\n${de}`,
    };
  };
  const buildInviteDefaults = (ev: DeloitteEvent): { subject: string; heading: string; subheading: string; body: string } => {
    const appUrl = `${siteUrl}/SitePages/DEX.aspx?env=WebView`;
    const linkHtml = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">${appUrl}</a>`;
    const orgList = (ev.organizers || []).map(s => (s || '').trim()).filter(Boolean);
    const teamLine = isDe ? `Das ${ev.title} Orga Team` : `The ${ev.title} Organizer Team`;
    const signatureNames = orgList.length > 0 ? `${teamLine}<br />${orgList.join('<br />')}` : teamLine;
    // v26.89: B2Run-Köln-Events erhalten den spezialisierten Vorschlag.
    if (isB2RunKoelnTitle(ev.title)) {
      return buildB2RunKoelnInviteDefaults(ev, appUrl, signatureNames);
    }
    const body = isDe
      ? `<p>Hallo,</p>\n<p>wir laden dich herzlich zum Event <strong>${ev.title}</strong> ein.</p>\n<p>Du kannst dich ab sofort über unsere Event-Plattform anmelden:</p>\n<p>${linkHtml}</p>\n<p>Falls du dich im Nachgang doch nicht beteiligen kannst, ist eine <strong>Abmeldung jederzeit über dieselbe Plattform</strong> möglich — bitte gib uns rechtzeitig Bescheid, damit Wartelisten-Plätze nachrücken können.</p>\n<p>Bei Rückfragen meld dich gern bei uns.</p>\n<p>Viele Grüße<br />${signatureNames}</p>`
      : `<p>Hello,</p>\n<p>we would like to invite you to the event <strong>${ev.title}</strong>.</p>\n<p>You can register via our event platform:</p>\n<p>${linkHtml}</p>\n<p>If you change your mind, you can <strong>cancel anytime via the same platform</strong> — please let us know early so people on the waitlist can move up.</p>\n<p>Feel free to reach out if you have any questions.</p>\n<p>Best regards<br />${signatureNames}</p>`;
    return {
      subject: isDe ? `Einladung: ${ev.title}` : `Invitation: ${ev.title}`,
      heading: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
      subheading: '',
      body,
    };
  };
  const applyInviteDraftOrDefaults = (ev: DeloitteEvent): void => {
    inviteHydratingRef.current = true;
    let loaded: { subject?: string; heading?: string; subheading?: string; body?: string; target?: string } | null = null;
    try {
      const raw = window.localStorage.getItem(inviteDraftKey(ev.id));
      if (raw) loaded = JSON.parse(raw);
    } catch { /* localStorage evtl. blockiert */ }
    const def = buildInviteDefaults(ev);
    setInviteSubject(loaded && typeof loaded.subject === 'string' ? loaded.subject : def.subject);
    setInviteHeading(loaded && typeof loaded.heading === 'string' ? loaded.heading : def.heading);
    setInviteSubheading(loaded && typeof loaded.subheading === 'string' ? loaded.subheading : def.subheading);
    setInviteBody(loaded && typeof loaded.body === 'string' ? loaded.body : def.body);
    setInviteTarget(loaded && loaded.target === 'audience' ? 'audience' : 'organizer');
    // v29.37: dito für Einladung und Erinnerung.
    setInviteHeaderImage(p => ({ ...p, ...eventHeaderImageLayout(ev.emailTemplateOverrides) }));
    // Hydration-Flag im nächsten Tick freigeben, damit das Auto-Speichern erst
    // auf echte Nutzer-Edits reagiert (nicht auf das initiale Laden).
    window.setTimeout(() => { inviteHydratingRef.current = false; }, 0);
  };
  /**
   * v29.32: Verteiler des Events in einzelne Personen auflösen — derselbe Weg
   * wie beim Einladungsversand (Graph-Gruppenmitglieder, Fallback auf die beim
   * Event-Speichern eingefrorene Liste). Ausgeschlossene Adressen fliegen raus,
   * denn wer auf der Ausschluss-Liste steht, sieht das Event nie.
   *
   * Wichtig: Ein reiner STANDORT-Filter lässt sich hier nicht abzählen — dafür
   * müsste man das ganze Verzeichnis lesen. Die Zeile sagt das dann auch, statt
   * eine Zahl zu erfinden.
   */
  const resolveAudienceEmails = async (ev: DeloitteEvent): Promise<AudiencePerson[]> => {
    const entries = (ev.audienceFilter || []).map(s => (s || '').trim()).filter(Boolean);
    const excluded = new Set((ev.excludedUsers || []).map(e => (e || '').toLowerCase().trim()).filter(Boolean));
    const out: AudiencePerson[] = [];
    const seen = new Set<string>();
    // v29.36: Name/Position/Standort mitnehmen, wenn die Verteiler-Abfrage sie
    // liefert — der Nachfass-Schritt zeigt Personen, keine Adressliste.
    const push = (e: string, extra?: Partial<AudiencePerson>): void => {
      const lc = (e || '').trim().toLowerCase();
      if (lc && lc.indexOf('@') > 0 && !seen.has(lc) && !excluded.has(lc)) {
        seen.add(lc);
        out.push({ email: lc, displayName: extra?.displayName || '', jobTitle: extra?.jobTitle || '', location: extra?.location || '' });
      }
    };
    for (const entry of entries) {
      if (entry.indexOf('@') < 0) continue; // Standort-Pattern — nicht auflösbar
      try {
        const grp = await getGroupMembers(entry);
        if (grp && grp.members && grp.members.length > 0) {
          grp.members.forEach(m => push(m.email, { displayName: m.displayName, jobTitle: m.jobTitle, location: m.location }));
        } else push(entry);
      } catch { push(entry); }
    }
    if (out.length === 0) (ev.audienceResolvedEmails || []).forEach(e => push(e));
    return out;
  };

  /**
   * v29.32: „Wer hat noch nicht geantwortet?" — die aufgelöste Sichtbarkeits-
   * Liste minus aller Personen, die sich bereits geäußert haben. Als Antwort
   * zählt JEDE Zeile im Event: Anmeldung, Warteliste, Check-in, Abmeldung und
   * die proaktive Absage („Ich nehme nicht teil"). Bei einer Klammer zählen
   * auch die Zeilen der Sub-Events — wer dort gebucht hat, hat geantwortet.
   *
   * Das Ergebnis geht in den bestehenden Einladungs-Dialog (Modus „Nachfassen",
   * Empfängerliste editierbar, Mailtext editierbar). Bewusst KEIN zweiter
   * Versand-Dialog daneben — der bestehende kann das alles bereits.
   */
  const openPendingReminder = async (): Promise<void> => {
    if (!selectedEvent || pendingCheckBusy) return;
    setPendingCheckBusy(true);
    try {
      const audience = visibilityResolved || await resolveAudienceEmails(selectedEvent);
      if (!visibilityResolved) setVisibilityResolved(audience);
      if (audience.length === 0) {
        showAlert(isDe
          ? 'Für dieses Event kennt DEX keine Namen. Das ist so, wenn die Sichtbarkeit nur über den Standort läuft — dahinter steht keine Liste einzelner Personen. Trage in Schritt 3 des Event-Edits zusätzlich einen Mailverteiler oder einzelne Personen ein, dann kann DEX nachfassen.'
          : 'DEX does not know any names for this event. That is the case when visibility runs via location only — there is no list of individual people behind it. Add a distribution list or individual people in step 3 of the event edit, then DEX can follow up.',
          { variant: 'info' });
        return;
      }
      // Wer hat schon geantwortet? Alle Zeilen des Events + (bei einer Klammer)
      // der Sub-Events, unabhängig vom Status.
      const decided = new Set<string>();
      registrations.forEach(r => { const e = (r.ParticipantEmail || '').toLowerCase().trim(); if (e) decided.add(e); });
      Object.keys(subEventRegsByEventId || {}).forEach(k => {
        (subEventRegsByEventId[k] || []).forEach(r => {
          const e = (r.ParticipantEmail || '').toLowerCase().trim(); if (e) decided.add(e);
        });
      });
      // Organizer-Team zählt nicht als offener Fall — es organisiert das Event.
      const team = new Set<string>([
        ...(selectedEvent.organizerEmails || []),
        ...(selectedEvent.coOrganizerEmails || []),
      ].map(e => (e || '').toLowerCase().trim()).filter(Boolean));
      const pending = audience.filter(p => !decided.has(p.email) && !team.has(p.email));
      if (pending.length === 0) {
        showAlert(isDe
          ? `Alle ${audience.length} Personen, die das Event sehen können, haben bereits geantwortet — angemeldet, abgemeldet oder abgesagt. Es gibt niemanden zum Erinnern.`
          : `All ${audience.length} people who can see this event have already responded — registered, cancelled or declined. There is nobody to remind.`,
          { variant: 'success' });
        return;
      }
      // v29.36: ERST die Übersicht, wer fehlt (Foto, Name, Position) — die Mail
      // kommt im zweiten Schritt. Vorher landete man direkt im Mail-Dialog und
      // musste einer Adressliste glauben, ohne zu sehen, wen man da anschreibt.
      setPendingPeople({ people: pending, reachable: audience.length });
    } catch (err) {
      showAlert((isDe ? 'Prüfung fehlgeschlagen: ' : 'Check failed: ') + String((err as Error)?.message || err), { variant: 'error' });
    } finally {
      setPendingCheckBusy(false);
    }
  };

  const openInviteModal = (): void => {
    if (!selectedEvent) return;
    applyInviteDraftOrDefaults(selectedEvent);
    // v28.37: Anpassungen der letzten Runde nicht mitschleppen und die
    // bereits verschickten Einladungen im Hintergrund nachladen (für den
    // Modus „Nur an noch nicht Eingeladene").
    setInviteCustomEmails(null);
    setInviteAddInput('');
    setInviteAudienceOpen(false);
    setInvitedLc(null);
    if (eventServiceRef) {
      eventServiceRef.getInvitedRecipients(selectedEvent.id)
        .then(list => setInvitedLc(new Set(list)))
        .catch(() => setInvitedLc(new Set<string>()));
    } else {
      setInvitedLc(new Set<string>());
    }
    setShowInviteModal(true);
  };
  const resetInviteDraft = (): void => {
    if (!selectedEvent) return;
    try { window.localStorage.removeItem(inviteDraftKey(selectedEvent.id)); } catch { /* */ }
    inviteHydratingRef.current = true;
    const def = buildInviteDefaults(selectedEvent);
    setInviteSubject(def.subject);
    setInviteHeading(def.heading);
    setInviteSubheading(def.subheading);
    setInviteBody(def.body);
    setInviteDraftSaved(false);
    window.setTimeout(() => { inviteHydratingRef.current = false; }, 0);
  };
  // v22.5/v22.6: expliziter „Entwurf speichern"-Klick — schreibt den aktuellen
  // Stand sofort in localStorage und zeigt kurz „Gespeichert".
  const saveInviteDraft = (): void => {
    if (!selectedEvent) return;
    try {
      window.localStorage.setItem(inviteDraftKey(selectedEvent.id), JSON.stringify({
        subject: inviteSubject, heading: inviteHeading, subheading: inviteSubheading,
        body: inviteBody, target: inviteTarget,
      }));
      setInviteDraftSaved(true);
      window.setTimeout(() => setInviteDraftSaved(false), 2500);
    } catch { /* localStorage evtl. blockiert */ }
  };
  // Auto-Speichern, solange das Modal offen ist — beim nächsten Öffnen wird der
  // Entwurf wiederhergestellt.
  React.useEffect(() => {
    if (!showInviteModal || !selectedEvent || inviteHydratingRef.current) return;
    try {
      window.localStorage.setItem(inviteDraftKey(selectedEvent.id), JSON.stringify({
        subject: inviteSubject, heading: inviteHeading, subheading: inviteSubheading,
        body: inviteBody, target: inviteTarget,
      }));
    } catch { /* */ }
  }, [showInviteModal, selectedEvent, inviteSubject, inviteHeading, inviteSubheading, inviteBody, inviteTarget]);

  // v22.9: Massenmail-Entwurf — Default-Texte, laden/speichern (localStorage pro
  // Event), Picker öffnen, zurücksetzen, Testmail an die Organizer.
  const massmailDraftKey = (id: string): string => `dex_massmail_draft_${id}`;
  const buildMassmailDefaults = (ev: DeloitteEvent): { subject: string; heading: string; body: string } => ({
    subject: `${ev.title} - Info`,
    heading: ev.title,
    body: '',
  });
  const applyMassmailDraftOrDefaults = (ev: DeloitteEvent): void => {
    massmailHydratingRef.current = true;
    let loaded: { subject?: string; heading?: string; subheading?: string; body?: string } | null = null;
    try {
      const raw = window.localStorage.getItem(massmailDraftKey(ev.id));
      if (raw) loaded = JSON.parse(raw);
    } catch { /* localStorage evtl. blockiert */ }
    const def = buildMassmailDefaults(ev);
    setEmailSubject(loaded && typeof loaded.subject === 'string' ? loaded.subject : def.subject);
    setEmailHeading(loaded && typeof loaded.heading === 'string' ? loaded.heading : def.heading);
    setMassmailSubheading(loaded && typeof loaded.subheading === 'string' ? loaded.subheading : '');
    setEmailBody(loaded && typeof loaded.body === 'string' ? loaded.body : def.body);
    // v29.37: Kopfbild-Größe aus dem Event übernehmen statt fest 180/30/30.
    setMassmailHeaderImage(p => ({ ...p, ...eventHeaderImageLayout(ev.emailTemplateOverrides) }));
    window.setTimeout(() => { massmailHydratingRef.current = false; }, 0);
  };
  const openMassmailPicker = (): void => {
    if (selectedEvent) applyMassmailDraftOrDefaults(selectedEvent);
    setMassmailAudience('active');
    setMassmailPasteRaw('');
    setMassmailTestMsg(null);
    setMassmailMode('pick');
  };
  const resetMassmailDraft = (): void => {
    if (!selectedEvent) return;
    try { window.localStorage.removeItem(massmailDraftKey(selectedEvent.id)); } catch { /* */ }
    massmailHydratingRef.current = true;
    const def = buildMassmailDefaults(selectedEvent);
    setEmailSubject(def.subject);
    setEmailHeading(def.heading);
    setMassmailSubheading('');
    setEmailBody(def.body);
    // v30.51: „Zurücksetzen" setzt auch das zusätzliche CC zurück — sonst
    // bliebe ein Verteiler stehen, den niemand mehr im Text sieht.
    setMassmailCc([]);
    setMassmailDraftSaved(false);
    window.setTimeout(() => { massmailHydratingRef.current = false; }, 0);
  };
  const saveMassmailDraft = (): void => {
    if (!selectedEvent) return;
    try {
      window.localStorage.setItem(massmailDraftKey(selectedEvent.id), JSON.stringify({
        subject: emailSubject, heading: emailHeading, subheading: massmailSubheading, body: emailBody,
      }));
      setMassmailDraftSaved(true);
      window.setTimeout(() => setMassmailDraftSaved(false), 2500);
    } catch { /* */ }
  };
  // Auto-Speichern, solange der Massenmail-Editor offen ist.
  React.useEffect(() => {
    if (massmailMode !== 'editor' || !showEmailModal || !selectedEvent || massmailHydratingRef.current) return;
    try {
      window.localStorage.setItem(massmailDraftKey(selectedEvent.id), JSON.stringify({
        subject: emailSubject, heading: emailHeading, subheading: massmailSubheading, body: emailBody,
      }));
    } catch { /* */ }
  }, [massmailMode, showEmailModal, selectedEvent, emailSubject, emailHeading, massmailSubheading, emailBody]);
  // v26.78: Event-Foto als Base64 vorladen, sobald der Massenmail-Editor
  // geöffnet wird — für die Live-Vorschau und den Versand (Bild-im-Kopf-Wahl).
  // Wechselt der Nutzer das Event, wird die Wahl auf „Standard" zurückgesetzt.
  React.useEffect(() => {
    if (!showEmailModal || !selectedEvent) return;
    setMassmailHeaderImage(p => ({ ...p, hero: 'logo' }));
    setMassmailEventPhotoB64('');
    const url = selectedEvent.imageUrl;
    if (!url) return;
    let cancelled = false;
    getCachedImage(url)
      .then(b64 => { if (!cancelled && b64 && b64.indexOf('data:') === 0) setMassmailEventPhotoB64(b64); })
      .catch(() => { /* Event-Foto nicht ladbar → Option bleibt ohne Vorschau/deaktiviert */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [showEmailModal, selectedEvent && selectedEvent.id]);
  // v26.88: Event-Foto AUCH für die Einladungsmail vorladen (Bild-im-Kopf-Wahl),
  // sobald das Einladungs-Fenster geöffnet wird. Beim Event-Wechsel zurück auf
  // „DEX-Logo".
  React.useEffect(() => {
    if (!showInviteModal || !selectedEvent) return;
    setInviteHeaderImage(p => ({ ...p, hero: 'logo' }));
    setInviteEventPhotoB64('');
    const url = selectedEvent.imageUrl;
    if (!url) return;
    let cancelled = false;
    getCachedImage(url)
      .then(b64 => { if (!cancelled && b64 && b64.indexOf('data:') === 0) setInviteEventPhotoB64(b64); })
      .catch(() => { /* Event-Foto nicht ladbar → Option bleibt deaktiviert */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [showInviteModal, selectedEvent && selectedEvent.id]);
  // v26.78: Ersetzt den {{ORB_URL}}-Platzhalter im gewickelten Mail-HTML durch
  // das gewählte Kopf-Bild. Bei „Event-Foto" (+ geladenem Foto) wird das
  // Event-Bild als Base64 fest eingebacken; sonst bleibt {{ORB_URL}} erhalten,
  // damit der Flow wie gehabt das Standard-Bild (DEX-Logo/Orb bzw. das
  // konfigurierte Mail-Logo des Events) einsetzt.
  // v30.52: EINE Umsetzung für beide (und die QR-Mail) — s. utils/mailHeaderImage.
  const applyMassmailHero = (wrappedHtml: string): string =>
    applyHeroImage(wrappedHtml, massmailHeaderImage, massmailEventPhotoB64);
  const applyInviteHero = (wrappedHtml: string): string =>
    applyHeroImage(wrappedHtml, inviteHeaderImage, inviteEventPhotoB64);
  // v29.37: Steht im Kopf ein eigenes Bild? Entweder das eingebackene Event-Foto
  // oder — wenn {{ORB_URL}} stehen bleibt — das Mail-Logo des Events, das der
  // Flow einsetzt. Nur dann darf die volle Breite gelten (sonst Orb-Deckel).
  const massmailHasOwnImage = hasOwnHeaderImage(massmailHeaderImage, massmailEventPhotoB64, selectedEvent && selectedEvent.mailImageBase64);
  const inviteHasOwnImage = hasOwnHeaderImage(inviteHeaderImage, inviteEventPhotoB64, selectedEvent && selectedEvent.mailImageBase64);
  const massmailHeaderOpts = mailHeaderOpts(massmailHeaderImage, massmailHasOwnImage);
  const inviteHeaderOpts = mailHeaderOpts(inviteHeaderImage, inviteHasOwnImage);
  // Testmail mit dem aktuellen Stand an die Organizer (zur Kontrolle vor dem
  // echten Massenversand). Geht NICHT an die Teilnehmer.
  const sendMassmailTestToOrganizers = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
    const to = orgEmails.length > 0 ? orgEmails.join(';') : (currentUser.email || '');
    if (!to) {
      setMassmailTestMsg(isDe ? 'Keine Organizer-E-Mail hinterlegt — Test nicht möglich.' : 'No organizer email available — test not possible.');
      return;
    }
    if (!emailSubject.trim() || !emailBody.trim()) {
      setMassmailTestMsg(isDe ? 'Bitte Betreff und Text ausfüllen.' : 'Please fill in subject and body.');
      return;
    }
    setMassmailTesting(true);
    setMassmailTestMsg(null);
    try {
      const previewVars: Record<string, string> = { EventTitle: selectedEvent.title, Organizer: (selectedEvent.organizers || []).join(', ') };
      const resolvedSubject = `[TEST] ${replacePlaceholders(emailSubject, previewVars)}`;
      const resolvedHeading = replacePlaceholders(emailHeading, previewVars);
      const resolvedBody = replacePlaceholders(emailBody, previewVars);
      const resolvedSub = massmailSubheading.trim() ? replacePlaceholders(massmailSubheading, previewVars) : `Event ${selectedEvent.title}`;
      const fullBody = applyMassmailHero(wrapTemplate('#86bc25', resolvedHeading, resolvedSub, resolvedBody, undefined, massmailHeaderOpts));
      await eventServiceRef.queueEmail(resolvedSubject, to, 'Organizer (Test)', fullBody, 'Massenmail', selectedEvent.title, selectedEvent.id);
      setMassmailTestMsg(isDe ? `Testmail an die Organizer (${to.split(';').length}) verschickt — bitte Postfach prüfen.` : `Test email sent to the organizers (${to.split(';').length}) — please check the mailbox.`);
    } catch (err) {
      setMassmailTestMsg((isDe ? 'Fehler beim Test-Versand: ' : 'Error during test send: ') + (err instanceof Error ? err.message : String(err)));
    }
    setMassmailTesting(false);
  };
  return {
    applyInviteHero, applyMassmailHero, inviteHeaderOpts, massmailHeaderOpts, openInviteModal,
    openMassmailPicker, openPendingReminder, resetInviteDraft, resetMassmailDraft,
    resolveAudienceEmails, saveInviteDraft, saveMassmailDraft, sendMassmailTestToOrganizers,
  };
}

