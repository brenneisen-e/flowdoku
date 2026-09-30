/**
 * v28.95: Aus `EventService` herausgeloest (dort 412 der 13.221 Zeilen).
 *
 * Die fertigen HTML-Koerper der System-Mails, die die App selbst in die
 * Warteschlange stellt (Outlook-Absage, Nachruecken, Team-Ereignisse,
 * Zimmerwunsch, Gruppenwechsel, Ueberbuchungs-Entschuldigung). Reine Daten:
 * `wrapTemplateForStorage` legt den Deloitte-Rahmen drumherum, der Flow
 * ersetzt später nur noch {{LOGO_URL}} und {{ORB_URL}}.
 *
 * v32.52: Alle Texte neu geschrieben (Nutzer-Auftrag 30.09.2026 „alle
 * Vorlagen neu"). Regeln, die dabei gelten und für neue Vorlagen weiter gelten:
 *  - Erst die Aussage in einem Satz, dann Fakten als Aufzählung, dann was die
 *    Person tun kann. Kein Satz, der den vorigen wiederholt.
 *  - Gruß einheitlich: Teilnehmer-Mails „Dein Event-Team", Hinweise an
 *    Organizer „Dein DEX-Team" (GRUSS_* unten).
 *  - Unterzeile nur `{{EventTitle}}` — das „Event " davor war doppelt.
 *  - „Organizer Center", nie „Admin Center"; neutral statt „der/die …/in".
 *  - Die Platzhalter der FLOW-Vorlagen (Nachrücken, Nachrücker-Info,
 *    Auto-Abmeldung, Outlook-Absage/-Weiterleitung/-Übersicht) sind
 *    unverändert: Der Flow ersetzt nur die, die er kennt — ein neuer
 *    Platzhalter stünde dort roh in der Mail.
 */
import { wrapTemplateForStorage } from './EmailTemplates';

const APP = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
const GRUSS_DE = '<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>';
const GRUSS_EN = '<p style="margin-top:24px;"><strong>Best regards</strong><br><br><strong>Your Event Team</strong></p>';
const GRUSS_DEX_DE = '<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein DEX-Team</strong></p>';
const GRUSS_DEX_EN = '<p style="margin-top:24px;"><strong>Best regards</strong><br><br><strong>Your DEX Team</strong></p>';
const UL = '<ul style="margin:8px 0 16px;padding-left:20px;line-height:1.7;">';
const knopf = (href: string, text: string, farbe = '#86bc25'): string =>
  `<p style="margin:24px 0;text-align:center;"><a href="${href}" style="display:inline-block;padding:12px 28px;background:${farbe};color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">${text}</a></p>`;
const leise = (text: string): string => `<p style="font-size:12px;color:#777;">${text}</p>`;

// Outlook-Absage, selbst abgelehnt (Flow DEX_OutlookDeclineHandler).
// Platzhalter: {{Name}}, {{EventTitle}}, {{CancelUrl}}.
export const OUTLOOK_DECLINE_BODY_EN = wrapTemplateForStorage(
  '#ed8b00',
  'You declined the Outlook invite',
  '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>you declined the Outlook invitation for <strong>{{EventTitle}}</strong> — but you are <strong>still registered</strong> and keep your spot.</p>
<p>Can’t attend? Then please cancel your registration as well, so your spot becomes free for someone else:</p>
${knopf('{{CancelUrl}}', 'Cancel my registration', '#da291c')}
${leise('Declined by mistake? Then simply ignore this email — your registration stays as it is.')}
${GRUSS_EN}`
);

export const OUTLOOK_DECLINE_BODY_DE = wrapTemplateForStorage(
  '#ed8b00',
  'Du hast den Outlook-Termin abgelehnt',
  '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>du hast die Outlook-Einladung für <strong>{{EventTitle}}</strong> abgelehnt — deine Anmeldung besteht aber <strong>weiterhin</strong>, dein Platz ist also noch belegt.</p>
<p>Kannst du nicht teilnehmen? Dann melde dich bitte auch hier ab, damit dein Platz für jemand anderen frei wird:</p>
${knopf('{{CancelUrl}}', 'Anmeldung stornieren', '#da291c')}
${leise('Versehentlich abgelehnt? Dann ignoriere diese Mail einfach — deine Anmeldung bleibt bestehen.')}
${GRUSS_DE}`
);

// Outlook-Absage in Vertretung (Assistenz lehnt für Partner/Director ab). Der
// Storno-Knopf wirkt nur für die angemeldete Person selbst (Item-Level-Security);
// die Assistenz bekommt einen zweiten Knopf, der die Organizer um die Abmeldung
// bittet. Platzhalter: {{Name}}, {{EventTitle}}, {{CancelUrl}}, {{AssistantForwardUrl}}.
export const OUTLOOK_DECLINE_BODY_ONBEHALF_EN = wrapTemplateForStorage(
  '#ed8b00',
  'Outlook invite declined on your behalf',
  '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>the Outlook invitation for <strong>{{EventTitle}}</strong> was declined on your behalf — but you are <strong>still registered</strong> and keep your spot.</p>
<p>If you can’t attend, please cancel your registration as well. Choose the button that fits:</p>
${UL}
<li><strong>You are the registered person:</strong> cancel directly.</li>
<li><strong>You are the assistant:</strong> forward the request to the organizers — they will cancel on the person’s behalf.</li>
</ul>
${knopf('{{CancelUrl}}', 'Cancel my registration', '#da291c')}
${knopf('{{AssistantForwardUrl}}', 'Forward to the organizers (as assistant)', '#0076a8')}
${leise('The first button only works for the registered person — cancelling someone else’s registration is not possible in the app. Declined by mistake? Then simply ignore this email.')}
${GRUSS_EN}`
);

export const OUTLOOK_DECLINE_BODY_ONBEHALF_DE = wrapTemplateForStorage(
  '#ed8b00',
  'Outlook-Termin in deinem Namen abgelehnt',
  '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>die Outlook-Einladung für <strong>{{EventTitle}}</strong> wurde in deinem Namen abgelehnt — die Anmeldung besteht aber <strong>weiterhin</strong>, der Platz ist also noch belegt.</p>
<p>Wenn die Teilnahme nicht stattfindet, melde die Anmeldung bitte auch hier ab. Nimm den passenden Knopf:</p>
${UL}
<li><strong>Du bist die angemeldete Person:</strong> direkt abmelden.</li>
<li><strong>Du bist die Assistenz:</strong> die Bitte an die Organizer weiterleiten — sie melden die Person dann ab.</li>
</ul>
${knopf('{{CancelUrl}}', 'Anmeldung stornieren', '#da291c')}
${knopf('{{AssistantForwardUrl}}', 'An die Organizer weiterleiten (als Assistenz)', '#0076a8')}
${leise('Der erste Knopf funktioniert nur für die angemeldete Person selbst — eine fremde Anmeldung lässt sich in der App nicht stornieren. Versehentlich abgelehnt? Dann ignoriere diese Mail einfach.')}
${GRUSS_DE}`
);

// Weiterleitung eines Outlook-Termins an eine nicht angemeldete Person (Flow,
// an die Organizer). Platzhalter: {{Forwarder}}, {{Recipient}},
// {{RecipientEmail}}, {{EventTitle}}.
// v30.60: Fakten-Tabelle statt Fließtext — bleibt so, nur die Texte sind neu.
const weiterleitungTabelle = (von: string, an: string, status: string): string =>
  `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;margin:20px 0;font-size:14px;">
<tr><td style="padding:8px 12px 8px 0;color:#63666A;width:32%;vertical-align:top;">${von}</td><td style="padding:8px 0;border-bottom:1px solid #eee;"><strong>{{Forwarder}}</strong></td></tr>
<tr><td style="padding:8px 12px 8px 0;color:#63666A;vertical-align:top;">${an}</td><td style="padding:8px 0;border-bottom:1px solid #eee;"><strong>{{Recipient}}</strong><br><span style="color:#63666A;">{{RecipientEmail}}</span></td></tr>
<tr><td style="padding:8px 12px 8px 0;color:#63666A;vertical-align:top;">Event</td><td style="padding:8px 0;border-bottom:1px solid #eee;">{{EventTitle}}</td></tr>
<tr><td style="padding:8px 12px 8px 0;color:#63666A;vertical-align:top;">Status</td><td style="padding:8px 0;border-bottom:1px solid #eee;"><span style="display:inline-block;padding:2px 10px;border-radius:999px;background:#fdeeee;color:#b3261e;font-weight:700;font-size:13px;">${status}</span></td></tr>
</table>`;

export const OUTLOOK_FORWARD_BODY_EN = wrapTemplateForStorage(
  '#0076a8',
  'Meeting was forwarded',
  '{{EventTitle}}',
  `<p>Hello,</p>
<p>the Outlook invitation for your event was forwarded to someone who is <strong>not registered in DEX</strong>.</p>
${weiterleitungTabelle('Forwarded by', 'To', 'not registered')}
<p>Without a registration, the person has no participant ID, no QR code and does not appear in the participant list. Your options:</p>
${UL}
<li>Ask the person to register via the app.</li>
<li>Register the person yourself — in the app via „Register for another person“.</li>
<li>Remove the person from the Outlook meeting if they should not attend.</li>
</ul>
${knopf(APP, 'Open the DEX app')}
${leise('If the forward is fine, you don’t need to do anything — the person just won’t count as a participant.')}
${GRUSS_DEX_EN}`
);

export const OUTLOOK_FORWARD_BODY_DE = wrapTemplateForStorage(
  '#0076a8',
  'Termin wurde weitergeleitet',
  '{{EventTitle}}',
  `<p>Hallo,</p>
<p>die Outlook-Einladung zu deinem Event wurde an eine Person weitergeleitet, die <strong>in DEX nicht angemeldet</strong> ist.</p>
${weiterleitungTabelle('Weitergeleitet von', 'An', 'nicht angemeldet')}
<p>Ohne Anmeldung hat die Person keine Teilnehmer-ID, keinen QR-Code und steht nicht in der Teilnehmerliste. Deine Möglichkeiten:</p>
${UL}
<li>Die Person bitten, sich selbst über die App anzumelden.</li>
<li>Die Person selbst eintragen — in der App über „Für andere Person registrieren“.</li>
<li>Die Person aus dem Outlook-Termin entfernen, falls sie nicht teilnehmen soll.</li>
</ul>
${knopf(APP, 'DEX-App öffnen')}
${leise('Ist die Weiterleitung in Ordnung, musst du nichts tun — die Person zählt dann nur nicht als Teilnehmer.')}
${GRUSS_DEX_DE}`
);

// Übersicht an die Organizer: angemeldet, aber Outlook abgelehnt (Flow).
// Platzhalter: {{DeclineCount}}, {{EventTitle}}, {{DeclineList}}.
export const OUTLOOK_DECLINE_DIGEST_BODY_EN = wrapTemplateForStorage(
  '#ed8b00',
  'Declined in Outlook, still registered',
  '{{EventTitle}}',
  `<p>Hello,</p>
<p><strong>{{DeclineCount}}</strong> people declined the Outlook invitation for <strong>{{EventTitle}}</strong> but are <strong>still registered</strong>:</p>
{{DeclineList}}
${UL}
<li>They have each received a reminder asking them to cancel their registration as well.</li>
<li>Until they do, they still count towards the capacity.</li>
<li>If someone’s attendance matters, it’s best to contact them directly.</li>
</ul>
${leise('This overview is sent automatically whenever someone declines and always shows the current state.')}
${GRUSS_DEX_EN}`
);

export const OUTLOOK_DECLINE_DIGEST_BODY_DE = wrapTemplateForStorage(
  '#ed8b00',
  'In Outlook abgelehnt, noch angemeldet',
  '{{EventTitle}}',
  `<p>Hallo,</p>
<p><strong>{{DeclineCount}}</strong> Personen haben die Outlook-Einladung für <strong>{{EventTitle}}</strong> abgelehnt, sind aber <strong>noch angemeldet</strong>:</p>
{{DeclineList}}
${UL}
<li>Alle haben bereits eine Erinnerung bekommen, sich auch in DEX abzumelden.</li>
<li>Bis dahin zählen sie weiter zur Kapazität.</li>
<li>Ist die Teilnahme einer Person wichtig, sprich sie am besten direkt an.</li>
</ul>
${leise('Diese Übersicht geht automatisch raus, sobald jemand absagt, und zeigt immer den aktuellen Stand.')}
${GRUSS_DEX_DE}`
);

// Nachrücken (App UND Flow DEX_IDReorder; der Flow ersetzt nur {{Name}} und
// {{EventTitle}}, deshalb feste App-URL statt {{AppUrl}}). Pre-wrapped, weil
// der Flow den BodyHtml roh verwendet; buildEmailFromTemplate erkennt das.
export const NACHRUECKEN_BODY_EN = wrapTemplateForStorage(
  '#86bc25',
  'You’ve got a spot!',
  '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>good news: a spot has become free — you have <strong>moved up from the waitlist</strong> and are now <strong>registered</strong> for <strong>{{EventTitle}}</strong>.</p>
${UL}
<li>Everything that was already sent out for this event (e.g. the invitation) is in the DEX app under <strong>„My Events“</strong>.</li>
<li>Can’t attend after all? Please cancel there as soon as possible so the next person can move up.</li>
</ul>
${knopf(APP, 'Open „My Events“')}
${GRUSS_EN}`
);

export const NACHRUECKEN_BODY_DE = wrapTemplateForStorage(
  '#86bc25',
  'Du hast einen Platz!',
  '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>gute Nachrichten: Ein Platz ist frei geworden — du bist von der <strong>Warteliste nachgerückt</strong> und jetzt für <strong>{{EventTitle}}</strong> <strong>angemeldet</strong>.</p>
${UL}
<li>Alles, was zu diesem Event schon verschickt wurde (z.&nbsp;B. die Einladung), findest du in der DEX-App unter <strong>„Meine Events“</strong>.</li>
<li>Kannst du doch nicht? Dann melde dich dort bitte zeitnah ab, damit die nächste Person nachrücken kann.</li>
</ul>
${knopf(APP, '„Meine Events“ öffnen')}
${GRUSS_DE}`
);

// Hinweis an die Organizer: Abmeldung mit Nachrücker (Flow DEX_IDReorder).
// Platzhalter: {{EventTitle}}, {{CancelledName}}, {{PromotedName}}.
export const ORG_NACHRUECKER_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Cancellation with move-up', '{{EventTitle}}',
  `<p>Hello,</p>
<p>someone cancelled for <strong>{{EventTitle}}</strong>, and the next person on the waitlist has moved up:</p>
${UL}
<li><strong>Cancelled:</strong> {{CancelledName}}</li>
<li><strong>Moved up:</strong> {{PromotedName}}</li>
</ul>
<p>Nothing to do — the participant list and participant IDs are already updated. The current state is in the <a href="${APP}">Organizer Center</a>.</p>
${GRUSS_DEX_EN}`
);
export const ORG_NACHRUECKER_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Abmeldung mit Nachrücker', '{{EventTitle}}',
  `<p>Hallo,</p>
<p>bei <strong>{{EventTitle}}</strong> hat sich jemand abgemeldet, und die nächste Person von der Warteliste ist nachgerückt:</p>
${UL}
<li><strong>Abgemeldet:</strong> {{CancelledName}}</li>
<li><strong>Nachgerückt:</strong> {{PromotedName}}</li>
</ul>
<p>Du musst nichts tun — Teilnehmerliste und Teilnehmer-IDs sind bereits aktualisiert. Den aktuellen Stand siehst du im <a href="${APP}">Organizer Center</a>.</p>
${GRUSS_DEX_DE}`
);

// v22.39: Roter Storno-Banner — Event-Titel ausgegraut und durchgestrichen,
// identisch zum Inline-Fallback `cancellationEmail()` in EmailTemplates.ts.
export const CANCEL_BANNER_HTML = '<div style="margin:16px 0 20px;padding:14px 18px;border:2px solid #da291c;background:rgba(218,41,28,0.06);border-radius:8px;text-align:center;"><div style="font-size:0.78rem;font-weight:700;color:#da291c;text-transform:uppercase;letter-spacing:1.5px;">Stornierung &middot; Cancellation</div><div style="margin-top:6px;font-size:1.15rem;font-weight:700;color:#888;text-decoration:line-through;">{{EventTitle}}</div></div>';

// Auto-Abmeldung nach Outlook-Absage (Flow DEX_OutlookDeclineHandler).
// Platzhalter: {{Name}}, {{EventTitle}}.
export const ABMELDUNG_AUTO_BODY_EN = wrapTemplateForStorage(
  '#da291c', 'Cancellation confirmed', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
${CANCEL_BANNER_HTML}
<p>because you declined the Outlook invitation, your registration has been <strong>cancelled</strong>. Your spot is now free for someone else.</p>
<p>Changed your mind? You can register again any time in the <a href="${APP}">DEX app</a> — as long as spots are available.</p>
${GRUSS_EN}`
);
export const ABMELDUNG_AUTO_BODY_DE = wrapTemplateForStorage(
  '#da291c', 'Abmeldung bestätigt', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
${CANCEL_BANNER_HTML}
<p>weil du die Outlook-Einladung abgelehnt hast, ist deine Anmeldung <strong>storniert</strong>. Dein Platz ist damit für jemand anderen frei.</p>
<p>Doch dabei? Du kannst dich jederzeit wieder in der <a href="${APP}">DEX-App</a> anmelden — solange es freie Plätze gibt.</p>
${GRUSS_DE}`
);

// Team-Mails (App). Platzhalter je Vorlage im Kommentar.

// {{Name}}, {{NewMemberName}}, {{TeamName}} (kann leer sein), {{EventTitle}}, {{AppUrl}}.
export const TEAM_MEMBER_JOINED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'New team member', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{NewMemberName}}</strong> has joined your team {{TeamName}}.</p>
<p>You can see who is in your team any time in the <a href="{{AppUrl}}">DEX app</a> under <strong>„My Events“</strong>.</p>
${GRUSS_EN}`
);
export const TEAM_MEMBER_JOINED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Neues Team-Mitglied', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{NewMemberName}}</strong> ist deinem Team {{TeamName}} beigetreten.</p>
<p>Wer in deinem Team ist, siehst du jederzeit in der <a href="{{AppUrl}}">DEX-App</a> unter <strong>„Meine Events“</strong>.</p>
${GRUSS_DE}`
);

// {{Name}} (Team-Lead), {{RequesterName}}, {{TeamName}}, {{EventTitle}}, {{ApproveUrl}}, {{RejectUrl}}.
const zweiKnoepfe = (ja: string, nein: string): string =>
  `<p style="text-align:center;margin:24px 0;"><a href="{{ApproveUrl}}" style="display:inline-block;padding:12px 24px;background:#86bc25;color:#fff;font-weight:700;text-decoration:none;border-radius:6px;margin-right:8px;">${ja}</a> <a href="{{RejectUrl}}" style="display:inline-block;padding:12px 24px;background:#63666A;color:#fff;font-weight:700;text-decoration:none;border-radius:6px;">${nein}</a></p>`;
export const TEAM_JOIN_REQUEST_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Team join request', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{RequesterName}}</strong> would like to join your team {{TeamName}}. As team lead, you decide:</p>
${zweiKnoepfe('Approve', 'Decline')}
${leise('The buttons open the DEX app. You can also find the request later under „My Events“.')}
${GRUSS_EN}`
);
export const TEAM_JOIN_REQUEST_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Team-Beitritts-Anfrage', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{RequesterName}}</strong> möchte deinem Team {{TeamName}} beitreten. Als Team-Lead entscheidest du:</p>
${zweiKnoepfe('Bestätigen', 'Ablehnen')}
${leise('Die Knöpfe öffnen die DEX-App. Die Anfrage findest du auch später noch unter „Meine Events“.')}
${GRUSS_DE}`
);

// {{Name}}, {{EventTitle}}.
export const TEAM_JOIN_REJECTED_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Team join request declined', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>the team lead has <strong>declined</strong> your request to join the team for <strong>{{EventTitle}}</strong>.</p>
<p>You still have two options on the registration page:</p>
${UL}
<li>register on your own, as long as spots are available, or</li>
<li>join another team that still has open slots.</li>
</ul>
${GRUSS_EN}`
);
export const TEAM_JOIN_REJECTED_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Team-Beitritts-Anfrage abgelehnt', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>der Team-Lead hat deine Anfrage, dem Team bei <strong>{{EventTitle}}</strong> beizutreten, <strong>abgelehnt</strong>.</p>
<p>Auf der Anmeldeseite hast du weiter zwei Möglichkeiten:</p>
${UL}
<li>dich allein anmelden, solange es freie Plätze gibt, oder</li>
<li>einem anderen Team mit freien Plätzen beitreten.</li>
</ul>
${GRUSS_DE}`
);

// {{Name}}, {{TeamName}}, {{NewLeadName}}, {{NewLeadBlock}} (nur beim neuen Lead gefüllt), {{EventTitle}}.
export const TEAM_LEAD_TRANSFERRED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'New team lead', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>your team {{TeamName}} has a new team lead: <strong>{{NewLeadName}}</strong>.</p>
{{NewLeadBlock}}
${GRUSS_EN}`
);
export const TEAM_LEAD_TRANSFERRED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Neuer Team-Lead', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>dein Team {{TeamName}} hat einen neuen Team-Lead: <strong>{{NewLeadName}}</strong>.</p>
{{NewLeadBlock}}
${GRUSS_DE}`
);

// {{Name}}, {{CancelledName}}, {{TeamName}}, {{EventTitle}}, {{ActiveCount}}, {{TeamSize}}, {{NewLeadBlock}}.
export const TEAM_MEMBER_CANCELLED_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Team member cancelled', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{CancelledName}}</strong> has cancelled and is no longer part of your team {{TeamName}}.</p>
${UL}
<li><strong>Team now:</strong> {{ActiveCount}} of {{TeamSize}} spots filled</li>
</ul>
{{NewLeadBlock}}
<p>What happens now:</p>
${UL}
<li>Nothing, if that’s fine — the free spot stays reserved for your team.</li>
<li>The team lead can add someone else under <strong>„My Events“</strong>.</li>
<li>If the organizers show open team spots, others can also join via the registration page.</li>
</ul>
${GRUSS_EN}`
);
export const TEAM_MEMBER_CANCELLED_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Team-Mitglied abgemeldet', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{CancelledName}}</strong> hat sich abgemeldet und ist nicht mehr in deinem Team {{TeamName}}.</p>
${UL}
<li><strong>Team jetzt:</strong> {{ActiveCount}} von {{TeamSize}} Plätzen belegt</li>
</ul>
{{NewLeadBlock}}
<p>Wie es weitergeht:</p>
${UL}
<li>Nichts, wenn es so passt — der freie Platz bleibt für euer Team reserviert.</li>
<li>Der Team-Lead kann unter <strong>„Meine Events“</strong> jemand anderen hinzufügen.</li>
<li>Zeigen die Organizer freie Team-Plätze an, können auch andere über die Anmeldeseite beitreten.</li>
</ul>
${GRUSS_DE}`
);

// {{Name}}, {{RegistrantName}}, {{EventTitle}}.
export const ROOMMATE_REQUEST_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Roommate request', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{RegistrantName}}</strong> would like to share a room with you at <strong>{{EventTitle}}</strong>.</p>
<p>To confirm: choose <strong>{{RegistrantName}}</strong> as your roommate when you register (or later via „Edit details“ under „My Events“). Once you have chosen each other, the organizers can see the match.</p>
${GRUSS_EN}`
);
export const ROOMMATE_REQUEST_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Zimmerpartner-Anfrage', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{RegistrantName}}</strong> möchte sich bei <strong>{{EventTitle}}</strong> ein Zimmer mit dir teilen.</p>
<p>So bestätigst du: Wähle bei deiner Anmeldung <strong>{{RegistrantName}}</strong> als Zimmerpartner aus (oder später über „Angaben bearbeiten“ unter „Meine Events“). Habt ihr euch gegenseitig gewählt, sehen die Organizer das Paar.</p>
${GRUSS_DE}`
);

// {{Name}}, {{GroupLabel}}, {{EventTitle}} (DE zusätzlich {{AppUrl}}).
export const GROUP_SWITCH_CONFIRMED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Group switch confirmed', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>your switch is confirmed: you are now registered in the group <strong>{{GroupLabel}}</strong>.</p>
<p>You can check your registration any time in the <a href="${APP}">DEX app</a> under <strong>„My Events“</strong>.</p>
${GRUSS_EN}`
);
export const GROUP_SWITCH_CONFIRMED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Gruppenwechsel bestätigt', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>dein Wechsel ist bestätigt: Du bist jetzt in der Gruppe <strong>{{GroupLabel}}</strong> angemeldet.</p>
<p>Deine Anmeldung siehst du jederzeit in der <a href="{{AppUrl}}">DEX-App</a> unter <strong>„Meine Events“</strong>.</p>
${GRUSS_DE}`
);

// {{Name}}, {{GroupLabel}}, {{EventTitle}}.
export const GROUP_SWITCH_WAITLIST_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Group switch — on the waitlist', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>the group <strong>{{GroupLabel}}</strong> is currently full — so you are now on its <strong>waitlist</strong>.</p>
<p>As soon as a spot frees up, you move up automatically and get a confirmation. Nothing else to do.</p>
${GRUSS_EN}`
);
export const GROUP_SWITCH_WAITLIST_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Gruppenwechsel — auf der Warteliste', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>die Gruppe <strong>{{GroupLabel}}</strong> ist gerade voll — du stehst deshalb auf ihrer <strong>Warteliste</strong>.</p>
<p>Sobald ein Platz frei wird, rückst du automatisch nach und bekommst eine Bestätigung. Du musst nichts weiter tun.</p>
${GRUSS_DE}`
);

// {{Name}}, {{EventTitle}}, {{WaitlistPositionBlock}} (optionaler HTML-Block).
export const OVERBOOK_APOLOGY_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Registration corrected', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>we have to apologise: because many people registered at the same moment, you were confirmed for <strong>{{EventTitle}}</strong> although all spots were already taken.</p>
<p>We have therefore moved your registration to the <strong>waitlist</strong>. This was a technical issue — not your fault.</p>
{{WaitlistPositionBlock}}
<p>As soon as a spot frees up, you move up automatically and get a confirmation right away. Nothing else to do.</p>
<p style="margin-top:24px;"><strong>Thank you for your understanding</strong><br><br><strong>Your Event Team</strong></p>`
);
export const OVERBOOK_APOLOGY_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Anmeldung korrigiert', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>wir müssen uns entschuldigen: Weil sich sehr viele Personen gleichzeitig angemeldet haben, wurdest du für <strong>{{EventTitle}}</strong> bestätigt, obwohl alle Plätze schon vergeben waren.</p>
<p>Deine Anmeldung steht deshalb jetzt auf der <strong>Warteliste</strong>. Das war ein technischer Fehler — nicht deiner.</p>
{{WaitlistPositionBlock}}
<p>Sobald ein Platz frei wird, rückst du automatisch nach und bekommst sofort eine Bestätigung. Du musst nichts weiter tun.</p>
<p style="margin-top:24px;"><strong>Danke für dein Verständnis</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// v32.52: Die Vorlagen, die die APP selbst in den Deloitte-Rahmen legt
// (buildEmailFromTemplate + wrapTemplate) — nur der Innenteil.
// {{Eckdaten}} ist ein HTML-Block (Wann/Wo), leer, wenn der Aufrufer ihn nicht
// liefert; {{OrganizerHtml}} setzt die Organizer-Namen fett ein.
export const ANMELDUNG_BODY_DE = `<p>Hallo {{Name}},</p>
<p>du bist für <strong>{{EventTitle}}</strong> <strong>angemeldet</strong>.</p>
{{Eckdaten}}
<p>Kannst du doch nicht teilnehmen? Dann melde dich bitte rechtzeitig in der <a href="{{AppUrl}}">DEX-App</a> unter „Meine Events“ ab — so wird dein Platz für jemand anderen frei.</p>
<p>Fragen zur Organisation beantworten dir {{OrganizerHtml}}.</p>
${GRUSS_DE}`;
export const ANMELDUNG_BODY_EN = `<p>Hello {{Name}},</p>
<p>you are <strong>registered</strong> for <strong>{{EventTitle}}</strong>.</p>
{{Eckdaten}}
<p>Can’t attend after all? Please cancel in good time in the <a href="{{AppUrl}}">DEX app</a> under „My Events“ — so your spot becomes free for someone else.</p>
<p>For organisational questions, please contact {{OrganizerHtml}}.</p>
${GRUSS_EN}`;

export const WARTELISTE_BODY_DE = `<p>Hallo {{Name}},</p>
<p>alle Plätze für <strong>{{EventTitle}}</strong> sind gerade vergeben — du stehst auf der <strong>Warteliste</strong>.</p>
<p>Dein aktueller Platz: <strong>{{WaitlistPosition}}</strong></p>
{{Eckdaten}}
<p>Wird ein Platz frei, rückst du automatisch nach und bekommst eine Mail. Deinen aktuellen Platz siehst du jederzeit in der <a href="{{AppUrl}}">DEX-App</a> unter „Meine Events“.</p>
${GRUSS_DE}`;
export const WARTELISTE_BODY_EN = `<p>Hello {{Name}},</p>
<p>all spots for <strong>{{EventTitle}}</strong> are currently taken — you are on the <strong>waitlist</strong>.</p>
<p>Your current position: <strong>{{WaitlistPosition}}</strong></p>
{{Eckdaten}}
<p>If a spot frees up, you move up automatically and get an email. You can see your current position any time in the <a href="{{AppUrl}}">DEX app</a> under „My Events“.</p>
${GRUSS_EN}`;

// v31.5: Geht an JEDE abgemeldete Zeile, auch an Wartelistler ohne
// Outlook-Termin — daher „falls du einen hattest".
export const ABMELDUNG_BODY_DE = `<p>Hallo {{Name}},</p>
${CANCEL_BANNER_HTML}
<p>deine Anmeldung ist <strong>storniert</strong>. Hattest du einen Outlook-Termin dafür, verschwindet er in Kürze aus deinem Kalender.</p>
<p>Doch dabei? Du kannst dich jederzeit wieder in der <a href="{{AppUrl}}">DEX-App</a> anmelden — solange es freie Plätze gibt.</p>
${GRUSS_DE}`;
export const ABMELDUNG_BODY_EN = `<p>Hello {{Name}},</p>
${CANCEL_BANNER_HTML}
<p>your registration has been <strong>cancelled</strong>. If you had an Outlook invitation for it, it will disappear from your calendar shortly.</p>
<p>Changed your mind? You can register again any time in the <a href="{{AppUrl}}">DEX app</a> — as long as spots are available.</p>
${GRUSS_EN}`;

export const EVENT_ERSTELLT_BODY_DE = `<p>Hallo {{Name}},</p>
<p>dein Event <strong>{{EventTitle}}</strong> ist angelegt.</p>
<p>Anmeldungen, Mails und Check-in verwaltest du im <a href="{{AppUrl}}">Organizer Center der DEX-App</a>.</p>
${GRUSS_DEX_DE}`;
export const EVENT_ERSTELLT_BODY_EN = `<p>Hello {{Name}},</p>
<p>your event <strong>{{EventTitle}}</strong> has been created.</p>
<p>You manage registrations, emails and check-in in the <a href="{{AppUrl}}">Organizer Center of the DEX app</a>.</p>
${GRUSS_DEX_EN}`;

// Fester Listenname auf jeder Subsite
