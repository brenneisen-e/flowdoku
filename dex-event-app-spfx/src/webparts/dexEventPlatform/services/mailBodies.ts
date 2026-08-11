/**
 * v28.95: Aus `EventService` herausgeloest (dort 412 der 13.221 Zeilen).
 *
 * Die fertigen HTML-Koerper der System-Mails, die die App selbst in die
 * Warteschlange stellt (Outlook-Absage, Nachruecken, Team-Ereignisse,
 * Zimmerwunsch, Gruppenwechsel, Ueberbuchungs-Entschuldigung). Reine Daten:
 * `wrapTemplateForStorage` legt den Deloitte-Rahmen drumherum, der Flow
 * ersetzt spaeter nur noch {{LOGO_URL}} und {{ORB_URL}}.
 *
 * Sie standen bisher vor der Klasse und machten den Einstieg in die Datei zu
 * 400 Zeilen Mailtext, bevor die erste Methode kam.
 */
import { wrapTemplateForStorage } from './EmailTemplates';


/**
 * HTML-Body für die OutlookDeclineReminder-Mail (EN) - komplett im
 * Deloitte-Design gewrappt, damit er vom DEX_SEND_MAIL-Flow direkt versendet
 * werden kann (der Flow ersetzt nur {{LOGO_URL}} und {{ORB_URL}}, wickelt aber
 * keinen Template-Wrapper mehr drumherum).
 *
 * Der Hinweis auf die Warteliste wurde bewusst durch eine neutrale Formulierung
 * ("your spot can be offered to someone else") ersetzt, damit die Mail auch
 * dann korrekt wirkt, wenn das konkrete Event keine Warteliste hat.
 */
// Standard-Decline-Reminder: Person hat selbst direkt abgelehnt (kein OnBehalfOf).
// Cancel-Button funktioniert direkt — die eingeloggte Person ist auch der Teilnehmer.
export const OUTLOOK_DECLINE_BODY_EN = wrapTemplateForStorage(
  '#ed8b00',
  'You declined the Outlook invite',
  'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
<p>we noticed that you declined the Outlook calendar invitation for <strong>{{EventTitle}}</strong>, but you are still listed as a confirmed participant.</p>
<p>If you no longer want to attend, please also cancel your registration.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Cancel my registration</a></p>
<p>If you clicked decline by accident, you can simply ignore this message.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

export const OUTLOOK_DECLINE_BODY_DE = wrapTemplateForStorage(
  '#ed8b00',
  'Du hast den Outlook-Termin abgelehnt',
  'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>wir haben gesehen, dass du die Outlook-Kalendereinladung f\u00FCr <strong>{{EventTitle}}</strong> abgelehnt hast \u2013 du bist aber noch als offiziell angemeldet gelistet.</p>
<p>Falls du nicht mehr teilnehmen m\u00F6chtest, melde dich bitte auch offiziell ab.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Anmeldung stornieren</a></p>
<p>Falls du versehentlich abgesagt hast, kannst du diese Mail einfach ignorieren.</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// OnBehalfOf-Decline-Reminder: Outlook-Decline kam von einer Assistenz, die den
// Termin im Namen einer anderen Person (Partner/Director) abgelehnt hat. Die Mail
// geht an die Mailbox der Assistenz/des Partners. Der Cancel-Button funktioniert
// nur für den registrierten Partner selbst (SP-Item-Level-Security). Für die
// Assistenz gibt's einen zweiten Button, der via mailto: eine Bitte um Abmeldung
// an die Event-Organizer schickt — die haben Full Control auf der Teilnehmerliste
// und können den Eintrag direkt im Admin Center löschen.
export const OUTLOOK_DECLINE_BODY_ONBEHALF_EN = wrapTemplateForStorage(
  '#ed8b00',
  'You declined the Outlook invite',
  'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
<p>we noticed that the Outlook calendar invitation for <strong>{{EventTitle}}</strong> was declined on your behalf, but you are still listed as a confirmed participant.</p>
<p>If you no longer want to attend, please also cancel your registration.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Cancel my registration</a></p>
<p style="font-size:12px;color:#666;text-align:center;margin:0 0 24px;">This button only works when clicked by the registered participant themselves (Partner/Director).</p>
<p style="margin:8px 0 12px;"><strong>Are you the assistant handling this for the participant?</strong> Use the button below to forward this request to the event organizer(s) — they will cancel the registration on the participant's behalf:</p>
<p style="margin:0 0 24px;text-align:center;"><a href="{{AssistantForwardUrl}}" style="display:inline-block;padding:10px 24px;background:#0d6efd;color:#fff;text-decoration:none;border-radius:6px;font-weight:600;font-size:0.95rem;">Forward to organizer (as assistant)</a></p>
<p>If the decline was sent by accident, you can simply ignore this message.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

export const OUTLOOK_DECLINE_BODY_ONBEHALF_DE = wrapTemplateForStorage(
  '#ed8b00',
  'Outlook-Termin abgelehnt',
  'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>die Outlook-Kalendereinladung f\u00FCr <strong>{{EventTitle}}</strong> wurde in deinem Namen abgelehnt — du bist aber noch als offiziell angemeldet gelistet.</p>
<p>Falls du nicht mehr teilnehmen m\u00F6chtest, melde dich bitte auch offiziell ab.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Anmeldung stornieren</a></p>
<p style="font-size:12px;color:#666;text-align:center;margin:0 0 24px;">Dieser Button funktioniert nur, wenn der/die registrierte Teilnehmer/in selbst (Partner/Director) klickt.</p>
<p style="margin:8px 0 12px;"><strong>Bist du die Assistenz der angemeldeten Person?</strong> Bitte nutze den folgenden Button, um diese Anfrage an die Event-Organisator:innen weiterzuleiten — sie melden den/die Teilnehmer/in dann ab:</p>
<p style="margin:0 0 24px;text-align:center;"><a href="{{AssistantForwardUrl}}" style="display:inline-block;padding:10px 24px;background:#0d6efd;color:#fff;text-decoration:none;border-radius:6px;font-weight:600;font-size:0.95rem;">An Organizer weiterleiten (als Assistenz)</a></p>
<p>Falls die Absage versehentlich verschickt wurde, kannst du diese Mail einfach ignorieren.</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// Meeting-Forward-Notification-FYI: Ein Teilnehmer hat den Outlook-Termin an eine
// dritte Person weitergeleitet, die NICHT in der SharePoint-Teilnehmerliste steht.
// Die Mail geht an den Organizer. Platzhalter:
//   {{OrganizerFirstName}} — Vorname des Organizers (Anrede)
//   {{Forwarder}}          — Person, die den Termin weitergeleitet hat
//   {{Recipient}}          — Name der hinzugefügten Person
//   {{RecipientEmail}}     — Email der hinzugefügten Person ('nicht aufgelöst' bei Externen)
//   {{EventTitle}}         — Event-Titel
//   {{AppUrl}}             — DEX-App-URL (Organizer kann dort manuell registrieren)
export const OUTLOOK_FORWARD_BODY_EN = wrapTemplateForStorage(
  '#0d6efd',
  'Meeting was forwarded',
  'Event {{EventTitle}}',
  `<p>Hi,</p>
<p>FYI: <strong>{{Forwarder}}</strong> forwarded the Outlook invitation for <strong>{{EventTitle}}</strong> to <strong>{{Recipient}}</strong> ({{RecipientEmail}}).</p>
<p><strong>{{Recipient}} is currently NOT registered in the DEX participant list.</strong> This person still needs to register via the app in order to get a ParticipantID and QR code and to appear in the official participant list.</p>
<p style="margin:24px 0;text-align:center;"><a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="display:inline-block;padding:12px 28px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Open DEX App</a></p>
<p style="font-size:13px;color:#555;margin-top:24px;">Possible next steps:</p>
<ul style="font-size:13px;color:#555;margin:0 0 24px 16px;padding:0;">
<li>Ask {{Recipient}} to register themselves via the app.</li>
<li>Or: register {{Recipient}} manually as organizer via "Register for another person".</li>
<li>Or: remove {{Recipient}} from the Outlook meeting if they should not attend.</li>
</ul>
<p style="font-size:12px;color:#999;">This message was generated automatically (Microsoft Outlook Meeting Forward Notification).</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

export const OUTLOOK_FORWARD_BODY_DE = wrapTemplateForStorage(
  '#0d6efd',
  'Termin wurde weitergeleitet',
  'Event {{EventTitle}}',
  `<p>Hallo,</p>
<p>zur Info: <strong>{{Forwarder}}</strong> hat die Outlook-Einladung f\u00FCr <strong>{{EventTitle}}</strong> an <strong>{{Recipient}}</strong> ({{RecipientEmail}}) weitergeleitet.</p>
<p><strong>{{Recipient}} ist aktuell NICHT in der DEX-Teilnehmerliste registriert.</strong> Die Person muss sich ggf. noch selbst \u00FCber die App anmelden, damit sie eine TeilnehmerID und einen QR-Code bekommt und in der offiziellen Teilnehmerliste erscheint.</p>
<p style="margin:24px 0;text-align:center;"><a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="display:inline-block;padding:12px 28px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">DEX-App \u00F6ffnen</a></p>
<p style="font-size:13px;color:#555;margin-top:24px;">M\u00F6gliche Handlungsoptionen:</p>
<ul style="font-size:13px;color:#555;margin:0 0 24px 16px;padding:0;">
<li>{{Recipient}} bitten, sich selbst \u00FCber die App zu registrieren.</li>
<li>Oder: {{Recipient}} als Organizer manuell \u00FCber "F\u00FCr andere Person registrieren" eintragen.</li>
<li>Oder: {{Recipient}} aus dem Outlook-Termin entfernen, falls nicht gew\u00FCnscht.</li>
</ul>
<p style="font-size:12px;color:#999;">Diese Mail wurde automatisch erzeugt (Microsoft Outlook Meeting Forward Notification).</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// v9.38: OutlookDeclineDigest — Info-Mail an Organizer, sobald jemand den
// Outlook-Termin abgelehnt hat. Listet alle Teilnehmer, die noch angemeldet
// sind, aber den Outlook-Termin abgelehnt haben. Wird nach jedem neuen
// Decline gequeued (Power-Automate-Flow DEX_OutlookDeclineHandler).
//   {{EventTitle}}    — Event-Titel
//   {{DeclineCount}}  — Anzahl der noch-angemeldeten Decliner
//   {{DeclineList}}   — HTML-Tabelle mit Vorname/Nachname/Mail/RegDate/Department
export const OUTLOOK_DECLINE_DIGEST_BODY_EN = wrapTemplateForStorage(
  '#ed8b00',
  'FYI: attendees declined the Outlook invite',
  'Event {{EventTitle}}',
  `<p>Hi,</p>
<p>The following <strong>{{DeclineCount}}</strong> attendees are still registered for <strong>{{EventTitle}}</strong> but have declined the Outlook calendar invite. They received a reminder mail asking them to also cancel their registration if they cannot attend — until then, they still count towards the capacity.</p>
{{DeclineList}}
<p>You may want to reach out to them directly if their attendance is critical for the event.</p>
<p style="font-size:12px;color:#999;margin-top:24px;">This summary is sent automatically every time someone declines the Outlook invite. The list always reflects the current state of registered-but-declined attendees.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

export const OUTLOOK_DECLINE_DIGEST_BODY_DE = wrapTemplateForStorage(
  '#ed8b00',
  'FYI: Teilnehmer haben den Outlook-Termin abgelehnt',
  'Event {{EventTitle}}',
  `<p>Hallo,</p>
<p>folgende <strong>{{DeclineCount}}</strong> Teilnehmer sind noch für <strong>{{EventTitle}}</strong> angemeldet, haben aber den Outlook-Termin <strong>abgelehnt</strong>. Sie haben bereits eine Erinnerungs-Mail bekommen mit der Bitte, sich auch offiziell abzumelden — bis dahin zählen sie aber zur Kapazität.</p>
{{DeclineList}}
<p>Falls die Teilnahme dieser Personen für das Event wichtig ist, sprich sie ggf. direkt an.</p>
<p style="font-size:12px;color:#999;margin-top:24px;">Diese Übersicht geht automatisch raus, sobald jemand den Outlook-Termin ablehnt. Die Liste zeigt immer den aktuellen Stand der noch-angemeldeten Decliner.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// Nachrücken-Mail (PA-Flow DEX_IDReorder queued sie) — muss pre-wrapped sein,
// weil die Flow-seite den BodyHtml raw verwendet (ohne wrapTemplate). Client-Code
// erkennt die Pre-Wrap in buildEmailFromTemplate() und skippt den Wrap dann.
// v12.11/v12.12: Nachr\u00FCcken-Mail-Text pr\u00E4zisiert \u2014 der alte \u201ESpot available"-
// Subject war missverst\u00E4ndlich (klang wie ein Angebot). Outlook-Verweis
// entfernt, weil nicht jedes Event Outlook-Termine versendet.
export const NACHRUECKEN_BODY_EN = wrapTemplateForStorage(
  '#86bc25',
  'You\u2019ve got a spot!',
  'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
<p>Great news! A spot has become available \u2014 you have <strong>moved up from the waitlist</strong> for the event <strong>{{EventTitle}}</strong> and are now a <strong>confirmed participant</strong>.</p>
<p>You are now on the official participant list.</p>
<p style="padding:10px 14px;background:#eef4fb;border:1px solid #0076a8;border-radius:8px;"><strong>Catch up on earlier updates:</strong> Any information already sent out for this event (e.g. an invitation or announcement) is available in the DEX App under <strong>\u201CMy Events\u201D</strong> on this event \u2014 so you\u2019re fully up to date.</p>
<p>You can review your participation any time in the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a> under <strong>\u201CMy Events\u201D</strong>.</p>
<p>If you are unable to attend after all, please cancel your registration as soon as possible via the App so that the next person on the waitlist can move up.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

export const NACHRUECKEN_BODY_DE = wrapTemplateForStorage(
  '#86bc25',
  'Du hast einen Platz!',
  'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>gute Nachrichten! Ein Platz ist frei geworden \u2014 du bist von der <strong>Warteliste nachger\u00FCckt</strong> f\u00FCr das Event <strong>{{EventTitle}}</strong> und bist jetzt <strong>fester Teilnehmer</strong>.</p>
<p>Du stehst nun auf der offiziellen Teilnehmerliste.</p>
<p style="padding:10px 14px;background:#eef4fb;border:1px solid #0076a8;border-radius:8px;"><strong>Bisherige Infos nachlesen:</strong> Alle bereits zu diesem Event versendeten Informationen (z.\u00A0B. Einladung oder Ank\u00FCndigung) findest du in der DEX App unter <strong>\u201EMeine Events\u201C</strong> beim Event \u2014 so bist du auf dem gleichen Stand.</p>
<p>Deine Teilnahme kannst du jederzeit in der <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a> unter <strong>\u201EMeine Events\u201C</strong> einsehen.</p>
<p>Falls du doch nicht teilnehmen kannst, melde dich bitte zeitnah \u00FCber die App ab, damit die n\u00E4chste Person von der Warteliste nachr\u00FCcken kann.</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// v18.63: Organizer-Benachrichtigung bei Abmeldung MIT Nachrücker. Wird vom
// DEX_IDReorder-Flow nach einem erfolgreichen Promote an die Organizer
// gequeued (nicht von der App). Pre-wrapped gespeichert wie Nachrücken; der
// Flow ersetzt nur {{EventTitle}} und {{PromotedName}} per replace(). Daher
// KEIN {{AppUrl}} (würde der Flow nicht auflösen) — feste App-URL eingebaut.
// Platzhalter: {{EventTitle}}, {{CancelledName}} (abgemeldete Person),
// {{PromotedName}} (voller Name des Nachrückers).
export const ORG_NACHRUECKER_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Cancellation — waitlist move-up', 'Event {{EventTitle}}',
  `<p>Hello,</p>
<p>There was a change for the event <strong>{{EventTitle}}</strong>:</p>
<ul style="margin:12px 0 16px; padding-left:20px; line-height:1.7;">
<li><strong>Cancellation:</strong> {{CancelledName}}</li>
<li><strong>Moved up from the waitlist:</strong> {{PromotedName}}</li>
</ul>
<p>You don't need to do anything — the participant list and participant IDs have already been updated automatically. You can review the current status in the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX Admin Center</a>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your DEX Team</strong></p>`
);
export const ORG_NACHRUECKER_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Abmeldung — Nachrücker', 'Event {{EventTitle}}',
  `<p>Hallo,</p>
<p>beim Event <strong>{{EventTitle}}</strong> gab es eine Änderung:</p>
<ul style="margin:12px 0 16px; padding-left:20px; line-height:1.7;">
<li><strong>Abmeldung:</strong> {{CancelledName}}</li>
<li><strong>Nachrücker:</strong> {{PromotedName}}</li>
</ul>
<p>Du musst nichts weiter tun — die Teilnehmerliste und die TeilnehmerIDs wurden bereits automatisch aktualisiert. Den aktuellen Stand siehst du im <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX Admin Center</a>.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein DEX-Team</strong></p>`
);

// v19.25: Pre-wrapped Abmelde-Bestätigung für die FLOW-getriebene Auto-Abmeldung
// bei Outlook-Absage (DEX_OutlookDeclineHandler queued sie). Muss pre-wrapped
// sein, weil der Flow den BodyHtml roh verwendet (kein wrapTemplate). Eigener
// TemplateType `AbmeldungAuto`, damit die App-eigene `Abmeldung` (unwrapped, mit
// Per-Event-Customizing) unangetastet bleibt. Flow ersetzt nur {{Name}} +
// {{EventTitle}}; KEIN {{AppUrl}} (würde der Flow nicht auflösen) — feste
// App-URL eingebaut. {{LOGO_URL}}/{{ORB_URL}} ersetzt DEX_SEND_MAIL beim Versand.
// v22.39: Roter Storno-Banner für die Standard-Abmelde-Templates — Event-Titel
// ausgegraut + durchgestrichen, identisch zum Inline-Fallback
// `cancellationEmail()` in EmailTemplates.ts (v17.20). Per Reseed kommt das
// Design in die Tenant-Templates; Event-Overrides bleiben unberührt.
export const CANCEL_BANNER_HTML = '<div style="margin:16px 0 20px;padding:14px 18px;border:2px solid #da291c;background:rgba(218,41,28,0.06);border-radius:8px;text-align:center;"><div style="font-size:0.78rem;font-weight:700;color:#da291c;text-transform:uppercase;letter-spacing:1.5px;">Stornierung &middot; Cancellation</div><div style="margin-top:6px;font-size:1.15rem;font-weight:700;color:#888;text-decoration:line-through;">{{EventTitle}}</div></div>';

export const ABMELDUNG_AUTO_BODY_EN = wrapTemplateForStorage(
  '#da291c', 'Cancellation confirmed', 'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
${CANCEL_BANNER_HTML}
<p>your registration for the event above has been <strong>cancelled</strong> because you declined the Outlook invitation.</p>
<p>If you change your mind, you can register again any time via the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const ABMELDUNG_AUTO_BODY_DE = wrapTemplateForStorage(
  '#da291c', 'Abmeldung bestätigt', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
${CANCEL_BANNER_HTML}
<p>deine Anmeldung für das oben genannte Event wurde <strong>storniert</strong>, weil du den Outlook-Termin abgelehnt hast.</p>
<p>Du kannst dich jederzeit erneut über die <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a> anmelden.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// v12.13: Team-bezogene Mail-Vorlagen — vorher inline in EventContext.tsx
// als ad-hoc-HTML zusammengebaut, jetzt zentral in DEX_EmailTemplates
// hinterlegt damit Admins sie genauso wie Anmeldung/Abmeldung/Nachrücken
// anpassen können. Platzhalter pro Template siehe Inline-Kommentare.

// {{Name}} (Empfänger-Vorname), {{NewMemberName}} (voller Name des
// neuen Mitglieds), {{TeamName}} (kann leer sein), {{EventTitle}}.
export const TEAM_MEMBER_JOINED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Team update', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{NewMemberName}}</strong> joined your team {{TeamName}} for the event <strong>{{EventTitle}}</strong>.</p>
<p>You can see the current team status in the <a href="{{AppUrl}}">DEX App</a> under <strong>“My Events”</strong>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const TEAM_MEMBER_JOINED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Team-Update', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{NewMemberName}}</strong> ist deinem Team {{TeamName}} beim Event <strong>{{EventTitle}}</strong> beigetreten.</p>
<p>Den aktuellen Team-Stand siehst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter <strong>„Meine Events"</strong>.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Lead-Vorname), {{RequesterName}} (voll), {{TeamName}},
// {{EventTitle}}, {{ApproveUrl}}, {{RejectUrl}}.
export const TEAM_JOIN_REQUEST_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Team join request', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{RequesterName}}</strong> would like to join your team {{TeamName}} for the event <strong>{{EventTitle}}</strong>. Please decide:</p>
<p style="text-align:center;margin:18px 0;"><a href="{{ApproveUrl}}" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Approve</a> <a href="{{RejectUrl}}" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Reject</a></p>
<p style="font-size:0.85rem;color:#666;">Note: the buttons lead you to the app; the request block lives under <strong>“My Events”</strong>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const TEAM_JOIN_REQUEST_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Team-Beitritts-Anfrage', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{RequesterName}}</strong> möchte deinem Team {{TeamName}} beim Event <strong>{{EventTitle}}</strong> beitreten. Bitte entscheide:</p>
<p style="text-align:center;margin:18px 0;"><a href="{{ApproveUrl}}" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Bestätigen</a> <a href="{{RejectUrl}}" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Ablehnen</a></p>
<p style="font-size:0.85rem;color:#666;">Hinweis: die Buttons führen dich in die App; den Beitritts-Anfragen-Block findest du unter <strong>„Meine Events"</strong>.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Anfrager-Vorname), {{EventTitle}}.
export const TEAM_JOIN_REJECTED_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Team join request declined', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>your join request for the team at the event <strong>{{EventTitle}}</strong> was declined by the team lead.</p>
<p>You can still register individually if capacity allows — or join another open team via the registration page.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const TEAM_JOIN_REJECTED_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Team-Beitritts-Anfrage abgelehnt', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>deine Beitritts-Anfrage zum Team beim Event <strong>{{EventTitle}}</strong> wurde vom Team-Lead abgelehnt.</p>
<p>Du kannst dich gerne einzeln anmelden, falls die Kapazität noch reicht — oder einem anderen offenen Team über die Anmeldeseite beitreten.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{NewLeadName}}, {{TeamName}},
// {{EventTitle}}, {{NewLeadBlock}} (HTML-Block — leer falls Empfänger
// nicht der neue Lead ist, sonst der zusätzliche Hinweis-Absatz).
export const TEAM_LEAD_TRANSFERRED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Team lead change', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>The team lead role in your team {{TeamName}} has been transferred to <strong>{{NewLeadName}}</strong>.</p>
{{NewLeadBlock}}
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const TEAM_LEAD_TRANSFERRED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Team-Lead-Wechsel', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>Die Team-Lead-Rolle in deinem Team {{TeamName}} wurde an <strong>{{NewLeadName}}</strong> übergeben.</p>
{{NewLeadBlock}}
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{CancelledName}}, {{TeamName}},
// {{EventTitle}}, {{ActiveCount}}, {{TeamSize}}, {{NewLeadBlock}}
// (leer, falls Empfänger nicht zum neuen Lead ernannt wurde).
export const TEAM_MEMBER_CANCELLED_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Team update', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>a member of your team {{TeamName}} has cancelled their registration for the event <strong>{{EventTitle}}</strong>:</p>
<p style="padding:8px 12px;background:#f7f7f7;border-left:3px solid #86bc25;font-weight:600;">{{CancelledName}}</p>
<p>Current team occupancy: <strong>{{ActiveCount}}/{{TeamSize}}</strong></p>
{{NewLeadBlock}}
<p>What you can do now:</p>
<ul>
<li>Do nothing — your seat stays reserved for the team for now.</li>
<li>As team lead: add a replacement person via <strong>“My Events”</strong>.</li>
<li>Other participants can join the open slot via the registration page (if the organizer enabled “Public open slots”).</li>
</ul>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const TEAM_MEMBER_CANCELLED_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Team-Update', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>ein Mitglied deines Teams {{TeamName}} hat sich vom Event <strong>{{EventTitle}}</strong> abgemeldet:</p>
<p style="padding:8px 12px;background:#f7f7f7;border-left:3px solid #86bc25;font-weight:600;">{{CancelledName}}</p>
<p>Aktuelle Team-Belegung: <strong>{{ActiveCount}}/{{TeamSize}}</strong></p>
{{NewLeadBlock}}
<p>Was du jetzt machen kannst:</p>
<ul>
<li>Nichts tun — euer Platz bleibt erstmal für das Team reserviert.</li>
<li>Als Team-Lead: über <strong>„Meine Events"</strong> eine andere Person nachträglich hinzufügen.</li>
<li>Andere Teilnehmer können ggf. den freien Slot über die Event-Anmeldeseite belegen (sofern der Organizer „Unvollständige Teams öffentlich sichtbar" aktiviert hat).</li>
</ul>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// v13.0: Vier weitere Templates aus dem Inline-Code geholt — analog zur
// Team-Migration in v12.13/v12.14.

// {{Name}} (Empfänger-Vorname), {{RegistrantName}} (voller Name dessen,
// der sie als Zimmerpartner ausgewählt hat), {{EventTitle}}, {{AppUrl}}.
export const ROOMMATE_REQUEST_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Roommate request', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{RegistrantName}}</strong> has selected you as their <strong>roommate</strong> for the event <strong>{{EventTitle}}</strong>.</p>
<p>To confirm the match, please pick <strong>{{RegistrantName}}</strong> as your roommate when registering. The organizers will then see a mutual match in the admin center.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const ROOMMATE_REQUEST_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Zimmerpartner-Anfrage', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{RegistrantName}}</strong> hat dich als <strong>Zimmerpartner</strong> für das Event <strong>{{EventTitle}}</strong> angegeben.</p>
<p>Wenn du das Match bestätigen möchtest, gib bei deiner Registrierung <strong>{{RegistrantName}}</strong> ebenfalls als Zimmerpartner an. Das Orga-Team sieht dann im Admin Center, dass ihr euch gegenseitig ausgewählt habt.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{GroupLabel}} (neue Gruppe), {{EventTitle}}, {{AppUrl}}.
export const GROUP_SWITCH_CONFIRMED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Group switch', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>Your group switch to <strong>{{GroupLabel}}</strong> for <strong>{{EventTitle}}</strong> is confirmed. You are now regularly registered in this group.</p>
<p>You can review your participation any time in the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a> under <strong>“My Events”</strong>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const GROUP_SWITCH_CONFIRMED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Gruppen-Wechsel', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>Dein Gruppen-Wechsel zu <strong>{{GroupLabel}}</strong> für <strong>{{EventTitle}}</strong> ist bestätigt. Du bist jetzt regulär in dieser Gruppe angemeldet.</p>
<p>Deine Teilnahme kannst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter <strong>„Meine Events"</strong> einsehen.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

export const GROUP_SWITCH_WAITLIST_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Group switch — on waitlist', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>You requested to switch to the <strong>{{GroupLabel}}</strong> group for <strong>{{EventTitle}}</strong>. The group is currently full, so your registration is on the <strong>{{GroupLabel}} waitlist</strong>.</p>
<p>You will be promoted automatically as soon as a spot frees up. You don't need to do anything else.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
export const GROUP_SWITCH_WAITLIST_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Gruppen-Wechsel — auf Warteliste', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>Du hast den Wechsel in die Gruppe <strong>{{GroupLabel}}</strong> für <strong>{{EventTitle}}</strong> angefragt. Diese Gruppe ist aktuell voll, daher steht deine Anmeldung auf der <strong>Warteliste der Gruppe {{GroupLabel}}</strong>.</p>
<p>Sobald jemand absagt, rückst du automatisch nach. Du musst nichts weiter tun.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{EventTitle}}, {{WaitlistPositionBlock}}
// (optionaler HTML-Block mit „Du stehst jetzt auf Warteliste-Platz X" —
// leer wenn keine Position bekannt).
export const OVERBOOK_APOLOGY_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Registration corrected', '{{EventTitle}}',
  `<p>Hi {{Name}},</p>
<p>We sincerely apologize for a technical problem: due to a large number of simultaneous registrations, you were mistakenly confirmed a spot for <strong>{{EventTitle}}</strong> although capacity was already full.</p>
<p>We therefore had to move your registration to the <strong>waitlist</strong>. We're truly sorry — this was not your fault but caused by a registration rush.</p>
{{WaitlistPositionBlock}}
<p>As soon as a spot opens up you will be promoted automatically and notified right away. Nothing else is needed from your side.</p>
<p style="margin-top:24px;"><strong>Thank you for your understanding</strong><br><br><strong>Your Event Team</strong></p>`
);
export const OVERBOOK_APOLOGY_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Anmeldung korrigiert', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>leider müssen wir uns für ein technisches Problem entschuldigen: durch sehr viele zeitgleiche Anmeldungen wurde dir für <strong>{{EventTitle}}</strong> versehentlich ein Platz bestätigt, obwohl die Kapazität bereits erschöpft war.</p>
<p>Wir mussten deine Anmeldung daher auf die <strong>Warteliste</strong> korrigieren. Das tut uns aufrichtig leid — es lag nicht an dir, sondern an einem Ansturm auf die Anmeldung.</p>
{{WaitlistPositionBlock}}
<p>Sobald ein Platz frei wird, rückst du automatisch nach und bekommst sofort eine Bestätigung. Du musst nichts weiter tun.</p>
<p style="margin-top:24px;"><strong>Vielen Dank für dein Verständnis</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// Fester Listenname auf jeder Subsite
