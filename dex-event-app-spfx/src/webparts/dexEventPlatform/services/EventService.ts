/**
 * Event Service - SharePoint-Operationen für Events und Teilnehmerlisten
 *
 * Erstellt DEX_Events-Liste automatisch beim ersten Start.
 * Erstellt pro Event eine Subsite mit einer "Teilnehmer"-Liste.
 *
 * Struktur auf SharePoint:
 *   DOL-c-DE-B2Run (Hauptsite)
 *   ├── DEX_Events (zentrale Event-Liste)
 *   ├── DEX_Roles (Rollenverwaltung)
 *   ├── b2run-frankfurt-2026 (Subsite)
 *   │   └── Teilnehmer (Registrierungsliste)
 *   ├── jpmorgan-muenchen-2026 (Subsite)
 *   │   └── Teilnehmer (Registrierungsliste)
 *   └── ...
 *
 * - Eike, Maerz 2026
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { wrapTemplateForStorage, buildEmailFromTemplate } from './EmailTemplates';
import { buildOutlookLocation } from '../utils/eventFormat';

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
const OUTLOOK_DECLINE_BODY_EN = wrapTemplateForStorage(
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

const OUTLOOK_DECLINE_BODY_DE = wrapTemplateForStorage(
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
const OUTLOOK_DECLINE_BODY_ONBEHALF_EN = wrapTemplateForStorage(
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

const OUTLOOK_DECLINE_BODY_ONBEHALF_DE = wrapTemplateForStorage(
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
const OUTLOOK_FORWARD_BODY_EN = wrapTemplateForStorage(
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

const OUTLOOK_FORWARD_BODY_DE = wrapTemplateForStorage(
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
const OUTLOOK_DECLINE_DIGEST_BODY_EN = wrapTemplateForStorage(
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

const OUTLOOK_DECLINE_DIGEST_BODY_DE = wrapTemplateForStorage(
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
const NACHRUECKEN_BODY_EN = wrapTemplateForStorage(
  '#86bc25',
  'You\u2019ve got a spot!',
  'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
<p>Great news! A spot has become available \u2014 you have <strong>moved up from the waitlist</strong> for the event <strong>{{EventTitle}}</strong> and are now a <strong>confirmed participant</strong>.</p>
<p>You are now on the official participant list.</p>
<p>You can review your participation any time in the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a> under <strong>\u201CMy Events\u201D</strong>.</p>
<p>If you are unable to attend after all, please cancel your registration as soon as possible via the App so that the next person on the waitlist can move up.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

const NACHRUECKEN_BODY_DE = wrapTemplateForStorage(
  '#86bc25',
  'Du hast einen Platz!',
  'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>gute Nachrichten! Ein Platz ist frei geworden \u2014 du bist von der <strong>Warteliste nachger\u00FCckt</strong> f\u00FCr das Event <strong>{{EventTitle}}</strong> und bist jetzt <strong>fester Teilnehmer</strong>.</p>
<p>Du stehst nun auf der offiziellen Teilnehmerliste.</p>
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
const ORG_NACHRUECKER_BODY_EN = wrapTemplateForStorage(
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
const ORG_NACHRUECKER_BODY_DE = wrapTemplateForStorage(
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
const CANCEL_BANNER_HTML = '<div style="margin:16px 0 20px;padding:14px 18px;border:2px solid #da291c;background:rgba(218,41,28,0.06);border-radius:8px;text-align:center;"><div style="font-size:0.78rem;font-weight:700;color:#da291c;text-transform:uppercase;letter-spacing:1.5px;">Stornierung &middot; Cancellation</div><div style="margin-top:6px;font-size:1.15rem;font-weight:700;color:#888;text-decoration:line-through;">{{EventTitle}}</div></div>';

const ABMELDUNG_AUTO_BODY_EN = wrapTemplateForStorage(
  '#da291c', 'Cancellation confirmed', 'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
${CANCEL_BANNER_HTML}
<p>your registration for the event above has been <strong>cancelled</strong> because you declined the Outlook invitation.</p>
<p>If you change your mind, you can register again any time via the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const ABMELDUNG_AUTO_BODY_DE = wrapTemplateForStorage(
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
const TEAM_MEMBER_JOINED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Team update', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{NewMemberName}}</strong> joined your team {{TeamName}} for the event <strong>{{EventTitle}}</strong>.</p>
<p>You can see the current team status in the <a href="{{AppUrl}}">DEX App</a> under <strong>“My Events”</strong>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const TEAM_MEMBER_JOINED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Team-Update', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{NewMemberName}}</strong> ist deinem Team {{TeamName}} beim Event <strong>{{EventTitle}}</strong> beigetreten.</p>
<p>Den aktuellen Team-Stand siehst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter <strong>„Meine Events"</strong>.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Lead-Vorname), {{RequesterName}} (voll), {{TeamName}},
// {{EventTitle}}, {{ApproveUrl}}, {{RejectUrl}}.
const TEAM_JOIN_REQUEST_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Team join request', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{RequesterName}}</strong> would like to join your team {{TeamName}} for the event <strong>{{EventTitle}}</strong>. Please decide:</p>
<p style="text-align:center;margin:18px 0;"><a href="{{ApproveUrl}}" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Approve</a> <a href="{{RejectUrl}}" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Reject</a></p>
<p style="font-size:0.85rem;color:#666;">Note: the buttons lead you to the app; the request block lives under <strong>“My Events”</strong>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const TEAM_JOIN_REQUEST_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Team-Beitritts-Anfrage', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{RequesterName}}</strong> möchte deinem Team {{TeamName}} beim Event <strong>{{EventTitle}}</strong> beitreten. Bitte entscheide:</p>
<p style="text-align:center;margin:18px 0;"><a href="{{ApproveUrl}}" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Bestätigen</a> <a href="{{RejectUrl}}" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Ablehnen</a></p>
<p style="font-size:0.85rem;color:#666;">Hinweis: die Buttons führen dich in die App; den Beitritts-Anfragen-Block findest du unter <strong>„Meine Events"</strong>.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Anfrager-Vorname), {{EventTitle}}.
const TEAM_JOIN_REJECTED_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Team join request declined', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>your join request for the team at the event <strong>{{EventTitle}}</strong> was declined by the team lead.</p>
<p>You can still register individually if capacity allows — or join another open team via the registration page.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const TEAM_JOIN_REJECTED_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Team-Beitritts-Anfrage abgelehnt', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>deine Beitritts-Anfrage zum Team beim Event <strong>{{EventTitle}}</strong> wurde vom Team-Lead abgelehnt.</p>
<p>Du kannst dich gerne einzeln anmelden, falls die Kapazität noch reicht — oder einem anderen offenen Team über die Anmeldeseite beitreten.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{NewLeadName}}, {{TeamName}},
// {{EventTitle}}, {{NewLeadBlock}} (HTML-Block — leer falls Empfänger
// nicht der neue Lead ist, sonst der zusätzliche Hinweis-Absatz).
const TEAM_LEAD_TRANSFERRED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Team lead change', 'Event {{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>The team lead role in your team {{TeamName}} has been transferred to <strong>{{NewLeadName}}</strong>.</p>
{{NewLeadBlock}}
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const TEAM_LEAD_TRANSFERRED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Team-Lead-Wechsel', 'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>Die Team-Lead-Rolle in deinem Team {{TeamName}} wurde an <strong>{{NewLeadName}}</strong> übergeben.</p>
{{NewLeadBlock}}
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{CancelledName}}, {{TeamName}},
// {{EventTitle}}, {{ActiveCount}}, {{TeamSize}}, {{NewLeadBlock}}
// (leer, falls Empfänger nicht zum neuen Lead ernannt wurde).
const TEAM_MEMBER_CANCELLED_BODY_EN = wrapTemplateForStorage(
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
const TEAM_MEMBER_CANCELLED_BODY_DE = wrapTemplateForStorage(
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
const ROOMMATE_REQUEST_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Roommate request', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p><strong>{{RegistrantName}}</strong> has selected you as their <strong>roommate</strong> for the event <strong>{{EventTitle}}</strong>.</p>
<p>To confirm the match, please pick <strong>{{RegistrantName}}</strong> as your roommate when registering. The organizers will then see a mutual match in the admin center.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const ROOMMATE_REQUEST_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Zimmerpartner-Anfrage', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p><strong>{{RegistrantName}}</strong> hat dich als <strong>Zimmerpartner</strong> für das Event <strong>{{EventTitle}}</strong> angegeben.</p>
<p>Wenn du das Match bestätigen möchtest, gib bei deiner Registrierung <strong>{{RegistrantName}}</strong> ebenfalls als Zimmerpartner an. Das Orga-Team sieht dann im Admin Center, dass ihr euch gegenseitig ausgewählt habt.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{GroupLabel}} (neue Gruppe), {{EventTitle}}, {{AppUrl}}.
const GROUP_SWITCH_CONFIRMED_BODY_EN = wrapTemplateForStorage(
  '#86bc25', 'Group switch', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>Your group switch to <strong>{{GroupLabel}}</strong> for <strong>{{EventTitle}}</strong> is confirmed. You are now regularly registered in this group.</p>
<p>You can review your participation any time in the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView">DEX App</a> under <strong>“My Events”</strong>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const GROUP_SWITCH_CONFIRMED_BODY_DE = wrapTemplateForStorage(
  '#86bc25', 'Gruppen-Wechsel', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>Dein Gruppen-Wechsel zu <strong>{{GroupLabel}}</strong> für <strong>{{EventTitle}}</strong> ist bestätigt. Du bist jetzt regulär in dieser Gruppe angemeldet.</p>
<p>Deine Teilnahme kannst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter <strong>„Meine Events"</strong> einsehen.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

const GROUP_SWITCH_WAITLIST_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Group switch — on waitlist', '{{EventTitle}}',
  `<p>Hello {{Name}},</p>
<p>You requested to switch to the <strong>{{GroupLabel}}</strong> group for <strong>{{EventTitle}}</strong>. The group is currently full, so your registration is on the <strong>{{GroupLabel}} waitlist</strong>.</p>
<p>You will be promoted automatically as soon as a spot frees up. You don't need to do anything else.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);
const GROUP_SWITCH_WAITLIST_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Gruppen-Wechsel — auf Warteliste', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>Du hast den Wechsel in die Gruppe <strong>{{GroupLabel}}</strong> für <strong>{{EventTitle}}</strong> angefragt. Diese Gruppe ist aktuell voll, daher steht deine Anmeldung auf der <strong>Warteliste der Gruppe {{GroupLabel}}</strong>.</p>
<p>Sobald jemand absagt, rückst du automatisch nach. Du musst nichts weiter tun.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// {{Name}} (Empfänger-Vorname), {{EventTitle}}, {{WaitlistPositionBlock}}
// (optionaler HTML-Block mit „Du stehst jetzt auf Warteliste-Platz X" —
// leer wenn keine Position bekannt).
const OVERBOOK_APOLOGY_BODY_EN = wrapTemplateForStorage(
  '#ed8b00', 'Registration corrected', '{{EventTitle}}',
  `<p>Hi {{Name}},</p>
<p>We sincerely apologize for a technical problem: due to a large number of simultaneous registrations, you were mistakenly confirmed a spot for <strong>{{EventTitle}}</strong> although capacity was already full.</p>
<p>We therefore had to move your registration to the <strong>waitlist</strong>. We're truly sorry — this was not your fault but caused by a registration rush.</p>
{{WaitlistPositionBlock}}
<p>As soon as a spot opens up you will be promoted automatically and notified right away. Nothing else is needed from your side.</p>
<p style="margin-top:24px;"><strong>Thank you for your understanding</strong><br><br><strong>Your Event Team</strong></p>`
);
const OVERBOOK_APOLOGY_BODY_DE = wrapTemplateForStorage(
  '#ed8b00', 'Anmeldung korrigiert', '{{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>leider müssen wir uns für ein technisches Problem entschuldigen: durch sehr viele zeitgleiche Anmeldungen wurde dir für <strong>{{EventTitle}}</strong> versehentlich ein Platz bestätigt, obwohl die Kapazität bereits erschöpft war.</p>
<p>Wir mussten deine Anmeldung daher auf die <strong>Warteliste</strong> korrigieren. Das tut uns aufrichtig leid — es lag nicht an dir, sondern an einem Ansturm auf die Anmeldung.</p>
{{WaitlistPositionBlock}}
<p>Sobald ein Platz frei wird, rückst du automatisch nach und bekommst sofort eine Bestätigung. Du musst nichts weiter tun.</p>
<p style="margin-top:24px;"><strong>Vielen Dank für dein Verständnis</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// Fester Listenname auf jeder Subsite
const REG_LIST_NAME = 'Teilnehmer';
const REG_LIST_ITEM_TYPE = 'SP.Data.TeilnehmerListItem';
// v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe (ETag-basiert).
// Pro Subsite eine Liste mit genau einem Item, dessen NextValue beim
// Anmelden via If-Match-Header inkrementiert wird. So können mehrere
// User parallel registrieren ohne dass IDs doppelt vergeben werden.
//
// Listenname mit Unterstrich → SharePoint kodiert das in der Item-Type-
// Bezeichnung als `_x005f_`. Genauso wie wir das schon bei DEX_Events
// machen ('SP.Data.DEX_x005f_EventsListItem'). Ohne dieses Encoding
// schlägt der POST stillschweigend mit HTTP 400 fehl, weil SP den
// Typ-Namen nicht auflösen kann.
const COUNTER_LIST_NAME = 'DEX_TeilnehmerCounter';
const COUNTER_LIST_ITEM_TYPE = 'SP.Data.DEX_x005f_TeilnehmerCounterListItem';

export interface DeclinedAttendee {
  email: string;
  name: string;
}

// v18.66: Ergebnis-Zusammenfassung des Template-Reseeds, damit die UI dem
// Admin konkret zurückmeldet, was passiert ist (statt nur "erfolgreich").
export interface ReseedSummary {
  created: number;
  updated: number;
  skipped: number;
  failed: number;
  errors: string[];
}

export type DeclineCheckReason = 'no-pointer' | 'not-found' | 'forbidden' | 'error';

// Flaches Ergebnis-Shape, weil der TS-4.7-Compiler bei Discriminated Unions
// ueber eine async-Funktion + React-Callback teils nicht korrekt narrowed.
// Erfolg: ok=true, attendees gefüllt, reason/message leer.
// Fehler: ok=false, reason gesetzt, attendees leeres Array.
export interface DeclineCheckResult {
  ok: boolean;
  attendees: DeclinedAttendee[];
  reason?: DeclineCheckReason;
  message?: string;
}

export interface SPEvent {
  Id: number;
  Title: string;
  EventStatus: string;
  EventType?: string; // @deprecated seit v5.2 — wird aus CustomFields abgeleitet; Spalte kann aus DEX_Events entfernt werden.
  EventNumber: number;
  Description: string;
  Location: string;
  LocationAddress: string; // JSON-String: { street, houseNo, zip, city }
  LocationFilter: string;
  Audience: string; // Zielgruppen-Filter (Gruppen + Emails, kommasepariert)
  FilterMode: string; // 'AND' | 'OR' - Verknüpfung Standort+Zielgruppe
  StartDate: string;
  EndDate: string;
  RegistrationDeadline: string;
  LastDeregisterDate: string;
  MaxParticipants: number;
  WaitlistEnabled: boolean;
  EventImageUrl: string;
  EmailImageBase64: string; // Base64 Event-Bild für E-Mails/Outlook
  Organizer: string;
  OrganizerEmail: string;
  /** v10.16: optionaler Ansprechpartner (Anzeige-Feld). */
  ContactName?: string;
  ContactEmail?: string;
  ContactInfo?: string;
  OutlookEventId: string;
  CalendarLink: string;
  OutlookBody: string; // Text für den Outlook-Kalendereintrag
  OutlookSubject?: string; // v18.42: Betreff des Outlook-Termins (leer = Event-Titel)
  OutlookStart?: string; // v18.44: abweichende Start-Zeit des Outlook-Termins (ISO, leer = Event-Start)
  OutlookEnd?: string;   // v18.44: abweichende End-Zeit des Outlook-Termins (ISO, leer = Event-Ende)
  OutlookLocation?: string; // v18.34: lesbarer Ort für das Location-Feld des Outlook-Termins
  EmailLanguage: string; // DE oder EN
  RegistrationLanguage?: string; // v18.35: erzwungene Sprache der Anmeldeseite ('de' | 'en' | '')
  EmailTemplateOverrides: string; // JSON mit Event-spezifischen Template-Anpassungen
  DisableEmails: boolean; // true = keine E-Mails bei An-/Abmeldung
  DisableRegistrationEmail?: boolean; // v19.21: true = keine Anmelde-Bestätigung (Master DisableEmails sticht weiterhin)
  DisableCancellationEmail?: boolean; // v19.21: true = keine Abmelde-Bestätigung
  AutoDeregisterOnDecline?: boolean; // v19.23: true = Outlook-Absage meldet automatisch vom Event ab
  DisableOutlook: boolean; // true = keine Outlook-Kalendereintraege
  OutlookDirty?: boolean; // v11.57: true = Outlook-relevante Felder geändert, Update an Teilnehmer-Termine noch nicht angestoßen
  AutoSendQRCode?: boolean; // v9.15: true = nach Anmeldung automatisch QR-Code-Mail versenden
  ActiveFrom?: string; // v9.21: ISO-Datum, ab dem ein "Active"-Event tatsächlich sichtbar wird
  // v8.5: Granulare Organizer-BCC-Modi
  NotifyOrgRegisterMode?: string; // 'Never' | 'Always' | 'FromDate'
  NotifyOrgRegisterFromDate?: string;
  NotifyOrgCancelMode?: string; // 'Never' | 'Always' | 'AfterDeadline'
  ExcludedUsers?: string; // v8.6: semikolon-separierte User-Mails die das Event NICHT sehen sollen
  IsFictive?: boolean; // true = Test-Event (nur Admin + eigene Organizer sichtbar)
  DurchstarterCapacity?: number; // Split-Capacity Gruppe A (historisch B2Run-Durchstarter)
  FunstarterCapacity?: number;   // Split-Capacity Gruppe B (historisch B2Run-Funstarter)
  SplitLabelA?: string; // v10.20: frei wählbare Bezeichnung Gruppe A
  SplitLabelB?: string; // v10.20: frei wählbare Bezeichnung Gruppe B
  SplitSharedWaitlist?: boolean; // v10.20: true = gemeinsame Warteliste, false = getrennt (Default)
  AllowAttendeeUpload?: boolean; // v11.0: Teilnehmer können PDF an ihre Anmeldung hängen
  AttendeeUploadHint?: string;   // v11.0: optionaler Hinweistext über dem Upload-Input
  AttendeeUploadLabel?: string;  // v11.0: Anzeige-Name des Upload-Felds in MyEvents
  AskSalutation?: boolean;       // v11.80: Anrede im Registrierungsformular abfragen
  ConfirmDialogEnabled?: boolean; // v18.75: Sicherheitshinweis vor dem Absenden
  ConfirmDialogMode?: string;     // v18.75: 'summary' | 'freetext'
  ConfirmDialogText?: string;     // v18.75: Freitext-Hinweis
  SelfCheckInEnabled?: boolean;  // v18.33: Self-Check-in per QR-Code erlauben
  SelfCheckInToken?: string;     // v18.33: Geheimer Token (statischer Link + HMAC-Schlüssel rotierender QR)
  SelfCheckInFrom?: string;      // v18.33: optionaler Start des Check-in-Fensters (ISO), leer = Default 2 Std. vor Start (v20.3)
  SelfCheckInTo?: string;        // v18.33: optionales Ende des Check-in-Fensters (ISO)
  TeamRegistrationEnabled?: boolean; // v11.80: Team-Anmeldung erlauben
  TeamSize?: number;             // v11.80: Maximale Teamgröße
  AskTeamName?: boolean;         // v11.80: Team-Namen abfragen
  TeamPartialAllowed?: boolean;       // v11.81: Auch Teil-Teams erlauben (statt nur komplette)
  TeamOpenSlotsVisible?: boolean;     // v11.81: Offene Slots öffentlich sichtbar für Beitritt
  TeamJoinRequiresApproval?: boolean; // v11.81: Beitritt erfordert Bestätigung durch Team-Kapitän
  BilingualFields?: boolean; // v17.20: Custom-Fields zweisprachig (DE + EN) anbieten
  CustomFields: string; // JSON-String mit konfigurierbaren Feldern
  Agenda: string; // JSON-Array mit Agenda-Einträgen
  Transfers: string; // JSON-Array mit Transferzeiten
  Documents: string; // JSON-Array mit Dokumenten
  FunZone: string; // JSON-Array mit Quiz-Fragen
  QuizClusterSize?: number; // 1..4 - wie viele Fragen pro Quiz-Ansicht. Optional, Default 1.
  ParentEventId?: string; // Seit v6.4: wenn gesetzt, ist dies ein Sub-Event und zeigt auf das Parent-Event. Leer = Top-Level-Event.
  RegistrationListName: string;
  SubsiteUrl: string; // Absolute URL der Event-Subsite
}

export interface CustomField {
  id: string;
  label: string;
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate' | 'document' | 'date'; // v19.0: document = Datei-Upload (Attachment); v24.25: date = Kalender-Auswahl
  required: boolean;
  options?: string[]; // für select-Felder
  /** v24.25: Nur für `type === 'date'` — zusätzlich die Uhrzeit abfragen. */
  withTime?: boolean;
  visible: boolean;
  /** v7.20: Optionaler Hilfe-/Beschreibungstext, der im Registrierungs-
   *  Formular als "i"-Tooltip neben dem Feld-Label sichtbar ist. */
  helpText?: string;
  /** v18.18: 'tooltip' (Default) oder 'inline' (Erklär-Text unter dem Label). */
  helpTextStyle?: 'tooltip' | 'inline';
  /** v18.41: People-Picker (user/roommate): ausgewählte Person bei An-/Abmelde-
   *  Mail auf CC (nicht im Outlook-Termin). */
  ccOnEmails?: boolean;
  /** v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn das Quell-Feld
   *  einen der `values` als Antwort hat. */
  showIf?: { fieldId: string; values: string[] };
  /** v10.24: Bei Split-Capacity-Events kann ein Feld auf eine der zwei
   *  Gruppen eingeschränkt werden ('A' = Durchstarter, 'B' = Funstarter).
   *  'all'/undefined = für beide Gruppen sichtbar. */
  onlyForGroup?: 'all' | 'A' | 'B';
  /** Optionale externe Links, die unter dem Feld als klickbare Links erscheinen.
   * Aktuell vor allem für B2Run-Zustimmung (AGB + Datenschutz von b2run.de). */
  externalLinks?: Array<{ label: string; url: string }>;
  /** v7.11: Bei type=select erlaubt true Mehrfachauswahl. Wert wird in der
   *  Teilnehmerliste " | "-getrennt gespeichert. */
  multi?: boolean;
  /** v11.94: Nur für type='checkbox' — Text neben der Checkbox (Default
   *  „Ja, bestätigen" / „Yes, confirm"). */
  confirmLabel?: string;
  /** v17.20: Englische Varianten — nur relevant wenn auf Event-Ebene
   *  `bilingualFields=true`. optionsEn ist positional zu options gemappt. */
  labelEn?: string;
  helpTextEn?: string;
  confirmLabelEn?: string;
  optionsEn?: string[];
}

export interface SPRegistration {
  Id: number;
  Title: string; // Email
  TeilnehmerID?: number;
  Anrede?: string;
  Vorname?: string;
  Nachname?: string;
  StarterType?: string; // B2Run: 'Durchstarter' | 'Funstarter'
  PreferredStarterType?: string; // B2Run: Wunsch (bei Fallback/Warteliste)
  QuizScore?: number; // Anzahl richtige Antworten
  QuizAnswers?: string; // JSON-Array der gegebenen Antworten
  QuizCompletedAt?: string; // ISO-DateTime
  ParticipantName: string;
  ParticipantEmail: string;
  Status: string;
  /** v11.36: '' = normal, 'Pending' = vom „Überbuchung prüfen"-Lauf als
   *  über Gruppen-/Event-Kapazität markiert; wartet auf Admin-Entscheidung. */
  OverbookReview?: string;
  RegistrationDate: string;
  RegisteredByName?: string;   // Audit: Name des Users der die Anmeldung durchführte
  RegisteredByEmail?: string;  // Audit: E-Mail des Users der die Anmeldung durchführte
  CancellationDate: string;
  CancelledByName?: string;    // Audit: Name des Users der die Abmeldung ausgelöst hat
  CancelledByEmail?: string;   // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
  /** v7.16: Check-In-Audit — gesetzt sobald checkInParticipant() aufgerufen wird. */
  CheckedInDate?: string;      // ISO-DateTime, wann der Teilnehmer eingecheckt wurde
  CheckedInByName?: string;    // Name des Helfers, der den Check-In ausgelöst hat
  CheckedInByEmail?: string;   // E-Mail des Helfers, der den Check-In ausgelöst hat
  /** v11.82: Team-Anmeldung — TeamId ist die UUID, die alle Mitglieder eines
   *  gemeinsam angemeldeten Teams gruppiert. TeamLead=true nur für die
   *  anmeldende Person. TeamName ist optional (nur wenn AskTeamName aktiv). */
  TeamId?: string;
  TeamLead?: boolean;
  TeamName?: string;
  /** v11.86: Standort des Teilnehmers (aus Anmeldeformular ableiten oder vom
   *  Profil uebernommen). Wird im Team-Badge in „Meine Events" zur
   *  Mitglieder-Identifikation angezeigt. Auf der SP-Liste seit jeher
   *  vorhanden — hier nur als TypeScript-Property nachgezogen. */
  Location?: string;
  /** v24.29: Unternehmenszugehörigkeit / Rechtsträger („Company name" aus dem
   *  Profil, z.B. „Deloitte GmbH" / „Deloitte Consulting"). */
  Company?: string;
  /** v17.15: Nachrück-Audit. PromotedDate ist die ISO-Zeit des Promote
   *  (Warteliste → Angemeldet). ReplacedParticipantEmail ist die E-Mail
   *  der Person, deren Cancel diesen Promote ausgelöst hat („Ersetzt
   *  wen"). ReplacedByParticipantEmail ist auf der cancelnden Person
   *  gesetzt und zeigt die E-Mail der nachrückenden Person („Ersetzt
   *  durch"). Beide Felder Single-Line-Text. */
  PromotedDate?: string;
  ReplacedParticipantEmail?: string;
  ReplacedByParticipantEmail?: string;
  CustomData: string; // JSON mit Custom Field Werten
}

export interface SPParticipant {
  Id: number;
  Title: string; // Email
  Vorname: string;
  Nachname: string;
  Email: string;
  EventRegistered: string; // Kommaseparierte EventNumbers
  EventOnWaitlist: string; // Kommaseparierte EventNumbers
}

export class EventService {
  private context: WebPartContext;
  public siteUrl: string;

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  // ==================== DEX_Emails Liste ====================

  /**
   * E-Mail-Queue-Liste erstellen falls nicht vorhanden.
   * Power Automate reagiert auf neue Einträge und versendet Mails.
   *
   * Spalten:
   * - Title: Betreff der E-Mail
   * - Recipient: Empfänger E-Mail-Adresse
   * - RecipientName: Name des Empfängers
   * - Body: HTML-Inhalt der E-Mail
   * - EmailType: Art der E-Mail (Anmeldung, Abmeldung, Warteliste, Nachrücken, Info)
   * - EventTitle: Name des Events
   * - EventId: ID des Events (Referenz)
   * - Status: Pending, Sent, Failed
   * - SentDate: Wann wurde die Mail versendet
   */
  public async ensureEmailsList(): Promise<void> {
    const listName = 'DEX_Emails';
    const exists = await this.listExists(listName);
    if (exists) {
      // Recipient-Feld auf Plain Text (RichText=false) setzen, falls es
      // noch im alten RichText-Modus ist. SharePoint wrappt sonst den Wert
      // in <div class="ExternalClassXXXX">...</div>, was den Power Automate
      // Flow "emailMessage/To must be String/email" Fehler auslöst.
      try {
        await this.setRecipientFieldPlainText(listName);
      } catch { /* ignore */ }

      // Cc-Feld nachträglich anlegen (für Anfrage-Mails von der Landing Page).
      // Bestehende Listen aus aelteren App-Versionen haben das Feld noch nicht.
      try {
        await this.ensureCcFieldExists(listName);
      } catch { /* ignore */ }

      // v8.5: Bcc-Feld nachträglich anlegen — wird genutzt um Organizer
      // automatisch in Anmelde-/Abmelde-Bestätigungen zu BCC'en, ohne den
      // Teilnehmer den Verteiler zu zeigen.
      try {
        await this.ensureBccFieldExists(listName);
      } catch { /* ignore */ }

      // v18.30: Importance-Feld nachträglich anlegen — der DEX_SEND_MAIL-Flow
      // setzt darauf basierend die Outlook-Wichtigkeit (High = rotes „!").
      // Leer/„Normal" = normale Wichtigkeit.
      try {
        await this.ensureImportanceFieldExists(listName);
      } catch { /* ignore */ }

      // Berechtigungen prüfen
      try {
        const listInfo = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
          SPHttpClient.configurations.v1
        );
        if (listInfo.ok) {
          const data = await listInfo.json();
          if (!data.HasUniqueRoleAssignments) {
            await this.setEmailsListPermissions(listName);
          }
        }
      } catch { /* ignore */ }
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'E-Mail-Queue für automatischen Versand via Power Automate',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fields: Array<Record<string, any>> = [
      // Recipient als Plain-Text Note-Feld (RichText=false), damit der Flow
      // die Email-Adresse(n) ohne HTML-Wrapping bekommt.
      { title: 'Recipient', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
      { title: 'Cc', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
      { title: 'Bcc', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
      { title: 'RecipientName', type: 2 },
      { title: 'Body', type: 3 }, // Body darf Rich/HTML bleiben (wird als HTML gerendert)
      { title: 'EmailType', type: 6, choices: ['Anmeldung', 'Abmeldung', 'Warteliste', 'Nachruecken', 'Info'], metaType: 'SP.FieldChoice' },
      { title: 'EventTitle', type: 2 },
      { title: 'EventId', type: 2 },
      { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
      { title: 'SentDate', type: 4 },
      // v18.30: Outlook-Wichtigkeit (leer/„Normal" = normal, „High" = rotes „!").
      { title: 'Importance', type: 2 },
    ];

    for (const f of fields) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        payload['Choices'] = { 'results': f.choices };
      }
      if (f.metaType === 'SP.FieldMultiLineText') {
        payload['RichText'] = !!f.richText;
        if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
      }
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    // Default View
    await this.configureDefaultView(listName, [
      'Recipient', 'RecipientName', 'EmailType', 'EventTitle', 'Status', 'SentDate',
    ]);

    await this.setEmailsListPermissions(listName);
  }

  /**
   * Recipient-Feld auf Plain Text (RichText=false) umstellen.
   * Idempotent: Wenn schon Plain Text, macht nichts.
   * Nur möglich wenn der Current User Manage Lists Rechte hat (Owner/Admin).
   */
  private async setRecipientFieldPlainText(listName: string): Promise<void> {
    const fieldUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Recipient')`;
    const resp = await this.context.spHttpClient.get(
      `${fieldUrl}?$select=FieldTypeKind,RichText`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return;
    const data = await resp.json();
    const kind = data.FieldTypeKind ?? data.d?.FieldTypeKind;
    const richText = data.RichText ?? data.d?.RichText;
    // Wenn bereits Note + Plain Text: nichts zu tun
    if (kind === 3 && richText === false) return;
    // Feld auf Note + RichText=false patchen
    await this._merge(
      fieldUrl,
      {
        '__metadata': { 'type': 'SP.FieldMultiLineText' },
        'FieldTypeKind': 3,
        'RichText': false,
        'NumberOfLines': 3,
      }
    );
  }

  /**
   * Cc-Feld auf DEX_Emails anlegen, falls noch nicht vorhanden.
   * Multi-line Plain-Text damit auch ;-separierte Mehrfach-Adressen passen.
   */
  private async ensureCcFieldExists(listName: string): Promise<void> {
    const probeUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Cc')?$select=Id`;
    const probe = await this.context.spHttpClient.get(probeUrl, SPHttpClient.configurations.v1);
    if (probe.ok) return;
    await this._post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
      {
        '__metadata': { 'type': 'SP.FieldMultiLineText' },
        'Title': 'Cc',
        'FieldTypeKind': 3,
        'Required': false,
        'RichText': false,
        'NumberOfLines': 3,
      }
    );
  }

  private async ensureBccFieldExists(listName: string): Promise<void> {
    const probeUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Bcc')?$select=Id`;
    const probe = await this.context.spHttpClient.get(probeUrl, SPHttpClient.configurations.v1);
    if (probe.ok) return;
    await this._post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
      {
        '__metadata': { 'type': 'SP.FieldMultiLineText' },
        'Title': 'Bcc',
        'FieldTypeKind': 3,
        'Required': false,
        'RichText': false,
        'NumberOfLines': 3,
      }
    );
  }

  // v18.30: Importance-Spalte (Single line text) idempotent anlegen. Der
  // DEX_SEND_MAIL-Flow liest sie und sendet bei „High" mit hoher Wichtigkeit.
  private async ensureImportanceFieldExists(listName: string): Promise<void> {
    const probeUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Importance')?$select=Id`;
    const probe = await this.context.spHttpClient.get(probeUrl, SPHttpClient.configurations.v1);
    if (probe.ok) return;
    await this._post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
      {
        '__metadata': { 'type': 'SP.Field' },
        'Title': 'Importance',
        'FieldTypeKind': 2,
        'Required': false,
      }
    );
  }

  /**
   * Berechtigungen für DEX_Emails: Owners Full Control, Members Contribute, Item-Level Security
   */
  private async setEmailsListPermissions(listName: string): Promise<void> {
    try {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );
      const ownersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (ownersResp.ok) {
        const d = await ownersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
        );
      }
      const deallId = await this.getVisitorsGroupId();
      if (deallId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${deallId}, roledefid=1073741827)`, {}
        );
      }
    } catch { /* */ }

    // Item-Level Security
    try {
      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'ReadSecurity': 2, 'WriteSecurity': 2,
          }),
        }
      );
    } catch { /* */ }
  }

  /**
   * Berechtigungen für Queue-Listen (DEX_Outlook, DEX_IDReorder):
   * Owners Full Control, Members Contribute, Item-Level Security
   */
  private async setQueueListPermissions(listName: string): Promise<void> {
    try {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );
      const ownersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (ownersResp.ok) {
        const d = await ownersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
        );
      }
      const deallId = await this.getVisitorsGroupId();
      if (deallId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${deallId}, roledefid=1073741827)`, {}
        );
      }
    } catch { /* */ }
  }

  /**
   * E-Mail in die Queue eintragen (wird von Power Automate versendet).
   */
  public async queueEmail(
    subject: string,
    recipient: string,
    recipientName: string,
    body: string,
    emailType: string,
    eventTitle: string,
    eventId: string,
    cc?: string,
    bcc?: string,
    // v18.30: 'High' = Outlook hohe Wichtigkeit (rotes „!"). Default normal.
    importance?: 'High' | 'Normal'
  ): Promise<boolean> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailsListItem' },
        'Title': subject,
        'Recipient': recipient,
        'RecipientName': recipientName,
        'Body': body,
        'EmailType': emailType,
        'EventTitle': eventTitle,
        'EventId': eventId,
        'Status': 'Pending',
      };
      if (cc) payload['Cc'] = cc;
      if (bcc) payload['Bcc'] = bcc;
      if (importance === 'High') payload['Importance'] = 'High';
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items`,
        payload
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== DEX_Outlook Liste ====================

  /**
   * Outlook-Termin-Queue-Liste erstellen falls nicht vorhanden.
   * Power Automate reagiert auf neue Einträge und lädt Teilnehmer
   * zum Outlook-Termin ein oder aus. Der Flow holt sich alle Event-Details
   * (Titel, Datum, Ort, CalendarLink) aus der DEX_Events-Liste via EventId.
   *
   * Spalten:
   * - Title: Kurzbeschreibung (z.B. "Einladung: B2Run")
   * - Attendee: E-Mail-Adresse des Teilnehmers
   * - EventId: ID des Events in DEX_Events (Referenz)
   * - ActionType: Einladen, Ausladen
   * - Status: Pending, Sent, Failed
   * - SentDate: Wann wurde die Aktion ausgeführt
   */
  public async ensureOutlookList(): Promise<void> {
    const listName = 'DEX_Outlook';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Outlook-Termin-Queue: Power Automate lädt Teilnehmer ein/aus',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'Attendee', type: 2 },
      { title: 'EventId', type: 2 },
      // ActionType:
      //  - Einladen / Ausladen: einzelnen Attendee zum Outlook-Termin hinzufügen/entfernen
      //  - UpdateEvent: Titel/Start/Ende aktualisieren (kein Attendee)
      //  - DeleteEvent: kompletten Kalender-Termin löschen (wird beim Löschen eines Events
      //    aus der App abgesetzt, inkl. CalendarLink damit der Flow nicht auf DEX_Events
      //    angewiesen ist - das Event-Item wird direkt danach aus DEX_Events gelöscht).
      { title: 'ActionType', type: 6, choices: ['Einladen', 'Ausladen', 'UpdateEvent', 'DeleteEvent'], metaType: 'SP.FieldChoice' },
      { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
      { title: 'SentDate', type: 4 },
      // CalendarLink (iCalUId) - nur für DeleteEvent nötig, damit der Flow das Outlook-
      // Event auch dann noch finden kann, wenn das DEX_Events-Item schon gelöscht wurde.
      { title: 'CalendarLink', type: 3 },
    ];

    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        payload['Choices'] = { 'results': f.choices };
      }
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    await this.configureDefaultView(listName, [
      'Attendee', 'EventId', 'SubEventId', 'ActionType', 'Status', 'SentDate', 'CalendarLink',
    ]);

    await this.setQueueListPermissions(listName);
  }

  // v18.34: pro App-Session je Event nur einmal den OutlookLocation-Backfill
  // versuchen — verhindert N Roundtrips bei N Einladen-Calls eines Events.
  private _outlookLocationBackfilled = new Set<string>();

  /**
   * v18.34: Backfill für Bestands-Events. Stellt sicher, dass das DEX_Events-Item
   * eine gefüllte OutlookLocation hat, BEVOR der Flow den Termin anlegt/aktualisiert.
   * Neue Events bekommen OutlookLocation bereits beim Anlegen/Bearbeiten — alte
   * (vor v18.34 erstellte) Events hätten sonst eine leere Spalte.
   */
  private async backfillOutlookLocation(eventId: string): Promise<void> {
    if (!eventId || this._outlookLocationBackfilled.has(eventId)) return;
    this._outlookLocationBackfilled.add(eventId);
    try {
      const numId = Number(eventId);
      if (isNaN(numId)) return;
      const ev = await this.getEvent(numId);
      if (!ev) return;
      if (ev.OutlookLocation && ev.OutlookLocation.trim() !== '') return; // schon gesetzt
      const loc = buildOutlookLocation(ev.Location, ev.LocationAddress);
      if (loc) {
        await this.updateEvent(numId, { 'OutlookLocation': loc });
      }
    } catch { /* best effort — Backfill darf den Queue-Eintrag nie blockieren */ }
  }

  /**
   * Outlook-Termin-Einladung in die Queue eintragen.
   * Flow holt Event-Details (Datum, Ort, CalendarLink) aus DEX_Events via EventId.
   */
  public async queueOutlookEvent(
    attendee: string,
    eventId: string,
    eventTitle: string,
    actionType: 'Einladen' | 'Ausladen' | 'UpdateEvent'
  ): Promise<boolean> {
    try {
      // v18.34: OutlookLocation für Bestands-Events nachziehen (einmal pro Event/Session).
      await this.backfillOutlookLocation(eventId);
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
        {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
          'Title': `${actionType}: ${eventTitle}`,
          'Attendee': attendee,
          'EventId': eventId,
          'ActionType': actionType,
          'Status': 'Pending',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * DeleteEvent in die DEX_Outlook-Queue eintragen. Wird vom deleteEvent-Flow
   * aufgerufen, BEVOR das DEX_Events-Item gelöscht wird. Der DEX_Outlook_Einladungen-
   * Flow findet den Outlook-Termin über CalendarLink (iCalUId) und löscht ihn.
   * Attendee bleibt leer - DeleteEvent wirkt event-weit.
   */
  public async queueOutlookDeleteEvent(
    eventId: string,
    eventTitle: string,
    calendarLink: string
  ): Promise<boolean> {
    try {
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
        {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
          'Title': `DeleteEvent: ${eventTitle}`,
          'Attendee': '',
          'EventId': eventId,
          'ActionType': 'DeleteEvent',
          'Status': 'Pending',
          'CalendarLink': calendarLink,
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Liefert alle Attendees, die den Outlook-Kalendertermin abgelehnt haben.
   *
   * Liest den Termin im Postfach der Shared Mailbox `no_reply.events@deloitte.de`
   * via Microsoft Graph. Der Admin-User braucht dafür delegate/shared access
   * auf das Postfach (Mailbox-Permission) + die SPFx-App muss `Calendars.Read.Shared`
   * im Admin Center genehmigt bekommen.
   *
   * Holt `OutlookEventId` + `CalendarLink` via `GET` auf DEX_Events/{id}. Primärer
   * Lookup des Outlook-Events ueber `OutlookEventId`. Wenn leer (alte Events):
   * Fallback ueber `iCalUId` per `$filter`.
   *
   * Rückgabe-Status:
   * - `ok: true`, `attendees: [...]` - Termin gefunden, Declines extrahiert
   * - `ok: false`, `reason: 'no-pointer'` - DEX_Events hat weder OutlookEventId noch CalendarLink
   * - `ok: false`, `reason: 'not-found'` - Outlook-Termin existiert nicht (mehr)
   * - `ok: false`, `reason: 'forbidden'` - Admin hat keine Mailbox-Permission oder Tenant-Admin hat Calendars.Read.Shared nicht genehmigt
   * - `ok: false`, `reason: 'error'` - unerwarteter Fehler
   */
  public async getDeclinedAttendees(
    eventId: number | string
  ): Promise<DeclineCheckResult> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = this.context as any;
    if (!ctx.msGraphClientFactory) return { ok: false, attendees: [], reason: 'error', message: 'Graph-Client nicht verfügbar.' };

    // 1. OutlookEventId + CalendarLink aus DEX_Events holen. Nutzt den bewährten
    // `getEvent()`-Path (gleiche Abfrage wie der Rest der App). Direktes
    // $select=OutlookEventId,CalendarLink hatte in v5.18 zu leeren Strings
    // geführt obwohl die Spalten in SharePoint gefüllt waren.
    let outlookEventId = '';
    let calendarLink = '';
    let loadedEvent = false;
    try {
      const numericId = Number(eventId);
      const spEvent = await this.getEvent(numericId);
      if (spEvent) {
        loadedEvent = true;
        outlookEventId = String(spEvent.OutlookEventId || '');
        calendarLink = String(spEvent.CalendarLink || '');
        console.warn('[DEX] getDeclinedAttendees: Event geladen', {
          id: numericId,
          outlookEventIdLen: outlookEventId.length,
          calendarLinkLen: calendarLink.length,
        });
      } else {
        console.warn('[DEX] getDeclinedAttendees: getEvent() lieferte null', { eventId });
      }
    } catch (err) {
      console.warn('[DEX] getDeclinedAttendees: getEvent() warf', err);
    }
    if (!outlookEventId && !calendarLink) {
      return {
        ok: false,
        attendees: [],
        reason: 'no-pointer',
        message: loadedEvent
          ? `Event-Item (Id=${eventId}) enthält weder OutlookEventId noch CalendarLink.`
          : `Event-Item (Id=${eventId}) konnte nicht aus DEX_Events geladen werden (403/404?). Details siehe Browser-Console.`,
      };
    }

    // 2. Outlook-Termin via Graph laden
    const mailbox = 'no_reply.events@deloitte.de';
    try {
      const client = await ctx.msGraphClientFactory.getClient('3');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let ev: any = null;
      if (outlookEventId) {
        ev = await client.api(`/users/${mailbox}/events/${outlookEventId}`)
          .select('id,subject,attendees')
          .get();
      } else {
        const escaped = calendarLink.replace(/'/g, "''");
        const resp = await client.api(`/users/${mailbox}/events`)
          .filter(`iCalUId eq '${escaped}'`)
          .select('id,subject,attendees')
          .top(1)
          .get();
        ev = (resp?.value || [])[0] || null;
      }
      if (!ev) return { ok: false, attendees: [], reason: 'not-found' };
      if (!Array.isArray(ev.attendees)) return { ok: true, attendees: [] };
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const declined = ev.attendees
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((a: any) => a?.status?.response === 'declined')
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .map((a: any) => ({
          email: String(a?.emailAddress?.address || '').toLowerCase(),
          name: String(a?.emailAddress?.name || ''),
        }))
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((a: any) => a.email);
      return { ok: true, attendees: declined };
    } catch (err) {
      console.warn('[DEX] getDeclinedAttendees failed:', err);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const status = (err as any)?.statusCode || (err as any)?.status;
      if (status === 403 || status === 401) return { ok: false, attendees: [], reason: 'forbidden' };
      if (status === 404) return { ok: false, attendees: [], reason: 'not-found' };
      return { ok: false, attendees: [], reason: 'error', message: err instanceof Error ? err.message : String(err) };
    }
  }

  /**
   * v22.7: Prüft per Microsoft Graph, ob die übergebenen (Deloitte-)E-Mail-
   * Adressen noch zu einem aktiven Konto gehören. Liefert die Liste der
   * Adressen, zu denen KEIN aktives Konto gefunden wurde — die Person hat
   * womöglich das Unternehmen verlassen oder das Konto ist deaktiviert.
   *
   * - Nur @deloitte-Adressen werden geprüft; externe/Nicht-Deloitte-Adressen
   *   werden übersprungen (nicht zuverlässig prüfbar) und nie gemeldet.
   * - Batches von je 8 Adressen pro Graph-Request (mail/UPN-OR-Filter).
   * - Best-effort: nur Adressen aus ERFOLGREICH abgefragten Batches können
   *   als inaktiv gemeldet werden — fehlgeschlagene Batches erzeugen keinen
   *   Fehlalarm. `ok=false`, wenn gar nichts geprüft werden konnte.
   */
  public async checkAccountsActive(
    emails: string[]
  ): Promise<{ ok: boolean; inactive: string[] }> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = this.context as any;
    if (!ctx.msGraphClientFactory) return { ok: false, inactive: [] };
    const candidates = Array.from(new Set(
      emails
        .map(e => (e || '').trim().toLowerCase())
        .filter(e => /@(.*\.)?deloitte\.(de|com)$/i.test(e))
    ));
    if (candidates.length === 0) return { ok: true, inactive: [] };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let client: any;
    try {
      client = await ctx.msGraphClientFactory.getClient('3');
    } catch {
      return { ok: false, inactive: [] };
    }
    const activeSet = new Set<string>();
    const checkedSet = new Set<string>();
    const esc = (s: string): string => s.replace(/'/g, "''");
    // v24.34 HOTFIX: Graph erlaubt max. 15 OR-Klauseln im $filter. Pro E-Mail
    // erzeugen wir 2 Klauseln (mail + userPrincipalName) → Batch 8 = 16 Klauseln
    // → JEDER Batch scheiterte mit HTTP 400. Batch 7 = 14 Klauseln.
    const BATCH = 7;
    for (let i = 0; i < candidates.length; i += BATCH) {
      const batch = candidates.slice(i, i + BATCH);
      const clauses = batch
        .map(e => `mail eq '${esc(e)}' or userPrincipalName eq '${esc(e)}'`)
        .join(' or ');
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resp: any = await client.api('/users')
          .filter(`(${clauses})`)
          .select('mail,userPrincipalName,accountEnabled')
          .top(999)
          .get();
        // Batch gilt als geprüft (egal ob jemand gefunden wurde).
        for (const e of batch) checkedSet.add(e);
        const found = (resp?.value || []) as Array<{ mail?: string; userPrincipalName?: string; accountEnabled?: boolean }>;
        for (const u of found) {
          if (u.accountEnabled === false) continue; // deaktiviert → zählt als inaktiv
          if (u.mail) activeSet.add(u.mail.toLowerCase());
          if (u.userPrincipalName) activeSet.add(u.userPrincipalName.toLowerCase());
        }
      } catch (err) {
        console.warn('[DEX] checkAccountsActive batch failed:', err);
      }
    }
    if (checkedSet.size === 0) return { ok: false, inactive: [] };
    const inactive = candidates.filter(e => checkedSet.has(e) && !activeSet.has(e));
    return { ok: true, inactive };
  }

  /**
   * v24.33: Unternehmenszugehörigkeit („Company name" / Graph `companyName`)
   * des eingeloggten Users via Microsoft Graph. Die SP-UserProfile-Property
   * „Company" ist im Tenant nicht zuverlässig gefüllt — Graph `/me` schon.
   */
  public async getMyCompanyViaGraph(): Promise<string> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = this.context as any;
    if (!ctx.msGraphClientFactory) return '';
    try {
      const client = await ctx.msGraphClientFactory.getClient('3');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const resp: any = await client.api('/me').select('companyName').get();
      return (resp?.companyName || '').trim();
    } catch {
      return '';
    }
  }

  /**
   * v24.33: Unternehmenszugehörigkeit für mehrere E-Mails via Graph (Batch à 8,
   * gleiches Muster wie checkAccountsActive). Liefert eine Map
   * lowercased-E-Mail → companyName. Für den Backfill bestehender Teilnehmer.
   */
  public async getCompaniesByEmails(emails: string[]): Promise<Record<string, string>> {
    const out: Record<string, string> = {};
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = this.context as any;
    if (!ctx.msGraphClientFactory) return out;
    const candidates = Array.from(new Set(
      emails.map(e => (e || '').trim().toLowerCase()).filter(e => /@(.*\.)?deloitte\.(de|com)$/i.test(e))
    ));
    if (candidates.length === 0) return out;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let client: any;
    try { client = await ctx.msGraphClientFactory.getClient('3'); } catch { return out; }
    const esc = (s: string): string => s.replace(/'/g, "''");
    // v24.34 HOTFIX: Graph erlaubt max. 15 OR-Klauseln im $filter. Pro E-Mail
    // erzeugen wir 2 Klauseln (mail + userPrincipalName) → Batch 8 = 16 Klauseln
    // → JEDER Batch scheiterte mit HTTP 400. Batch 7 = 14 Klauseln.
    const BATCH = 7;
    for (let i = 0; i < candidates.length; i += BATCH) {
      const batch = candidates.slice(i, i + BATCH);
      const clauses = batch.map(e => `mail eq '${esc(e)}' or userPrincipalName eq '${esc(e)}'`).join(' or ');
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resp: any = await client.api('/users').filter(`(${clauses})`).select('mail,userPrincipalName,companyName').top(999).get();
        const found = (resp?.value || []) as Array<{ mail?: string; userPrincipalName?: string; companyName?: string }>;
        for (const u of found) {
          const comp = (u.companyName || '').trim();
          if (!comp) continue;
          if (u.mail) out[u.mail.toLowerCase()] = comp;
          if (u.userPrincipalName) out[u.userPrincipalName.toLowerCase()] = comp;
        }
      } catch (err) {
        console.warn('[DEX] getCompaniesByEmails batch failed:', err);
      }
    }
    return out;
  }

  /**
   * v24.33: Trägt die Unternehmenszugehörigkeit für bestehende Teilnehmer einer
   * Liste nach (Backfill) — lädt alle Zeilen, holt `companyName` via Graph und
   * setzt `Company` per MERGE, aber nur dort, wo es noch leer ist. Best-effort:
   * fehlt die Spalte/das Recht, wird die Zeile übersprungen.
   */
  public async backfillCompanyForList(subsiteUrl: string): Promise<{ updated: number; checked: number }> {
    let regs: SPRegistration[] = [];
    try { regs = await this.getAllRegistrations(subsiteUrl); } catch { return { updated: 0, checked: 0 }; }
    const emails = Array.from(new Set(regs.map(r => (r.ParticipantEmail || '').toLowerCase()).filter(Boolean)));
    if (emails.length === 0) return { updated: 0, checked: regs.length };
    const compMap = await this.getCompaniesByEmails(emails);
    let updated = 0;
    for (const r of regs) {
      const comp = compMap[(r.ParticipantEmail || '').toLowerCase()];
      if (!comp) continue;
      if ((r.Company || '').trim()) continue; // schon gesetzt
      if (!r.Id) continue;
      try {
        await this._merge(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${r.Id})`, { Company: comp });
        updated++;
      } catch { /* Spalte evtl. nicht da / keine Rechte — überspringen */ }
    }
    return { updated, checked: regs.length };
  }

  // ==================== DEX_IDReorder Queue ====================

  /**
   * Queue-Liste für ID-Neuvergabe erstellen falls nicht vorhanden.
   * Power Automate reagiert auf neue Einträge und vergibt TeilnehmerIDs
   * auf der jeweiligen Subsite-Teilnehmerliste lückenlos neu.
   *
   * Spalten:
   * - Title: Kurzbeschreibung (z.B. "Reorder: Test_20260408")
   * - EventId: SP Item-ID des Events in DEX_Events
   * - EventNumber: Hochlaufende EventNumber
   * - SubsiteUrl: Absolute URL der Event-Subsite
   * - Status: Pending, Processing, Done, Failed
   */
  public async ensureIDReorderList(): Promise<void> {
    const listName = 'DEX_IDReorder';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Queue für TeilnehmerID-Neuvergabe nach Abmeldungen',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields = [
      { title: 'EventId', type: 2 },
      { title: 'EventNumber', type: 9 },
      { title: 'SubsiteUrl', type: 2 },
      { title: 'Status', type: 6, choices: ['Pending', 'Processing', 'Done', 'Failed'], metaType: 'SP.FieldChoice' },
      // v18.65: Name der abgemeldeten Person (für die Organizer-Nachrücker-Mail).
      { title: 'CancelledName', type: 2 },
    ];

    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if ((f as { choices?: string[] }).choices) {
        payload['Choices'] = { 'results': (f as { choices: string[] }).choices };
      }
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    await this.configureDefaultView(listName, [
      'EventId', 'EventNumber', 'SubsiteUrl', 'Status',
    ]);

    await this.setQueueListPermissions(listName);
  }

  // ==================== DEX_ChangeLog (v9.0) ====================
  // Audit-Liste für alle Aenderungen an Events und Teilnehmern. Read-
  // Berechtigung für Organizer/Admin (gleiche Permission-Pattern wie
  // DEX_Roles), Schreibrechte für alle (damit User-Aktionen wie
  // Anmeldung/Abmeldung mitloggen können).
  public async ensureChangeLogList(): Promise<void> {
    const listName = 'DEX_ChangeLog';
    try {
      const exists = await this.listExists(listName);
      if (exists) {
        try { await this.ensureChangeLogPermissions(listName); } catch { /* */ }
        return;
      }
      // Liste existiert nicht — versuchen sie anzulegen. Schlägt fehl
      // wenn der aktuelle User keine Owner-Permissions hat. Dann wird der
      // App-Start nicht blockiert (Audit-Log ist best-effort für User
      // ohne Schreibrechte auf der Liste-Erstellung).
      const createResp = await this._post(`${this.siteUrl}/_api/web/lists`, {
        '__metadata': { 'type': 'SP.List' },
        'Title': listName,
        'Description': 'Audit-Log für alle Aenderungen an Events und Teilnehmern (v9.0)',
        'BaseTemplate': 100,
        'AllowContentTypes': false,
      });
      if (!createResp.ok) {
        console.warn('[DEX] DEX_ChangeLog konnte nicht angelegt werden — vermutlich fehlen dem User Owner-Rechte. App läuft weiter, Audit-Einträge fehlen aber.');
        return;
      }
      const fields = [
        { title: 'Action', type: 6, choices: ['EventCreated', 'EventUpdated', 'EventArchived', 'EventRestored', 'EventDeletedPermanent', 'EventDeletedTest', 'ParticipantRegistered', 'ParticipantCancelled', 'ParticipantReactivated', 'ParticipantUpdated', 'ParticipantCheckedIn', 'ParticipantCheckedOut', 'IDReorder', 'Other'], metaType: 'SP.FieldChoice' },
        { title: 'TargetType', type: 6, choices: ['Event', 'Participant', 'Subsite', 'Other'], metaType: 'SP.FieldChoice' },
        { title: 'TargetId', type: 2 },
        { title: 'TargetName', type: 2 },
        { title: 'EventId', type: 2 },
        { title: 'EventTitle', type: 2 },
        { title: 'ActorName', type: 2 },
        { title: 'ActorEmail', type: 2 },
        { title: 'Details', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 6 },
      ];
      for (const f of fields) {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const payload: Record<string, any> = {
            '__metadata': { 'type': f.metaType || 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          };
          if (f.choices) payload['Choices'] = { 'results': f.choices };
          if (f.metaType === 'SP.FieldMultiLineText') {
            payload['RichText'] = !!f.richText;
            if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
          }
          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
        } catch { /* einzelne Feld-Fehler ignorieren */ }
      }
      try {
        await this.configureDefaultView(listName, [
          'Created', 'Action', 'TargetType', 'TargetName', 'EventTitle', 'ActorName', 'Details',
        ]);
      } catch { /* View-Setup ist optional */ }
      try { await this.ensureChangeLogPermissions(listName); } catch { /* */ }
    } catch (err) {
      console.warn('[DEX] ensureChangeLogList failed (best-effort, App läuft weiter):', err);
    }
  }

  // Berechtigungen: Site-Members und alle authentifizierten User können
  // Einträge HINZUFUEGEN (damit Self-Reg/Cancel mitschreibt), aber nur
  // Organizer/Admin können LESEN. Setzt Item-Level-Read = "Only their own".
  private async ensureChangeLogPermissions(listName: string): Promise<void> {
    try {
      // 1. Inheritance brechen
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );
      // 2. Owners (Admin-Group) → Full Control
      const ownersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (ownersResp.ok) {
        const d = await ownersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
        );
      }
      // 3. Members (Organizer-Group typischerweise) → Contribute (sollen
      //    auch schreiben können, damit Organizer-Aktionen wie Event-
      //    Updates protokolliert werden).
      const membersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedmembergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (membersResp.ok) {
        const d = await membersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741827)`, {} // 1073741827 = Contribute
        );
      }
      // 4. Visitors (DEALL / Authenticated Users) → Contribute (damit
      //    User-Aktionen wie Self-Anmeldung mitloggen können).
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`, {} // Contribute
        );
      }
      // 5. Item-Level-Read = "ReadAllItems" (1) — Organizer und Admin
      //    müssen ALLE Einträge sehen, nicht nur eigene. Eigene
      //    Lese-Beschränkung wäre für Audit-Log nutzlos.
      // v22.2 FIX: war vorher ein nackter POST (kein MERGE) → SharePoint
      // antwortete bei jedem Boot mit HTTP 400 (Console-Rauschen); der Wert
      // wurde nie gesetzt (Default ist ohnehin 1, daher ohne Folgen).
      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({ '__metadata': { 'type': 'SP.List' }, 'ReadSecurity': 1 }),
        }
      );
    } catch (e) {
      console.warn('[DEX] ensureChangeLogPermissions failed:', e);
    }
  }

  /**
   * v9.0: Audit-Eintrag schreiben. Best-effort — Fehler werden nur
   * geloggt, blocken die aufrufende Aktion nie. Wird automatisch von
   * createEvent / updateEvent / deleteEvent / registerForEvent /
   * cancelRegistration / adminUpdateRegistration / etc. gerufen.
   */
  public async writeChangeLog(entry: {
    action: string;
    targetType: 'Event' | 'Participant' | 'Subsite' | 'Other';
    targetId?: string;
    targetName?: string;
    eventId?: string;
    eventTitle?: string;
    actorName?: string;
    actorEmail?: string;
    details?: string | Record<string, unknown>;
  }): Promise<void> {
    try {
      const me = this.context.pageContext.user;
      const actorName = entry.actorName || me.displayName || '';
      const actorEmail = (entry.actorEmail || me.email || '').toLowerCase();
      const detailsStr = typeof entry.details === 'string'
        ? entry.details
        : entry.details
          ? JSON.stringify(entry.details)
          : '';
      // CLAUDE.md-Hinweis: bei odata=nometadata KEIN __metadata im Body —
      // SP leitet den Typ aus der URL ab. Robust gegen List-Type-Encoding-
      // Quirks (Bug-Story v7.28 → v7.29). Nutzen wir hier statt verbose-POST
      // damit ein verschmierter Type-Name den ChangeLog-Insert nicht
      // stillschweigend killt.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        'Title': `${entry.action}: ${entry.targetName || entry.targetId || '-'}`,
        'Action': entry.action,
        'TargetType': entry.targetType,
        'TargetId': entry.targetId || '',
        'TargetName': entry.targetName || '',
        'EventId': entry.eventId || '',
        'EventTitle': entry.eventTitle || '',
        'ActorName': actorName,
        'ActorEmail': actorEmail,
        'Details': detailsStr.substring(0, 30000),
      };
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_ChangeLog')/items`;
      const options: ISPHttpClientOptions = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
        body: JSON.stringify(payload),
      };
      await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    } catch (err) {
      console.warn('[DEX] writeChangeLog failed:', err);
    }
  }

  /**
   * Audit-Log lesen — Organizer/Admin only (durch SP-Permissions
   * geschützt). Liefert die letzten N Einträge, optional gefiltert
   * nach EventId.
   */
  public async readChangeLog(opts?: { eventId?: string; top?: number }): Promise<Array<{
    Id: number; Created: string; Action: string; TargetType: string;
    TargetId: string; TargetName: string; EventId: string; EventTitle: string;
    ActorName: string; ActorEmail: string; Details: string;
  }>> {
    const top = opts?.top || 200;
    const base = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_ChangeLog')/items`;
    const sel = `$select=Id,Created,Action,TargetType,TargetId,TargetName,EventId,EventTitle,ActorName,ActorEmail,Details`;
    const filter = opts?.eventId
      ? `&$filter=EventId eq '${String(opts.eventId).replace(/'/g, "''")}'`
      : '';
    // v19.13 BUG-FIX: Das Audit-Log lud mit HTTP 400 (→ „0 Einträge"). Zwei
    // typische Ursachen, gegen beide robust:
    //  (a) `$orderby=Created` auf einer großen Liste ohne Index auf `Created`
    //      → „List View Threshold"-Fehler (400). Deshalb jetzt nach `Id desc`
    //      sortieren — `Id` ist IMMER indiziert, und da auto-increment entspricht
    //      die Reihenfolge chronologisch absteigend.
    //  (b) Ein Feld im `$select` existiert auf einer Bestands-/Legacy-Liste nicht
    //      (Feld-Anlage best-effort) → 400. Deshalb Fallbacks ohne `$select` und
    //      notfalls ohne Server-Filter (dann client-seitig nach EventId filtern).
    const candidates = [
      `${base}?${sel}&$orderby=Id desc&$top=${top}${filter}`,
      `${base}?$orderby=Id desc&$top=${top}${filter}`,
      `${base}?$orderby=Id desc&$top=${top}`,
    ];
    for (let i = 0; i < candidates.length; i++) {
      try {
        const resp = await this.context.spHttpClient.get(candidates[i], SPHttpClient.configurations.v1);
        if (!resp.ok) continue;
        const data = await resp.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let items: any[] = data.value || data.d?.results || [];
        // Letzter Kandidat hat keinen Server-Filter (falls die EventId-Spalte
        // im Filter das 400 ausgelöst hat) → client-seitig nachfiltern.
        if (i === candidates.length - 1 && opts?.eventId) {
          items = items.filter(it => String((it && it.EventId) || '') === String(opts.eventId));
        }
        return items;
      } catch { /* nächsten Kandidaten versuchen */ }
    }
    console.warn('[DEX] readChangeLog: alle Abfrage-Varianten fehlgeschlagen.');
    return [];
  }

  /**
   * ID-Reorder in Queue eintragen (nach Abmeldung).
   */
  // v18.65: einmal pro Session die CancelledName-Spalte auf DEX_IDReorder
  // nachrüsten (Bestands-Listen). Selbstheilend, weil der zentrale
  // initEvents-ensure-Pfad bei gesetztem ENSURE_FLAG übersprungen wird.
  private _idReorderCancelledFieldEnsured = false;
  private async ensureIDReorderCancelledNameField(): Promise<void> {
    if (this._idReorderCancelledFieldEnsured) return;
    this._idReorderCancelledFieldEnsured = true;
    // v19.5: CancelledName UND CancelledEmail nachrüsten. CancelledEmail erlaubt
    // dem Nachrück-Flow, die abgemeldete Person eindeutig zu adressieren
    // (Replaced-Audit: ReplacedByParticipantEmail auf der abgemeldeten Person +
    // ReplacedParticipantEmail auf der nachrückenden Person).
    for (const fieldTitle of ['CancelledName', 'CancelledEmail']) {
      try {
        const resp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/fields/getbytitle('${fieldTitle}')?$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (resp.ok) continue; // existiert bereits
      } catch { /* anlegen */ }
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/fields`, {
          '__metadata': { 'type': 'SP.Field' }, 'Title': fieldTitle, 'FieldTypeKind': 2, 'Required': false,
        });
      } catch { /* best-effort — Retry-ohne-Feld unten fängt es ab */ }
    }
  }

  public async queueIDReorder(
    eventId: string,
    eventNumber: number,
    subsiteUrl: string,
    eventTitle: string,
    // v18.65: Name der abgemeldeten Person — wird in die Queue geschrieben,
    // damit der DEX_IDReorder-Flow ihn direkt aus dem Trigger lesen kann (statt
    // die „jüngste Abmeldung" abzufragen, was bei gleichzeitigen Abmeldungen
    // während des Flow-Laufs falsch sein könnte). Genutzt für die
    // Organizer-Nachrücker-Mail (OrgNachruecker-Template).
    cancelledName?: string,
    // v19.5: E-Mail der abgemeldeten Person — der Nachrück-Flow nutzt sie für
    // das Replaced-Audit (Hat ersetzt / Wurde ersetzt durch).
    cancelledEmail?: string
  ): Promise<boolean> {
    try {
      // ListItemEntityTypeFullName dynamisch ermitteln
      let listItemType = 'SP.Data.DEX_x005f_IDReorderListItem';
      try {
        const typeResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')?$select=ListItemEntityTypeFullName`,
          SPHttpClient.configurations.v1
        );
        if (typeResp.ok) {
          const typeData = await typeResp.json();
          listItemType = typeData.ListItemEntityTypeFullName || typeData.d?.ListItemEntityTypeFullName || listItemType;
        }
      } catch { /* Fallback auf Standard-Name */ }

      if (cancelledName || cancelledEmail) { try { await this.ensureIDReorderCancelledNameField(); } catch { /* */ } }

      const baseBody: Record<string, unknown> = {
        '__metadata': { 'type': listItemType },
        'Title': `Reorder: ${eventTitle}`,
        'EventId': eventId,
        'EventNumber': eventNumber,
        'SubsiteUrl': subsiteUrl,
        'Status': 'Pending',
      };
      // v19.5: CancelledName + CancelledEmail als optionale Zusatzfelder.
      const extra: Record<string, unknown> = {};
      if (cancelledName) extra['CancelledName'] = cancelledName;
      if (cancelledEmail) extra['CancelledEmail'] = cancelledEmail;
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/items`;
      let response = await this._post(url, Object.keys(extra).length ? { ...baseBody, ...extra } : baseBody);
      // Falls die Zusatz-Spalten (noch) fehlen, schlägt der erste POST fehl —
      // dann ohne die Felder erneut posten, damit der Reorder NIEMALS verloren
      // geht (kritischer Pfad).
      if (!response.ok && Object.keys(extra).length) {
        response = await this._post(url, baseBody);
      }
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== DEX_EmailTemplates Liste ====================

  /**
   * Email-Templates-Liste erstellen und Default-Templates einfügen.
   * Templates können pro Event ueberschrieben werden (im Event JSON).
   *
   * Platzhalter: {{Name}}, {{EventTitle}}, {{AppUrl}}
   */
  public async ensureEmailTemplatesList(): Promise<void> {
    const listName = 'DEX_EmailTemplates';
    const exists = await this.listExists(listName);
    if (exists) {
      // Liste existiert - prüfen ob _Config Zeile und Logo-Spalten vorhanden
      await this.ensureEmailTemplatesConfig(listName);
      // Neuere Templates nachrüsten (falls die Liste vor v3.0.27 angelegt wurde
      // und OutlookDeclineReminder noch nicht existiert)
      await this.ensureMissingEmailTemplates(listName);
      // Standard-Templates auf aktuelle Version upgraden (uerschreibt User-Customizing!)
      // Damit Platzhalter wie {{WaitlistPosition}} bei aelteren Tenants nachgezogen werden.
      await this.upgradeStandardEmailTemplates(listName);
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Email-Vorlagen für die DEX Event Experience Platform (DE + EN)',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields = [
      { title: 'TemplateType', type: 2 },
      { title: 'Language', type: 2 },
      { title: 'Subject', type: 2 },
      { title: 'HeadingColor', type: 2 },
      { title: 'Heading', type: 2 },
      // v15.17: Subheading editierbar (vorher hart als „Event {{EventTitle}}").
      // Leer/nicht-gesetzt → Fallback im Code auf {{EventTitle}} ohne Präfix.
      { title: 'Subheading', type: 2 },
      { title: 'BodyHtml', type: 3 },
      { title: 'LogoBase64', type: 3 },           // Base64 Deloitte Logo (Deloitte_Logo.png)
      { title: 'DefaultImageBase64', type: 3 },    // Base64 Default-Bild (dex-orb.png)
    ];

    for (const f of fields) {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    }

    // Default-Templates: DE + EN für jeden Typ
    const defaults = [
      // ===== ENGLISCH =====
      { TemplateType: 'Anmeldung', Language: 'EN', Subject: 'Registration confirmation: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Registration successful',
        BodyHtml: '<p>Dear {{Name}},</p><p>you have successfully registered for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">DEX App</a> (\u201EMy Events\u201C).</p><p>For organizational questions about the event, please contact {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Waitlist confirmation',
        BodyHtml: '<p>Dear {{Name}},</p><p>you have been placed on the <strong>waitlist</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>Your current position: <strong>#{{WaitlistPosition}}</strong></p><p>We will notify you as soon as a spot becomes available. You can always check your current position in the <a href="{{AppUrl}}">DEX App</a> under \u201EMy Events\u201C.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
        BodyHtml: '<p>Dear {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>your registration for the event above has been <strong>cancelled</strong>. The Outlook calendar entry will be removed from your calendar shortly.</p><p>If you change your mind, you can register again via the <a href="{{AppUrl}}">DEX App</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'You’ve got a spot: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'You’ve got a spot!',
        BodyHtml: NACHRUECKEN_BODY_EN },
      { TemplateType: 'EventErstellt', Language: 'EN', Subject: '[Deloitte Eventmanager] - New event created: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event Created',
        BodyHtml: '<p>Dear {{Name}},</p><p>your event <strong>{{EventTitle}}</strong> has been successfully created.</p><p>You can manage participants in the <a href="{{AppUrl}}">DEX App</a>.</p><p>Regards,<br>Team DEX App</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_BODY_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      // ===== DEUTSCH =====
      { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Anmeldung erfolgreich',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du hast dich erfolgreich für das Event <strong>{{EventTitle}}</strong> angemeldet.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">DEX App</a> (\u201EMeine Events\u201C) ab.</p><p>Zu organisatorischen Fragen zum Event wende dich bitte an {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Warteliste-Bestätigung',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du stehst auf der <strong>Warteliste</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Deine aktuelle Position: <strong>#{{WaitlistPosition}}</strong></p><p>Wir benachrichtigen dich, sobald ein Platz frei wird. Deinen aktuellen Warteliste-Platz kannst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter \u201EMeine Events\u201C sehen.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
        BodyHtml: '<p>Hallo {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>deine Anmeldung für das oben genannte Event wurde <strong>storniert</strong>. Der Outlook-Termin wird in Kürze aus deinem Kalender entfernt.</p><p>Du kannst dich jederzeit erneut über die <a href="{{AppUrl}}">DEX App</a> anmelden.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Du hast einen Platz: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Du hast einen Platz!',
        BodyHtml: NACHRUECKEN_BODY_DE },
      // v19.25: pre-wrapped Abmelde-Bestätigung für die Flow-getriebene
      // Auto-Abmeldung (DEX_OutlookDeclineHandler), eigener Type damit die
      // App-eigene `Abmeldung` (unwrapped) unberührt bleibt.
      { TemplateType: 'AbmeldungAuto', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
        BodyHtml: ABMELDUNG_AUTO_BODY_EN },
      { TemplateType: 'AbmeldungAuto', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
        BodyHtml: ABMELDUNG_AUTO_BODY_DE },
      // v18.63: Organizer-Benachrichtigung bei Abmeldung mit Nachrücker (vom DEX_IDReorder-Flow gequeued).
      { TemplateType: 'OrgNachruecker', Language: 'EN', Subject: 'Cancellation with waitlist move-up: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Cancellation — waitlist move-up',
        BodyHtml: ORG_NACHRUECKER_BODY_EN },
      { TemplateType: 'OrgNachruecker', Language: 'DE', Subject: 'Abmeldung mit Nachrücker: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Abmeldung — Nachrücker',
        BodyHtml: ORG_NACHRUECKER_BODY_DE },
      { TemplateType: 'EventErstellt', Language: 'DE', Subject: '[Deloitte Eventmanager] - Neues Event erstellt: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event erstellt',
        BodyHtml: '<p>Hallo {{Name}},</p><p>dein Event <strong>{{EventTitle}}</strong> wurde erfolgreich erstellt.</p><p>Du kannst die Teilnehmer in der <a href="{{AppUrl}}">DEX App</a> verwalten.</p><p>Viele Grüße,<br>Team DEX App</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'DE', Subject: 'Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Du hast den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_DE },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'DE', Subject: 'Action Required: Anmeldung für {{EventTitle}} stornieren?', HeadingColor: '#ed8b00', Heading: 'Outlook-Termin in deinem Namen abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_DE },
      // Meeting-Forward-Notification: FYI an Organizer wenn weitergeleitete Person nicht registriert ist
      { TemplateType: 'OutlookForwardNotification', Language: 'EN', Subject: 'FYI: Meeting was forwarded — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Meeting was forwarded',
        BodyHtml: OUTLOOK_FORWARD_BODY_EN },
      { TemplateType: 'OutlookForwardNotification', Language: 'DE', Subject: 'FYI: Termin wurde weitergeleitet — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Termin wurde weitergeleitet',
        BodyHtml: OUTLOOK_FORWARD_BODY_DE },
      // v9.38: OutlookDeclineDigest — geht an Organizer nach jedem Decline mit Liste aller noch-angemeldeten Decliner.
      { TemplateType: 'OutlookDeclineDigest', Language: 'EN', Subject: 'FYI: {{DeclineCount}} attendees declined Outlook — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: attendees declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_EN },
      { TemplateType: 'OutlookDeclineDigest', Language: 'DE', Subject: 'FYI: {{DeclineCount}} Teilnehmer haben Outlook abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: Teilnehmer haben den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_DE },
      // v12.13: Team-bezogene Templates (vorher inline in EventContext.tsx).
      { TemplateType: 'TeamMemberJoined', Language: 'EN', Subject: 'New team member — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team update',
        BodyHtml: TEAM_MEMBER_JOINED_BODY_EN },
      { TemplateType: 'TeamMemberJoined', Language: 'DE', Subject: 'Neues Team-Mitglied — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Update',
        BodyHtml: TEAM_MEMBER_JOINED_BODY_DE },
      { TemplateType: 'TeamJoinRequest', Language: 'EN', Subject: 'Team join request — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team join request',
        BodyHtml: TEAM_JOIN_REQUEST_BODY_EN },
      { TemplateType: 'TeamJoinRequest', Language: 'DE', Subject: 'Team-Beitritts-Anfrage — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Beitritts-Anfrage',
        BodyHtml: TEAM_JOIN_REQUEST_BODY_DE },
      { TemplateType: 'TeamJoinRejected', Language: 'EN', Subject: 'Team join request declined — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team join request declined',
        BodyHtml: TEAM_JOIN_REJECTED_BODY_EN },
      { TemplateType: 'TeamJoinRejected', Language: 'DE', Subject: 'Team-Beitritts-Anfrage abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Beitritts-Anfrage abgelehnt',
        BodyHtml: TEAM_JOIN_REJECTED_BODY_DE },
      { TemplateType: 'TeamLeadTransferred', Language: 'EN', Subject: 'Team lead change — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team lead change',
        BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_EN },
      { TemplateType: 'TeamLeadTransferred', Language: 'DE', Subject: 'Team-Lead-Wechsel — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Lead-Wechsel',
        BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_DE },
      { TemplateType: 'TeamMemberCancelled', Language: 'EN', Subject: 'Team update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team update',
        BodyHtml: TEAM_MEMBER_CANCELLED_BODY_EN },
      { TemplateType: 'TeamMemberCancelled', Language: 'DE', Subject: 'Team-Update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Update',
        BodyHtml: TEAM_MEMBER_CANCELLED_BODY_DE },
      // v13.0: Restliche bisher-inline-Mails (Zimmerpartner, Gruppen-Wechsel, Überbuchung).
      { TemplateType: 'RoommateRequest', Language: 'EN', Subject: '{{RegistrantName}} selected you as roommate — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Roommate request',
        BodyHtml: ROOMMATE_REQUEST_BODY_EN },
      { TemplateType: 'RoommateRequest', Language: 'DE', Subject: '{{RegistrantName}} hat dich als Zimmerpartner gewählt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Zimmerpartner-Anfrage',
        BodyHtml: ROOMMATE_REQUEST_BODY_DE },
      { TemplateType: 'GroupSwitchConfirmed', Language: 'EN', Subject: 'Group switch confirmed — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Group switch',
        BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_EN },
      { TemplateType: 'GroupSwitchConfirmed', Language: 'DE', Subject: 'Gruppen-Wechsel bestätigt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Gruppen-Wechsel',
        BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_DE },
      { TemplateType: 'GroupSwitchWaitlist', Language: 'EN', Subject: 'Group switch — on waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Group switch — on waitlist',
        BodyHtml: GROUP_SWITCH_WAITLIST_BODY_EN },
      { TemplateType: 'GroupSwitchWaitlist', Language: 'DE', Subject: 'Gruppen-Wechsel — auf Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Gruppen-Wechsel — auf Warteliste',
        BodyHtml: GROUP_SWITCH_WAITLIST_BODY_DE },
      { TemplateType: 'OverbookingApology', Language: 'EN', Subject: 'Important: correction of your registration — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Registration corrected',
        BodyHtml: OVERBOOK_APOLOGY_BODY_EN },
      { TemplateType: 'OverbookingApology', Language: 'DE', Subject: 'Wichtig: Korrektur deiner Anmeldung — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Anmeldung korrigiert',
        BodyHtml: OVERBOOK_APOLOGY_BODY_DE },
    ];

    let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
    try {
      const typeResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      );
      if (typeResp.ok) {
        const typeData = await typeResp.json();
        listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
      }
    } catch { /* Fallback */ }

    for (const t of defaults) {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
        '__metadata': { 'type': listItemType },
        'Title': `${t.TemplateType}_${t.Language}`,
        'TemplateType': t.TemplateType,
        'Language': t.Language,
        'Subject': t.Subject,
        'HeadingColor': t.HeadingColor,
        'Heading': t.Heading,
        'BodyHtml': t.BodyHtml,
      });
    }

    // _Config Eintrag für Logos erstellen (Base64 muss manuell in SharePoint eingetragen werden)
    await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
      '__metadata': { 'type': listItemType },
      'Title': '_Config',
      'TemplateType': '_Config',
      'Language': '',
      'Subject': '',
      'HeadingColor': '',
      'Heading': '',
      'BodyHtml': '',
      'LogoBase64': '',           // Manuell: Base64 Data-URI von Deloitte_Logo.png eintragen
      'DefaultImageBase64': '',   // Manuell: Base64 Data-URI von dex-orb.png eintragen
    });

    await this.configureDefaultView(listName, ['TemplateType', 'Language', 'Subject', 'Heading', 'HeadingColor']);
  }

  /**
   * Sicherstellen dass LogoBase64/DefaultImageBase64 Spalten und _Config Zeile existieren.
   * Für Tenants wo DEX_EmailTemplates schon vor v3.0.27 angelegt wurde:
   * neuere Templates (z.B. OutlookDeclineReminder DE+EN) nachrüsten, ohne
   * bestehende zu ueberschreiben.
   */
  private async ensureMissingEmailTemplates(listName: string): Promise<void> {
    const newTemplates = [
      { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_BODY_EN },
      { TemplateType: 'OutlookDeclineReminder', Language: 'DE', Subject: 'Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Du hast den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_DE },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'DE', Subject: 'Action Required: Anmeldung für {{EventTitle}} stornieren?', HeadingColor: '#ed8b00', Heading: 'Outlook-Termin in deinem Namen abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_DE },
      { TemplateType: 'OutlookForwardNotification', Language: 'EN', Subject: 'FYI: Meeting was forwarded — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Meeting was forwarded',
        BodyHtml: OUTLOOK_FORWARD_BODY_EN },
      { TemplateType: 'OutlookForwardNotification', Language: 'DE', Subject: 'FYI: Termin wurde weitergeleitet — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Termin wurde weitergeleitet',
        BodyHtml: OUTLOOK_FORWARD_BODY_DE },
      // v9.38: OutlookDeclineDigest — wird bei bestehenden Tenants nachgerüstet.
      { TemplateType: 'OutlookDeclineDigest', Language: 'EN', Subject: 'FYI: {{DeclineCount}} attendees declined Outlook — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: attendees declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_EN },
      { TemplateType: 'OutlookDeclineDigest', Language: 'DE', Subject: 'FYI: {{DeclineCount}} Teilnehmer haben Outlook abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: Teilnehmer haben den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_DE },
    ];

    let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
    try {
      const typeResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      );
      if (typeResp.ok) {
        const typeData = await typeResp.json();
        listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
      }
    } catch { /* Fallback */ }

    for (const t of newTemplates) {
      try {
        // Existiert das Template bereits? (TemplateType + Language)
        const checkResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '${t.TemplateType}' and Language eq '${t.Language}'&$top=1&$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (checkResp.ok) {
          const checkData = await checkResp.json();
          const items = checkData.value || checkData.d?.results || [];
          if (items.length > 0) continue; // Schon vorhanden - nicht ueberschreiben
        }
        // Template fehlt - nachlegen
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
          '__metadata': { 'type': listItemType },
          'Title': `${t.TemplateType}_${t.Language}`,
          'TemplateType': t.TemplateType,
          'Language': t.Language,
          'Subject': t.Subject,
          'HeadingColor': t.HeadingColor,
          'Heading': t.Heading,
          'BodyHtml': t.BodyHtml,
        });
      } catch { /* Einzelnen Fehler nicht kritisch */ }
    }
  }

  /**
   * Standard-Email-Templates auf die aktuelle Version aktualisieren.
   * Wird bei jedem App-Start aufgerufen, wenn die Liste schon existiert.
   *
   * !! ACHTUNG !! Ueberschreibt User-Customizing.
   *
   * Hintergrund: Templates wie 'Warteliste' wurden ueber die Zeit erweitert
   * (z.B. {{WaitlistPosition}}-Platzhalter). Aelter angelegte Tenants haben
   * noch die OOTB-Version ohne diese Felder. Diese Funktion zieht den BodyHtml
   * (sowie Subject + Heading) auf den aktuellen Code-Stand nach.
   */
  /**
   * v12.12: Öffentliche Re-Seed-Funktion für Admins. Stößt das Update aller
   * Standard-Templates an — überschreibt eventuelle individuelle Änderungen
   * in DEX_EmailTemplates mit den aktuellen Default-Texten aus dem Code.
   */
  public async reseedDefaultEmailTemplates(): Promise<ReseedSummary> {
    return this.upgradeStandardEmailTemplates('DEX_EmailTemplates');
  }

  private async upgradeStandardEmailTemplates(listName: string): Promise<ReseedSummary> {
    const summary: ReseedSummary = { created: 0, updated: 0, skipped: 0, failed: 0, errors: [] };
    const APP_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
    void APP_URL; // Reserviert für spätere Templates die {{AppUrl}} hardcoden
    const standards = [
      // ========== EN ==========
      { TemplateType: 'Anmeldung', Language: 'EN', Subject: 'Registration confirmation: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Registration successful',
        BodyHtml: '<p>Dear {{Name}},</p><p>you have successfully registered for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">DEX App</a> (\u201EMy Events\u201C).</p><p>For organizational questions about the event, please contact {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Waitlist confirmation',
        BodyHtml: '<p>Dear {{Name}},</p><p>you have been placed on the <strong>waitlist</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>Your current position: <strong>#{{WaitlistPosition}}</strong></p><p>We will notify you as soon as a spot becomes available. You can always check your current position in the <a href="{{AppUrl}}">DEX App</a> under \u201EMy Events\u201C.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
        BodyHtml: '<p>Dear {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>your registration for the event above has been <strong>cancelled</strong>. The Outlook calendar entry will be removed from your calendar shortly.</p><p>If you change your mind, you can register again via the <a href="{{AppUrl}}">DEX App</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'You’ve got a spot: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'You’ve got a spot!',
        BodyHtml: NACHRUECKEN_BODY_EN },
      { TemplateType: 'EventErstellt', Language: 'EN', Subject: '[Deloitte Eventmanager] - New event created: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event Created',
        BodyHtml: '<p>Dear {{Name}},</p><p>your event <strong>{{EventTitle}}</strong> has been successfully created.</p><p>You can manage participants in the <a href="{{AppUrl}}">DEX App</a>.</p><p>Regards,<br>Team DEX App</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_BODY_EN },
      // ========== DE ==========
      { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Anmeldung erfolgreich',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du hast dich erfolgreich für das Event <strong>{{EventTitle}}</strong> angemeldet.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">DEX App</a> (\u201EMeine Events\u201C) ab.</p><p>Zu organisatorischen Fragen zum Event wende dich bitte an {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Warteliste-Bestätigung',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du stehst auf der <strong>Warteliste</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Deine aktuelle Position: <strong>#{{WaitlistPosition}}</strong></p><p>Wir benachrichtigen dich, sobald ein Platz frei wird. Deinen aktuellen Warteliste-Platz kannst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter \u201EMeine Events\u201C sehen.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
        BodyHtml: '<p>Hallo {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>deine Anmeldung für das oben genannte Event wurde <strong>storniert</strong>. Der Outlook-Termin wird in Kürze aus deinem Kalender entfernt.</p><p>Du kannst dich jederzeit erneut über die <a href="{{AppUrl}}">DEX App</a> anmelden.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Du hast einen Platz: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Du hast einen Platz!',
        BodyHtml: NACHRUECKEN_BODY_DE },
      // v19.25: pre-wrapped Abmelde-Bestätigung für die Flow-getriebene
      // Auto-Abmeldung (DEX_OutlookDeclineHandler), eigener Type damit die
      // App-eigene `Abmeldung` (unwrapped) unberührt bleibt.
      { TemplateType: 'AbmeldungAuto', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
        BodyHtml: ABMELDUNG_AUTO_BODY_EN },
      { TemplateType: 'AbmeldungAuto', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
        BodyHtml: ABMELDUNG_AUTO_BODY_DE },
      // v18.63: Organizer-Benachrichtigung bei Abmeldung mit Nachrücker (vom DEX_IDReorder-Flow gequeued).
      { TemplateType: 'OrgNachruecker', Language: 'EN', Subject: 'Cancellation with waitlist move-up: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Cancellation — waitlist move-up',
        BodyHtml: ORG_NACHRUECKER_BODY_EN },
      { TemplateType: 'OrgNachruecker', Language: 'DE', Subject: 'Abmeldung mit Nachrücker: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Abmeldung — Nachrücker',
        BodyHtml: ORG_NACHRUECKER_BODY_DE },
      { TemplateType: 'EventErstellt', Language: 'DE', Subject: '[Deloitte Eventmanager] - Neues Event erstellt: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event erstellt',
        BodyHtml: '<p>Hallo {{Name}},</p><p>dein Event <strong>{{EventTitle}}</strong> wurde erfolgreich erstellt.</p><p>Du kannst die Teilnehmer in der <a href="{{AppUrl}}">DEX App</a> verwalten.</p><p>Viele Grüße,<br>Team DEX App</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'DE', Subject: 'Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Du hast den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_DE },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'DE', Subject: 'Action Required: Anmeldung für {{EventTitle}} stornieren?', HeadingColor: '#ed8b00', Heading: 'Outlook-Termin in deinem Namen abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_DE },
      { TemplateType: 'OutlookForwardNotification', Language: 'EN', Subject: 'FYI: Meeting was forwarded — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Meeting was forwarded',
        BodyHtml: OUTLOOK_FORWARD_BODY_EN },
      { TemplateType: 'OutlookForwardNotification', Language: 'DE', Subject: 'FYI: Termin wurde weitergeleitet — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Termin wurde weitergeleitet',
        BodyHtml: OUTLOOK_FORWARD_BODY_DE },
      // v9.38: OutlookDeclineDigest
      { TemplateType: 'OutlookDeclineDigest', Language: 'EN', Subject: 'FYI: {{DeclineCount}} attendees declined Outlook — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: attendees declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_EN },
      { TemplateType: 'OutlookDeclineDigest', Language: 'DE', Subject: 'FYI: {{DeclineCount}} Teilnehmer haben Outlook abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: Teilnehmer haben den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_DE },
      // v12.13: Team-Templates auch im Re-Seed-Pfad, sonst greift der Admin-
      // Reseed-Button die Texte nicht.
      { TemplateType: 'TeamMemberJoined', Language: 'EN', Subject: 'New team member — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team update',
        BodyHtml: TEAM_MEMBER_JOINED_BODY_EN },
      { TemplateType: 'TeamMemberJoined', Language: 'DE', Subject: 'Neues Team-Mitglied — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Update',
        BodyHtml: TEAM_MEMBER_JOINED_BODY_DE },
      { TemplateType: 'TeamJoinRequest', Language: 'EN', Subject: 'Team join request — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team join request',
        BodyHtml: TEAM_JOIN_REQUEST_BODY_EN },
      { TemplateType: 'TeamJoinRequest', Language: 'DE', Subject: 'Team-Beitritts-Anfrage — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Beitritts-Anfrage',
        BodyHtml: TEAM_JOIN_REQUEST_BODY_DE },
      { TemplateType: 'TeamJoinRejected', Language: 'EN', Subject: 'Team join request declined — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team join request declined',
        BodyHtml: TEAM_JOIN_REJECTED_BODY_EN },
      { TemplateType: 'TeamJoinRejected', Language: 'DE', Subject: 'Team-Beitritts-Anfrage abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Beitritts-Anfrage abgelehnt',
        BodyHtml: TEAM_JOIN_REJECTED_BODY_DE },
      { TemplateType: 'TeamLeadTransferred', Language: 'EN', Subject: 'Team lead change — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team lead change',
        BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_EN },
      { TemplateType: 'TeamLeadTransferred', Language: 'DE', Subject: 'Team-Lead-Wechsel — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Lead-Wechsel',
        BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_DE },
      { TemplateType: 'TeamMemberCancelled', Language: 'EN', Subject: 'Team update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team update',
        BodyHtml: TEAM_MEMBER_CANCELLED_BODY_EN },
      { TemplateType: 'TeamMemberCancelled', Language: 'DE', Subject: 'Team-Update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Update',
        BodyHtml: TEAM_MEMBER_CANCELLED_BODY_DE },
      // v13.0: Zimmerpartner, Gruppen-Wechsel, Überbuchung (vorher inline).
      { TemplateType: 'RoommateRequest', Language: 'EN', Subject: '{{RegistrantName}} selected you as roommate — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Roommate request',
        BodyHtml: ROOMMATE_REQUEST_BODY_EN },
      { TemplateType: 'RoommateRequest', Language: 'DE', Subject: '{{RegistrantName}} hat dich als Zimmerpartner gewählt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Zimmerpartner-Anfrage',
        BodyHtml: ROOMMATE_REQUEST_BODY_DE },
      { TemplateType: 'GroupSwitchConfirmed', Language: 'EN', Subject: 'Group switch confirmed — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Group switch',
        BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_EN },
      { TemplateType: 'GroupSwitchConfirmed', Language: 'DE', Subject: 'Gruppen-Wechsel bestätigt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Gruppen-Wechsel',
        BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_DE },
      { TemplateType: 'GroupSwitchWaitlist', Language: 'EN', Subject: 'Group switch — on waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Group switch — on waitlist',
        BodyHtml: GROUP_SWITCH_WAITLIST_BODY_EN },
      { TemplateType: 'GroupSwitchWaitlist', Language: 'DE', Subject: 'Gruppen-Wechsel — auf Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Gruppen-Wechsel — auf Warteliste',
        BodyHtml: GROUP_SWITCH_WAITLIST_BODY_DE },
      { TemplateType: 'OverbookingApology', Language: 'EN', Subject: 'Important: correction of your registration — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Registration corrected',
        BodyHtml: OVERBOOK_APOLOGY_BODY_EN },
      { TemplateType: 'OverbookingApology', Language: 'DE', Subject: 'Wichtig: Korrektur deiner Anmeldung — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Anmeldung korrigiert',
        BodyHtml: OVERBOOK_APOLOGY_BODY_DE },
    ];

    let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
    try {
      const typeResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      );
      if (typeResp.ok) {
        const typeData = await typeResp.json();
        listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
      }
    } catch { /* Fallback */ }

    for (const t of standards) {
      const label = `${t.TemplateType}_${t.Language}`;
      try {
        // Bestehendes Item finden
        const checkResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '${t.TemplateType}' and Language eq '${t.Language}'&$top=1&$select=Id,BodyHtml`,
          SPHttpClient.configurations.v1
        );
        if (!checkResp.ok) {
          summary.failed++;
          summary.errors.push(`${label}: Prüfung fehlgeschlagen (HTTP ${checkResp.status})`);
          continue;
        }
        const checkData = await checkResp.json();
        const items = checkData.value || checkData.d?.results || [];
        if (items.length === 0) {
          // existiert nicht -> anlegen (uebernimmt ensureMissingEmailTemplates für einige; hier sicherheitshalber auch)
          const postResp = await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
            '__metadata': { 'type': listItemType },
            'Title': label,
            'TemplateType': t.TemplateType,
            'Language': t.Language,
            'Subject': t.Subject,
            'HeadingColor': t.HeadingColor,
            'Heading': t.Heading,
            'BodyHtml': t.BodyHtml,
          });
          // v18.66: _post wirft NICHT bei HTTP 4xx/5xx — Status explizit prüfen,
          // sonst scheitern Inserts (z.B. neue Templates wie OrgNachruecker)
          // stillschweigend und der Reseed meldet fälschlich Erfolg.
          if (postResp.ok || postResp.status === 201 || postResp.status === 204) {
            summary.created++;
          } else {
            summary.failed++;
            let detail = '';
            try { detail = (await postResp.text()).slice(0, 200); } catch { /* ignore */ }
            summary.errors.push(`${label}: Anlegen fehlgeschlagen (HTTP ${postResp.status})${detail ? ' — ' + detail : ''}`);
          }
        } else {
          // existiert -> updaten falls BodyHtml vom Default abweicht
          const item = items[0];
          if (item.BodyHtml !== t.BodyHtml) {
            const mergeResp = await this._merge(
              `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${item.Id})`,
              {
                'Title': label,
                'TemplateType': t.TemplateType,
                'Language': t.Language,
                'Subject': t.Subject,
                'HeadingColor': t.HeadingColor,
                'Heading': t.Heading,
                'BodyHtml': t.BodyHtml,
              }
            );
            if (mergeResp.ok || mergeResp.status === 204) {
              summary.updated++;
            } else {
              summary.failed++;
              summary.errors.push(`${label}: Aktualisieren fehlgeschlagen (HTTP ${mergeResp.status})`);
            }
          } else {
            summary.skipped++;
          }
        }
      } catch (e) {
        summary.failed++;
        summary.errors.push(`${label}: ${e instanceof Error ? e.message : 'Unbekannter Fehler'}`);
      }
    }
    return summary;
  }

  /**
   * Wird aufgerufen wenn die Liste bereits existiert (nachträgliches Upgrade).
   */
  private async ensureEmailTemplatesConfig(listName: string): Promise<void> {
    try {
      // 1. Logo-Spalten nachträglich anlegen falls fehlend
      // v9.16: TestTeamEmails ergänzt — globale Liste (";"-separiert) der
      // User die Test-Events sehen + sich anmelden dürfen, auch wenn sie
      // keine Organizer/Admin-Rolle haben.
      const logoFields = [
        { title: 'LogoBase64', type: 3 },
        { title: 'DefaultImageBase64', type: 3 },
        { title: 'TestTeamEmails', type: 3 }, // Note (multi-line text), ";"-separiert
        // v11.47: App-Aufruf-Counter für die KPI-Boxen auf der LandingPage.
        // Wird pro Browser-Session genau einmal inkrementiert (Session-Guard
        // in LandingPage), ETag-CAS-Retry im incrementAppViewCount().
        { title: 'AppViewCount', type: 9 }, // Number
        // v11.52: gecachter Total-Teilnehmer-Counter für das LandingPage-KPI.
        // Live-Zählung ueber alle Event-Subsites war zu langsam — stattdessen
        // liest der Boot-Loader diesen einen Wert (schneller REST-Call), und
        // sobald die App fertig geladen hat, schreiben wir den frischen Wert
        // im Hintergrund zurück. Eventual consistency, für KPI-Anzeige ok.
        { title: 'TotalParticipantsCount', type: 9 }, // Number
        { title: 'TotalEventsCount', type: 9 }, // Number — analog für 'Events'
        // v15.17: Subheading-Spalte für die untere Headline-Zeile pro
        // Template (vorher hart als „Event {{EventTitle}}" im Code).
        // Leerwert = Fallback im Code auf {{EventTitle}} ohne Präfix.
        { title: 'Subheading', type: 2 }, // Single line text
      ];
      for (const f of logoFields) {
        try {
          const check = await this.context.spHttpClient.get(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${f.title}')`,
            SPHttpClient.configurations.v1
          );
          if (!check.ok) {
            await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
              '__metadata': { 'type': 'SP.Field' },
              'Title': f.title,
              'FieldTypeKind': f.type,
              'Required': false,
            });
          }
        } catch { /* Spalte existiert oder Fehler - ignorieren */ }
      }

      // 2. _Config Zeile prüfen und ggf. anlegen
      const configResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (configResp.ok) {
        const configData = await configResp.json();
        const items = configData.value || configData.d?.results || [];
        if (items.length === 0) {
          // _Config Zeile fehlt - anlegen
          let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
          try {
            const typeResp = await this.context.spHttpClient.get(
              `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
              SPHttpClient.configurations.v1
            );
            if (typeResp.ok) {
              const typeData = await typeResp.json();
              listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
            }
          } catch { /* Fallback */ }

          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
            '__metadata': { 'type': listItemType },
            'Title': '_Config',
            'TemplateType': '_Config',
          });
        }
      }
    } catch (err) { console.warn('[DEX] ensureEmailTemplatesConfig fehlgeschlagen:', err); }
  }

  /**
   * v11.52: Gecachte KPI-Werte (TotalParticipantsCount + TotalEventsCount)
   * aus der _Config-Zeile von DEX_EmailTemplates lesen. Ein einziger REST-
   * Call, kein Subsite-Roundtrip — Boot-Loader zeigt das innerhalb von ms.
   * Liefert null bei Fehler, sonst { participants, events } mit 0 als
   * Default für leere Felder.
   */
  public async getKpiCache(): Promise<{ participants: number; events: number } | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id,TotalParticipantsCount,TotalEventsCount`,
        SPHttpClient.configurations.v1,
      );
      if (!resp.ok) return null;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return { participants: 0, events: 0 };
      const it = items[0];
      const pRaw = it.TotalParticipantsCount;
      const eRaw = it.TotalEventsCount;
      const p = (pRaw === null || pRaw === undefined) ? 0 : (typeof pRaw === 'number' ? pRaw : (parseInt(String(pRaw), 10) || 0));
      const e = (eRaw === null || eRaw === undefined) ? 0 : (typeof eRaw === 'number' ? eRaw : (parseInt(String(eRaw), 10) || 0));
      return { participants: p, events: e };
    } catch { return null; }
  }

  /**
   * v11.53: KPI-Counter um delta hochzählen (Anmeldung +1, Cancel -1,
   * createEvent +1, deleteEvent -N). ETag-CAS-Retry, race-safe bei 10k+
   * parallelen Usern. Liefert den neuen Wert oder null bei Fehler.
   */
  public async bumpKpiParticipants(delta: number): Promise<number | null> {
    return this.bumpKpiField('TotalParticipantsCount', delta);
  }
  public async bumpKpiEvents(delta: number): Promise<number | null> {
    return this.bumpKpiField('TotalEventsCount', delta);
  }
  private async bumpKpiField(field: string, delta: number): Promise<number | null> {
    if (!Number.isFinite(delta) || delta === 0) return null;
    const itemUrl = await this.getConfigItemUrl();
    if (!itemUrl) return null;
    const MAX_RETRIES = 8;
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this.context.spHttpClient.get(itemUrl, SPHttpClient.configurations.v1);
      } catch { return null; }
      if (!getResp.ok) return null;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return null;
      let data;
      try { data = await getResp.json(); } catch { return null; }
      const raw = data?.[field] ?? data?.d?.[field];
      const current = (raw === null || raw === undefined)
        ? 0
        : (typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0));
      const next = Math.max(0, current + delta);
      const patchResp = await this._mergeIfMatch(itemUrl, { [field]: next }, etag);
      if (patchResp.ok) return next;
      if (patchResp.status !== 412) return null;
      await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
    }
    return null;
  }

  /**
   * v11.52: Gecachte KPI-Werte zurückschreiben. Wird nach vollem App-Load
   * im Hintergrund aufgerufen (DexEventPlatform), damit der nächste Boot-
   * Loader frische Zahlen sieht. Best-effort, kein ETag-CAS nötig — bei
   * gleichzeitigen Schreibern gewinnt der letzte, was für KPI ok ist.
   */
  public async updateKpiCache(values: { participants: number; events: number }): Promise<boolean> {
    const itemUrl = await this.getConfigItemUrl();
    if (!itemUrl) return false;
    try {
      const resp = await this._mergeIfMatch(itemUrl, {
        'TotalParticipantsCount': Math.max(0, Math.floor(values.participants || 0)),
        'TotalEventsCount': Math.max(0, Math.floor(values.events || 0)),
      }, '*');
      return resp.ok;
    } catch { return false; }
  }

  /**
   * v11.50: Anzahl Items in DEX_Participants (= unique User, die jemals für
   * irgendein Event angemeldet/auf Warteliste waren). Liest nur das ItemCount-
   * Metadatum der Liste, nicht alle Items — schnell und cheap. Liefert null
   * bei Fehler.
   */
  public async getParticipantsListCount(): Promise<number | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')?$select=ItemCount`,
        SPHttpClient.configurations.v1,
      );
      if (!resp.ok) return null;
      const data = await resp.json();
      const raw = data?.ItemCount ?? data?.d?.ItemCount;
      if (raw === null || raw === undefined) return null;
      return typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0);
    } catch { return null; }
  }

  /**
   * v11.47: Aktuellen App-Aufruf-Counter aus der _Config-Zeile von
   * DEX_EmailTemplates lesen. Liefert 0 wenn das Feld leer / nicht
   * vorhanden ist. null bei Lese-Fehler.
   */
  public async getAppViewCount(): Promise<number | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id,AppViewCount`,
        SPHttpClient.configurations.v1,
      );
      if (!resp.ok) return null;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return 0;
      const raw = items[0].AppViewCount;
      if (raw === null || raw === undefined) return 0;
      return typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0);
    } catch { return null; }
  }

  /**
   * v11.47: App-Aufruf-Counter um 1 inkrementieren — ETag-CAS-Retry analog
   * zum reserveSeat-Muster. Liefert den neuen Wert nach Inkrement, oder
   * null bei Fehler / Retry-Erschöpfung.
   */
  public async incrementAppViewCount(): Promise<number | null> {
    const itemUrl = await this.getConfigItemUrl();
    if (!itemUrl) return null;
    const MAX_RETRIES = 8;
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this.context.spHttpClient.get(itemUrl, SPHttpClient.configurations.v1);
      } catch { return null; }
      if (!getResp.ok) return null;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return null;
      let data;
      try { data = await getResp.json(); } catch { return null; }
      const raw = data?.AppViewCount ?? data?.d?.AppViewCount;
      const current = (raw === null || raw === undefined)
        ? 0
        : (typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0));
      const next = current + 1;
      const patchResp = await this._mergeIfMatch(itemUrl, { 'AppViewCount': next }, etag);
      if (patchResp.ok) return next;
      if (patchResp.status !== 412) return null;
      // 412 = stale ETag → kurzer Backoff + neu lesen
      await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
    }
    return null;
  }

  /**
   * v11.47: Helper — URL des _Config-Items in DEX_EmailTemplates ermitteln.
   * Liefert null, wenn die Liste/Zeile noch nicht existiert.
   */
  private async getConfigItemUrl(): Promise<string | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1,
      );
      if (!resp.ok) return null;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return null;
      const id = items[0].Id;
      return `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${id})`;
    } catch { return null; }
  }

  /**
   * Logos als Base64 in die _Config Zeile schreiben (für Power Automate Flows).
   * Lädt Deloitte_Logo.png und dex-orb.png aus SiteAssets/DEX_Logos,
   * konvertiert zu Base64 Data-URI und speichert in LogoBase64/DefaultImageBase64.
   */
  public async ensureLogosInConfig(): Promise<void> {
    try {
      // 1. _Config Zeile lesen
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      const configItem = items[0];
      if (!configItem) return;

      // 2. Wenn LogoBase64 schon korrekt befüllt ist (mit image/ MIME-Type), nichts tun
      if (configItem.LogoBase64 && configItem.LogoBase64.startsWith('data:image/')) return;

      // 3. Bilder aus SiteAssets laden
      const logoBase64 = await this.loadFileAsBase64('DEX_Logos/Deloitte_Logo.png');
      const orbBase64 = await this.loadFileAsBase64('DEX_Logos/dex-orb.png');
      if (!logoBase64 && !orbBase64) return;

      // 4. In _Config Zeile schreiben (ueber die getestete _post/_merge Methode)
      const configId = configItem.Id || configItem.d?.Id;
      if (!configId) return;

      const updatePayload: Record<string, unknown> = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailTemplatesListItem' },
      };
      if (logoBase64) updatePayload['LogoBase64'] = logoBase64;
      if (orbBase64) updatePayload['DefaultImageBase64'] = orbBase64;

      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${configId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            'odata-version': '',
          },
          body: JSON.stringify(updatePayload),
        }
      );
    } catch (err) {
      console.warn('[DEX] ensureLogosInConfig fehlgeschlagen:', err);
    }
  }

  /**
   * Datei aus SiteAssets als Base64 Data-URI laden.
   */
  private async loadFileAsBase64(path: string): Promise<string> {
    try {
      const serverRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
      const fileUrl = `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}/SiteAssets/${path}')/$value`;

      // SPHttpClient mit binaryStringResponseBody für Binary-Downloads
      const resp = await this.context.spHttpClient.get(fileUrl, SPHttpClient.configurations.v1, {
        headers: { 'Accept': '*/*' },
      } as ISPHttpClientOptions);
      if (!resp.ok) {
        console.warn('[DEX] loadFileAsBase64 fehlgeschlagen:', path, resp.status);
        return '';
      }
      const blob = await resp.blob();
      if (!blob || blob.size === 0) return '';
      // MIME-Type aus Dateiendung ableiten (SPHttpClient gibt oft application/octet-stream)
      const ext = path.split('.').pop()?.toLowerCase();
      const mimeType = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : blob.type;
      const correctBlob = (blob.type !== mimeType) ? new Blob([blob], { type: mimeType }) : blob;
      return await new Promise<string>((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string || '');
        reader.onerror = () => resolve('');
        reader.readAsDataURL(correctBlob);
      });
    } catch (err) {
      console.warn('[DEX] loadFileAsBase64 Error:', path, err);
      return '';
    }
  }

  /**
   * Email-Template aus DEX_EmailTemplates laden.
   * Fallback auf eingebautes Template wenn nicht gefunden.
   */
  // v9.16: Test-Team — globale ";"-separierte E-Mail-Liste, gespeichert auf
  // dem _Config-Eintrag der DEX_EmailTemplates-Liste. Erlaubt nicht-Admin/
  // -Organizer-Usern Test-Events zu sehen + sich anzumelden.
  public async getTestTeamEmails(): Promise<string[]> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=TestTeamEmails`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return [];
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return [];
      const raw: string = (items[0].TestTeamEmails || '').toString();
      return raw.split(/[;,\n]/).map(s => s.trim().toLowerCase()).filter(s => !!s && s.includes('@'));
    } catch { return []; }
  }

  public async setTestTeamEmails(emails: string[]): Promise<boolean> {
    try {
      const cleaned = (emails || []).map(s => (s || '').trim()).filter(s => !!s && s.includes('@'));
      const value = cleaned.join(';');
      // _Config-Item-ID lookup
      const lookup = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (!lookup.ok) return false;
      const data = await lookup.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return false;
      const itemId = items[0].Id;
      const resp = await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${itemId})`,
        { 'TestTeamEmails': value }
      );
      return resp.ok;
    } catch { return false; }
  }

  public async getEmailTemplate(templateType: string, language: string = 'EN'): Promise<{ subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string } | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '${templateType.replace(/'/g, "''")}' and Language eq '${language.replace(/'/g, "''")}'&$select=Subject,HeadingColor,Heading,Subheading,BodyHtml&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0) {
          return {
            subject: items[0].Subject || '',
            headingColor: items[0].HeadingColor || '#86bc25',
            heading: items[0].Heading || '',
            // v15.17: Subheading (untere Headline-Zeile, vorher hart als
            // „Event {{EventTitle}}" geschrieben) jetzt aus dem Template.
            // Leer = Fallback auf reinen EventTitle ohne „Event "-Präfix.
            subheading: items[0].Subheading || '',
            bodyHtml: items[0].BodyHtml || '',
          };
        }
      }
    } catch { /* */ }
    return null;
  }

  /**
   * Alle Email-Templates laden (für Event-Erstellung / Admin).
   */
  public async getAllEmailTemplates(): Promise<Array<{ id: number; templateType: string; language: string; subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string }>> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$select=Id,TemplateType,Language,Subject,HeadingColor,Heading,Subheading,BodyHtml&$orderby=TemplateType,Language&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (data.value || data.d?.results || []).map((item: any) => ({
          id: item.Id,
          templateType: item.TemplateType || '',
          language: item.Language || 'EN',
          subject: item.Subject || '',
          headingColor: item.HeadingColor || '#86bc25',
          heading: item.Heading || '',
          subheading: item.Subheading || '',
          bodyHtml: item.BodyHtml || '',
        }));
      }
    } catch { /* */ }
    return [];
  }

  // ==================== DEX_Participants Liste ====================

  /**
   * Zentrale Teilnehmer-Liste erstellen falls nicht vorhanden.
   * Speichert pro Person die EventNumbers für Registrierung und Warteliste.
   */
  public async ensureParticipantsList(): Promise<void> {
    const listName = 'DEX_Participants';
    const exists = await this.listExists(listName);
    if (exists) {
      await this.ensureMissingParticipantsFields(listName);
      await this.configureDefaultView(listName, [
        'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
      ]);
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Zentrale Teilnehmerliste der DEX Event Experience Platform',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields = [
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'Email', type: 2 },
      { title: 'EventRegistered', type: 3 }, // Note für beliebig viele EventNumbers
      { title: 'EventOnWaitlist', type: 3 }, // Note für beliebig viele EventNumbers
    ];

    for (const f of fields) {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    }

    await this.configureDefaultView(listName, [
      'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
    ]);

    await this.setEmailsListPermissions(listName);
  }

  private async ensureMissingParticipantsFields(listName: string): Promise<void> {
    const requiredFields = [
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'Email', type: 2 },
      { title: 'EventRegistered', type: 3 },
      { title: 'EventOnWaitlist', type: 3 },
    ];

    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title&$filter=Hidden eq false&$top=200`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;
      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const existingNames = new Set((data.value || []).flatMap((f: any) => [f.InternalName, f.Title]));

      for (const f of requiredFields) {
        if (!existingNames.has(f.title)) {
          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
            '__metadata': { 'type': 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          });
        }
      }
    } catch { /* optional */ }
  }

  /**
   * Teilnehmer-Eintrag per Email suchen
   */
  public async getParticipantByEmail(email: string): Promise<SPParticipant | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$filter=Email eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      if (data.value && data.value.length > 0) return data.value[0];
      return null;
    } catch {
      return null;
    }
  }

  /**
   * Teilnehmer anlegen oder aktualisieren bei Registrierung.
   * Fügt eventNumber zu EventRegistered oder EventOnWaitlist hinzu.
   */
  public async upsertParticipant(
    vorname: string,
    nachname: string,
    email: string,
    eventNumber: number,
    status: string // 'Angemeldet' | 'Warteliste'
  ): Promise<boolean> {
    try {
      const existing = await this.getParticipantByEmail(email);

      if (existing) {
        // EventNumber zu richtigem Feld hinzufügen
        let registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s) : [];
        let waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s) : [];
        const en = eventNumber.toString();

        // Erst aus beiden entfernen
        registered = registered.filter(n => n !== en);
        waitlist = waitlist.filter(n => n !== en);

        if (status === 'Warteliste') {
          waitlist.push(en);
        } else {
          registered.push(en);
        }

        await this._merge(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
          {
            'Vorname': vorname,
            'Nachname': nachname,
            'EventRegistered': registered.join(','),
            'EventOnWaitlist': waitlist.join(','),
          }
        );
      } else {
        // Neuen Eintrag erstellen
        const isWaitlist = status === 'Warteliste';
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items`, {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_ParticipantsListItem' },
          'Title': email,
          'Vorname': vorname,
          'Nachname': nachname,
          'Email': email,
          'EventRegistered': isWaitlist ? '' : eventNumber.toString(),
          'EventOnWaitlist': isWaitlist ? eventNumber.toString() : '',
        });
      }
      return true;
    } catch {
      return false;
    }
  }

  /**
   * EventNumber aus den Feldern eines Teilnehmers entfernen (bei Abmeldung).
   */
  public async removeParticipantEvent(email: string, eventNumber: number): Promise<boolean> {
    try {
      const existing = await this.getParticipantByEmail(email);
      if (!existing) return false;

      const en = eventNumber.toString();
      const registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];
      const waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];

      await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
        {
          'EventRegistered': registered.join(','),
          'EventOnWaitlist': waitlist.join(','),
        }
      );
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Alle Teilnehmer laden (für Admin-Seite).
   */
  public async getAllParticipants(): Promise<SPParticipant[]> {
    const allItems: SPParticipant[] = [];
    // $top=5000 statt 500 — bei mehr als 500 jemals registrierten Personen
    // im Tenant lieferte SharePoint mit $orderby+$top in Kombination mit
    // ILS nicht zuverlässig nextLink, sodass Einträge fehlten.
    let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist&$orderby=Nachname,Vorname&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    return allItems;
  }

  // ==================== DEX_Events Liste ====================

  /**
   * Events-Liste erstellen falls nicht vorhanden
   */
  public async ensureEventsList(): Promise<void> {
    const listName = 'DEX_Events';
    const exists = await this.listExists(listName);
    if (exists) {
      await this.ensureMissingFields(listName);

      // Default-View komplett neu aufbauen: ID, Title, EventImageUrl, dann Rest
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
          {}
        );
      } catch { /* ignore */ }
      await this.configureDefaultView(listName, [
        'ID', 'LinkTitle', 'EventImageUrl',
        'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
        'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
        'WaitlistEnabled', 'Organizer', 'DisableEmails', 'DisableOutlook',
        'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
      ], undefined, { rebuild: true });
      await this.setColumnFormatting(listName, 'EventImageUrl', {
        '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        'elmType': 'img',
        'attributes': { 'src': '@currentField' },
        'style': { 'max-height': '60px', 'max-width': '120px', 'border-radius': '6px', 'box-shadow': '0 1px 3px rgba(0,0,0,0.15)' },
      });
      try {
        const listInfo = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
          SPHttpClient.configurations.v1
        );
        if (listInfo.ok) {
          const data = await listInfo.json();
          if (!data.HasUniqueRoleAssignments) {
            await this.setEventsListPermissions(listName);
          }
        }
      } catch { /* ignore */ }
      return;
    }

    // Liste erstellen
    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Events der DEX Event Experience Platform',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // Spalten hinzufügen
    const fields = this.getEventsFieldDefinitions();
    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        payload['Choices'] = { 'results': f.choices };
      }
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    await this.configureDefaultView(listName, [
      'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
      'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
      'WaitlistEnabled', 'Organizer', 'EventImageUrl', 'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
    ], undefined, { rebuild: true });
    await this.setColumnFormatting(listName, 'EventImageUrl', {
      '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
      'elmType': 'img',
      'attributes': { 'src': '@currentField' },
      'style': { 'max-height': '50px', 'max-width': '120px', 'border-radius': '4px' },
    });

    await this.setEventsListPermissions(listName);
  }

  /**
   * Feld-Definitionen für DEX_Events Liste
   */
  private getEventsFieldDefinitions(): Array<{ title: string; type: number; choices?: string[]; metaType?: string; richText?: boolean; numberOfLines?: number }> {
    return [
      { title: 'EventStatus', type: 6, choices: ['Under Construction', 'Active', 'Completed', 'Cancelled'], metaType: 'SP.FieldChoice' },
      // EventType-Spalte ab v5.2 deprecated (Feld wird nicht mehr angelegt/aktualisiert).
      // Typ wird beim Laden aus CustomFields abgeleitet. Bestehende Spalte in DEX_Events
      // kann manuell entfernt werden.
      { title: 'Description', type: 3 },
      { title: 'Location', type: 2 },
      { title: 'LocationAddress', type: 2 }, // JSON-String: { street, houseNo, zip, city }
      { title: 'LocationFilter', type: 2 },
      // Audience ist Multi-Line-Text (Note) damit es auch bei 100+ E-Mail-Adressen
      // nicht abgeschnitten wird (Single-Line-Text ist auf 255 Zeichen limitiert).
      // Für bestehende Events siehe upgradeAudienceFieldToNote() — migriert die
      // alte Text-Spalte zu Note ohne Datenverlust.
      { title: 'Audience', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      // v16.4: Vor-aufgelöste E-Mails der Audience-DLs (Multi-Line, ';'-
      // separiert). Wird beim Event-Save vom EventCreationPage-Flow
      // gesetzt; matchesAudience im EventListPage checkt zusätzlich
      // gegen diese Liste. Damit funktioniert die Sichtbarkeit auch für
      // verschachtelte DLs, die /me/memberOf nicht zurückliefert.
      { title: 'AudienceResolvedEmails', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 8 },
      { title: 'FilterMode', type: 6, choices: ['AND', 'OR'], metaType: 'SP.FieldChoice' },
      { title: 'StartDate', type: 4 },
      { title: 'EndDate', type: 4 },
      { title: 'RegistrationDeadline', type: 4 },
      { title: 'LastDeregisterDate', type: 4 },
      { title: 'MaxParticipants', type: 9 },
      { title: 'WaitlistEnabled', type: 8 },
      { title: 'EventImageUrl', type: 2 },
      { title: 'EmailImageBase64', type: 3 }, // Base64 Event-Bild für E-Mails/Outlook (Flow ersetzt {{ORB_URL}})
      // Organizer + OrganizerEmail sind Multi-Line-Text (Note) damit sie auch bei
      // 10+ Co-Organizern nicht abgeschnitten werden (Single-Line-Text ist auf 255
      // Zeichen limitiert — bei ~17 Personen mit Format `vorname.nachname@deloitte.de;`
      // wird das überschritten und SP antwortet mit „Invalid text value" beim Update).
      // Für bestehende Events siehe upgradeOrganizerFieldsToNote() — migriert die
      // alten Text-Spalten ohne Datenverlust.
      { title: 'Organizer', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      { title: 'OrganizerEmail', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      // v10.16: Optionaler Ansprechpartner pro Event (Anzeige-Feld, kein App-Login).
      { title: 'ContactName', type: 2 },
      { title: 'ContactEmail', type: 2 },
      { title: 'ContactInfo', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
      { title: 'EventNumber', type: 9 },
      { title: 'OutlookEventId', type: 2 },
      { title: 'CalendarLink', type: 2 },
      { title: 'OutlookBody', type: 3 }, // Multiline - Text für Outlook-Termin
      { title: 'OutlookSubject', type: 2 }, // v18.42: Single line - Betreff des Outlook-Termins (leer = Titel)
      { title: 'OutlookStart', type: 4 }, // v18.44: DateTime - abweichende Start-Zeit (leer = Event-Start)
      { title: 'OutlookEnd', type: 4 },   // v18.44: DateTime - abweichende End-Zeit (leer = Event-Ende)
      { title: 'OutlookLocation', type: 2 }, // v18.34: Single line - lesbarer Ort für den Outlook-Termin
      { title: 'EmailLanguage', type: 2 }, // DE oder EN
      { title: 'RegistrationLanguage', type: 2 }, // v18.35: erzwungene Anmeldeseiten-Sprache ('de'|'en'|'')
      { title: 'EmailTemplateOverrides', type: 3 }, // JSON mit Event-spezifischen Template-Anpassungen
      { title: 'DisableEmails', type: 8, metaType: 'SP.Field' }, // Boolean - keine E-Mails versenden
      { title: 'DisableRegistrationEmail', type: 8, metaType: 'SP.Field' }, // v19.21 Boolean - keine Anmelde-Bestätigung
      { title: 'DisableCancellationEmail', type: 8, metaType: 'SP.Field' }, // v19.21 Boolean - keine Abmelde-Bestätigung
      { title: 'AutoDeregisterOnDecline', type: 8, metaType: 'SP.Field' }, // v19.23 Boolean - Outlook-Absage = Auto-Abmeldung
      { title: 'DisableOutlook', type: 8, metaType: 'SP.Field' }, // Boolean - keine Outlook-Kalendereintraege
      { title: 'OutlookDirty', type: 8, metaType: 'SP.Field' }, // v11.57 Boolean - Outlook-Update ausstehend nach Bearbeitung
      { title: 'AutoSendQRCode', type: 8, metaType: 'SP.Field' }, // v9.15 Boolean - QR-Code automatisch nach Anmeldung versenden
      { title: 'ActiveFrom', type: 4, metaType: 'SP.Field' }, // v9.21 DateTime - Auto-Aktivierungs-Datum
      { title: 'NotifyOrgRegisterMode', type: 6, choices: ['Never', 'Always', 'FromDate'], metaType: 'SP.FieldChoice' }, // v8.5
      { title: 'NotifyOrgRegisterFromDate', type: 4 }, // v8.5: ISO-Date, nur für Mode='FromDate'
      { title: 'NotifyOrgCancelMode', type: 6, choices: ['Never', 'Always', 'AfterDeadline'], metaType: 'SP.FieldChoice' }, // v8.5
      { title: 'ExcludedUsers', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 }, // v8.6: explizit ausgeschlossene User
      { title: 'IsFictive', type: 8, metaType: 'SP.Field' }, // Boolean - Test-Event (nur Admin + eigene Organizer sichtbar)
      { title: 'DurchstarterCapacity', type: 9 }, // Split-Capacity Gruppe A (historisch B2Run-Durchstarter)
      { title: 'FunstarterCapacity', type: 9 }, // Split-Capacity Gruppe B (historisch B2Run-Funstarter)
      { title: 'SplitLabelA', type: 2 }, // v10.20: frei wählbare Bezeichnung Gruppe A (Single line text)
      { title: 'SplitLabelB', type: 2 }, // v10.20: frei wählbare Bezeichnung Gruppe B (Single line text)
      { title: 'SplitSharedWaitlist', type: 8, metaType: 'SP.Field' }, // v10.20: Boolean - true = gemeinsame Warteliste
      { title: 'AllowAttendeeUpload', type: 8, metaType: 'SP.Field' }, // v11.0: Boolean - Teilnehmer-PDF-Upload erlauben
      { title: 'AttendeeUploadHint', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 }, // v11.0: Hinweistext
      { title: 'AttendeeUploadLabel', type: 2 }, // v11.0: Single-line Label für den Upload-Block in MyEvents
      { title: 'AskSalutation', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Anrede im Registrierungsformular abfragen
      { title: 'ConfirmDialogEnabled', type: 8, metaType: 'SP.Field' }, // v18.75: Boolean - Sicherheitshinweis vor dem Absenden anzeigen
      { title: 'ConfirmDialogMode', type: 2 }, // v18.75: Single line text - 'summary' (Auswahl-Übersicht) | 'freetext'
      { title: 'ConfirmDialogText', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 }, // v18.75: Note - Freitext-Hinweis
      { title: 'SelfCheckInEnabled', type: 8, metaType: 'SP.Field' }, // v18.33: Boolean - Self-Check-in per QR-Code erlauben
      { title: 'SelfCheckInToken', type: 2 }, // v18.33: Single line text - geheimer Token (statischer Link + HMAC-Schlüssel)
      { title: 'SelfCheckInFrom', type: 4 }, // v18.33: DateTime - optionaler Start des Check-in-Fensters
      { title: 'SelfCheckInTo', type: 4 }, // v18.33: DateTime - optionales Ende des Check-in-Fensters
      { title: 'TeamRegistrationEnabled', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Team-Anmeldung erlauben
      { title: 'TeamSize', type: 9 }, // v11.80: Number - Maximale Teamgröße (0 = nicht gesetzt)
      { title: 'AskTeamName', type: 8, metaType: 'SP.Field' }, // v11.80: Boolean - Team-Name abfragen
      { title: 'TeamPartialAllowed', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - Auch Teil-Teams erlauben
      { title: 'TeamOpenSlotsVisible', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - offene Slots öffentlich sichtbar
      { title: 'TeamJoinRequiresApproval', type: 8, metaType: 'SP.Field' }, // v11.81: Boolean - Lead muss Beitritt bestätigen
      { title: 'BilingualFields', type: 8, metaType: 'SP.Field' }, // v17.20: Boolean - Custom-Fields zweisprachig (DE + EN) anbieten
      { title: 'CustomFields', type: 3 },
      { title: 'Agenda', type: 3 }, // JSON-Array mit Agenda-Einträgen
      { title: 'Transfers', type: 3 }, // JSON-Array mit Transferzeiten
      { title: 'Documents', type: 3 }, // JSON-Array mit Dokumenten
      { title: 'FunZone', type: 3 }, // JSON-Array mit Quiz-Fragen
      { title: 'QuizClusterSize', type: 9 }, // Number - 1..4 Fragen pro Quiz-Ansicht
      { title: 'ParentEventId', type: 2 }, // Seit v6.4: ID des Parent-Events (wenn dies ein Sub-Event ist)
      { title: 'RegistrationListName', type: 2 },
      { title: 'RegistrationListUrl', type: 2 },
      { title: 'SubsiteUrl', type: 2 },
    ];
  }

  /**
   * Fehlende Spalten auf einer bestehenden DEX_Events-Liste nachträglich hinzufügen.
   */
  private async ensureMissingFields(listName: string): Promise<void> {
    const requiredFields = this.getEventsFieldDefinitions();

    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$filter=Hidden eq false&$top=200`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;

      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const existingFields = new Set((data.value || []).map((f: any) => f.InternalName));

      for (const f of requiredFields) {
        if (!existingFields.has(f.title)) {
          // Fehlende Spalte nachträglich hinzufügen
          const payload: Record<string, unknown> = {
            '__metadata': { 'type': f.metaType || 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          };
          if (f.choices) {
            payload['Choices'] = { 'results': f.choices };
          }
          if (f.metaType === 'SP.FieldMultiLineText') {
            payload['RichText'] = !!f.richText;
            if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
          }
          try {
            await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
          } catch {
            console.warn('[DEX] Konnte Spalte nicht hinzufügen:', f.title);
          }
        }
      }
    } catch (e) {
      console.warn('[DEX] ensureMissingFields Error:', e);
    }
  }

  /**
   * Migration: alte Audience-Spalte (Type 2, Single-Line-Text, 255 Zeichen Limit)
   * auf Multi-Line-Text (Type 3, Plain-Text, 63.999 Zeichen) umstellen.
   *
   * Nötig weil bei Zielgruppen mit vielen Email-Adressen (~10+) der 255-Zeichen-
   * Cutoff schon griff und Adressen stumm abgeschnitten wurden.
   *
   * Ablauf (idempotent):
   *   1. Check TypeAsString der Audience-Spalte. Wenn schon 'Note' -> skip.
   *   2. Backup aller Event-Werte (id -> audience) im Speicher.
   *   3. Alte Spalte löschen.
   *   4. Neue Spalte als SP.FieldMultiLineText (RichText=false, NumberOfLines=4) anlegen.
   *   5. Werte aus dem Backup zurückschreiben (MERGE pro Event).
   *
   * Läuft beim App-Start (nur für Admins, weil wir Write-Rechte auf DEX_Events brauchen).
   */
  public async upgradeAudienceFieldToNote(): Promise<void> {
    const listName = 'DEX_Events';
    try {
      // 1. TypeAsString abfragen
      const fieldResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')?$select=TypeAsString,FieldTypeKind`,
        SPHttpClient.configurations.v1
      );
      if (!fieldResp.ok) return;
      const fieldData = await fieldResp.json();
      const typeAsString: string = fieldData.TypeAsString || fieldData.d?.TypeAsString || '';
      const fieldTypeKind: number = fieldData.FieldTypeKind ?? fieldData.d?.FieldTypeKind ?? 0;
      if (typeAsString === 'Note' || fieldTypeKind === 3) {
        // Schon migriert
        return;
      }
      if (typeAsString !== 'Text' && fieldTypeKind !== 2) {
        // Unerwarteter Typ - nicht anfassen
        console.warn(`[DEX] upgradeAudienceFieldToNote: Audience hat unerwarteten Typ '${typeAsString}' (kind=${fieldTypeKind}) - skip.`);
        return;
      }

      // 2. Alle Event-Werte laden und backuppen
      const itemsResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Audience&$top=2000`,
        SPHttpClient.configurations.v1
      );
      if (!itemsResp.ok) return;
      const itemsData = await itemsResp.json();
      const items: Array<{ Id: number; Audience: string }> = itemsData.value || itemsData.d?.results || [];
      const backup: Record<number, string> = {};
      for (const it of items) {
        if (it.Audience) backup[it.Id] = it.Audience;
      }
      console.warn(`[DEX] upgradeAudienceFieldToNote: Backup ${Object.keys(backup).length} von ${items.length} Event-Audience-Werten.`);

      // 3. Alte Spalte löschen
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')/deleteObject`,
          {}
        );
      } catch (e) {
        console.warn('[DEX] upgradeAudienceFieldToNote: Delete alte Audience-Spalte fehlgeschlagen, Migration abgebrochen:', e);
        return;
      }

      // 4. Neue Spalte als Multi-Line-Text anlegen
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          {
            '__metadata': { 'type': 'SP.FieldMultiLineText' },
            'Title': 'Audience',
            'FieldTypeKind': 3,
            'Required': false,
            'RichText': false,
            'NumberOfLines': 4,
          }
        );
      } catch (e) {
        console.error('[DEX] upgradeAudienceFieldToNote: Konnte neue Audience-Note-Spalte nicht anlegen - Daten könnten verloren gehen:', e, backup);
        return;
      }

      // 5. Werte zurückschreiben per _merge (odata=nometadata, daher kein __metadata im Body nötig)
      let restored = 0;
      let failed = 0;
      for (const idStr of Object.keys(backup)) {
        const id = Number(idStr);
        try {
          const resp = await this._merge(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
            { 'Audience': backup[id] }
          );
          if (resp.ok) restored += 1;
          else failed += 1;
        } catch { failed += 1; }
      }
      console.warn(`[DEX] upgradeAudienceFieldToNote: Migration fertig — ${restored} Werte zurückgeschrieben, ${failed} Fehler.`);
    } catch (e) {
      console.warn('[DEX] upgradeAudienceFieldToNote Error:', e);
    }
  }

  /**
   * Migration: alte `Organizer` + `OrganizerEmail`-Spalten (Type 2, Single-Line-Text,
   * 255 Zeichen Limit) auf Multi-Line-Text (Type 3, Plain-Text, 63.999 Zeichen) umstellen.
   *
   * Nötig weil bei Events mit 10+ Co-Organizern der 255-Zeichen-Cutoff greift und
   * SharePoint beim Update mit „Invalid text value. A text field contains invalid data."
   * (HTTP 500, Microsoft.SharePoint.SPException) antwortet — der Save bricht komplett ab.
   *
   * Beispiel-Overflow: 17 × `vorname.nachname@deloitte.de;` ≈ 425 Zeichen.
   *
   * Ablauf pro Feld (idempotent, parallel für beide Felder):
   *   1. Check TypeAsString. Wenn schon 'Note' -> skip.
   *   2. Backup aller Event-Werte (id -> wert) im Speicher.
   *   3. Alte Spalte löschen.
   *   4. Neue Spalte als SP.FieldMultiLineText (RichText=false, NumberOfLines=4) anlegen.
   *   5. Werte aus dem Backup zurückschreiben (MERGE pro Event).
   *
   * Läuft beim App-Start (nur für Admins, weil wir Write-Rechte auf DEX_Events brauchen).
   */
  public async upgradeOrganizerFieldsToNote(): Promise<void> {
    await this._upgradeTextFieldToNote('DEX_Events', 'Organizer');
    await this._upgradeTextFieldToNote('DEX_Events', 'OrganizerEmail');
  }

  /**
   * Generischer Helper: migriert ein einzelnes Single-Line-Text-Feld einer Liste auf
   * Multi-Line-Text (Note). Idempotent — wenn das Feld schon Note ist, no-op. Wenn das
   * Feld einen anderen Typ hat (Choice/Number/etc.), no-op mit Warnung.
   *
   * Wird von `upgradeAudienceFieldToNote()` (existierendes Audience-Feld) und
   * `upgradeOrganizerFieldsToNote()` (Organizer + OrganizerEmail) genutzt. Die alte
   * `upgradeAudienceFieldToNote()`-Implementierung ist aus Kompatibilitätsgründen
   * unberührt geblieben — neue Migrationen sollten diesen Helper nutzen.
   */
  private async _upgradeTextFieldToNote(listName: string, fieldName: string): Promise<void> {
    const tag = `_upgradeTextFieldToNote(${listName}.${fieldName})`;
    try {
      const fieldResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${fieldName}')?$select=TypeAsString,FieldTypeKind`,
        SPHttpClient.configurations.v1
      );
      if (!fieldResp.ok) return;
      const fieldData = await fieldResp.json();
      const typeAsString: string = fieldData.TypeAsString || fieldData.d?.TypeAsString || '';
      const fieldTypeKind: number = fieldData.FieldTypeKind ?? fieldData.d?.FieldTypeKind ?? 0;
      if (typeAsString === 'Note' || fieldTypeKind === 3) return;
      if (typeAsString !== 'Text' && fieldTypeKind !== 2) {
        console.warn(`[DEX] ${tag}: unerwarteter Typ '${typeAsString}' (kind=${fieldTypeKind}) — skip.`);
        return;
      }

      const itemsResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,${fieldName}&$top=2000`,
        SPHttpClient.configurations.v1
      );
      if (!itemsResp.ok) return;
      const itemsData = await itemsResp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const items: Array<any> = itemsData.value || itemsData.d?.results || [];
      const backup: Record<number, string> = {};
      for (const it of items) {
        const v = it[fieldName];
        if (v) backup[it.Id] = v;
      }
      console.warn(`[DEX] ${tag}: Backup ${Object.keys(backup).length} von ${items.length} Werten.`);

      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${fieldName}')/deleteObject`,
          {}
        );
      } catch (e) {
        console.warn(`[DEX] ${tag}: Delete alte Spalte fehlgeschlagen, Migration abgebrochen:`, e);
        return;
      }

      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          {
            '__metadata': { 'type': 'SP.FieldMultiLineText' },
            'Title': fieldName,
            'FieldTypeKind': 3,
            'Required': false,
            'RichText': false,
            'NumberOfLines': 4,
          }
        );
      } catch (e) {
        console.error(`[DEX] ${tag}: Konnte neue Note-Spalte nicht anlegen — Daten könnten verloren gehen:`, e, backup);
        return;
      }

      let restored = 0;
      let failed = 0;
      for (const idStr of Object.keys(backup)) {
        const id = Number(idStr);
        try {
          const resp = await this._merge(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
            { [fieldName]: backup[id] }
          );
          if (resp.ok) restored += 1;
          else failed += 1;
        } catch { failed += 1; }
      }
      console.warn(`[DEX] ${tag}: Migration fertig — ${restored} Werte zurückgeschrieben, ${failed} Fehler.`);
    } catch (e) {
      console.warn(`[DEX] ${tag} Error:`, e);
    }
  }

  /**
   * Berechtigungen für DEX_Events setzen
   */
  private async setEventsListPermissions(listName: string): Promise<void> {
    try {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      const ownersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
          {}
        );
      }

      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
          {}
        );
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden
    }
  }

  /**
   * Default View einer Liste konfigurieren
   */
  private async configureDefaultView(listName: string, fieldNames: string[], baseUrl?: string, opts?: { rebuild?: boolean }): Promise<void> {
    const url = baseUrl || this.siteUrl;
    try {
      let existingFields: string[] = [];
      if (opts?.rebuild) {
        // Komplett neu aufbauen — SP-Defaults (Modified, Created, ID, Type,
        // Compliance-Tag, App Created By, ...) werden rausgeworfen.
        try {
          await this._post(
            `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
            {}
          );
        } catch { /* ignore */ }
      } else {
        // Nur hinzufügen (behält bestehende SP-Felder bei). Duplikate vermeiden.
        const existingResponse = await this.context.spHttpClient.get(
          `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields`,
          SPHttpClient.configurations.v1
        );
        if (existingResponse.ok) {
          const existingData = await existingResponse.json();
          if (existingData.Items) existingFields = existingData.Items;
          else if (existingData.d?.Items) existingFields = existingData.d.Items;
          else if (existingData.value) existingFields = existingData.value;
        }
      }

      for (const fieldName of fieldNames) {
        // Nur hinzufügen wenn noch nicht in der View
        if (existingFields.indexOf(fieldName) < 0) {
          await this._post(
            `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/addviewfield('${fieldName}')`,
            {}
          );
        }
      }
    } catch {
      // View-Konfiguration ist optional
    }
  }

  /**
   * Column Formatting auf ein Feld setzen (z.B. Bild-Vorschau für URL-Spalten)
   */
  private async setColumnFormatting(listName: string, fieldName: string, formatJson: object): Promise<void> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq '${fieldName}'&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;
      const data = await response.json();
      const field = data.value?.[0];
      if (!field) return;

      await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields('${field.Id}')`,
        { CustomFormatter: JSON.stringify(formatJson) }
      );
      // Column Formatting gesetzt
    } catch {
      // Column Formatting ist optional
    }
  }

  // ==================== Events CRUD ====================

  private static readonly EVENT_SELECT = 'Id,Title,EventStatus,EventNumber,Description,Location,LocationAddress,LocationFilter,Audience,AudienceResolvedEmails,FilterMode,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,WaitlistEnabled,EventImageUrl,EmailImageBase64,Organizer,OrganizerEmail,ContactName,ContactEmail,ContactInfo,OutlookEventId,CalendarLink,OutlookBody,OutlookSubject,OutlookStart,OutlookEnd,OutlookLocation,EmailLanguage,RegistrationLanguage,EmailTemplateOverrides,DisableEmails,DisableRegistrationEmail,DisableCancellationEmail,AutoDeregisterOnDecline,DisableOutlook,OutlookDirty,AutoSendQRCode,ActiveFrom,NotifyOrgRegisterMode,NotifyOrgRegisterFromDate,NotifyOrgCancelMode,ExcludedUsers,IsFictive,DurchstarterCapacity,FunstarterCapacity,SplitLabelA,SplitLabelB,SplitSharedWaitlist,AllowAttendeeUpload,AttendeeUploadHint,AttendeeUploadLabel,AskSalutation,ConfirmDialogEnabled,ConfirmDialogMode,ConfirmDialogText,SelfCheckInEnabled,SelfCheckInToken,SelfCheckInFrom,SelfCheckInTo,TeamRegistrationEnabled,TeamSize,AskTeamName,TeamPartialAllowed,TeamOpenSlotsVisible,TeamJoinRequiresApproval,BilingualFields,CustomFields,Agenda,Transfers,Documents,FunZone,QuizClusterSize,ParentEventId,RegistrationListName,SubsiteUrl';

  /**
   * Strip SharePoint-Note-Field-Wrapper.
   *
   * Seit der Note-Migration wickelt SP die Werte für `Organizer` + `OrganizerEmail`
   * beim REST-Read in `<div class="ExternalClassXXXX">…</div>`. Vor dem Splitten
   * via `;` muss der Wrapper raus, sonst landen die Tag-Reste in den
   * Email-Listen → falsche Match-Vergleiche, kaputte Permissions.
   *
   * Idempotent: Werte ohne Wrapper bleiben unverändert.
   */
  private static stripNoteWrapper(value: string | null | undefined): string {
    if (!value) return '';
    let v = value.trim();
    v = v.replace(/^<div\b[^>]*>/i, '');
    v = v.replace(/<\/div>\s*$/i, '');
    return v.trim();
  }

  /**
   * Seed-Events anlegen falls sie nicht existieren (einmalig beim ersten Start).
   */
  public async seedEvents(): Promise<void> {
    try {
      // Prüfen ob "Assistenz Meeting 2026" schon existiert
      const check = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=Title eq 'Assistenz Meeting 2026'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (check.ok) {
        const data = await check.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0) return; // Existiert bereits
      }

      // Event anlegen
      await this.createEvent({
        title: 'Assistenz Meeting 2026',
        type: 'Other',
        status: 'Active',
        description: 'Assistenz Meeting Mai 2026 - Frankfurt am Main',
        location: 'Frankfurt am Main',
        locationFilter: '',
        audience: 'All',
        filterMode: 'OR',
        startDate: '2026-05-07T11:00:00.000Z',
        endDate: '2026-05-08T15:00:00.000Z',
        registrationDeadline: '2026-04-09T00:00:00.000Z',
        lastDeregisterDate: '',
        maxParticipants: 130,
        waitlistEnabled: true,
        eventImageUrl: '',
        organizer: 'Maerzluft, Petra; Schwartz, Eva',
        organizerEmail: 'pmaerzluft@deloitte.de',
        outlookEventId: '',
        outlookBody: '',
        emailLanguage: 'EN',
        emailTemplateOverrides: '',
        customFields: [
          { id: 'travel', label: 'You will travel with?', type: 'select', required: false, visible: true, options: ['Train', 'Car', 'Public Transport'] },
          { id: 'deutschlandticket', label: 'Do you own a Deutschlandticket?', type: 'select', required: false, visible: true, options: ['Yes', 'No'] },
          { id: 'expenses', label: 'Please insert the total amount of your travel expenses!', type: 'text', required: false, visible: true },
        ],
        agenda: '[]',
        transfers: '[]',
        documents: '[]',
      });
    } catch { /* Seed fehlgeschlagen - nicht kritisch */ }
  }

  /**
   * Alle Events laden
   */
  public async getEvents(): Promise<SPEvent[]> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EventService.EVENT_SELECT}&$orderby=StartDate desc&$top=100`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return [];
      const data = await response.json();
      return data.value || [];
    } catch {
      return [];
    }
  }

  /**
   * Einzelnes Event laden
   */
  public async getEvent(eventId: number): Promise<SPEvent | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=${EventService.EVENT_SELECT}`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      return await response.json();
    } catch {
      return null;
    }
  }

  /**
   * v18.33: Event anhand des Self-Check-in-Tokens finden (für den statischen
   * Check-in-Link ?action=selfcheckin&token=…). Liefert das erste Event mit
   * passendem Token. Alle eingeloggten User dürfen DEX_Events lesen.
   */
  public async getEventBySelfCheckInToken(token: string): Promise<SPEvent | null> {
    try {
      const safe = token.replace(/'/g, "''");
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EventService.EVENT_SELECT}&$filter=SelfCheckInToken eq '${safe}'&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      const arr = data.value || (data.d && data.d.results) || [];
      return arr.length > 0 ? arr[0] : null;
    } catch {
      return null;
    }
  }

  /**
   * v18.33: Event anhand der Event-Nummer finden (für den rotierenden Live-QR
   * ?action=selfcheckin&event=<Nr>&code=…&t=…).
   */
  public async getEventByEventNumber(eventNumber: number): Promise<SPEvent | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EventService.EVENT_SELECT}&$filter=EventNumber eq ${eventNumber}&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      const arr = data.value || (data.d && data.d.results) || [];
      return arr.length > 0 ? arr[0] : null;
    } catch {
      return null;
    }
  }

  /**
   * Neues Event erstellen + Subsite mit Teilnehmerliste anlegen
   */
  public async createEvent(event: {
    title: string;
    status: string;
    type: string;
    description: string;
    location: string;
    locationAddress?: string; // JSON-String: { street, houseNo, zip, city }
    outlookSubject?: string; // v18.42: Betreff des Outlook-Termins (leer = Titel)
    outlookStart?: string; // v18.44: abweichende Start-Zeit (ISO, leer = Event-Start)
    outlookEnd?: string;   // v18.44: abweichende End-Zeit (ISO, leer = Event-Ende)
    outlookLocation?: string; // v18.40: manueller Outlook-Ort (leer = Auto aus Ort + Adresse)
    locationFilter: string;
    audience: string;
    /** v16.4: Vor-aufgelöste E-Mails der Audience-DLs, ';'-separiert, lowercase. */
    audienceResolvedEmails?: string;
    filterMode: string;
    startDate: string;
    endDate: string;
    registrationDeadline: string;
    lastDeregisterDate: string;
    maxParticipants: number;
    waitlistEnabled: boolean;
    eventImageUrl: string;
    organizer: string;
    organizerEmail: string;
    /** v10.16: optionaler Ansprechpartner (Anzeige-Feld). */
    contactName?: string;
    contactEmail?: string;
    contactInfo?: string;
    outlookEventId: string;
    outlookBody: string;
    agenda?: string; // JSON-Array mit Agenda-Einträgen
    transfers?: string; // JSON-Array mit Transferzeiten
    documents?: string; // JSON-Array mit Dokumenten
    funZone?: string; // JSON-Array mit Quiz-Fragen
    quizClusterSize?: number; // 1..4 - Fragen pro Quiz-Ansicht
    /** Seit v6.4: wenn gesetzt, wird dieses Event als Sub-Event angelegt und zeigt auf das angegebene Parent-Event. */
    parentEventId?: string;
    emailLanguage?: string;
    registrationLanguage?: 'de' | 'en';
    emailTemplateOverrides?: string;
    disableEmails?: boolean;
    disableRegistrationEmail?: boolean;
    disableCancellationEmail?: boolean;
    autoDeregisterOnDecline?: boolean;
    disableOutlook?: boolean;
    notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
    notifyOrgRegisterFromDate?: string;
    notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';
    excludedUsers?: string[];
    isFictive?: boolean;
    durchstarterCapacity?: number;
    funstarterCapacity?: number;
    splitLabelA?: string;
    splitLabelB?: string;
    splitSharedWaitlist?: boolean;
    allowAttendeeUpload?: boolean;
    attendeeUploadHint?: string;
    attendeeUploadLabel?: string;
    /** v11.80: Anrede im Registrierungsformular abfragen (Default false). */
    askSalutation?: boolean;
    /** v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung. */
    confirmDialogEnabled?: boolean;
    confirmDialogMode?: string; // 'summary' | 'freetext'
    confirmDialogText?: string;
    /** v18.33: Self-Check-in per QR-Code erlauben (Default false). */
    selfCheckInEnabled?: boolean;
    /** v18.33: Geheimer Token (statischer Link + HMAC-Schlüssel rotierender QR). */
    selfCheckInToken?: string;
    /** v18.33: optionaler Start des Check-in-Fensters (ISO). */
    selfCheckInFrom?: string;
    /** v18.33: optionales Ende des Check-in-Fensters (ISO). */
    selfCheckInTo?: string;
    /** v11.80: Team-Anmeldung erlauben (Default false). */
    teamRegistrationEnabled?: boolean;
    /** v11.80: Maximale Teamgröße (0 = nicht gesetzt). */
    teamSize?: number;
    /** v11.80: Team-Name abfragen (Default false). */
    askTeamName?: boolean;
    /** v11.81: Auch Teil-Teams zulassen (Default false = nur komplette Teams). */
    teamPartialAllowed?: boolean;
    /** v11.81: Offene Slots öffentlich für Beitritt sichtbar (Default false). */
    teamOpenSlotsVisible?: boolean;
    /** v11.81: Beitritt erfordert Bestätigung durch Team-Kapitän (Default false). */
    teamJoinRequiresApproval?: boolean;
    /** v17.20: Custom-Fields zweisprachig anbieten (DE+EN). */
    bilingualFields?: boolean;
    customFields: CustomField[];
    /** v11.69: Wenn `existingSubsiteUrl` UND `existingRegistrationListName`
     *  gesetzt sind, wird KEINE neue Subsite und KEINE neue Teilnehmer-
     *  liste angelegt — stattdessen werden die mitgegebenen Werte direkt in
     *  das neue DEX_Events-Item geschrieben. Hintergrund: Outlook-Termin
     *  nachträglich aktivieren ohne Verlust bestehender Anmeldungen — das
     *  Sub-Event wird mit `deleteEventItemOnly()` aus DEX_Events entfernt
     *  und hier neu angelegt, wobei die alte Subsite + Teilnehmerliste
     *  unangetastet bleiben und an die neue Event-Zeile angehängt werden.
     *  Damit triggert der `DEX_CreateOutlookEvent`-Flow (GetOnNewItems) auf
     *  dem neuen Item und legt den Outlook-Termin an. */
    existingSubsiteUrl?: string;
    existingRegistrationListName?: string;
    /** v11.87: Optionaler Progress-Callback. Wird zu Beginn jeder Teil-
     *  Operation aufgerufen — die UI kann darauf den Fortschrittsbalken
     *  und die Unter-Caption sichtbar bewegen, statt minutenlang auf
     *  „Event wird vorbereitet..." stehen zu bleiben. Stages decken
     *  die langsamen SP-Operationen ab (Subsite-Create, Listen-Create,
     *  Permissions, Counter, Views). */
    onProgress?: (stage:
      | 'start'
      | 'subsite-creating'
      | 'subsite-done'
      | 'permissions'
      | 'list-creating'
      | 'list-done'
      | 'item-insert'
      | 'done'
    ) => void;
  }): Promise<number | null> {
    const reportProgress = (stage:
      | 'start'
      | 'subsite-creating'
      | 'subsite-done'
      | 'permissions'
      | 'list-creating'
      | 'list-done'
      | 'item-insert'
      | 'done'
    ): void => {
      try { event.onProgress?.(stage); } catch { /* */ }
    };
    try {
      reportProgress('start');
      // 0. Nächste EventNumber ermitteln
      let nextEventNumber = 1;
      try {
        const enResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=EventNumber&$orderby=EventNumber desc&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (enResp.ok) {
          const enData = await enResp.json();
          if (enData.value && enData.value.length > 0 && enData.value[0].EventNumber) {
            nextEventNumber = enData.value[0].EventNumber + 1;
          }
        }
      } catch { /* Fallback: 1 */ }

      // v11.69: Reuse-Pfad — wenn `existingSubsiteUrl` UND
      // `existingRegistrationListName` mitgegeben wurden, ueberspringen wir
      // 1) Subsite-Anlegen, 2) Subsite-Permissions, 3) Teilnehmerliste
      // anlegen. Die mitgegebene Subsite bleibt unangetastet inkl. aller
      // Teilnehmer-Anmeldungen. Custom-Fields werden ohne spInternalName-
      // Anreicherung uebernommen — die Felder existieren bereits auf der
      // alten Teilnehmerliste mit den korrekten Internal-Names.
      const reuseSubsite = !!(event.existingSubsiteUrl && event.existingRegistrationListName);
      let subsiteUrl: string;
      let enrichedCustomFields: CustomField[];
      const coOrgEmailsForPerm: string[] = (() => {
        try {
          const o = JSON.parse(event.emailTemplateOverrides || '{}');
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const list = (o as any)._coOrganizers;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          if (Array.isArray(list)) return list.map((x: any) => String(x?.email || '')).filter(Boolean);
        } catch { /* */ }
        return [];
      })();
      const allOrgEmails = [event.organizerEmail || '', ...coOrgEmailsForPerm].filter(Boolean).join(';');
      const regListName = reuseSubsite ? (event.existingRegistrationListName as string) : REG_LIST_NAME;

      if (reuseSubsite) {
        // v11.69: bestehende Subsite + Teilnehmerliste wiederverwenden.
        subsiteUrl = event.existingSubsiteUrl as string;
        // Custom-Fields unverändert uebernehmen — die Liste existiert
        // bereits, kein neues Schema nötig.
        enrichedCustomFields = event.customFields.map(cf => ({ ...cf }));
      } else {
        // 1. Subsite für das Event erstellen
        reportProgress('subsite-creating');
        const createdSubsite = await this.createEventSubsite(event.title, event.description);
        if (!createdSubsite) {
          console.error('[DEX] Subsite konnte nicht erstellt werden');
          throw new Error('Subsite konnte nicht erstellt werden. Fehlende Berechtigung? Bitte wende dich an einen Site-Administrator.');
        }
        subsiteUrl = createdSubsite;
        reportProgress('subsite-done');

        // 2. Subsite-Berechtigungen: Members der Parent-Site auf der Subsite berechtigen.
        // v9.18: Co-Organizer-Emails aus emailTemplateOverrides._coOrganizers extrahieren
        // und mit dem Hauptorganizer zusammen Full Control erteilen.
        reportProgress('permissions');
        await this.setSubsitePermissions(subsiteUrl, allOrgEmails);

        // 3. Teilnehmerliste auf der Subsite erstellen
        reportProgress('list-creating');
        const fieldMap: Record<string, string> = await this.createRegistrationList(subsiteUrl, event.customFields, allOrgEmails);
        reportProgress('list-done');

        // Custom Fields mit SP InternalName anreichern
        enrichedCustomFields = event.customFields.map(cf => ({
          ...cf,
          spInternalName: fieldMap[cf.id] || '',
        }));
      }

      // 3. Event in DEX_Events eintragen
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
        'Title': event.title,
        'EventNumber': nextEventNumber,
        'EventStatus': event.status,
        'Description': event.description,
        'Location': event.location,
        'LocationAddress': event.locationAddress || '',
        // v18.42: Outlook-Betreff (leer = Flow fällt auf Titel zurück via coalesce).
        'OutlookSubject': (event.outlookSubject && event.outlookSubject.trim()) ? event.outlookSubject.trim() : '',
        // v18.44: abweichendes Outlook-Datum (leer = Flow nutzt StartDate/EndDate).
        'OutlookStart': event.outlookStart || null,
        'OutlookEnd': event.outlookEnd || null,
        // v18.34/v18.40: Outlook-Ort = manuelle Überschreibung, sonst
        // automatisch aus Veranstaltungsort + Adresse. Flow mappt OutlookLocation 1:1.
        'OutlookLocation': (event.outlookLocation && event.outlookLocation.trim())
          ? event.outlookLocation.trim()
          : buildOutlookLocation(event.location, event.locationAddress),
        'LocationFilter': event.locationFilter,
        'Audience': event.audience,
        'AudienceResolvedEmails': event.audienceResolvedEmails || '',
        'FilterMode': event.filterMode || 'OR',
        'StartDate': event.startDate || null,
        'EndDate': event.endDate || null,
        'RegistrationDeadline': event.registrationDeadline || null,
        'LastDeregisterDate': event.lastDeregisterDate || null,
        'MaxParticipants': event.maxParticipants,
        'WaitlistEnabled': event.waitlistEnabled,
        'EventImageUrl': event.eventImageUrl,
        // Custom-Event-Logo aus emailTemplateOverrides._eventLogo extrahieren (falls
        // vorhanden) und als EmailImageBase64 persistieren — damit der Power-Automate-Flow
        // es als {{ORB_URL}} in Mail + Outlook-Termin einsetzt.
        'EmailImageBase64': (() => {
          try {
            const o = JSON.parse(event.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })(),
        'Organizer': event.organizer,
        'OrganizerEmail': event.organizerEmail,
        // v10.16: optionaler Ansprechpartner (Anzeige-Feld). Strings können
        // leer sein — leer = kein Ansprechpartner gepflegt.
        'ContactName': event.contactName || '',
        'ContactEmail': event.contactEmail || '',
        'ContactInfo': event.contactInfo || '',
        'OutlookEventId': event.outlookEventId,
        // outlookBody kommt bereits vollständig gewickelt + mit aufgelösten Variablen
        // aus EventCreationPage — hier nur durchreichen.
        'OutlookBody': event.outlookBody || '',
        'EmailLanguage': event.emailLanguage || 'EN',
        'RegistrationLanguage': event.registrationLanguage || '',
        'EmailTemplateOverrides': event.emailTemplateOverrides || '',
        'DisableEmails': !!event.disableEmails,
        'DisableRegistrationEmail': !!event.disableRegistrationEmail,
        'DisableCancellationEmail': !!event.disableCancellationEmail,
        'AutoDeregisterOnDecline': !!event.autoDeregisterOnDecline,
        'DisableOutlook': !!event.disableOutlook,
        'NotifyOrgRegisterMode': (() => {
          const m = event.notifyOrgRegisterMode || 'never';
          return m === 'always' ? 'Always' : m === 'fromDate' ? 'FromDate' : 'Never';
        })(),
        'NotifyOrgRegisterFromDate': event.notifyOrgRegisterFromDate || null,
        'NotifyOrgCancelMode': (() => {
          const m = event.notifyOrgCancelMode || 'never';
          return m === 'always' ? 'Always' : m === 'afterDeadline' ? 'AfterDeadline' : 'Never';
        })(),
        'ExcludedUsers': (event.excludedUsers || []).filter(Boolean).join(';'),
        'IsFictive': !!event.isFictive,
        'DurchstarterCapacity': typeof event.durchstarterCapacity === 'number' ? event.durchstarterCapacity : null,
        'FunstarterCapacity': typeof event.funstarterCapacity === 'number' ? event.funstarterCapacity : null,
        'SplitLabelA': event.splitLabelA || '',
        'SplitLabelB': event.splitLabelB || '',
        'SplitSharedWaitlist': !!event.splitSharedWaitlist,
        'AllowAttendeeUpload': !!event.allowAttendeeUpload,
        'AttendeeUploadHint': event.attendeeUploadHint || '',
        'AttendeeUploadLabel': event.attendeeUploadLabel || '',
        'AskSalutation': !!event.askSalutation,
        'ConfirmDialogEnabled': !!event.confirmDialogEnabled,
        'ConfirmDialogMode': event.confirmDialogMode || '',
        'ConfirmDialogText': event.confirmDialogText || '',
        'SelfCheckInEnabled': !!event.selfCheckInEnabled,
        'SelfCheckInToken': event.selfCheckInToken || '',
        'SelfCheckInFrom': event.selfCheckInFrom || null,
        'SelfCheckInTo': event.selfCheckInTo || null,
        'TeamRegistrationEnabled': !!event.teamRegistrationEnabled,
        'TeamSize': typeof event.teamSize === 'number' && event.teamSize > 0 ? event.teamSize : null,
        'AskTeamName': !!event.askTeamName,
        'TeamPartialAllowed': !!event.teamPartialAllowed,
        'TeamOpenSlotsVisible': !!event.teamOpenSlotsVisible,
        'TeamJoinRequiresApproval': !!event.teamJoinRequiresApproval,
        'BilingualFields': !!event.bilingualFields,
        'CustomFields': JSON.stringify(enrichedCustomFields),
        'Agenda': event.agenda || '[]',
        'Transfers': event.transfers || '[]',
        'Documents': event.documents || '[]',
        'FunZone': event.funZone || '[]',
        'QuizClusterSize': typeof event.quizClusterSize === 'number' ? event.quizClusterSize : null,
        'ParentEventId': event.parentEventId || '',
        'RegistrationListName': regListName,
        'RegistrationListUrl': `${subsiteUrl}/Lists/${regListName}/AllItems.aspx`,
        'SubsiteUrl': subsiteUrl,
      };

      reportProgress('item-insert');
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`,
        payload
      );

      if (!response.ok) return null;
      const result = await response.json();
      reportProgress('done');
      return result.d?.Id || result.Id;
    } catch (err) {
      if (err instanceof Error) throw err;
      return null;
    }
  }


  /**
   * Admin-Cleanup beim App-Start: alle Events mit EventStatus='Active' und EndDate < jetzt
   * werden automatisch auf 'Completed' gesetzt. Liefert die Anzahl der aktualisierten Events.
   */
  public async markExpiredEventsAsCompleted(): Promise<number> {
    try {
      // SharePoint OData Filter: Active + EndDate < jetzt
      const nowIso = new Date().toISOString();
      const filter = `EventStatus eq 'Active' and EndDate lt datetime'${nowIso}'`;
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=${encodeURIComponent(filter)}&$select=Id,Title,EndDate&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return 0;
      const data = await resp.json();
      const items: Array<{ Id: number; Title: string }> = data.value || data.d?.results || [];
      if (items.length === 0) return 0;

      let updated = 0;
      for (const it of items) {
        try {
          const ok = await this.updateEvent(it.Id, { 'EventStatus': 'Completed' });
          if (ok) updated += 1;
        } catch { /* einzelnes Update ueberspringen */ }
      }
      return updated;
    } catch (err) {
      console.warn('[DEX] markExpiredEventsAsCompleted failed:', err);
      return 0;
    }
  }

  /**
   * Event aktualisieren
   */
  /**
   * v11.11: Versionsverlauf des Event-Items aus DEX_Events lesen, um
   * versehentlich gelöschte Custom-Fields (z.B. b2run_*-Felder nach
   * der zu aggressiven v11.9-Migration) wieder zurückzuholen.
   *
   * Liefert eine Liste der Versionen, jeweils mit dem geparsten
   * `CustomFields`-Array (sortiert: neueste zuerst). Werte ohne
   * CustomFields oder mit leerem Array fallen einfach mit raus, sind
   * aber nicht gefiltert — der Caller entscheidet, welche Version
   * relevant ist.
   */
  public async getEventCustomFieldsHistory(eventId: number): Promise<Array<{
    versionLabel: string;
    modified: string;
    customFields: Array<Record<string, unknown>>;
  }>> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/versions?$select=VersionLabel,Modified,CustomFields`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
      });
      if (!response.ok) {
        console.warn('[DEX] getEventCustomFieldsHistory failed:', response.status);
        return [];
      }
      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const versions: any[] = data.value || [];
      return versions.map(v => {
        let parsed: Array<Record<string, unknown>> = [];
        try {
          const raw = (v.CustomFields || '').toString();
          if (raw.trim()) {
            const obj = JSON.parse(raw);
            if (Array.isArray(obj)) parsed = obj as Array<Record<string, unknown>>;
          }
        } catch { /* invalid JSON in old version → leer */ }
        return {
          versionLabel: String(v.VersionLabel || ''),
          modified: String(v.Modified || ''),
          customFields: parsed,
        };
      });
    } catch (err) {
      console.warn('[DEX] getEventCustomFieldsHistory error:', err);
      return [];
    }
  }

  public async updateEvent(eventId: number, updates: Record<string, unknown>): Promise<boolean> {
    try {
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
        ...updates,
      };

      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            'odata-version': '',
          },
          body: JSON.stringify(payload),
        }
      );
      if (!response.ok) {
        const errText = await response.text().catch(() => '');
        console.warn('[DEX] updateEvent failed:', response.status, errText.substring(0, 200));
      }
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Event vollständig löschen:
   * 1. Subsite löschen (inkl. Teilnehmerliste) - für neue Events
   * 2. Alte Registrierungsliste löschen (DEX_Reg_*) - für alte Events
   * 3. Event-Eintrag aus DEX_Events löschen
   */
  public async deleteEvent(eventId: number): Promise<boolean> {
    try {
      // Event-Daten laden um SubsiteUrl und RegistrationListName zu bekommen
      const event = await this.getEvent(eventId);
      if (!event) return false;

      // 0. Outlook-Kalendereintrag per Queue löschen (VOR allem anderen, damit
      //    CalendarLink noch vorhanden ist). Der DEX_Outlook_Einladungen-Flow
      //    greift den DeleteEvent-Eintrag auf und löscht den Kalender-Termin
      //    im Shared Mailbox ueber den Flow-Service-Account.
      //    Fehler hier ignorieren - Event-Delete soll trotzdem durchlaufen.
      if (event.CalendarLink) {
        try {
          await this.queueOutlookDeleteEvent(String(eventId), event.Title || '', event.CalendarLink);
        } catch { /* Queue-Fehler ignorieren */ }
      }
      // 1. Subsite RECYCEN (v9.0: nicht mehr per DELETE, sonst landet die
      //    Subsite permanent weg ohne Recycle-Bin-Eintrag. recycle() legt
      //    die Subsite mitsamt Teilnehmerliste 93 Tage in den Site
      //    Collection Recycle Bin → ein Tenant-Admin / Site Collection
      //    Admin kann sie dort wiederherstellen falls nötig.
      if (event.SubsiteUrl) {
        try {
          await this._post(`${event.SubsiteUrl}/_api/web/recycle`, {});
        } catch {
          console.warn('[DEX] Subsite konnte nicht in den Recycle Bin verschoben werden:', event.SubsiteUrl);
        }
      }

      // 2. Event-Bild ebenfalls RECYCEN statt löschen.
      if (event.EventImageUrl) {
        try {
          const url = new URL(event.EventImageUrl);
          const serverRelUrl = url.pathname;
          if (serverRelUrl.indexOf('DEX_EventImages') >= 0) {
            await this._post(
              `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelUrl}')/recycle`,
              {}
            );
          }
        } catch {
          console.warn('[DEX] Event-Bild konnte nicht in den Recycle Bin verschoben werden');
        }
      }

      // 3. Alte Registrierungsliste recyceln (legacy Events ohne Subsite).
      if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
        try {
          await this._post(
            `${this.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')/recycle`,
            {}
          );
        } catch {
          console.warn('[DEX] Alte Registrierungsliste konnte nicht recycelt werden:', event.RegistrationListName);
        }
      }

      // 3. DEX_Participants aufräumen: EventNumber aus allen Teilnehmern entfernen
      if (event.EventNumber) {
        try {
          const allParticipants = await this.getAllParticipants();
          // Parallelize participant cleanup for better performance
          const updatePromises = allParticipants
            .filter(p => {
              const en = String(event.EventNumber);
              const hasRegistered = p.EventRegistered?.split(',').map(s => s.trim()).includes(en);
              const hasWaitlist = p.EventOnWaitlist?.split(',').map(s => s.trim()).includes(en);
              return hasRegistered || hasWaitlist;
            })
            .map(p => this.removeParticipantEvent(p.Email, event.EventNumber));
          // Promise.all mit individueller Fehlerbehandlung (Promise.allSettled nicht verfügbar in ES2017)
          const safePromises = updatePromises.map(p => p.catch(() => null));
          await Promise.all(safePromises);
        } catch {
          console.warn('[DEX] DEX_Participants konnte nicht aufgeräumt werden');
        }
      }

      // 4. Event-Dokumente löschen (SiteAssets/DEX_EventDocs/Event_{number}_*)
      if (event.EventNumber) {
        try {
          const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
          const safeName = (event.Title || '').replace(/[#%&*:<>?/\\|"']/g, '').replace(/\s+/g, '_').substring(0, 50);
          const folderName = safeName ? `Event_${event.EventNumber}_${safeName}` : `Event_${event.EventNumber}`;
          await this._delete(`${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/${folderName}')`);
        } catch {
          // Fallback: alten Ordnernamen ohne Titel probieren
          try {
            const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
            await this._delete(`${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/Event_${event.EventNumber}')`);
          } catch { /* Ordner nicht gefunden */ }
        }
      }

      // 5. Event-Eintrag aus DEX_Events RECYCEN (v9.0: per recycle() statt
      //    delete(), damit ein Admin via SharePoint-Recycle-Bin das Item
      //    bei Bedarf 93 Tage lang wiederherstellen kann).
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/recycle`,
        {}
      );

      // 6. Audit-Eintrag in DEX_ChangeLog (v9.0). Best-effort, blockt
      //    den Lösch-Vorgang nicht falls Logging fehlschlägt.
      try {
        await this.writeChangeLog({
          action: 'EventDeletedTest', // wird vom Aufrufer ueberschrieben
          targetType: 'Event',
          targetId: String(eventId),
          targetName: event.Title || '',
          eventId: String(eventId),
          eventTitle: event.Title || '',
          details: {
            subsiteUrl: event.SubsiteUrl || '',
            eventNumber: event.EventNumber,
            recycledTo: 'SharePoint Recycle Bin (93 Tage)',
          },
        });
      } catch { /* */ }

      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * v11.69: Löscht NUR das DEX_Events-Listenitem — KEIN Cascade auf Subsite,
   * KEIN Outlook-DeleteEvent in die Queue, KEIN EventImage-Recycle, KEIN
   * DEX_Participants-Cleanup. Gegenstück zu `deleteEvent()`, das alles
   * mit-aufräumt.
   *
   * Nutzungs-Szenario: Outlook-Termin nachträglich auf einem bereits
   * angelegten Sub-Event aktivieren. Der `DEX_CreateOutlookEvent`-Flow
   * triggert ausschliesslich auf NEUE DEX_Events-Items (GetOnNewItems) —
   * ein MERGE-Update reicht nicht aus. Statt das ganze Sub-Event komplett
   * delete+recreate zu machen (was kaskadierend Subsite + Teilnehmer-
   * anmeldungen mitlöschen würde), wird hier nur die DEX_Events-Zeile
   * entfernt und gleich darauf eine neue mit `createEvent({ existingSubsiteUrl,
   * existingRegistrationListName })` angelegt. Die alte Subsite mit allen
   * Anmeldungen bleibt unangetastet und wird einfach an die neue Zeile
   * gekoppelt.
   *
   * **Garantie:** Diese Methode ruft KEIN `recycle()` auf der Subsite, KEIN
   * `recycle()` auf der Teilnehmerliste und KEIN `removeParticipantEvent()`.
   * Nur das DEX_Events-Item wird per REST-DELETE entfernt — alles andere
   * bleibt 1:1 erhalten.
   */
  public async deleteEventItemOnly(eventId: string | number): Promise<boolean> {
    try {
      const idNum = typeof eventId === 'string' ? parseInt(eventId, 10) : eventId;
      if (!idNum || Number.isNaN(idNum)) return false;
      const response = await this._delete(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${idNum})`
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== Subsites ====================

  /**
   * URL-Suffix aus Event-Titel generieren.
   * "B2Run Frankfurt 2026" → "b2run-frankfurt-2026-k8f3a"
   */
  private generateSubsiteUrl(title: string): string {
    const slug = title
      .toLowerCase()
      .replace(/[äÄ]/g, 'ae').replace(/[öÖ]/g, 'oe').replace(/[üÜ]/g, 'ue').replace(/ß/g, 'ss')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .substring(0, 40);
    const suffix = Date.now().toString(36).slice(-5);
    return `${slug}-${suffix}`;
  }

  /**
   * Subsite für ein Event erstellen.
   * Versucht mehrere Templates falls eines fehlschlägt.
   * Gibt die absolute URL der neuen Subsite zurück.
   */
  private async createEventSubsite(title: string, description: string): Promise<string | null> {
    const urlSuffix = this.generateSubsiteUrl(title);
    const desc = description || `Event-Subsite: ${title}`;

    // Templates in Reihenfolge versuchen:
    // STS#3 = Modern ohne Group, STS#0 = Classic Team Site, STS = Blank
    const templates = ['STS#3', 'STS#0', 'STS'];

    for (const template of templates) {
      try {
        const payload = {
          'parameters': {
            '__metadata': { 'type': 'SP.WebCreationInformation' },
            'Title': title,
            'Url': urlSuffix,
            'Description': desc,
            'Language': 1031,
            'WebTemplate': template,
            'UseSamePermissionsAsParentSite': false,
          }
        };

        const response = await this._post(`${this.siteUrl}/_api/web/webs/add`, payload);
        if (response.ok) {
          const result = await response.json();
          const subsiteAbsoluteUrl = result.d?.Url || result.Url;
          // Subsite erfolgreich erstellt
          return subsiteAbsoluteUrl || `${this.siteUrl}/${urlSuffix}`;
        }

        // Fehlerdetails loggen
        try {
          const err = await response.json();
          console.warn(`[DEX] Template ${template} fehlgeschlagen (${response.status}):`, err.error?.message?.value || err);
        } catch {
          console.warn(`[DEX] Template ${template} fehlgeschlagen: ${response.status}`);
        }
      } catch (e) {
        console.warn(`[DEX] Template ${template} Fehler:`, e);
      }
    }

    console.error('[DEX] Subsite konnte mit keinem Template erstellt werden');
    return null;
  }

  // ==================== Teilnehmerlisten (auf Subsites) ====================

  /**
   * Teilnehmerliste auf einer Subsite erstellen.
   * Liste heisst immer "Teilnehmer".
   */
  private async createRegistrationList(
    subsiteUrl: string,
    customFields: CustomField[],
    organizerEmail: string
  ): Promise<Record<string, string>> {
    // Liste erstellen
    await this._post(`${subsiteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': REG_LIST_NAME,
      'Description': 'Teilnehmerliste für dieses Event',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // Basis-Spalten
    const baseFields = [
      { title: 'TeilnehmerID', type: 9 }, // Number - fortlaufende ID
      { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'ParticipantName', type: 2 }, // Backward compat
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Department', type: 2 },
      { title: 'Location', type: 2 },
      { title: 'JobTitle', type: 2 },
      { title: 'Phone', type: 2 },
      // v24.29: Unternehmenszugehörigkeit / Rechtsträger (aus dem Profil).
      { title: 'Company', type: 2 },
      { title: 'Status', type: 6, choices: ['Angemeldet', 'QR versendet', 'Warteliste', 'Eingecheckt', 'No-Show', 'Abgemeldet'], metaType: 'SP.FieldChoice' },
      { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Typ-Auswahl
      { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Wunsch-Typ (wenn Fallback oder Warteliste)
      // v10.13: B2Run-Leistungsnachweis-Bestätigung. Virtuelles Feld der
      // RegistrationPage, das nur durchläuft wenn durchstarterRequiresProof
      // aktiv ist — die SP-Spalte muss aber existieren sonst kippt die
      // Anmeldung mit HTTP 400. Wird auf jeder neuen Teilnehmerliste angelegt
      // damit B2Run-Events nicht später nochmal manuell repariert werden müssen.
      { title: 'b2run_leistungsnachweis', type: 2 },
      { title: 'QuizScore', type: 9 }, // Number - Anzahl richtiger Antworten
      { title: 'QuizAnswers', type: 3 }, // Note - JSON der Antworten (für Statistik)
      { title: 'QuizCompletedAt', type: 4 }, // DateTime
      { title: 'RegistrationDate', type: 4 },
      { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgeführt hat
      { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgeführt hat
      { title: 'ProxyConsent', type: 3 },      // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung (Note)
      { title: 'LastModifiedDate', type: 4 },
      { title: 'ChangeLog', type: 3 }, // Note (multiline) - Aenderungshistorie
      { title: 'CancellationDate', type: 4 },
      { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgelöst hat
      { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
      { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
      { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
      { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
      // v17.15: Nachrück-Audit (siehe SPRegistration-Interface):
      // - PromotedDate: gesetzt beim Promote auf die nachrückende Person.
      // - ReplacedParticipantEmail: E-Mail der Person, deren Cancel den
      //   Promote ausgelöst hat („Ersetzt wen") — auf der promoteten Person.
      // - ReplacedByParticipantEmail: E-Mail der nachrückenden Person
      //   („Ersetzt durch") — auf der cancelnden Person.
      { title: 'PromotedDate', type: 4 },
      { title: 'ReplacedParticipantEmail', type: 2 },
      { title: 'ReplacedByParticipantEmail', type: 2 },
      // v11.36: Überbuchungs-Review-Marker. '' = normal, 'Pending' = vom
      // „Überbuchung prüfen"-Lauf als über Kapazität erkannt; der Admin
      // entscheidet pro Person (auf Warteliste / Platz behalten).
      { title: 'OverbookReview', type: 2 },
      // v11.82: Team-Anmeldung — drei Spalten gruppieren Mitglieder eines
      // gemeinsam angemeldeten Teams. TeamId = UUID (gleicher Wert für alle
      // Mitglieder), TeamLead = true nur für die anmeldende Person, TeamName
      // = optionaler frei wählbarer Name (nur wenn das Event AskTeamName an
      // hat). Bei Nicht-Team-Anmeldungen bleiben alle drei Felder leer.
      { title: 'TeamId', type: 2 },
      { title: 'TeamLead', type: 8 },
      { title: 'TeamName', type: 2 },
      { title: 'CustomData', type: 3 },
    ];

    for (const f of baseFields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if ((f as { choices?: string[] }).choices) {
        payload['Choices'] = { 'results': (f as { choices: string[] }).choices };
      }
      await this._post(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload);
    }

    // Custom Fields als eigene Spalten anlegen + InternalName merken
    const customFieldViewNames: string[] = [];
    const fieldMap: Record<string, string> = {}; // cf.id -> SP InternalName
    for (const cf of customFields) {
      if (!cf.label) continue;
      // v19.0: Dokument-Felder bekommen KEINE Spalte — die Datei wird als
      // Attachment an die Teilnehmer-Zeile gehängt, nicht als Spaltenwert.
      if (cf.type === 'document') continue;
      let fieldPayload: Record<string, unknown>;

      if (cf.type === 'select' && cf.options && cf.options.length > 0) {
        fieldPayload = {
          '__metadata': { 'type': 'SP.FieldChoice' },
          'Title': cf.label,
          'FieldTypeKind': 6,
          'Required': false,
          'Choices': { 'results': cf.options },
        };
      } else if (cf.type === 'number') {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 9,
          'Required': false,
        };
      } else if (cf.type === 'checkbox') {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 8,
          'Required': false,
        };
      } else {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 2,
          'Required': false,
        };
      }

      try {
        const fieldResponse = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
          fieldPayload
        );
        if (fieldResponse.ok) {
          const fieldResult = await fieldResponse.json();
          const internalName = fieldResult.d?.InternalName || fieldResult.InternalName || cf.label;
          fieldMap[cf.id] = internalName;
          customFieldViewNames.push(internalName);
        }
      } catch {
        console.warn('[DEX] Custom Field konnte nicht angelegt werden:', cf.label);
      }
    }

    // FieldMap wird als Rückgabewert an den Caller zurückgegeben

    // Default View komplett neu aufbauen (Basis + Custom Fields). Mit rebuild:true
    // werden alle SP-Default-Spalten (Modified, Created, ID, Type, Compliance Asset,
    // App Created By, ...) aus der View rausgeworfen — nur funktionelle Felder.
    await this.configureDefaultView(REG_LIST_NAME, [
      'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail', 'Department', 'Location', 'JobTitle', 'Company', 'Phone', 'StarterType', 'PreferredStarterType', 'Status', 'RegistrationDate', 'RegisteredByName', 'RegisteredByEmail', 'ProxyConsent', 'CancellationDate', 'CancelledByName', 'CancelledByEmail',
      ...customFieldViewNames,
      // v11.82: Team-Spalten am Ende der View (nach allen Custom Fields, vor
      // System-Spalten). So bleibt die View bei Nicht-Team-Events unauffällig
      // und bei Team-Events sieht der Organizer auf einen Blick, wer mit wem
      // angemeldet ist.
      'TeamId', 'TeamLead', 'TeamName',
    ], subsiteUrl, { rebuild: true });

    // Item-Level Permissions
    await this.setItemLevelPermissions(subsiteUrl);

    // Berechtigungen
    await this.setRegistrationListPermissions(subsiteUrl, organizerEmail);

    // v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe anlegen
    // (Race-Condition-Schutz bei parallelen Anmeldungen).
    try {
      await this.ensureCounterList(subsiteUrl);
    } catch {
      // Nicht kritisch — falls das schiefgeht, fallback auf max+1 in upsertParticipant.
    }

    return fieldMap;
  }

  /**
   * v9.35: Berechtigungs-Sync für nachträglich hinzugefügte Organizer/Co-Organizer.
   *
   * Wird im Wizard im Edit-Modus nach updateEvent aufgerufen. Geht über die
   * komma-/semikolon-separierte Liste aller Organizer-Mails und stellt sicher,
   * dass jede Person Full Control auf der Subsite + auf der Teilnehmerliste hat.
   *
   * Idempotent: Personen, die bereits Full Control haben, werden von SharePoints
   * `addroleassignment` einfach durchgereicht (kein Fehler, kein Doppel-Eintrag).
   * Existierende Item-Level-Permissions auf der Liste bleiben unangetastet — wir
   * brechen die Inheritance hier NICHT erneut, sondern fügen nur fehlende Principals
   * obendrauf hinzu.
   */
  public async ensureOrganizerPermissions(subsiteUrl: string, organizerEmails: string): Promise<void> {
    if (!subsiteUrl || !organizerEmails) return;
    const emails = organizerEmails.split(/[;,]/).map(s => s.trim()).filter(Boolean);
    for (const em of emails) {
      try {
        const userResponse = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (!userResponse.ok) continue;
        const userData = await userResponse.json();
        const userId = userData.d?.Id || userData.Id;
        if (!userId) continue;
        // Subsite Full Control (Web-Level)
        try {
          await this._post(
            `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
            {}
          );
        } catch { /* idempotent — Person hatte schon Rechte */ }
        // Teilnehmerliste Full Control (List-Level)
        try {
          await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
            {}
          );
        } catch { /* idempotent */ }
      } catch { /* skip einzelne User-Fehler, mit nächstem weiter */ }
    }
  }

  /**
   * Subsite-Berechtigungen: Owners Full Control, Members Read (damit User die Subsite betreten können).
   */
  private async setSubsitePermissions(subsiteUrl: string, organizerEmail: string): Promise<void> {
    try {
      // Owners der Hauptsite: Full Control auf der Subsite
      const ownersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        const ownersId = ownersData.d?.Id || ownersData.Id;
        await this._post(
          `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${ownersId}, roledefid=1073741829)`,
          {}
        );
      }

      // Visitors der Hauptsite: Read auf der Subsite (damit User die Subsite betreten können)
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
          {}
        );
      }

      // Organizer: Full Control auf der Subsite. v9.18: organizerEmail kann
      // ";"-separiert mehrere Emails enthalten — Hauptorganizer + Co-Organizer
      // bekommen alle Full Control auf der Subsite.
      if (organizerEmail) {
        const emails = organizerEmail.split(/[;,]/).map(s => s.trim()).filter(Boolean);
        for (const em of emails) {
          try {
            const userResponse = await this.context.spHttpClient.get(
              `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
              SPHttpClient.configurations.v1
            );
            if (userResponse.ok) {
              const userData = await userResponse.json();
              const userId = userData.d?.Id || userData.Id;
              await this._post(
                `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
                {}
              );
            }
          } catch { /* Organizer-Berechtigung optional */ }
        }
      }
    } catch {
      console.warn('[DEX] Subsite-Berechtigungen konnten nicht gesetzt werden');
    }
  }

  /**
   * Item-Level Permissions auf der Teilnehmerliste setzen.
   */
  private async setItemLevelPermissions(subsiteUrl: string): Promise<void> {
    try {
      await this.context.spHttpClient.post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'ReadSecurity': 2,
            'WriteSecurity': 2,
          }),
        }
      );
    } catch {
      // Item-Level Permissions konnten nicht gesetzt werden
    }
  }

  /**
   * v20.5: Setzt nachträglich den Autor (Created By / SharePoint-Ersteller)
   * einer Teilnehmer-Zeile auf den TEILNEHMER selbst.
   *
   * Hintergrund: Die Teilnehmerlisten laufen mit Item-Level-Security
   * (ReadSecurity=2 / WriteSecurity=2) — ein User darf nur Items LESEN und
   * BEARBEITEN, die ER ERSTELLT hat (geprüft am Autor, NICHT am Feld
   * ParticipantEmail). Bei einer stellvertretenden Anmeldung (Organizer/Admin
   * meldet eine andere Person an) wäre sonst der Akteur der Autor — der
   * angemeldete Teilnehmer sähe seine eigene Anmeldung NICHT in "Meine
   * Events" und könnte sich nicht selbst abmelden. Indem der Teilnehmer zum
   * Autor wird, bekommt er Lese- + Abmelde-Zugriff auf SEINE Zeile.
   *
   * Best-effort: das Umsetzen von AuthorId erfordert "Listen verwalten" /
   * Full Control auf der Liste. Organizer (eigenes Event) und Admin haben das;
   * ein normaler Contribute-User (z.B. eine Assistenz) NICHT — dort schlägt
   * der MERGE mit 403 fehl und wird still ignoriert (die Zeile bleibt beim
   * Akteur als Autor). Der eigentliche Akteur ist ohnehin separat im Feld
   * RegisteredByEmail protokolliert, der Audit-Nachweis bleibt also erhalten.
   */
  /**
   * v24.40: Eine Teilnehmer-Zeile einer **Assistenz** zuordnen — der Admin
   * übergibt damit die Verwaltung der (Fremd-)Anmeldung an eine bestimmte
   * Assistenz. Setzt ZWEI Dinge:
   *  1. `RegisteredByEmail`/`RegisteredByName` auf die Assistenz (Audit + Filter
   *     der „Assistenz"-Kachel).
   *  2. Den **Zeilen-Autor** (`Created By` / `AuthorId`) auf die Assistenz —
   *     unter Item-Level-Security („nur eigene Elemente") ist das die
   *     Voraussetzung, damit eine NORMALE Assistenz die Zeile überhaupt
   *     lesen/bearbeiten darf (sonst sieht sie sie in ihrer Kachel nicht).
   * Best-effort: Schlägt der Autor-Wechsel mangels Rechten fehl, landet er in
   * der `DEX_AccessFix`-Queue (Flow setzt ihn nach). Gibt zurück, ob der
   * RegisteredBy-Schreibvorgang gelang.
   */
  public async assignRegistrationToAssistant(
    subsiteUrl: string,
    itemId: number,
    assistantEmail: string,
    assistantName: string
  ): Promise<boolean> {
    try {
      const merge = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'RegisteredByEmail': assistantEmail, 'RegisteredByName': assistantName }
      );
      // Zeilen-Autor auf die Assistenz (Voraussetzung für ILS-Lesezugriff).
      await this.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, itemId, assistantEmail);
      return merge.ok;
    } catch (err) {
      console.warn('[DEX] assignRegistrationToAssistant error:', err);
      return false;
    }
  }

  private async trySetItemAuthor(subsiteUrl: string, listName: string, itemId: number, participantEmail: string): Promise<void> {
    try {
      // 1. Teilnehmer als SP-User der Subsite sicherstellen + dessen Id holen.
      const ensureResp = await this.context.spHttpClient.post(
        `${subsiteUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
          },
          body: JSON.stringify({ logonName: participantEmail }),
        }
      );
      if (!ensureResp.ok) {
        // v20.7: z.B. Assistenz ohne ausreichende Rechte auf der Subsite →
        // Auftrag in die DEX_AccessFix-Queue, der Flow setzt den Autor.
        await this.queueAccessFix(subsiteUrl, itemId, participantEmail);
        return;
      }
      const u = await ensureResp.json();
      const userId: number = u?.Id || u?.d?.Id || 0;
      if (!userId) {
        await this.queueAccessFix(subsiteUrl, itemId, participantEmail);
        return;
      }
      // 2. AuthorId der Zeile auf den Teilnehmer setzen (nometadata-MERGE,
      //    daher KEIN __metadata im Body). 403 = fehlende Rechte → Queue.
      const m = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`,
        { 'AuthorId': userId }
      );
      if (!m.ok) {
        // v20.7: typischer Assistenz-Fall — Contribute reicht nicht, um den
        // Autor zu setzen (braucht "Listen verwalten"). Flow übernimmt.
        await this.queueAccessFix(subsiteUrl, itemId, participantEmail);
      }
    } catch {
      // Best-effort: auch hier den Flow-Auftrag versuchen — scheitert auch
      // der, bleibt die Zeile beim Akteur (Verhalten wie vor v20.5).
      try { await this.queueAccessFix(subsiteUrl, itemId, participantEmail); } catch { /* */ }
    }
  }

  /**
   * v20.6: Reparatur-Werkzeug (Admin) — prüft EINE Teilnehmerliste und
   * repariert den Zugriff bei Fremd-Anmeldungen. Zwei Schritte:
   *
   * 1. **Item-Level-Security verifizieren:** liest ReadSecurity/WriteSecurity
   *    der Liste. Steht sie NICHT auf 2/2 („nur eigene Elemente" — z.B. weil
   *    der Set beim Anlegen still fehlschlug, siehe Security-Audit v20.x),
   *    wird sie neu gesetzt und per Read-back verifiziert.
   * 2. **Fremd-Anmeldungen (Anmeldung durch Dritte):** lädt alle Items mit
   *    Autor und setzt bei jedem Item, dessen `RegisteredByEmail` von der
   *    `ParticipantEmail` abweicht UND dessen Autor noch nicht der Teilnehmer
   *    ist, den Autor auf den Teilnehmer (`ensureuser` → `AuthorId`-MERGE,
   *    pro E-Mail gecacht). Damit sieht die angemeldete Person ihre eigene
   *    Zeile in „Meine Events" und kann sich selbst abmelden (v20.5-Logik,
   *    rückwirkend für Bestands-Anmeldungen).
   *
   * Läuft sequentiell (SP-Throttling-Schonung). `onProgress` meldet den
   * Item-Fortschritt für die UI. Externe Teilnehmer (kein Tenant-Login)
   * scheitern am `ensureuser` und landen in `authorFailed` — erwartbar.
   */
  public async repairProxyRegistrationAccess(
    subsiteUrl: string,
    onProgress?: (done: number, total: number) => void
  ): Promise<{ ilsWasWrong: boolean; ilsFixed: boolean; itemsTotal: number; proxyFound: number; authorFixed: number; authorFailed: number }> {
    const result = { ilsWasWrong: false, ilsFixed: false, itemsTotal: 0, proxyFound: 0, authorFixed: 0, authorFailed: 0 };

    // ---- Schritt 1: Listen-Sicherheit „nur eigene Elemente" sicherstellen ----
    // v21 FIX: Der v20.6-Check meldete fälschlich ALLE Listen als unsicher
    // („27 falsch, 0 repariert"), obwohl die ILS nachweislich aktiv war
    // (Fremd-Zeilen unsichtbar). Ursache: Antwortformat/Typ der
    // ReadSecurity-Property nicht deterministisch behandelt. Jetzt: explizit
    // nometadata anfordern, Werte hart zu Zahlen koerzieren und bei
    // Unplausibilität die ROHE Antwort loggen statt „falsch" zu raten —
    // nur ein KLARER numerischer Wert ungleich 2 zählt als unsicher.
    const readSecurity = async (): Promise<{ rs: number; ws: number } | null> => {
      try {
        const resp = await this.context.spHttpClient.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')?$select=ReadSecurity,WriteSecurity`,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (!resp.ok) return null;
        const d = await resp.json();
        const rawRs = d.ReadSecurity ?? d.d?.ReadSecurity;
        const rawWs = d.WriteSecurity ?? d.d?.WriteSecurity;
        const rs = Number(rawRs);
        const ws = Number(rawWs);
        if (!Number.isFinite(rs) || !Number.isFinite(ws)) {
          console.warn('[DEX][ILS-Check] ReadSecurity/WriteSecurity nicht lesbar — rohe Antwort:', subsiteUrl, JSON.stringify(d).slice(0, 400));
          return null;
        }
        return { rs, ws };
      } catch (e) {
        console.warn('[DEX][ILS-Check] Lesen fehlgeschlagen:', subsiteUrl, e);
        return null;
      }
    };
    const before = await readSecurity();
    if (before && (before.rs !== 2 || before.ws !== 2)) {
      console.warn(`[DEX][ILS-Check] Liste meldet ReadSecurity=${before.rs}/WriteSecurity=${before.ws} (erwartet 2/2):`, subsiteUrl);
      result.ilsWasWrong = true;
      await this.setItemLevelPermissions(subsiteUrl);
      const after = await readSecurity();
      result.ilsFixed = !!after && after.rs === 2 && after.ws === 2;
      if (!result.ilsFixed) {
        console.warn('[DEX][ILS-Check] Read-back nach Fix weiterhin abweichend:', subsiteUrl, after);
      }
    }

    // ---- Schritt 2: Items mit Autor laden (paged) ----
    type Row = { Id: number; ParticipantEmail?: string; RegisteredByEmail?: string; Author?: { EMail?: string } };
    const items: Row[] = [];
    let url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,RegisteredByEmail,Author/EMail&$expand=Author&$top=500`;
    while (url) {
      const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      const arr: Row[] = data.value || data.d?.results || [];
      items.push(...arr);
      url = data['odata.nextLink'] || data['@odata.nextLink'] || data.d?.__next || '';
    }
    result.itemsTotal = items.length;

    // ---- Schritt 3: Fremd-Anmeldungen → Autor auf Teilnehmer setzen ----
    const userIdCache: Record<string, number> = {};
    let done = 0;
    for (const it of items) {
      done++;
      const pe = (it.ParticipantEmail || '').toLowerCase().trim();
      const rb = (it.RegisteredByEmail || '').toLowerCase().trim();
      const au = (it.Author?.EMail || '').toLowerCase().trim();
      // Nur Fremd-Anmeldungen: RegisteredByEmail vorhanden UND != Teilnehmer.
      // (Alt-Bestand ohne RegisteredByEmail = vor v3.x — dort ist der Autor
      // ohnehin der Teilnehmer selbst, weil es nur Selbst-Anmeldung gab.)
      if (!pe || !rb || pe === rb) { if (onProgress) onProgress(done, items.length); continue; }
      result.proxyFound++;
      // Autor stimmt schon (z.B. v20.5-Anmeldung oder früherer Lauf) → ok.
      if (au === pe) { if (onProgress) onProgress(done, items.length); continue; }
      try {
        let uid = userIdCache[pe] || 0;
        if (!uid) {
          const er = await this.context.spHttpClient.post(
            `${subsiteUrl}/_api/web/ensureuser`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata',
              },
              body: JSON.stringify({ logonName: pe }),
            }
          );
          if (er.ok) {
            const u = await er.json();
            uid = u?.Id || u?.d?.Id || 0;
            if (uid) userIdCache[pe] = uid;
          }
        }
        if (!uid) {
          result.authorFailed++;
          if (onProgress) onProgress(done, items.length);
          continue;
        }
        const m = await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${it.Id})`,
          { 'AuthorId': uid }
        );
        if (m.ok) result.authorFixed++; else result.authorFailed++;
      } catch {
        result.authorFailed++;
      }
      if (onProgress) onProgress(done, items.length);
    }
    return result;
  }

  /**
   * Berechtigungen für Teilnehmerliste auf der Subsite setzen.
   */
  private async setRegistrationListPermissions(subsiteUrl: string, organizerEmail: string): Promise<void> {
    try {
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      // Site Owners der Hauptsite: Full Control
      const ownersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
          {}
        );
      }

      // Visitors: Contribute (damit User sich registrieren können)
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`,
          {}
        );
      }

      // Organizer: Full Control. v9.18: organizerEmail kann ";"-separiert
      // mehrere Emails enthalten (Hauptorganizer + Co-Organizer).
      if (organizerEmail) {
        const emails = organizerEmail.split(/[;,]/).map(s => s.trim()).filter(Boolean);
        for (const em of emails) {
          try {
            const userResponse = await this.context.spHttpClient.get(
              `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
              SPHttpClient.configurations.v1
            );
            if (userResponse.ok) {
              const userData = await userResponse.json();
              await this._post(
                `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userData.Id}, roledefid=1073741829)`,
                {}
              );
            }
          } catch { /* Organizer-Berechtigung optional */ }
        }
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden
    }
  }

  // ==================== Registrierungen ====================

  /**
   * Registrierung für ein Event erstellen.
   * Operiert auf der Subsite des Events.
   */
  public async registerForEvent(
    subsiteUrl: string,
    firstName: string,
    surname: string,
    participantEmail: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string>, // cf.id -> SP InternalName
    starterType?: string, // B2Run: effektiver Typ (nach Fallback)
    preferredStarterType?: string, // B2Run: Wunsch-Typ (was der User eigentlich wollte)
    registeredByName?: string, // Audit: Name des Users der die Anmeldung auslöst
    registeredByEmail?: string, // Audit: E-Mail des Users der die Anmeldung auslöst
    // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung — wird in
    // die SP-Spalte ProxyConsent geschrieben (leer bei Selbst-Anmeldung).
    proxyConsent?: string,
    // v19.9: Der CLIENT hat bereits zuverlässig festgestellt, dass der/die
    // Anmeldende Haupt- oder Co-Organizer dieses Events ist (aus dem geladenen
    // Event-Objekt: organizerEmails/coOrganizerEmails ⊇ aktueller User). Diese
    // Information wird hier durchgereicht und hat Vorrang vor der fragilen
    // serverseitigen Ableitung (SubsiteUrl-Filter + Note-Feld-Parsing +
    // pageContext-Identität), die im Tenant gelegentlich fehlschlug und
    // legitime Organizer mit „bereits angemeldet" ablehnte.
    actorIsEventOrganizer: boolean = false,
    // v23.10: Der CLIENT hat bereits validiert, dass der/die Anmeldende eine
    // Assistenz ist UND das Ziel Partner/Director (RegistrationPage prüft das
    // beim Submit gegen den Picker-JobTitle). Diesem Flag vertrauen wir — die
    // serverseitige Ableitung in canRegisterForOthers hängt an
    // Profil-Lookups (Title/SPS-JobTitle), die im Tenant unzuverlässig leer
    // zurückkamen und legitime Assistenzen mit „nicht berechtigt" ablehnten.
    // Gilt NUR für Check A (Berechtigung) — NICHT für die Deadline (Assistenz
    // darf wie ein normaler User nicht nach Frist anmelden).
    clientAssistantAllowed: boolean = false
  // v23.9: Statt nacktem boolean ein konkreter Grund bei Misserfolg, damit die
  // UI nicht mehr pauschal „bereits registriert" anzeigt (irreführend, wenn der
  // echte Grund Berechtigung/Deadline/Insert-Fehler war).
  ): Promise<{ ok: boolean; reason?: 'not-allowed' | 'deadline' | 'insert-failed' | 'error' }> {
    try {
      // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
      // Serverseitige Prüfungen — nicht perfekt (SPFx läuft im Browser),
      // aber fangt naiven App-Bypass (F12, direkter Service-Aufruf) ab.
      const sessionEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const targetEmail = (participantEmail || '').toLowerCase();

      // Event-Metadaten laden (Deadline + OrganizerEmail) ueber SubsiteUrl.
      // Beide Checks nutzen die gleiche Abfrage — einmal laden, mehrfach prüfen.
      let eventDeadline = '';
      let eventOrganizerEmails: string[] = [];
      try {
        const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
        const evResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail`,
          SPHttpClient.configurations.v1
        );
        if (evResp.ok) {
          const evData = await evResp.json();
          const items = evData.value || evData.d?.results || [];
          if (items.length > 0) {
            eventDeadline = items[0].RegistrationDeadline || '';
            const orgStr: string = EventService.stripNoteWrapper(items[0].OrganizerEmail);
            eventOrganizerEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
          }
        }
      } catch { /* Bei Load-Fehler konservativ weitermachen — andere Checks greifen */ }

      // Check A: Darf der User für eine andere Person registrieren?
      // v19.9: Wenn der Client bereits bestätigt hat, dass der/die Anmeldende
      // Organizer/Co-Organizer dieses Events ist, vertrauen wir dem (gleiche
      // Datengrundlage wie die Button-Sichtbarkeit) und überspringen die
      // fragile serverseitige Ableitung. Sonst Fallback auf canRegisterForOthers
      // (deckt Admin-Rolle + Assistant-Ausnahme zuverlässig ab).
      if (targetEmail && targetEmail !== sessionEmail) {
        const allowed = actorIsEventOrganizer || clientAssistantAllowed || await this.canRegisterForOthers(subsiteUrl, participantEmail);
        if (!allowed) {
          console.warn(`[DEX] registerForEvent DENIED: ${sessionEmail} versuchte ${targetEmail} zu registrieren — weder Organizer noch Admin noch erlaubter Assistant-Fall.`);
          return { ok: false, reason: 'not-allowed' };
        }
      }

      // Check B: Deadline abgelaufen? Nur Event-Organizer + Admin dürfen nach
      // Deadline registrieren (auch für sich selbst). Assistant NICHT — das ist
      // wie ein normaler User.
      if (eventDeadline) {
        const deadlineDate = new Date(eventDeadline);
        if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
          const isEventOrganizer = eventOrganizerEmails.indexOf(sessionEmail) >= 0;
          let isAdmin = false;
          try {
            const esc = sessionEmail.replace(/'/g, "''");
            const roleResp = await this.context.spHttpClient.get(
              `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
              SPHttpClient.configurations.v1
            );
            if (roleResp.ok) {
              const rd = await roleResp.json();
              const rItems = rd.value || rd.d?.results || [];
              if (rItems.length > 0 && rItems[0].Role === 'Admin') isAdmin = true;
            }
          } catch { /* ignore */ }

          if (!isEventOrganizer && !isAdmin) {
            console.warn(`[DEX] registerForEvent DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${eventDeadline} zu registrieren — weder Event-Organizer noch Admin.`);
            return { ok: false, reason: 'deadline' };
          }
        }
      }
      // ---- Ende Permission-Checks ----

      // v7.28 / v9.10: Nächste TeilnehmerID atomar ueber den Subsite-Counter
      // holen (ETag-CAS, verhindert Race-Conditions bei parallelen Anmeldungen).
      // Counter wird bei Bedarf on-demand angelegt + geseeded.
      //
      // v9.10: Der alte race-anfällige Fallback "max+1" wurde entfernt — bei
      // Massen-Anmeldungen (Go-Live grosse Events) hat er Duplikate produziert,
      // weil zwei Clients gleichzeitig den gleichen Max-Wert lesen und beide
      // mit Max+1 schreiben. Wenn der atomare Counter ausnahmsweise gar nicht
      // erreichbar ist, lassen wir TeilnehmerID undefined und der Admin
      // lädt anschliessend "IDs neu vergeben" — Lückenfreiheit ist nicht
      // hart kritisch, Eindeutigkeit ist es.
      const nextId = await this.getNextTeilnehmerId(subsiteUrl);

      // Profildaten laden - für den TATSAECHLICHEN Teilnehmer (nicht den eingeloggten User!)
      // Wenn jemand für eine andere Person registriert, muss deren Profil geladen werden,
      // sonst wird der eigene JobTitle/Department/Office in deren Teilnehmer-Eintrag geschrieben.
      const myEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const profile = participantEmail.toLowerCase() === myEmail
        ? await this.getCurrentUserProfile()
        : await this.getUserProfileByEmail(participantEmail);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        // v9.10: TeilnehmerID nur setzen wenn der atomare Counter sie geliefert hat.
        // Bei Counter-Outage bleibt das Feld leer — Admin kann nachträglich
        // "IDs neu vergeben" laufen lassen, was sequentielle IDs setzt.
        ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
        'Anrede': customData.salutation || '',
        'Vorname': firstName,
        'Nachname': surname,
        'ParticipantName': `${firstName} ${surname}`,
        'ParticipantEmail': participantEmail,
        'Department': profile.department,
        'Location': profile.location,
        'JobTitle': profile.jobTitle,
        'Phone': profile.phone,
        // v24.29: Unternehmenszugehörigkeit / Rechtsträger mitschreiben.
        'Company': profile.company,
        'Status': status,
        'RegistrationDate': new Date().toISOString(),
        'CustomData': JSON.stringify(customData),
      };

      // Audit: wer hat die Anmeldung ausgelöst?
      // Bei Self-Registration = der User selbst. Bei "Für andere Person registrieren"
      // = der Organizer/Admin der geklickt hat. Fallback wenn nichts uebergeben: aus pageContext.
      const auditName = registeredByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (registeredByEmail || this.context.pageContext.user.email || '').toLowerCase();
      if (auditName) payload['RegisteredByName'] = auditName;
      if (auditEmail) payload['RegisteredByEmail'] = auditEmail;
      // v18.74: Zustimmungs-Nachweis bei stellvertretender Anmeldung.
      if (proxyConsent) payload['ProxyConsent'] = proxyConsent;

      // B2Run: Starter-Typ + Wunsch-Typ schreiben (bei normalen Events null)
      if (starterType) payload['StarterType'] = starterType;
      if (preferredStarterType) payload['PreferredStarterType'] = preferredStarterType;

      // Custom Field Werte in die echten SP-Spalten schreiben.
      // Wichtig: Wenn spInternalName fehlt (z.B. weil der Admin das Feld später
      // ergänzt hat ohne Spalte in der Teilnehmerliste), würde der Wert
      // SILENT VERLOREN GEHEN — deshalb ein console.warn damit der Admin im
      // Admin Center per "Custom Fields prüfen" das Mapping fixen kann.
      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            payload[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
          }
        }
      }

      let response = await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
        payload
      );
      // v19.9 BUG-FIX: Stellvertretende Anmeldung schlug auf älteren
      // Teilnehmerlisten fehl, weil die erst mit v18.74 eingeführte Spalte
      // `ProxyConsent` dort noch nicht existiert — der Insert mit
      // `ProxyConsent` im Body wird dann von SharePoint mit HTTP 400
      // ("property does not exist") abgewiesen. Da `ProxyConsent` NUR bei
      // Stellvertreter-Anmeldungen im Body steht, scheiterte ausschließlich der
      // Proxy-Pfad (Selbst-Anmeldung lief, weil dort kein ProxyConsent gesetzt
      // wird) — für den User sah es aus wie „Person bereits angemeldet", obwohl
      // gar nichts gespeichert wurde. Fix: Insert einmal OHNE das optionale
      // Audit-Feld wiederholen, damit die Anmeldung nicht an einer fehlenden
      // Spalte scheitert. Der Zustimmungs-Nachweis geht dann verloren (der
      // Admin kann die Spalte per „Spalten fixen" nachrüsten), die Anmeldung
      // selbst gelingt aber.
      // v24.32: Gleiches Schutz-Muster jetzt AUCH für die v24.29-Spalte
      // `Company` — auf Teilnehmerlisten, auf denen „Spalten fixen" noch nicht
      // lief, existiert sie nicht → der Insert mit `Company` im Body würde sonst
      // mit HTTP 400 scheitern und die GANZE Anmeldung kaputtmachen. Deshalb bei
      // einem fehlgeschlagenen Insert das optionale Feld strippen und erneut
      // versuchen. Folge ohne Spalte: Anmeldung gelingt, Unternehmens-Wert wird
      // nicht getrackt (Admin kann „Spalten fixen" nachziehen).
      if (!response.ok && (payload['ProxyConsent'] || payload['Company'])) {
        console.warn('[DEX] registerForEvent: Insert fehlgeschlagen — Retry OHNE ProxyConsent/Company (Spalte evtl. nicht vorhanden). Bitte im Admin Center "Spalten fixen" ausführen.');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const retryPayload: Record<string, any> = { ...payload };
        delete retryPayload['ProxyConsent'];
        delete retryPayload['Company'];
        response = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
          retryPayload
        );
      }
      if (!response.ok) return { ok: false, reason: 'insert-failed' };

      // Inserted-Item-Id EINMALIG aus der Response lesen (der Body lässt sich
      // nur einmal konsumieren) — wird sowohl für die Dedup-Prüfung als auch
      // für das Setzen des Autors (stellvertretende Anmeldung) gebraucht.
      let insertedId = 0;
      try {
        const respJson = await response.json();
        insertedId = respJson?.d?.Id || respJson?.Id || 0;
      } catch { /* Body nicht lesbar — Dedup/Autor-Set entfallen, Insert war ok */ }

      // v9.10: Post-Insert Safety Net — bei Massen-Anmeldungen (Go-Live)
      // gab es trotz ETag-Counter vereinzelt Duplikate. Ursache war der
      // alte max+1-Fallback (jetzt entfernt) und ggf. Edge-Cases im
      // Counter-Pfad. Als zusätzliche Versicherung: nach dem Insert
      // prüfen, ob jetzt zwei Einträge dieselbe TeilnehmerID haben.
      // Wenn ja: der mit der HOEHEREN SP-Item-Id verliert (= der spätere
      // Insert), holt sich frisch eine ID am Counter und patcht sich.
      // So bleiben die zuerst eingetroffenen Anmeldungen stabil.
      if (typeof nextId === 'number' && nextId > 0 && insertedId > 0) {
        try {
          const dupResp = await this.context.spHttpClient.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,TeilnehmerID&$filter=TeilnehmerID eq ${nextId}&$top=10`,
            SPHttpClient.configurations.v1
          );
          if (dupResp.ok) {
            const dupData = await dupResp.json();
            const dupItems: Array<{ Id: number; TeilnehmerID: number }> = dupData.value || dupData.d?.results || [];
            if (dupItems.length > 1) {
              const minId = Math.min(...dupItems.map(d => d.Id));
              if (insertedId !== minId) {
                // Wir haben verloren — fresh ID holen + patchen
                const fresh = await this.getNextTeilnehmerId(subsiteUrl);
                if (typeof fresh === 'number' && fresh > 0) {
                  await this._merge(
                    `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${insertedId})`,
                    { 'TeilnehmerID': fresh }
                  );
                  console.warn(`[DEX] Post-insert dedup: TeilnehmerID ${nextId} kollidierte, Item ${insertedId} hat jetzt #${fresh}.`);
                } else {
                  console.warn(`[DEX] Post-insert dedup: kollidierende TeilnehmerID ${nextId} entdeckt, aber Counter lieferte keine fresh ID. Admin sollte "IDs neu vergeben" laufen lassen.`);
                }
              }
            }
          }
        } catch (err) {
          // Safety-Net-Fehler nicht kritisch — Insert war erfolgreich
          console.warn('[DEX] Post-insert dedup check fehlgeschlagen:', err);
        }
      }

      // v20.5: Stellvertretende Anmeldung (Akteur != Teilnehmer) → den
      // Teilnehmer zum Autor der Zeile machen, damit er seine eigene Anmeldung
      // in "Meine Events" sieht und sich selbst abmelden kann. auditEmail ist
      // bereits lowercased; bei Selbst-Anmeldung sind beide gleich → kein Set.
      if (insertedId > 0 && auditEmail && participantEmail && participantEmail.toLowerCase().trim() !== auditEmail) {
        await this.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, insertedId, participantEmail);
      }

      return { ok: true };
    } catch {
      return { ok: false, reason: 'error' };
    }
  }

  /**
   * v11.82: Ein einzelnes Teilnehmer-Item im Team-Modus anlegen.
   *
   * Unterschied zu `registerForEvent`: kein eigener Permission-Check (der
   * Aufrufer hat schon im Team-Submit alle Mitglieder validiert), kein
   * Post-Insert Dedup-Loop (der ist im Team-Pfad überflüssig — wenn ein
   * Member mit Kollision verliert, fixt es der Folge-IDReorder). Nimmt
   * Profil-Daten und Anzeige-Namen direkt entgegen, weil der Lead-Submit
   * pro Member ohnehin schon das Graph-Profil geladen hat.
   */
  public async registerTeamMember(
    subsiteUrl: string,
    args: {
      firstName: string;
      lastName: string;
      email: string;
      profile: { department: string; location: string; jobTitle: string; phone: string; company?: string };
      status: 'Angemeldet' | 'Warteliste';
      teamId: string;
      teamLead: boolean;
      teamName?: string;
      customData?: Record<string, string>;
      customFieldMap?: Record<string, string>;
      starterType?: string;
      preferredStarterType?: string;
      registeredByName?: string;
      registeredByEmail?: string;
      salutation?: string;
    }
  ): Promise<{ ok: boolean; teilnehmerId?: number; itemId?: number }> {
    try {
      const nextId = await this.getNextTeilnehmerId(subsiteUrl);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': args.email,
        ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
        'Anrede': args.salutation || '',
        'Vorname': args.firstName,
        'Nachname': args.lastName,
        'ParticipantName': `${args.firstName} ${args.lastName}`.trim(),
        'ParticipantEmail': args.email,
        'Department': args.profile.department,
        'Location': args.profile.location,
        'JobTitle': args.profile.jobTitle,
        'Phone': args.profile.phone,
        'Company': args.profile.company || '',
        'Status': args.status,
        'RegistrationDate': new Date().toISOString(),
        'TeamId': args.teamId,
        'TeamLead': !!args.teamLead,
        'TeamName': args.teamName || '',
        'CustomData': JSON.stringify(args.customData || {}),
      };
      if (args.registeredByName) payload['RegisteredByName'] = args.registeredByName;
      if (args.registeredByEmail) payload['RegisteredByEmail'] = args.registeredByEmail;
      if (args.starterType) payload['StarterType'] = args.starterType;
      if (args.preferredStarterType) payload['PreferredStarterType'] = args.preferredStarterType;
      if (args.customFieldMap && args.customData) {
        for (const cfId of Object.keys(args.customData)) {
          if (cfId === 'salutation') continue;
          const v = args.customData[cfId];
          if (!v) continue;
          const spName = args.customFieldMap[cfId];
          if (spName) payload[spName] = v;
        }
      }
      let response = await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
        payload
      );
      // v24.32: Retry OHNE Company, falls die Spalte auf der Liste fehlt (s.
      // registerForEvent) — sonst bräche die Team-Anmeldung auf Alt-Listen.
      if (!response.ok && payload['Company']) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const retryPayload: Record<string, any> = { ...payload };
        delete retryPayload['Company'];
        response = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
          retryPayload
        );
      }
      if (!response.ok) return { ok: false };
      try {
        const respJson = await response.json();
        const itemId: number = respJson?.d?.Id || respJson?.Id || 0;
        return { ok: true, teilnehmerId: typeof nextId === 'number' ? nextId : undefined, itemId };
      } catch {
        return { ok: true, teilnehmerId: typeof nextId === 'number' ? nextId : undefined };
      }
    } catch {
      return { ok: false };
    }
  }

  /**
   * v11.82: Alle Mitglieder eines Teams (per TeamId) zu einer Registrierung
   * laden — wird in „Meine Events" zum Rendern des Team-Badges genutzt.
   */
  public async getTeamMembers(subsiteUrl: string, teamId: string): Promise<SPRegistration[]> {
    if (!teamId) return [];
    try {
      const tidEsc = teamId.replace(/'/g, "''");
      const response = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=TeamId eq '${tidEsc}'&$top=100&$orderby=TeamLead desc,Id asc`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return [];
      const data = await response.json();
      return data.value || data.d?.results || [];
    } catch {
      return [];
    }
  }

  /**
   * v11.83: Auf einer existierenden Teilnehmer-Zeile das Feld TeamLead
   * auf true setzen (Auto-Promote nach Lead-Cancel). MERGE auf der
   * Teilnehmerliste — die Subsite kennt das Item ueber `itemId`.
   */
  public async promoteToTeamLead(subsiteUrl: string, itemId: number): Promise<boolean> {
    try {
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
      const resp = await this._merge(url, { TeamLead: true });
      return !!resp.ok;
    } catch {
      return false;
    }
  }

  /**
   * v11.84: Lead-Rolle innerhalb eines Teams von einer Person auf eine andere
   * uebergeben. Wird im Admin Center per Dropdown im Teams-Block ausgelöst.
   * Best-effort transaktional: erst die neue Lead-Zeile auf TeamLead=true
   * setzen, danach die alte auf TeamLead=false. Schlägt der zweite MERGE
   * fehl, gibt es kurzfristig zwei Leads — der Aufrufer kann dann erneut
   * versuchen oder die Liste manuell reparieren. Keine echte Transaktion,
   * SharePoint bietet sowas auf Listen-Ebene nicht.
   */
  public async transferTeamLead(
    subsiteUrl: string,
    fromLeadItemId: number,
    toNewLeadItemId: number
  ): Promise<boolean> {
    if (!subsiteUrl || !fromLeadItemId || !toNewLeadItemId || fromLeadItemId === toNewLeadItemId) {
      return false;
    }
    try {
      const newUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${toNewLeadItemId})`;
      const r1 = await this._merge(newUrl, { TeamLead: true });
      if (!r1.ok) return false;
      const oldUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${fromLeadItemId})`;
      const r2 = await this._merge(oldUrl, { TeamLead: false });
      return !!r2.ok;
    } catch {
      return false;
    }
  }

  /**
   * v11.83: Prüfen, ob eine bestimmte Email-Adresse bereits aktiv beim
   * Event angemeldet ist (Status in Angemeldet/QR versendet/Eingecheckt/
   * Warteliste). Wird vor jedem Team-Add (Initial, Add-Member, Beitritt)
   * benutzt, um Doppel-Anmeldungen sauber abzuweisen, bevor ein Sitzplatz
   * reserviert wird.
   *
   * Rückgabe: true = blockieren, false = frei (auch bei SP-Fehlern, weil
   * der Aufrufer dann auf die strikteren Stellen-internen Checks zurück-
   * fällt; ein lauter Throw würde den Pfad unnötig abbrechen).
   */
  public async isUserAlreadyOnEvent(subsiteUrl: string, email: string): Promise<boolean> {
    if (!subsiteUrl || !email) return false;
    try {
      const emEsc = email.trim().replace(/'/g, "''");
      const blockingStatuses = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
      const statusClause = blockingStatuses.map(s => `Status eq '${s}'`).join(' or ');
      const filter = `(ParticipantEmail eq '${emEsc}') and (${statusClause})`;
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${encodeURIComponent(filter)}&$top=1&$select=Id,Status,ParticipantEmail`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return false;
      const data = await response.json();
      const items = data.value || data.d?.results || [];
      return items.length > 0;
    } catch {
      return false;
    }
  }

  // ==================== DEX_TeamJoinRequests (v11.83) ====================

  /**
   * v11.83: Globale Liste für Team-Beitritts-Anfragen (Approve-Queue).
   * Liegt auf der Site-Collection-Ebene (nicht pro Subsite), damit alle
   * Events darauf zugreifen können und der Team-Lead alle ausstehenden
   * Anfragen in einer einzigen Query findet.
   *
   * Spalten:
   * - Title: Anzeige-Zusammenfassung "RequesterName -> Event-Title"
   * - EventId: ID des Events in DEX_Events
   * - TeamId: UUID der Team-Anmeldung
   * - RequesterEmail: Email des Anfragenden
   * - RequesterDisplayName: Anzeigename des Anfragenden
   * - Status: Pending / Approved / Rejected
   * - DecidedDate: Wann hat der Team-Lead entschieden
   * - DecidedByEmail: Email des entscheidenden Leads
   */
  public async ensureTeamJoinRequestsList(): Promise<void> {
    const listName = 'DEX_TeamJoinRequests';
    const exists = await this.listExists(listName);
    if (exists) {
      // v13.0: Backfill für ältere Installationen, die die Liste vor
      // v11.83 angelegt haben (DecidedDate/DecidedByEmail damals nicht
      // vorhanden). Ohne diesen Patch schlägt decideTeamJoinRequest
      // beim MERGE auf die fehlenden Felder mit HTTP 400 fehl.
      await this.ensureMissingTeamJoinRequestsFields(listName);
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Approve-Queue für Team-Beitritts-Anfragen (v11.83+).',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'EventId', type: 2 },
      { title: 'TeamId', type: 2 },
      { title: 'RequesterEmail', type: 2 },
      { title: 'RequesterDisplayName', type: 2 },
      { title: 'Status', type: 6, choices: ['Pending', 'Approved', 'Rejected'], metaType: 'SP.FieldChoice' },
      { title: 'DecidedDate', type: 4 },
      { title: 'DecidedByEmail', type: 2 },
      // v18.73: event-spezifische Antworten des Anfragenden als JSON — werden
      // beim Approve auf den neuen Member angewandt (Note = multi-line text).
      { title: 'CustomData', type: 3 },
    ];

    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) payload['Choices'] = { 'results': f.choices };
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    await this.configureDefaultView(listName, [
      'EventId', 'TeamId', 'RequesterEmail', 'RequesterDisplayName',
      'Status', 'Created', 'DecidedDate', 'DecidedByEmail',
    ]);

    // Schreibrechte für alle Authentifizierten (analog zu DEX_Emails-Queue):
    // jeder darf eine Anfrage erstellen, aber Item-Level-Security greift
    // sowieso ueber den Lead-Check beim Approve-Pfad.
    try {
      await this.setQueueListPermissions(listName);
    } catch { /* best-effort */ }
  }

  /**
   * v13.0: Backfill fehlender Felder in einer bestehenden DEX_TeamJoinRequests-
   * Liste. Greift bei Tenants die die Liste vor v11.83 angelegt haben.
   */
  private async ensureMissingTeamJoinRequestsFields(listName: string): Promise<void> {
    const wanted = [
      { title: 'DecidedDate', type: 4 },
      { title: 'DecidedByEmail', type: 2 },
      // v18.73: CustomData-Spalte auf Bestands-Listen nachziehen.
      { title: 'CustomData', type: 3 },
    ];
    for (const f of wanted) {
      try {
        const resp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${f.title}')?$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (resp.ok) continue; // existiert
      } catch { /* anlegen */ }
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        });
      } catch (e) {
        console.warn(`[DEX] ensureMissingTeamJoinRequestsFields: failed to add '${f.title}':`, e);
      }
    }
  }

  /**
   * v18.48: Sperr-Liste für den Outlook-Einladungs-Flow (DEX_Outlook_Einladungen).
   *
   * Hintergrund: der Einladungs-Flow patcht pro Anmeldung/Abmeldung die
   * KOMPLETTE Teilnehmerliste eines Outlook-Termins (bei Grossevents bis zu
   * 1500 Personen) an Microsoft Graph. Lief der Flow seriell (Concurrency 1),
   * standen Anmeldungen für völlig UNTERSCHIEDLICHE Events stundenlang in
   * der Warteschlange. Lösung „Option B": die Trigger-Concurrency wird hoch-
   * gesetzt (z.B. 25 parallele Läufe), und ein Pro-Event-Lock verhindert,
   * dass zwei Läufe für dasSELBE Event gleichzeitig die Teilnehmerliste
   * lesen-und-schreiben (Race -> verlorene Einträge).
   *
   * Die Liste hat eine eindeutige (Enforce-Unique) Spalte `EventId`. Der Flow
   * „erwirbt" den Lock per Create-Item: gelingt das Create, hat er den Lock;
   * schlägt es wegen der Eindeutigkeits-Prüfung fehl, hält gerade ein
   * anderer Lauf desselben Events den Lock -> kurz warten und erneut
   * versuchen. Am Ende löscht der Flow das Lock-Item wieder (Release).
   *
   * Die UI-Schritt-für-Schritt-Anleitung steht in `docs/flow-jsons.md` unter
   * „UI-Anleitung 2026-06-02 (v18.48) — Option B: Pro-Event-Lock für
   * parallele Outlook-Läufe".
   */
  /**
   * v20.7: Queue-Liste `DEX_AccessFix` (Site-Collection-Root) für den
   * Assistenz-Fall der Fremd-Anmeldung. Wenn `trySetItemAuthor` mangels
   * „Listen verwalten"-Rechten scheitert (normaler Contribute-User, z.B.
   * Assistenz meldet einen Partner an), schreibt die App hier einen
   * Auftrag — der Power-Automate-Flow `DEX_AccessFix_Autor` (läuft mit
   * Service-Identität, hat Full Control) setzt dann den Zeilen-Autor auf
   * den Teilnehmer und markiert den Auftrag als Done/Failed.
   * Spalten: SubsiteUrl (Text), ItemId (Number), ParticipantEmail (Text),
   * Status (Text: Pending/Done/Failed). Schreibrechte für alle User via
   * setQueueListPermissions (analog DEX_Emails).
   */
  public async ensureAccessFixList(): Promise<void> {
    const listName = 'DEX_AccessFix';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Queue: Zeilen-Autor bei Fremd-Anmeldungen auf den Teilnehmer setzen (v20.7, Flow DEX_AccessFix_Autor).',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    const fields: Array<{ title: string; type: number }> = [
      { title: 'SubsiteUrl', type: 2 },
      { title: 'ItemId', type: 9 },
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Status', type: 2 },
    ];
    for (const f of fields) {
      try {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        });
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
    try {
      await this.configureDefaultView(listName, ['SubsiteUrl', 'ItemId', 'ParticipantEmail', 'Status', 'Created']);
    } catch { /* View optional */ }
    try {
      await this.setQueueListPermissions(listName);
    } catch { /* best-effort */ }
    // Item-Level-Security: User sehen/ändern nur EIGENE Aufträge (sonst wäre
    // ablesbar, wer wen angemeldet hat). Der Flow läuft als Site-Owner und
    // sieht alle Items („Listen verwalten" hebelt die Item-Beschränkung aus).
    try {
      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'ReadSecurity': 2, 'WriteSecurity': 2,
          }),
        }
      );
    } catch { /* best-effort */ }
  }

  /**
   * v20.7: Auftrag für den DEX_AccessFix_Autor-Flow einreihen (siehe
   * ensureAccessFixList). Best-effort — Fehler blocken die Anmeldung nie.
   */
  public async queueAccessFix(subsiteUrl: string, itemId: number, participantEmail: string): Promise<void> {
    try {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_AccessFix')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_AccessFixListItem' },
        'Title': `${participantEmail} -> Item ${itemId}`.slice(0, 250),
        'SubsiteUrl': subsiteUrl,
        'ItemId': itemId,
        'ParticipantEmail': participantEmail,
        'Status': 'Pending',
      });
    } catch (err) {
      console.warn('[DEX] queueAccessFix fehlgeschlagen (best-effort):', err);
    }
  }

  public async ensureOutlookLocksList(): Promise<void> {
    const listName = 'DEX_OutlookLocks';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Pro-Event-Sperre für den Outlook-Einladungs-Flow (v18.48) — verhindert gleichzeitige Läufe für dasselbe Event.',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // EventId: der Lock-Schlüssel. Muss eindeutig + indiziert sein, damit das
    // Create-als-Lock-Erwerb-Muster funktioniert (zweiter gleichzeitiger
    // Create für dieselbe EventId schlägt fehl -> der Lauf wartet & retryt).
    await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': 'EventId',
      'FieldTypeKind': 2,
      'Required': false,
    });
    // LockedAt: rein informativ (Debugging hängengebliebener Locks).
    await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': 'LockedAt',
      'FieldTypeKind': 4,
      'Required': false,
    });

    // EventId indizieren und Eindeutigkeit erzwingen. Reihenfolge wichtig:
    // erst Indexed, dann EnforceUniqueValues (SP verlangt eine indizierte
    // Spalte für die Eindeutigkeits-Prüfung). Auf einer frischen, leeren
    // Liste ist das unkritisch (keine Duplikate vorhanden).
    const fieldUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('EventId')`;
    try {
      await this._merge(fieldUrl, { 'Indexed': true });
      await this._merge(fieldUrl, { 'EnforceUniqueValues': true });
    } catch (e) {
      console.warn('[DEX] ensureOutlookLocksList: EnforceUniqueValues konnte nicht gesetzt werden:', e);
    }

    await this.configureDefaultView(listName, ['EventId', 'LockedAt']);

    // Schreibrechte wie bei den anderen Queue-Listen (DEX_Emails etc.) —
    // der Flow-Connection-Account muss Lock-Items anlegen/löschen können.
    try {
      await this.setQueueListPermissions(listName);
    } catch { /* best-effort */ }
  }

  /**
   * v11.83: Neue Team-Beitritts-Anfrage anlegen.
   */
  public async createTeamJoinRequest(args: {
    eventId: string;
    eventTitle: string;
    teamId: string;
    requesterEmail: string;
    requesterDisplayName: string;
    // v18.73: event-spezifische Antworten als JSON (optional).
    customData?: string;
  }): Promise<{ ok: boolean; itemId?: number }> {
    try {
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_TeamJoinRequestsListItem' },
        'Title': `${args.requesterDisplayName} -> ${args.eventTitle}`.slice(0, 250),
        'EventId': args.eventId,
        'TeamId': args.teamId,
        'RequesterEmail': args.requesterEmail,
        'RequesterDisplayName': args.requesterDisplayName,
        'Status': 'Pending',
        ...(args.customData ? { 'CustomData': args.customData } : {}),
      };
      const resp = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items`,
        payload
      );
      if (!resp.ok) return { ok: false };
      try {
        const j = await resp.json();
        const id: number = j?.d?.Id || j?.Id || 0;
        return { ok: true, itemId: id };
      } catch {
        return { ok: true };
      }
    } catch {
      return { ok: false };
    }
  }

  /**
   * v11.83: Alle Pending-Beitritts-Anfragen — optional gefiltert nach
   * Event und/oder Team. Wird für die "Beitritts-Anfragen"-Box im
   * Team-Lead-UI in MyEventsPage aufgerufen.
   */
  public async listTeamJoinRequests(args: {
    eventId?: string;
    teamId?: string;
    status?: 'Pending' | 'Approved' | 'Rejected';
  }): Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string; DecidedDate?: string; DecidedByEmail?: string; CustomData?: string }>> {
    try {
      const clauses: string[] = [];
      if (args.eventId) clauses.push(`EventId eq '${args.eventId.replace(/'/g, "''")}'`);
      if (args.teamId) clauses.push(`TeamId eq '${args.teamId.replace(/'/g, "''")}'`);
      clauses.push(`Status eq '${args.status || 'Pending'}'`);
      const filter = clauses.join(' and ');
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items?$filter=${encodeURIComponent(filter)}&$top=200&$orderby=Created asc`;
      const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return [];
      const data = await resp.json();
      return data.value || data.d?.results || [];
    } catch {
      return [];
    }
  }

  /**
   * v11.83: Approve/Reject einer Beitritts-Anfrage — schreibt Status,
   * DecidedDate und DecidedByEmail. Die Folge-Logik (Member-Insert,
   * Mails) liegt im EventContext, weil dort die Subsite-/Event-Lookups
   * verfügbar sind.
   */
  public async decideTeamJoinRequest(
    requestId: number,
    decision: 'Approved' | 'Rejected',
    decidedByEmail: string
  ): Promise<boolean> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items(${requestId})`;
      const body = {
        'Status': decision,
        'DecidedDate': new Date().toISOString(),
        'DecidedByEmail': decidedByEmail || '',
      };
      const resp = await this._merge(url, body);
      return !!resp.ok;
    } catch {
      return false;
    }
  }

  /**
   * Bestehende abgemeldete Registrierung reaktivieren.
   * Setzt Status zurück auf Angemeldet/Warteliste, löscht CancellationDate,
   * aktualisiert RegistrationDate und CustomData.
   */
  public async reactivateRegistration(
    subsiteUrl: string,
    itemId: number,
    firstName: string,
    surname: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string>,
    registeredByName?: string, // Audit: Name des Users der die Re-Anmeldung auslöst
    registeredByEmail?: string, // Audit: E-Mail des Users der die Re-Anmeldung auslöst
    proxyConsent?: string // v18.74: Zustimmungs-Nachweis bei stellvertretender Re-Anmeldung
  ): Promise<boolean> {
    try {
      // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
      // Lade die ParticipantEmail aus dem zu reaktivierenden Item und prüfe,
      // ob der aktuelle User dafür berechtigt ist. Plus Deadline-Check.
      // v7.30: Wir laden hier zusätzlich die existierende TeilnehmerID, damit
      // beim Reaktivieren die alte ID erhalten bleibt — Counter wird NUR
      // dann angefasst, wenn die alte ID null/0 ist (Legacy-Edge).
      let existingTeilnehmerId = 0;
      // v20.5: Teilnehmer-E-Mail in den aeusseren Scope heben, damit der
      // Autor-Set am Ende (stellvertretende Re-Anmeldung) sie nutzen kann.
      let reactivateParticipantEmail = '';
      try {
        const itemResp = await this.context.spHttpClient.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${itemId})?$select=ParticipantEmail,TeilnehmerID`,
          SPHttpClient.configurations.v1
        );
        const sessionEmail = (this.context.pageContext.user.email || '').toLowerCase();
        let targetEmail = '';
        if (itemResp.ok) {
          const itemData = await itemResp.json();
          targetEmail = (itemData.ParticipantEmail || itemData.d?.ParticipantEmail || '').toLowerCase();
          reactivateParticipantEmail = targetEmail;
          const tnId = itemData.TeilnehmerID ?? itemData.d?.TeilnehmerID;
          if (typeof tnId === 'number' && tnId > 0) existingTeilnehmerId = tnId;
        }

        // Check A: für andere Person registrieren?
        if (targetEmail && targetEmail !== sessionEmail) {
          const allowed = await this.canRegisterForOthers(subsiteUrl, targetEmail);
          if (!allowed) {
            console.warn(`[DEX] reactivateRegistration DENIED: ${sessionEmail} versuchte ${targetEmail} zu reaktivieren — nicht berechtigt.`);
            return false;
          }
        }

        // Check B: Deadline-Check (Event ueber SubsiteUrl finden)
        const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
        const evResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail`,
          SPHttpClient.configurations.v1
        );
        if (evResp.ok) {
          const evData = await evResp.json();
          const items = evData.value || evData.d?.results || [];
          if (items.length > 0) {
            const deadline = items[0].RegistrationDeadline || '';
            const orgStr: string = EventService.stripNoteWrapper(items[0].OrganizerEmail);
            const orgEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
            if (deadline) {
              const deadlineDate = new Date(deadline);
              if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
                const isEventOrganizer = orgEmails.indexOf(sessionEmail) >= 0;
                let isAdmin = false;
                try {
                  const esc = sessionEmail.replace(/'/g, "''");
                  const roleResp = await this.context.spHttpClient.get(
                    `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
                    SPHttpClient.configurations.v1
                  );
                  if (roleResp.ok) {
                    const rd = await roleResp.json();
                    const rItems = rd.value || rd.d?.results || [];
                    if (rItems.length > 0 && rItems[0].Role === 'Admin') isAdmin = true;
                  }
                } catch { /* ignore */ }
                if (!isEventOrganizer && !isAdmin) {
                  console.warn(`[DEX] reactivateRegistration DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${deadline} zu reaktivieren.`);
                  return false;
                }
              }
            }
          }
        }
      } catch { /* bei Load-Fehler konservativ: weitermachen */ }
      // ---- Ende Permission-Checks ----

      // Reaktivierung = funktional eine Neuanmeldung mit existierendem Listen-
      // Item. Deshalb wird hier — analog zu registerForEvent — atomar eine
      // neue TeilnehmerID am Counter gezogen. Wer mal #12 war und reaktiviert,
      // bekommt jetzt z.B. die #87, also die nächst-freie ID am Ende der
      // Liste — exakt wie ein Neuzugang. Ohne diesen Schritt blieb der
      // Eintrag mit TeilnehmerID=null hängen, weil im Reaktivierungs-Pfad
      // niemand den DEX_IDReorder-Flow triggert.
      // v9.10: race-anfälliger max+1-Fallback entfernt (siehe Kommentar in
      // registerForEvent). Bei Counter-Outage bleibt TeilnehmerID undefined.
      void existingTeilnehmerId;
      const nextId = await this.getNextTeilnehmerId(subsiteUrl);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'Vorname': firstName,
        'Nachname': surname,
        'ParticipantName': `${firstName} ${surname}`,
        'Status': status,
        // v9.10: TeilnehmerID nur setzen wenn Counter sie geliefert hat.
        ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
        'RegistrationDate': new Date().toISOString(),
        'CancellationDate': null,
        'CustomData': JSON.stringify(customData),
      };

      // Audit: wer hat die Re-Anmeldung ausgelöst? (ueberschreibt den Wert von
      // der ursprünglichen Anmeldung, weil das faktisch eine neue Anmeldung ist)
      const auditName = registeredByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (registeredByEmail || this.context.pageContext.user.email || '').toLowerCase();
      if (auditName) body['RegisteredByName'] = auditName;
      if (auditEmail) body['RegisteredByEmail'] = auditEmail;
      // v18.74: Zustimmungs-Nachweis bei stellvertretender Re-Anmeldung.
      if (proxyConsent) body['ProxyConsent'] = proxyConsent;

      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            body[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
          }
        }
      }

      let response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      // v19.9: Wie bei registerForEvent — falls die ProxyConsent-Spalte auf
      // diesem (älteren) Event noch fehlt, scheitert der MERGE mit HTTP 400.
      // Einmal ohne das optionale Audit-Feld wiederholen, damit die
      // Re-Aktivierung nicht an einer fehlenden Spalte scheitert.
      if (!response.ok && body['ProxyConsent']) {
        console.warn('[DEX] reactivateRegistration: MERGE fehlgeschlagen — Retry OHNE ProxyConsent (Spalte evtl. nicht vorhanden). Bitte im Admin Center "Spalten fixen" ausführen.');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const retryBody: Record<string, any> = { ...body };
        delete retryBody['ProxyConsent'];
        response = await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
          retryBody
        );
      }
      if (!response.ok) return false;

      // v9.10: Post-Update Safety Net (siehe registerForEvent). Bei
      // Counter-Edge-Cases könnte der nächste Wert kollidieren — der
      // aeltere Eintrag (kleinere SP-Id) gewinnt, der spätere bekommt
      // fresh ID.
      if (typeof nextId === 'number' && nextId > 0) {
        try {
          const dupResp = await this.context.spHttpClient.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,TeilnehmerID&$filter=TeilnehmerID eq ${nextId}&$top=10`,
            SPHttpClient.configurations.v1
          );
          if (dupResp.ok) {
            const dupData = await dupResp.json();
            const dupItems: Array<{ Id: number; TeilnehmerID: number }> = dupData.value || dupData.d?.results || [];
            if (dupItems.length > 1) {
              const minId = Math.min(...dupItems.map(d => d.Id));
              if (itemId !== minId) {
                const fresh = await this.getNextTeilnehmerId(subsiteUrl);
                if (typeof fresh === 'number' && fresh > 0) {
                  await this._merge(
                    `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
                    { 'TeilnehmerID': fresh }
                  );
                  console.warn(`[DEX] Post-update dedup (reactivate): TeilnehmerID ${nextId} kollidierte, Item ${itemId} hat jetzt #${fresh}.`);
                }
              }
            }
          }
        } catch (err) {
          console.warn('[DEX] Post-update dedup check (reactivate) fehlgeschlagen:', err);
        }
      }

      // v20.5: Stellvertretende Re-Anmeldung → Teilnehmer zum Autor der Zeile
      // machen (analog registerForEvent), damit er sie in "Meine Events" sieht
      // und sich selbst abmelden kann. Best-effort (nur mit "Listen verwalten").
      if (reactivateParticipantEmail && auditEmail && reactivateParticipantEmail !== auditEmail) {
        await this.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, itemId, reactivateParticipantEmail);
      }

      return true;
    } catch {
      return false;
    }
  }

  /**
   * Schlanker MERGE-Helper auf ein einzelnes Teilnehmerlisten-Item — baut
   * keine ChangeLog-Logik, keine FieldMap-Auflösung, keine Default-Felder ein.
   * Genutzt für One-Shot-Migrationen (z.B. T-Shirt-Größen-Import), die direkt
   * bestimmte Felder (inkl. CustomData-JSON + einzelne SP-Spalten) setzen wollen.
   */
  public async mergeRegistrationFields(
    subsiteUrl: string,
    itemId: number,
    body: Record<string, unknown>
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Custom Data einer Registrierung aktualisieren (Teilnehmer aendert eigene Angaben).
   */
  /**
   * v17.2: Bestehende Teilnehmer-Registrierung einem Team zuordnen
   * (PATCH der TeamId/TeamName/TeamLead-Felder auf einem schon existierenden
   * Item). Wird vom Admin-Center-Team-Management genutzt, wenn der
   * Organizer einen schon Angemeldeten ohne Team einem (neuen) Team
   * zuweist — vermeidet doppelte Anmeldung + Mail/Outlook-Spam.
   */
  public async assignRegistrationToTeam(
    subsiteUrl: string,
    itemId: number,
    teamId: string,
    teamName: string | undefined,
    isLead: boolean,
  ): Promise<boolean> {
    try {
      const body: Record<string, unknown> = {
        TeamId: teamId,
        TeamName: teamName || '',
        TeamLead: !!isLead,
      };
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] assignRegistrationToTeam failed:', err);
      return false;
    }
  }

  public async updateRegistrationData(
    subsiteUrl: string,
    itemId: number,
    customData: Record<string, string>,
    customFieldMap?: Record<string, string>,
    oldCustomData?: Record<string, string>,
    fieldLabelMap?: Record<string, string> // cf.id -> label
  ): Promise<boolean> {
    try {
      // Änderungen ermitteln
      const changes: string[] = [];
      const now = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      if (oldCustomData && fieldLabelMap) {
        for (const key of Object.keys(customData)) {
          if (key === 'salutation') continue;
          const label = fieldLabelMap[key] || key;
          const oldVal = oldCustomData[key] || '';
          const newVal = customData[key] || '';
          if (oldVal !== newVal) {
            changes.push(`${label}: "${oldVal}" → "${newVal}"`);
          }
        }
      }
      const changeEntry = changes.length > 0 ? `[${now}] ${changes.join(', ')}` : '';

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'CustomData': JSON.stringify(customData),
        'LastModifiedDate': new Date().toISOString(),
      };

      // ChangeLog anhängen (bestehenden Log behalten)
      if (changeEntry) {
        try {
          const existing = await this.context.spHttpClient.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
            SPHttpClient.configurations.v1
          );
          if (existing.ok) {
            const data = await existing.json();
            const oldLog = data.ChangeLog || '';
            body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
          }
        } catch {
          body['ChangeLog'] = changeEntry;
        }
      }

      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            body[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
          }
        }
      }

      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * v8.0: Admin/Organizer kann Teilnehmerdaten direkt aus dem Admin Center
   * editieren. Erlaubt das Aendern von Anrede/Vorname/Nachname/Email/Phone/
   * Department/Location/JobTitle/Status sowie aller Custom-Felder. Schreibt
   * automatisch ChangeLog-Eintrag mit Wer/Wann/Was-Diff und setzt
   * LastModifiedDate.
   *
   * patch: Nur die echten Spalten-Werte (keine __metadata nötig — _merge
   * sendet odata=nometadata).
   * actor: Audit-Info des aufrufenden Users.
   * oldValues: zum Diff-Bauen, nur Felder mit oldValues[key] !== patch[key]
   * landen im ChangeLog.
   * fieldLabelMap: optional, mappt internal column name -> display label.
   */
  public async adminUpdateRegistration(
    subsiteUrl: string,
    itemId: number,
    patch: Record<string, unknown>,
    actor: { name: string; email: string },
    oldValues?: Record<string, unknown>,
    fieldLabelMap?: Record<string, string>
  ): Promise<boolean> {
    try {
      const changes: string[] = [];
      const now = new Date().toLocaleString('de-DE', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit',
      });
      if (oldValues) {
        for (const key of Object.keys(patch)) {
          const oldV = oldValues[key];
          const newV = patch[key];
          // Vergleich als String, damit number vs string nicht stört
          const oldStr = oldV === null || oldV === undefined ? '' : String(oldV);
          const newStr = newV === null || newV === undefined ? '' : String(newV);
          if (oldStr !== newStr) {
            const label = (fieldLabelMap && fieldLabelMap[key]) || key;
            changes.push(`${label}: "${oldStr}" → "${newStr}"`);
          }
        }
      }
      const changeEntry = changes.length > 0
        ? `[${now}] ${actor.name || actor.email}: ${changes.join(', ')}`
        : '';

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = { ...patch, 'LastModifiedDate': new Date().toISOString() };

      // ChangeLog anhängen (bestehenden Log behalten, neuestes oben)
      if (changeEntry) {
        try {
          const existing = await this.context.spHttpClient.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
            SPHttpClient.configurations.v1
          );
          if (existing.ok) {
            const data = await existing.json();
            const oldLog = data.ChangeLog || data.d?.ChangeLog || '';
            body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
          } else {
            body['ChangeLog'] = changeEntry;
          }
        } catch {
          body['ChangeLog'] = changeEntry;
        }
      }

      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch (err) {
      console.warn('[DEX] adminUpdateRegistration error:', err);
      return false;
    }
  }

  /**
   * Eigene Registrierung für ein Event laden
   */
  public async getMyRegistration(
    subsiteUrl: string,
    email: string
  ): Promise<SPRegistration | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.replace(/'/g, "''")}'&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      if (data.value && data.value.length > 0) return data.value[0];
      return null;
    } catch {
      return null;
    }
  }

  /**
   * v24.36: Alle Anmeldungen einer Teilnehmerliste, die der eingeloggte User
   * STELLVERTRETEND für eine andere Person durchgeführt hat (Assistenz-Funktion).
   * Filter: `RegisteredByEmail = Akteur` UND `ParticipantEmail ≠ Akteur`.
   *
   * Hinweis zur Item-Level-Security: Auf den Teilnehmerlisten greift
   * ReadSecurity=2 („nur eigene Elemente"). Eine Assistenz (Contribute-User)
   * sieht daher nur die Zeilen, deren SharePoint-Autor sie selbst ist. Nach dem
   * v20.5-Autor-Wechsel (AuthorId → Teilnehmer) verliert die Assistenz den
   * Lesezugriff auf genau diese Zeilen — solange der `DEX_AccessFix`-Autor-Flow
   * NICHT eingerichtet ist (häufigster Fall), bleibt die Assistenz Autor und
   * sieht ihre Fremd-Anmeldungen weiterhin. Admin/Organizer eigener Events
   * sehen ohnehin alle Zeilen. Best-effort: liefert nur, was der Server
   * tatsächlich zurückgibt.
   */
  public async getProxyRegistrationsByActor(
    subsiteUrl: string,
    actorEmail: string
  ): Promise<SPRegistration[]> {
    const me = (actorEmail || '').toLowerCase().trim();
    if (!me) return [];
    const out: SPRegistration[] = [];
    const esc = me.replace(/'/g, "''");
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=RegisteredByEmail eq '${esc}'&$orderby=Id asc&$top=5000`;
    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        const page: SPRegistration[] = data.value || data.d?.results || [];
        for (const r of page) {
          const pe = (r.ParticipantEmail || '').toLowerCase().trim();
          if (pe && pe !== me) out.push(r);
        }
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    return out;
  }

  /**
   * Alle Registrierungen für ein Event laden (nur für Organizer/Admin)
   */
  public async getAllRegistrations(subsiteUrl: string): Promise<SPRegistration[]> {
    const allItems: SPRegistration[] = [];
    // $top=5000 ist das SP-REST-Maximum pro Page. Damit fallen Events bis zu
    // 5000 Teilnehmern in einen einzigen Response — keine Pagination-Edgecases
    // mit fehlendem nextLink. Bei größeren Listen folgen wir dem nextLink
    // weiter (Schleife unten). Vorher stand hier $top=500, was bei Events mit
    // ≥500 Teilnehmern zu fehlenden Einträgen führte: SharePoint liefert
    // bei $orderby+$top in Kombination mit Item-Level-Security nicht
    // zuverlässig nextLink, wenn die erste Page exakt voll ist.
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$orderby=Id asc&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        // Beide OData-Formate abdecken: nometadata (data.value) UND verbose
        // (data.d.results). Vorher nur data.value — bei verbose-Response
        // wären null Items dazugekommen.
        const page = data.value || data.d?.results || [];
        allItems.push(...page);
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    return allItems;
  }

  /**
   * TeilnehmerIDs sequentiell neu vergeben (1, 2, 3, ...).
   *
   * Seit v6.5: Zwei-Pass-Reorder.
   * 1. Erst alle **Angemeldeten** (Status ∈ Angemeldet / QR versendet / Eingecheckt)
   *    in Reihenfolge der Registrierung (SP-ItemId asc) → bekommen IDs 1..N.
   * 2. Danach alle **Warteliste**-Teilnehmer in Reihenfolge der Registrierung
   *    → bekommen IDs N+1..N+M (Warteliste hängt lückenlos hinten an).
   * 3. Abgemeldete bekommen TeilnehmerID = null.
   *
   * Damit stehen im Teilnehmerlisten-Grid die Angemeldeten sauber oben, die
   * Warteliste sauber unten — kein Durchmischen mehr.
   */
  public async reorderParticipantIDs(
    subsiteUrl: string,
    onProgress?: (pct: number) => void
  ): Promise<{ success: number; errors: number }> {
    // Alle Items laden, sortiert nach SP Id (Erstellungsreihenfolge = Reihenfolge der Registrierung)
    const allItems: Array<{ Id: number; Status: string; TeilnehmerID: number | null }> = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Status,TeilnehmerID&$orderby=Id asc&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch { break; }
    }

    // Ziel-IDs in einem ersten Durchlauf berechnen: erst Angemeldete, dann Warteliste.
    // v22.20: Innerhalb der Gruppen wird primär nach der VORHANDENEN TeilnehmerID
    // sortiert (aufsteigend, ohne ID ans Ende), Gleichstand nach Item-Id. Damit ist
    // die Client-Renummerierung deckungsgleich mit dem DEX_IDReorder-Flow (der
    // ebenfalls nach TeilnehmerID sortiert) — egal welcher Pfad zuletzt lief, das
    // Ergebnis ist identisch und die IDs „springen" nicht mehr zwischen
    // Admin-Button und Flow-Lauf hin und her.
    const targetIds = new Map<number, number | null>();
    let nextId = 1;
    const NO_TID = Number.MAX_SAFE_INTEGER;
    const byTidThenId = (a: { Id: number; TeilnehmerID: number | null }, b: { Id: number; TeilnehmerID: number | null }): number =>
      ((a.TeilnehmerID ?? NO_TID) - (b.TeilnehmerID ?? NO_TID)) || (a.Id - b.Id);
    // Pass 1: Angemeldete / QR versendet / Eingecheckt
    const activeItems = allItems
      .filter(item => item.Status === 'Angemeldet' || item.Status === 'QR versendet' || item.Status === 'Eingecheckt')
      .sort(byTidThenId);
    for (const item of activeItems) {
      targetIds.set(item.Id, nextId++);
    }
    // Pass 2: Warteliste
    const waitlistItems = allItems
      .filter(item => item.Status === 'Warteliste')
      .sort(byTidThenId);
    for (const item of waitlistItems) {
      targetIds.set(item.Id, nextId++);
    }
    // Pass 3: Abgemeldete (TeilnehmerID=null)
    for (const item of allItems) {
      if (!targetIds.has(item.Id)) {
        targetIds.set(item.Id, null);
      }
    }

    let success = 0;
    let errors = 0;
    const totalItems = allItems.length || 1;
    let processed = 0;
    if (onProgress) { try { onProgress(0); } catch { /* */ } }
    for (const item of allItems) {
      const newId = targetIds.get(item.Id) ?? null;
      if (newId === item.TeilnehmerID) {
        success++;
      } else {
        try {
          const resp = await this._merge(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${item.Id})`,
            { 'TeilnehmerID': newId }
          );
          if (resp.ok || resp.status === 406) { success++; } else { errors++; }
        } catch { errors++; }
      }
      processed++;
      if (onProgress) {
        // 0..95 % während der Merges; die letzten 5 % für syncCounterToMax.
        try { onProgress(Math.min(95, Math.round((processed / totalItems) * 95))); } catch { /* */ }
      }
    }

    // v7.31 / v9.14: Counter konsistent halten — syncCounterToMax patcht
    // Counter (monotonic up-only). ensureCounterList wurde hier ursprünglich
    // (v9.13) ebenfalls gerufen, hat aber Race-Conditions ausgelöst. Die
    // Counter-Liste sollte zum Zeitpunkt eines Reorders ohnehin existieren —
    // sonst hat die App ein anderes Problem das ein expliziter Klick auf
    // "Counter zurücksetzen" löst.
    try { await this.syncCounterToMax(subsiteUrl); } catch { /* best-effort */ }
    if (onProgress) { try { onProgress(100); } catch { /* */ } }

    return { success, errors };
  }

  // ==================== v11.36: Überbuchungs-Schutz + Bereinigung ====================

  private static readonly ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

  /**
   * Zählt die aktiven (= nicht Warteliste/Abgemeldet) Anmeldungen, gesamt und
   * pro Starter-Gruppe. Quelle ist die echte Teilnehmerliste — wird zum Seeden
   * und Reconcilen der Sitzplatz-Counter genutzt.
   */
  private async getActiveCounts(subsiteUrl: string): Promise<{ total: number; durch: number; fun: number }> {
    const regs = await this.getAllRegistrations(subsiteUrl);
    const active = regs.filter(r => EventService.ACTIVE_STATI.indexOf(r.Status) >= 0);
    return {
      total: active.length,
      durch: active.filter(r => r.StarterType === 'Durchstarter').length,
      fun: active.filter(r => r.StarterType === 'Funstarter').length,
    };
  }

  private seatFieldFor(group: string): 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun' {
    if (group === 'Durchstarter') return 'SeatsTakenDurch';
    if (group === 'Funstarter') return 'SeatsTakenFun';
    return 'SeatsTaken';
  }

  /**
   * v11.36: Atomare Sitzplatz-Reservierung pro Gruppe via ETag-CAS auf der
   * DEX_TeilnehmerCounter-Liste — exakt dasselbe bewährte Muster wie
   * getNextTeilnehmerId (IF-MATCH, 412-Retry mit Backoff).
   *
   * Verhindert die Überbuchung bei zeitgleichen Anmeldungen: zwei parallele
   * Anmeldungen können nicht beide den letzten Platz greifen — die CAS
   * serialisiert das Increment, der Verlierer liest neu und sieht „voll".
   *
   * Rückgabe:
   * - 'reserved' → Platz wurde atomar belegt, Aufrufer darf 'Angemeldet' setzen
   * - 'full'     → Gruppe/Event ist voll → Aufrufer setzt 'Warteliste'
   * - 'error'    → Counter nicht nutzbar (Liste fehlt, Permission, Retries
   *                 erschöpft). Aufrufer MUSS fail-closed handeln (Warteliste),
   *                 NICHT optimistisch 'Angemeldet'.
   *
   * Self-Seed: ist das Seat-Feld noch nie gesetzt (null), wird es einmalig aus
   * der echten aktiven Anzahl der Gruppe initialisiert, bevor entschieden wird.
   */
  public async reserveSeat(
    subsiteUrl: string,
    group: '' | 'Durchstarter' | 'Funstarter',
    cap: number,
    count: number = 1
  ): Promise<'reserved' | 'full' | 'error'> {
    // cap <= 0 = unbegrenzt → kein Reservieren nötig.
    if (!cap || cap <= 0) return 'reserved';
    const inc = Math.max(1, Math.floor(count));
    const field = this.seatFieldFor(group);
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    // v18.8 (Überbuchungs-Fix): Der Counter allein ist NICHT verlässlich.
    // Der Power-Automate-Nachrück-Flow promotet Warteliste→Angemeldet, ohne
    // SeatsTaken zu erhöhen; läuft der app-seitige syncSeatsToActiveCount
    // zeitlich VOR dieser asynchronen Promotion, steht der Counter unter dem
    // echten Aktiv-Bestand. Folge (real beobachtet): trotz voller Warteliste
    // sah der nächste Registrant einen Phantom-Platz und überbuchte. Deshalb
    // lesen wir EINMAL pro Aufruf den echten Aktiv-Bestand der Gruppe und
    // floor-en den Counter-Wert dagegen (max). Das schließt die Drift-Lücke,
    // erhält die atomare CAS-Serialisierung paralleler Anmeldungen UND heilt
    // den Counter nach oben. Bei Lesefehler (Throttling): kein Floor (-1) →
    // Fallback auf reines Counter-Verhalten, nicht schlechter als vorher.
    let realActive = -1;
    try {
      const rc = await this.getActiveCounts(subsiteUrl);
      realActive = group === 'Durchstarter' ? rc.durch : group === 'Funstarter' ? rc.fun : rc.total;
    } catch { realActive = -1; }
    const MAX_RETRIES = 40;
    let triedLazyCreate = false;
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this.context.spHttpClient.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return 'error';
      }
      if (!getResp.ok) {
        if (getResp.status === 404 && !triedLazyCreate) {
          triedLazyCreate = true;
          try { await this.ensureCounterList(subsiteUrl); continue; } catch { return 'error'; }
        }
        return 'error';
      }
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return 'error';
      let data;
      try { data = await getResp.json(); } catch { return 'error'; }
      const rawVal = data?.[field] ?? data?.d?.[field];
      let current: number;
      if (rawVal === null || rawVal === undefined) {
        // Feld noch nie initialisiert → aus echtem Bestand seeden. v18.8:
        // den bereits oben gelesenen realActive wiederverwenden (kein zweiter
        // getActiveCounts-Roundtrip); nur falls der Read fehlschlug (-1),
        // konservativ auf 0.
        current = realActive >= 0 ? realActive : 0;
      } else {
        current = typeof rawVal === 'number' ? rawVal : (parseInt(String(rawVal), 10) || 0);
      }
      // v18.8: gegen echten Aktiv-Bestand floor-en (siehe Kommentar oben) —
      // fängt eine durch den Nachrück-Flow nach unten gedriftete Zählung ab.
      if (realActive >= 0 && realActive > current) current = realActive;
      // v11.82: Team-Anmeldungen reservieren N Plätze atomar. Wenn nicht alle
      // N in dieselbe Gruppe passen, schlägt die Reservierung als „full" fehl —
      // der Aufrufer setzt das gesamte Team auf Warteliste (kein Teil-Team
      // aktivieren). Bei count=1 (Solo) ist das Verhalten identisch zu vorher.
      if (current + inc > cap) return 'full';
      const patchResp = await this._mergeIfMatch(counterItemUrl, { [field]: current + inc }, etag);
      if (patchResp.ok) return 'reserved';
      if (patchResp.status !== 412) return 'error';
      const baseDelay = Math.min(500, 50 * Math.pow(1.4, attempt));
      await new Promise(res => setTimeout(res, Math.floor(baseDelay * (0.5 + Math.random()))));
    }
    // Retries erschöpft → fail-closed (Aufrufer setzt Warteliste).
    console.warn('[DEX] reserveSeat: 40 Retries erschöpft — fail-closed (Warteliste).');
    return 'error';
  }

  /**
   * v11.36: Sitzplatz-Counter mit dem echten aktiven Bestand abgleichen.
   * Nach Abmeldungen, Reorder und Überbuchungs-Bereinigung aufrufen. Die
   * Power-Automate-Nachrück-Promotion fasst den Counter nicht an — dieser
   * Reconcile (aktive Anzahl aus der Liste) hält ihn ehrlich. Best-effort,
   * ETag-CAS, blockiert nie den aufrufenden Flow.
   */
  public async syncSeatsToActiveCount(
    subsiteUrl: string,
    opts: { isSplit: boolean }
  ): Promise<void> {
    let counts: { total: number; durch: number; fun: number };
    try { counts = await this.getActiveCounts(subsiteUrl); } catch { return; }
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    const desired = opts.isSplit
      ? { SeatsTakenDurch: counts.durch, SeatsTakenFun: counts.fun, SeatsTaken: counts.total }
      : { SeatsTaken: counts.total };
    for (let attempt = 0; attempt < 6; attempt++) {
      try {
        const getResp = await this.context.spHttpClient.get(counterItemUrl, SPHttpClient.configurations.v1);
        if (!getResp.ok) return;
        const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
        if (!etag) return;
        const patchResp = await this._mergeIfMatch(counterItemUrl, desired, etag);
        if (patchResp.ok) return;
        if (patchResp.status !== 412) return;
        await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
      } catch { return; }
    }
  }

  /**
   * v11.36: Überbuchung erkennen + markieren (ändert KEINEN Status).
   * Pro Gruppe (bzw. gesamt bei Nicht-Split) werden die zuletzt angemeldeten
   * Einträge über Kapazität (höchste SP-Id = zuletzt registriert) mit
   * OverbookReview='Pending' markiert. First-come-first-served: wer zuerst
   * da war, behält den Platz.
   */
  public async detectOverbooking(
    subsiteUrl: string,
    opts: { isSplit: boolean; maxParticipants?: number; durchstarterCapacity?: number; funstarterCapacity?: number }
  ): Promise<{ groups: Array<{ group: string; cap: number; activeBefore: number; marked: number }>; total: number; errors: number }> {
    const regs = await this.getAllRegistrations(subsiteUrl); // Id asc
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
            const resp = await this._merge(
              `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${it.Id})`,
              { 'OverbookReview': 'Pending' }
            );
            if (resp.ok || resp.status === 406) { marked++; total++; } else { errors++; }
          } catch { errors++; }
        }
      }
      groups.push({ group: label, cap, activeBefore: before, marked });
    };
    const isActive = (r: SPRegistration): boolean => EventService.ACTIVE_STATI.indexOf(r.Status) >= 0;
    if (opts.isSplit) {
      await markExcess(regs.filter(r => isActive(r) && r.StarterType === 'Durchstarter'), opts.durchstarterCapacity || 0, 'Durchstarter');
      await markExcess(regs.filter(r => isActive(r) && r.StarterType === 'Funstarter'), opts.funstarterCapacity || 0, 'Funstarter');
    } else {
      await markExcess(regs.filter(isActive), opts.maxParticipants || 0, 'all');
    }
    return { groups, total, errors };
  }

  /** v11.36: Review-Marker einer Zeile entfernen (ohne Status-Änderung). */
  public async clearOverbookMark(subsiteUrl: string, itemId: number): Promise<boolean> {
    try {
      const resp = await this._merge(
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
  public async resolveOverbookToWaitlist(
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
        const ex = await this.context.spHttpClient.get(
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
      const resp = await this._merge(
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
  public async resolveOverbookKeepActive(subsiteUrl: string, itemId: number): Promise<boolean> {
    return this.clearOverbookMark(subsiteUrl, itemId);
  }

  /**
   * v11.36: „Platz behalten" — Variante (b): Person wird Erste(r) auf der
   * gruppeneigenen Warteliste. Der Nachrück-Flow sortiert Warteliste nach
   * RegistrationDate asc — daher setzen wir RegistrationDate knapp VOR den
   * frühesten aktuellen Wartelisten-Eintrag derselben Gruppe. Original-Datum
   * wird im ChangeLog vermerkt.
   */
  public async resolveOverbookKeepAsFirstWaitlist(
    subsiteUrl: string,
    itemId: number,
    group: string
  ): Promise<boolean> {
    try {
      const all = await this.getAllRegistrations(subsiteUrl);
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
        const ex = await this.context.spHttpClient.get(
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
      const resp = await this._merge(
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
  public async buildOverbookApologyEmail(
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
    const tpl = await this.getEmailTemplate('OverbookingApology', lang).catch(() => null);
    const vars: Record<string, string> = {
      Name: first || name,
      EventTitle: eventTitle,
      WaitlistPositionBlock: posBlock,
      WaitlistPosition: hasPos ? String(waitlistPos) : '',
      AppUrl: `${this.siteUrl}/SitePages/DEX.aspx?env=WebView`,
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

  /**
   * Spalten der Teilnehmerliste fixen: fehlende Spalten anlegen, View-Reihenfolge korrigieren.
   * Kann auf bestehenden Events ausgeführt werden um die Liste nachträglich zu aktualisieren.
   */
  public async fixRegistrationListColumns(
    subsiteUrl: string,
    eventContext?: {
      isB2Run?: boolean;  // Event hat Durchstarter/Funstarter Kapazität
      hasQuiz?: boolean;  // Event hat Quizfragen
      customFields?: CustomField[]; // Event-spezifische Custom-Fields — fehlende SP-Spalten werden angelegt
    },
    // v11.56: Optionaler Confirm-Callback. Wird aufgerufen, wenn Duplikat-Spalten
    // erkannt wurden, BEVOR irgendetwas gelöscht wird. Liefert der Callback false,
    // werden die Duplikate uebersprungen (die Hauptfix-Logik läuft trotzdem).
    confirmDeleteDuplicates?: (count: number, titles: string[]) => boolean | Promise<boolean>
  ): Promise<{ added: string[]; removed: string[]; viewFixed: boolean; customFieldMap?: Record<string, string>; duplicatesRemoved?: string[]; duplicatesWithData?: string[] }> {
    const added: string[] = [];
    const removed: string[] = [];
    const duplicatesRemoved: string[] = [];
    const duplicatesWithData: string[] = [];

    // Bestehende Felder laden — InternalName + Title beide nehmen, damit wir per Title
    // dedupen können (siehe v11.56: alte Builds haben durch fehlgeschlagene Existenz-
    // checks beim wiederholten "Spalten fixen" pro Custom-Field 50+ Duplikate angelegt).
    const fieldsResp = await this.context.spHttpClient.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=500&$select=InternalName,Title`,
      SPHttpClient.configurations.v1
    );
    const existingFieldsList: string[] = [];
    const existingByInternal: Set<string> = new Set();
    const existingByTitle: Map<string, Array<{ internalName: string }>> = new Map();
    if (fieldsResp.ok) {
      const fieldsData = await fieldsResp.json();
      const fields = fieldsData.value || fieldsData.d?.results || [];
      for (const f of fields) {
        const internalName: string = String(f.InternalName || '');
        const title: string = String(f.Title || '');
        if (!internalName) continue;
        existingFieldsList.push(internalName);
        existingByInternal.add(internalName);
        if (title) {
          const arr = existingByTitle.get(title) || [];
          arr.push({ internalName });
          existingByTitle.set(title, arr);
        }
      }
    }

    // ===== DEDUPE-PASS (v11.56) =====
    // Pro Title: wenn mehr als ein Feld diesen Titel hat → die uebrigen Felder löschen,
    // sofern sie leer sind (keine Items mit Wert in der Spalte). Erstes Feld bleibt
    // immer erhalten. Felder mit Daten werden gemeldet (duplicatesWithData) und nicht
    // automatisch gelöscht — der User soll sie manuell prüfen.
    const duplicateTitles: Array<{ title: string; entries: Array<{ internalName: string }> }> = [];
    existingByTitle.forEach((entries, title) => {
      if (entries.length > 1) {
        duplicateTitles.push({ title, entries });
      }
    });
    if (duplicateTitles.length > 0) {
      // Vor dem Löschen den Aufrufer fragen — Operation ist irreversibel.
      if (confirmDeleteDuplicates) {
        const count = duplicateTitles.reduce((sum, d) => sum + (d.entries.length - 1), 0);
        const titles = duplicateTitles.map(d => d.title);
        const ok = await Promise.resolve(confirmDeleteDuplicates(count, titles));
        if (!ok) {
          // Cleanup ueberspringen — nur den Hauptfix laufen lassen
          duplicateTitles.length = 0;
        }
      }
    }
    if (duplicateTitles.length > 0) {
      // Pro Duplikat-Set: den ersten Eintrag behalten, für alle weiteren prüfen ob leer.
      for (const dup of duplicateTitles) {
        // entries[0] bleibt erhalten
        for (let i = 1; i < dup.entries.length; i++) {
          const candidate = dup.entries[i];
          let isEmpty = false;
          // v11.67: SP truncated InternalNames auf 32 Zeichen. Wenn die Truncation
          // mitten in einer `_xXXXX_`-Encoding-Sequenz liegt (z.B.
          // `ADMIN_x0020__x002d__x0020_Who_x00` — die letzten Zeichen `_x00`
          // sind eine angeschnittene `_x0020_`-Sequenz), wirft SP HTTP 400 auf
          // jeden OData-`$filter`-Versuch. Solche Spalten werden hier nicht
          // geprüft → konservativ als „hat Daten" behandelt (kein Auto-
          // Löschen). Der Admin kann sie ueber die SP-Listen-UI manuell
          // entfernen, wenn sie wirklich leer sind.
          const looksTruncated = candidate.internalName.length === 32
            && /_x[0-9a-f]{1,3}$/i.test(candidate.internalName);
          if (looksTruncated) {
            if (duplicatesWithData.indexOf(dup.title) < 0) duplicatesWithData.push(dup.title);
            continue;
          }
          try {
            const checkUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${candidate.internalName} ne null&$top=1&$select=ID`;
            const checkResp = await this.context.spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
            if (checkResp.ok) {
              const data = await checkResp.json();
              const items = data.value || data.d?.results || [];
              isEmpty = items.length === 0;
            }
          } catch { isEmpty = false; }
          if (isEmpty) {
            try {
              const delResp = await this._post(
                `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields/getbyinternalnameortitle('${candidate.internalName}')/deleteObject`,
                {}
              );
              if (delResp.ok) {
                duplicatesRemoved.push(candidate.internalName);
                // Aus existingByInternal + existingFieldsList rausziehen
                existingByInternal.delete(candidate.internalName);
                const idx = existingFieldsList.indexOf(candidate.internalName);
                if (idx >= 0) existingFieldsList.splice(idx, 1);
              }
            } catch { /* löschen fehlgeschlagen — weiter */ }
          } else {
            // Daten vorhanden — Title für manuelle Prüfung melden (nur einmal pro Title)
            if (duplicatesWithData.indexOf(dup.title) < 0) duplicatesWithData.push(dup.title);
          }
        }
        // existingByTitle entsprechend bereinigen: nur die nicht-gelöschten Einträge behalten
        const remaining = dup.entries.filter(e => existingByInternal.has(e.internalName));
        existingByTitle.set(dup.title, remaining);
      }
    }

    // Fehlende Basis-Spalten anlegen. StarterType/Quiz-Felder sind feature-spezifisch:
    // - StarterType/PreferredStarterType: nur für B2Run-Events mit Split-Kapazität
    // - QuizScore/QuizAnswers/QuizCompletedAt: nur für Events mit Quizfragen
    // Wird das Event ohne eventContext gefixt (kein Aufrufer-seitiger Flag), lassen wir
    // feature-spezifische Spalten raus, damit sie nicht unbegründet auf jedem
    // Teilnehmerlisten-Schema auftauchen.
    const requiredFields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
      { title: 'Company', type: 2 },           // v24.29: Unternehmenszugehörigkeit / Rechtsträger (aus dem Profil)
      { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgeführt hat
      { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgeführt hat
      { title: 'ProxyConsent', type: 3 },      // v18.74: Zustimmungs-Nachweis bei stellvertretender Anmeldung (Note)
      { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgelöst hat
      { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
      { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
      { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
      { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
      // v19.3: Nachrück-Audit-Spalten auch beim „Spalten fixen" nachziehen, damit
      // der DEX_IDReorder-Flow (und der App-Button) sie auf Bestands-Events
      // beschreiben kann → „Nachgerückt am / Ersetzt / Ersetzt durch" in der App.
      { title: 'PromotedDate', type: 4 },              // DateTime — Zeitpunkt des Nachrückens
      { title: 'ReplacedParticipantEmail', type: 2 },  // E-Mail der Person, deren Cancel den Platz freigab
      { title: 'ReplacedByParticipantEmail', type: 2 },// E-Mail der nachrückenden Person (Spiegelbild)
      { title: 'OverbookReview', type: 2 },    // v11.36: Überbuchungs-Review-Marker
      { title: 'TeamId', type: 2 },            // v11.82: UUID einer Team-Anmeldung (leer = Solo)
      { title: 'TeamLead', type: 8 },          // v11.82: Boolean — true für die anmeldende Person
      { title: 'TeamName', type: 2 },          // v11.82: optionaler frei wählbarer Team-Name
    ];
    if (eventContext?.isB2Run) {
      requiredFields.push(
        { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' },
        { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' },
        // v10.13: b2run_leistungsnachweis ist ein virtuelles Feld das die
        // RegistrationPage hardcoded hinzufügt wenn durchstarterRequiresProof
        // aktiv ist — es ist NICHT Teil der regulären customFields, daher
        // muss die SP-Spalte hier explizit angelegt werden, sonst kippt die
        // Anmeldung mit HTTP 400 'Field not found'. Wird auf jedem B2Run-Event
        // angelegt (egal ob proof-flag aktuell aktiv ist) — die Spalte ist
        // klein und stört nicht wenn ungenutzt.
        { title: 'b2run_leistungsnachweis', type: 2 }
      );
    }
    if (eventContext?.hasQuiz) {
      requiredFields.push(
        { title: 'QuizScore', type: 9 },
        { title: 'QuizAnswers', type: 3 },
        { title: 'QuizCompletedAt', type: 4 }
      );
    }

    // Feature-spezifische Spalten, die auf diesem Event NICHT gebraucht werden,
    // aktiv löschen (z.B. StarterType auf einem Nicht-B2Run-Event). Das ist
    // irreversibel — eventuelle Daten in diesen Spalten gehen verloren. Ist aber
    // vom User explizit gewünscht, damit die Teilnehmerliste pro Event-Typ
    // sauber bleibt.
    const deletableFields: string[] = [];
    if (!eventContext?.isB2Run) {
      deletableFields.push('StarterType', 'PreferredStarterType');
    }
    if (!eventContext?.hasQuiz) {
      deletableFields.push('QuizScore', 'QuizAnswers', 'QuizCompletedAt');
    }
    for (const fieldName of deletableFields) {
      if (existingFieldsList.indexOf(fieldName) >= 0) {
        try {
          const delResp = await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields/getbytitle('${fieldName}')/deleteObject`,
            {}
          );
          if (delResp.ok) {
            removed.push(fieldName);
            // aus existingFieldsList rausziehen damit die View-Logik weiter unten
            // den Feldnamen nicht mehr als "noch vorhanden" betrachtet.
            const idx = existingFieldsList.indexOf(fieldName);
            if (idx >= 0) existingFieldsList.splice(idx, 1);
          }
        } catch { /* Feld konnte nicht gelöscht werden - weitermachen */ }
      }
    }

    for (const f of requiredFields) {
      if (existingFieldsList.indexOf(f.title) < 0) {
        const payload: Record<string, unknown> = {
          '__metadata': { 'type': f.metaType || 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        };
        if (f.choices) {
          payload['Choices'] = { 'results': f.choices };
        }
        const resp = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload
        );
        if (resp.ok) added.push(f.title);
      }
    }

    // Custom Fields pro Event: wenn spInternalName leer oder die Spalte fehlt,
    // jetzt anlegen. Der Aufrufer bekommt customFieldMap zurück und kann das
    // Event-Item persistieren (spInternalName für jede cf.id).
    const customFieldMap: Record<string, string> = {};
    if (eventContext?.customFields && eventContext.customFields.length > 0) {
      // Post-Fix Felder-Snapshot nach Basis-Anlage
      let currentFields = [...existingFieldsList, ...added];
      for (const cf of eventContext.customFields) {
        if (!cf.label || !cf.label.trim()) continue;
        // v19.0: Dokument-Felder bekommen keine Spalte (Datei = Attachment).
        if (cf.type === 'document') continue;
        // Wenn spInternalName schon gesetzt und Feld existiert: uebernehmen.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const existingSp: string = String((cf as any).spInternalName || '');
        if (existingSp && currentFields.indexOf(existingSp) >= 0) {
          customFieldMap[cf.id] = existingSp;
          continue;
        }
        // v11.56: Wenn spInternalName fehlt oder nicht zur Liste passt, aber ein
        // Feld mit demselben Title bereits existiert: dieses InternalName uebernehmen,
        // statt eine Duplikat-Spalte anzulegen. Das ist die Hauptursache der
        // 100x-Duplikate-Misere (P/D MEETING0, P/D MEETING1, ...).
        const titleMatches = existingByTitle.get(cf.label) || [];
        if (titleMatches.length > 0) {
          const firstInternal = titleMatches[0].internalName;
          customFieldMap[cf.id] = firstInternal;
          if (currentFields.indexOf(firstInternal) < 0) currentFields.push(firstInternal);
          continue;
        }
        // Feld-Payload je nach Typ
        let fieldPayload: Record<string, unknown>;
        if (cf.type === 'select' && cf.options && cf.options.length > 0) {
          fieldPayload = { '__metadata': { 'type': 'SP.FieldChoice' }, 'Title': cf.label, 'FieldTypeKind': 6, 'Required': false, 'Choices': { 'results': cf.options } };
        } else if (cf.type === 'number') {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 9, 'Required': false };
        } else if (cf.type === 'checkbox') {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 8, 'Required': false };
        } else if (cf.type === 'user') {
          // user-Picker wird als Text gespeichert ("Name <email>").
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 3, 'Required': false };
        } else {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 2, 'Required': false };
        }
        try {
          const resp = await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, fieldPayload
          );
          if (resp.ok) {
            const createdField = await resp.json().catch(() => null);
            const internalName = createdField?.InternalName || createdField?.d?.InternalName || '';
            if (internalName) {
              customFieldMap[cf.id] = internalName;
              added.push(internalName);
              currentFields = currentFields.concat([internalName]);
              // Title-Map aktualisieren, damit ein zweites cf mit gleichem Label
              // im selben Durchlauf (z.B. zwei Custom-Fields mit identischem Title)
              // das gerade angelegte Feld wiederverwendet, statt erneut zu erzeugen.
              const arr = existingByTitle.get(cf.label) || [];
              arr.push({ internalName });
              existingByTitle.set(cf.label, arr);
            }
          }
        } catch { /* nächstes Feld */ }
      }
    }

    // Default View komplett neu aufbauen (Reihenfolge: TeilnehmerID, Anrede, Vorname, Nachname, ...)
    let viewFixed = false;
    try {
      // Alle bestehenden Felder aus der View entfernen
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/removeallviewfields`,
        {}
      );

      // Felder in gewünschter Reihenfolge hinzufügen. StarterType/Quiz-Spalten
      // werden nur eingebaut, wenn sie tatsächlich auf der Liste existieren —
      // auf Nicht-B2Run- bzw. Nicht-Quiz-Events sollen sie nicht auftauchen.
      const viewFieldsCore = [
        'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail',
        'Department', 'Location', 'JobTitle', 'Company', 'Phone',
      ];
      const viewFields: string[] = [...viewFieldsCore];
      // Post-Fix-Feldliste (bestehende + gerade hinzugefügte) für die Existenz-Checks
      const postFixFields = existingFieldsList.concat(added);
      if (postFixFields.indexOf('StarterType') >= 0) viewFields.push('StarterType');
      if (postFixFields.indexOf('PreferredStarterType') >= 0) viewFields.push('PreferredStarterType');
      viewFields.push('Status', 'RegistrationDate');
      if (postFixFields.indexOf('RegisteredByName') >= 0) viewFields.push('RegisteredByName');
      if (postFixFields.indexOf('RegisteredByEmail') >= 0) viewFields.push('RegisteredByEmail');
      if (postFixFields.indexOf('ProxyConsent') >= 0) viewFields.push('ProxyConsent');
      viewFields.push('CancellationDate');

      // Wir blenden SP-System-Spalten komplett aus (Modified, Created, ID, Type,
      // Compliance Asset Id, Retention Label, etc.) — nur funktionelle Felder der
      // App + Custom Fields kommen in die View.
      const systemBlocklist = new Set([
        'ID', '_UIVersionString', 'Edit', 'LinkTitle', 'LinkTitleNoMenu',
        'LinkFilename', 'LinkFilenameNoMenu', 'DocIcon', 'FileLeafRef',
        'Modified', 'Created', 'Editor', 'Author', 'CreatedBy', 'ModifiedBy',
        'Title', 'ParticipantName',
        'ContentType', 'ContentTypeId', 'Attachments',
        'AppAuthor', 'AppEditor', 'App Created By', 'App Modified By',
        'Type', 'ItemChildCount', 'FolderChildCount',
        'ComplianceAssetId', '_ComplianceTag', '_ComplianceTagWrittenTime',
        '_ComplianceTagUserId', 'TaxCatchAll', 'TaxCatchAllLabel',
        'SMTotalFileStreamSize', 'SMTotalSize', 'SortBehavior',
        'OData__UIVersionString', 'OData__HasCopyDestinations',
        'LastModifiedDate', 'ChangeLog', 'CustomData',
        '_CopySource', 'owshiddenversion', 'WorkflowVersion', 'WorkflowInstanceID',
        'ItemIsRecord', '_HasEncryptedContent', '_IsRecord', '_IsRecordApplied',
        'InstanceID', 'Order', 'GUID', 'FileSizeDisplay', 'MetaInfo',
        'ParentUniqueId', 'AccessPolicy', 'HasUniqueRoleAssignments',
        'Restricted', 'Type0', 'ServerUrl', 'EncodedAbsUrl', 'BaseName',
        'FileType', 'HTML_x0020_File_x0020_Type', '_EditMenuTableStart',
        '_EditMenuTableStart2', '_EditMenuTableEnd', 'PermMask',
      ]);
      // Bereits zur View hinzugefügt — nicht doppelt anfassen
      const alreadyAdded = new Set(viewFields);
      // v11.82: Team-Spalten kommen ans Ende der View — nach allen
      // Custom-Fields, damit sie nicht zwischen den event-spezifischen
      // Antwortspalten landen. Hier merken und im Post-Loop ueberspringen.
      const teamTailFields = ['TeamId', 'TeamLead', 'TeamName'];
      const teamTailSet = new Set(teamTailFields);
      // Kompletter Feld-Stand NACH dem Fix (bestehende + neu angelegte),
      // damit neu angelegte Custom-Fields auch in die View kommen.
      for (const fn of postFixFields) {
        if (alreadyAdded.has(fn)) continue;
        if (systemBlocklist.has(fn)) continue;
        if (fn.charAt(0) === '_') continue;
        if (teamTailSet.has(fn)) continue; // ans Ende
        viewFields.push(fn);
        alreadyAdded.add(fn);
      }
      // Team-Spalten jetzt am Ende anhängen (nur die, die tatsächlich existieren).
      for (const fn of teamTailFields) {
        if (alreadyAdded.has(fn)) continue;
        if (postFixFields.indexOf(fn) < 0) continue;
        viewFields.push(fn);
        alreadyAdded.add(fn);
      }

      for (const fn of viewFields) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/addviewfield('${fn}')`,
          {}
        );
      }
      viewFixed = true;
    } catch {
      console.warn('[DEX] View-Reihenfolge konnte nicht gesetzt werden');
    }

    // v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe anlegen
    // (oder seeden mit dem aktuellen Max-Wert wenn schon vorhanden).
    try {
      const counterResult = await this.ensureCounterList(subsiteUrl);
      if (counterResult.created) {
        added.push(`Counter-Liste ${COUNTER_LIST_NAME} (atomare TeilnehmerID-Vergabe, seeded mit ${counterResult.seededValue})`);
      } else if (counterResult.seededValue !== undefined) {
        added.push(`Counter-Item nachgeseedet (NextValue=${counterResult.seededValue})`);
      }
    } catch {
      console.warn('[DEX] Counter-Liste konnte nicht angelegt werden');
    }

    return { added, removed, viewFixed, customFieldMap: Object.keys(customFieldMap).length > 0 ? customFieldMap : undefined };
  }


  /**
   * Quiz-Fortschritt in die Registrierung eines Teilnehmers schreiben.
   *
   * - answers: ausgewählte Antwort-Indices pro Frage (Array von Arrays, weil
   *   Fragen mehrere richtige Antworten haben können). Unbeantwortete Fragen
   *   bleiben als leeres Array `[]` stehen, damit der Index-Offset erhalten bleibt.
   * - score: aktuell erreichte Punkte (Anzahl korrekt beantworteter Fragen).
   * - isComplete: true wenn alle Fragen beantwortet sind — dann wird auch
   *   `QuizCompletedAt` gesetzt. Andernfalls bleibt QuizCompletedAt unverändert
   *   (null/leer), sodass der Teilnehmer als "teilweise beantwortet" gelistet wird.
   *
   * Ersetzt die früher nur-am-Ende aufgerufene `saveQuizResult()`. Wird jetzt
   * bei jedem "Weiter"-Klick im QuizPlayer aufgerufen (Auto-Save), damit der
   * Teilnehmer beim späteren Wiedereintritt an derselben Stelle weitermachen kann.
   */
  public async saveQuizProgress(
    subsiteUrl: string,
    itemId: number,
    score: number,
    answers: number[][],
    isComplete: boolean
  ): Promise<boolean> {
    try {
      // Vor dem Schreiben sicherstellen, dass die Quiz-Spalten auf der
      // Teilnehmer-Liste existieren. Bei Bestandsevents (vor Quiz-Feature
      // angelegt) fehlen sie oft; _merge mit odata=nometadata schluckt
      // unbekannte Felder stumm und das Save wirkt wie "gespeichert",
      // persistiert aber nichts.
      // Silent: wenn der aktuelle User keine Manage-Lists-Permission auf
      // der Subsite hat, schlägt das Anlegen fehl (Regular User). Dann
      // kann nur ein Admin/Organizer die Spalten anlegen — dafür gibt's
      // die "Spalten fixen"-Funktion im Admin Center.
      await this.ensureQuizColumnsOnRegList(subsiteUrl);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: { [key: string]: any } = {
        'QuizScore': score,
        'QuizAnswers': JSON.stringify(answers),
      };
      if (isComplete) {
        payload.QuizCompletedAt = new Date().toISOString();
      }
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        payload
      );
      return resp.ok || resp.status === 406;
    } catch {
      return false;
    }
  }

  /**
   * Quiz-Spalten auf der Teilnehmer-Liste einer Event-Subsite anlegen,
   * falls sie fehlen. Idempotent und silent: bei fehlender Permission
   * kein Crash, einfach kein-op.
   */
  private async ensureQuizColumnsOnRegList(subsiteUrl: string): Promise<void> {
    try {
      const fieldsResp = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=200&$select=InternalName`,
        SPHttpClient.configurations.v1
      );
      if (!fieldsResp.ok) return;
      const fieldsData = await fieldsResp.json();
      const existing = new Set<string>(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (fieldsData.value || fieldsData.d?.results || []).map((f: any) => f.InternalName)
      );
      const required: Array<{ title: string; type: number }> = [
        { title: 'QuizScore', type: 9 },      // Number
        { title: 'QuizAnswers', type: 3 },    // Note (multiline)
        { title: 'QuizCompletedAt', type: 4 } // DateTime
      ];
      for (const f of required) {
        if (existing.has(f.title)) continue;
        try {
          await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
            {
              '__metadata': { 'type': 'SP.Field' },
              'Title': f.title,
              'FieldTypeKind': f.type,
              'Required': false,
            }
          );
          console.warn(`[DEX] ensureQuizColumnsOnRegList: ${f.title} nachgelegt auf ${subsiteUrl}`);
        } catch {
          // keine Permission -> silent. User braucht Admin der "Spalten fixen" macht.
        }
      }
    } catch { /* silent */ }
  }

  /**
   * Ersten Warteliste-Teilnehmer nachrücken: Status -> Angemeldet.
   * Wenn inheritStarterType uebergeben wird (B2Run Split-Capacity), wird dieser Typ
   * dem Nachrücker zugewiesen (er erbt den Platz des Abgemeldeten).
   *
   * Wird **client-seitig** ausgeführt (von der App beim Abmelden), damit der
   * Power Automate DEX_IDReorder-Flow keinen doppelten Nachrück-Versuch macht.
   * Liefert den nachgerückten Teilnehmer (Email + Name) zurück für die
   * Nachrück-E-Mail.
   *
   * Schutz gegen Ueberbuchung: Wenn maxParticipants gesetzt ist und die Anzahl
   * der aktuell Angemeldeten (nach der Abmeldung) >= maxParticipants ist, wird
   * NICHT nachgerückt. Das verhindert, dass nach einer früheren Ueberbuchung
   * der Abbruch der Abmeldung nicht zu einer weiteren Ueberbuchung führt.
   */
  public async promoteFirstWaitlistItem(
    subsiteUrl: string,
    inheritStarterType?: string,
    maxParticipants?: number,
    /** Seit v6.5: Bei B2Run-Events mit getrennten Durchstarter-/Funstarter-Wartelisten
     * hier den freigewordenen Starter-Typ mitgeben — dann wird NUR der erste
     * Warteliste-Teilnehmer mit passendem PreferredStarterType nachgerückt.
     * Wenn leer: Default-Verhalten (beliebiger Warteliste-Teilnehmer). */
    onlyWithPreferredType?: string,
    /** v17.15: Audit-Tracking — wenn der Promote durch das Cancel einer
     *  konkreten Person ausgelöst wurde (in der App-Pfad), die E-Mail
     *  und Item-Id dieser Person mitgeben. Wird auf der nachrückenden
     *  Person als ReplacedParticipantEmail + PromotedDate gespeichert,
     *  und zusätzlich auf der cancelnden Person als
     *  ReplacedByParticipantEmail (zweite MERGE-PATCH). */
    replacedByCancel?: { itemId: number; participantEmail: string },
  ): Promise<{ success: boolean; email?: string; name?: string; itemId?: number; skippedOverbooked?: boolean }> {
    try {
      // Ueberbuchungs-Schutz: Nur nachrücken, wenn tatsächlich ein Platz frei ist.
      // Bei unlimited (maxParticipants === 0 oder undefined) immer nachrücken.
      //
      // WICHTIG: '>' statt '>='. Die Abmeldung (Status->Abgemeldet) ist kurz vor
      // diesem Call passiert; falls SharePoint den Statuswechsel noch nicht in
      // getRegistrationCount reflektiert (stale read), würden wir bei einem
      // vollen Event (z.B. 128/128) mit '>=' fälschlich skippen. Mit '>' ist
      // 'registered == max' noch erlaubt (= genau ein Platz wird nachgerückt),
      // und eine echte Ueberbuchung (401 > 128) wird weiterhin abgefangen.
      if (maxParticipants && maxParticipants > 0) {
        const counts = await this.getRegistrationCount(subsiteUrl);
        if (counts.registered > maxParticipants) {
          console.warn(`[DEX] promoteFirstWaitlistItem: skipping promotion - event is overbooked (${counts.registered}/${maxParticipants} registered).`);
          return { success: false, skippedOverbooked: true };
        }
      }

      // v12.10: Nachrück-Sortierung jetzt nach TeilnehmerID asc statt
      // RegistrationDate. Hintergrund: nach dem IDReorder-Flow sind die
      // TeilnehmerIDs durchlaufend (1..N aktiv, N+1.. Warteliste). Wenn
      // also Platz 100 frei wird, soll TID 101 (= erster auf der Liste)
      // nachrücken — unabhängig davon, ob TID 103 zeitlich gesehen vor
      // TID 101 registriert war (z.B. nach Re-Registration oder Wechsel
      // der Gruppe). RegistrationDate sortierte chronologisch, was bei
      // umverteilten IDs zur falschen Reihenfolge führte.
      // Bei B2Run-Split-Kapazitäten: nur die passende Warteliste durchsuchen
      // (PreferredStarterType == onlyWithPreferredType).
      let filter = `Status eq 'Warteliste'`;
      if (onlyWithPreferredType) {
        const esc = onlyWithPreferredType.replace(/'/g, "''");
        filter += ` and PreferredStarterType eq '${esc}'`;
      }
      const resp = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${encodeURIComponent(filter)}&$orderby=TeilnehmerID asc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return { success: false };
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return { success: false };

      const firstWaiting = items[0];
      // v18.71: Kern-Update (Status -> Angemeldet, ggf. StarterType) STRIKT
      // getrennt von den optionalen Audit-Feldern. Hintergrund: bei Legacy-
      // Teilnehmerlisten, die noch nie per „Spalten fixen" aktualisiert wurden,
      // fehlt die Spalte PromotedDate (erst seit v17.15). Wenn PromotedDate im
      // selben MERGE-Body steht, lehnt SharePoint den GESAMTEN Request mit
      // HTTP 400 ab („The property 'PromotedDate' does not exist…") — der
      // Nachrück-Status wird dann gar nicht gesetzt und der Button „macht
      // nichts". Deshalb zuerst nur die Pflichtfelder schreiben.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mergeBody: Record<string, any> = { 'Status': 'Angemeldet' };
      // v19.14 AUDIT-FIX: StarterType IMMER mitsetzen — entweder den vom
      // Abgemeldeten geerbten Typ (inheritStarterType, Admin-Cancel-Pfad) ODER,
      // falls keiner mitgegeben wurde (z.B. der manuelle „Nachrücken"-Button, der
      // inheritStarterType=undefined übergibt), den EIGENEN Wunsch der
      // nachrückenden Person (PreferredStarterType). Vorher blieb StarterType bei
      // Promotes ohne inheritStarterType auf Split-Events LEER → es entstanden
      // angemeldete „typlose" Personen (Audit-Befund: Andreas Jehle), die aus den
      // Gruppen-Zahlen fielen und als „Wunsch: …" angezeigt wurden. Bei
      // Nicht-Split-Events ist PreferredStarterType leer → StarterType bleibt leer
      // (korrekt, da es dort keine Gruppen gibt).
      const effectiveStarter = inheritStarterType || firstWaiting.PreferredStarterType || '';
      if (effectiveStarter) {
        mergeBody['StarterType'] = effectiveStarter;
      }
      const mergeResp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${firstWaiting.Id})`,
        mergeBody
      );
      if (!(mergeResp.ok || mergeResp.status === 406)) return { success: false };

      // v17.15: Nachrück-Audit auf der promoteten Person — best-effort, in
      // einem SEPARATEN MERGE, damit eine fehlende Audit-Spalte (Legacy-Liste)
      // den eigentlichen Promote oben nicht kaputtmacht.
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const auditBody: Record<string, any> = { 'PromotedDate': new Date().toISOString() };
        if (replacedByCancel && replacedByCancel.participantEmail) {
          auditBody['ReplacedParticipantEmail'] = replacedByCancel.participantEmail;
        }
        await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${firstWaiting.Id})`,
          auditBody
        );
      } catch (err) {
        console.warn('[DEX] promoteFirstWaitlistItem: Nachrück-Audit (PromotedDate) konnte nicht geschrieben werden — Spalte fehlt evtl. auf einer Legacy-Liste:', err);
      }

      const vorname = firstWaiting.Vorname || '';
      const nachname = firstWaiting.Nachname || '';
      const name = (vorname && nachname) ? `${vorname} ${nachname}` : (firstWaiting.ParticipantName || '');
      const email = firstWaiting.ParticipantEmail || firstWaiting.Title || '';

      // v17.15: zweite PATCH auf die cancelnde Person — „Ersetzt durch".
      if (replacedByCancel && replacedByCancel.itemId && email) {
        try {
          await this._merge(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${replacedByCancel.itemId})`,
            { 'ReplacedByParticipantEmail': email }
          );
        } catch (err) {
          console.warn('[DEX] Nachrück-Audit auf cancelnder Person fehlgeschlagen:', err);
        }
      }

      console.warn(`[DEX] promoteFirstWaitlistItem: promoted ${name} <${email}> (item ${firstWaiting.Id}) to Angemeldet.`);
      return { success: true, email, name, itemId: firstWaiting.Id };
    } catch {
      return { success: false };
    }
  }

  /**
   * v11.0: Item-Attachments einer Teilnehmer-Registrierung listen.
   * Liefert ein Array mit FileName + ServerRelativeUrl, sodass die App
   * Download-Links rendern kann. Subsite-spezifisch (jede Teilnehmerliste
   * lebt in der Event-Subsite).
   */
  public async listRegistrationAttachments(
    subsiteUrl: string,
    itemId: number,
  ): Promise<Array<{ fileName: string; serverRelativeUrl: string }>> {
    try {
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles`;
      const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
      });
      if (!resp.ok) return [];
      const data = await resp.json();
      const items = data.value || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return items.map((a: any) => ({
        fileName: a.FileName || '',
        serverRelativeUrl: a.ServerRelativeUrl || '',
      })).filter((x: { fileName: string }) => !!x.fileName);
    } catch (err) {
      console.warn('[DEX] listRegistrationAttachments failed:', err);
      return [];
    }
  }

  /**
   * v11.0: PDF / Datei als Item-Attachment an eine Teilnehmer-Zeile
   * hängen. SharePoint erlaubt mehrere Attachments pro Item; bei
   * gleichem Namen wirft die API einen 409, daher prefixen wir den
   * Dateinamen mit einem Timestamp wenn die App das aufruft.
   */
  public async addRegistrationAttachment(
    subsiteUrl: string,
    itemId: number,
    file: File,
    // v19.0: optionaler Präfix, um ein Attachment einem Dokument-Custom-Field
    // zuzuordnen (z.B. 'dxf-<fieldId>--'). Leer = generischer Attendee-Upload.
    fieldPrefix: string = '',
  ): Promise<boolean> {
    try {
      const buf = await file.arrayBuffer();
      // Dateiname säubern + Timestamp-prefix für Eindeutigkeit
      const safeName = (file.name || 'upload.pdf').replace(/[^a-zA-Z0-9._-]+/g, '_');
      const ts = new Date().toISOString().replace(/[:.]/g, '-').replace(/T/, '_').slice(0, 19);
      const finalName = `${fieldPrefix}${ts}_${safeName}`;
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(finalName)}')`;
      const resp = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
        body: buf,
      });
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] addRegistrationAttachment failed:', err);
      return false;
    }
  }

  /**
   * v11.0: Item-Attachment löschen. Wird sowohl vom User (eigener
   * Upload zurückziehen) als auch vom Admin (im Admin Center) genutzt.
   */
  public async deleteRegistrationAttachment(
    subsiteUrl: string,
    itemId: number,
    fileName: string,
  ): Promise<boolean> {
    try {
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`;
      const resp = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'Accept': 'application/json;odata=nometadata' },
      });
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] deleteRegistrationAttachment failed:', err);
      return false;
    }
  }

  /**
   * v10.27: User wechselt seine Split-Capacity-Gruppe.
   *
   * Logik:
   * - Lädt aktuelle aktive Registrierungen der Subsite und zählt die Anzahl
   *   pro StarterType.
   * - Wenn die Ziel-Gruppe noch unter ihrer Kapazität liegt: User wird mit
   *   neuem StarterType direkt als 'Angemeldet' eingetragen.
   * - Wenn die Ziel-Gruppe bereits voll ist: User wandert auf die Warteliste
   *   mit PreferredStarterType=newType. StarterType bleibt leer (wie bei
   *   Erst-Anmeldung auf Warteliste). Nachgerückt wird er erst, wenn ein
   *   Platz in der Ziel-Gruppe frei wird (siehe Power-Automate-Flow).
   *
   * Liefert { ok, status, full } zurück — die App nutzt das, um die richtige
   * Mail (Anmeldung vs. Warteliste) zu queuen und dem User Feedback zu geben.
   */
  /**
   * v11.24: Tauscht StarterType (und PreferredStarterType) bei ALLEN
   * Registrierungen einer Subsite: jeder 'Durchstarter' wird zu
   * 'Funstarter' und umgekehrt. Wird vom Admin-Center aufgerufen, wenn
   * der Organizer im Wizard die Reihenfolge der Gruppen-Labels +
   * -Kapazitäten getauscht hat — die existierenden Anmeldungen sind
   * dann technisch noch in der „alten" Slot-Bedeutung. Dieser Flip
   * synchronisiert sie mit der neuen Reihenfolge.
   *
   * Liefert die Anzahl erfolgreich aktualisierter Items zurück.
   */
  public async flipAllStarterTypes(subsiteUrl: string): Promise<{ ok: boolean; updated: number; failed: number }> {
    try {
      const all = await this.getAllRegistrations(subsiteUrl);
      let updated = 0;
      let failed = 0;
      for (const r of all) {
        const flip = (t: string | undefined): string => {
          if (t === 'Durchstarter') return 'Funstarter';
          if (t === 'Funstarter') return 'Durchstarter';
          return t || '';
        };
        const newStarter = flip(r.StarterType);
        const newPref = flip(r.PreferredStarterType);
        if (newStarter === (r.StarterType || '') && newPref === (r.PreferredStarterType || '')) continue;
        try {
          const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${r.Id})`;
          const body: Record<string, unknown> = {};
          if (newStarter !== (r.StarterType || '')) body['StarterType'] = newStarter;
          if (newPref !== (r.PreferredStarterType || '')) body['PreferredStarterType'] = newPref;
          const resp = await this._merge(url, body);
          if (resp.ok) updated++;
          else failed++;
        } catch { failed++; }
      }
      return { ok: failed === 0, updated, failed };
    } catch (err) {
      console.warn('[DEX] flipAllStarterTypes error:', err);
      return { ok: false, updated: 0, failed: 0 };
    }
  }

  public async switchSplitGroup(
    subsiteUrl: string,
    itemId: number,
    newType: 'Durchstarter' | 'Funstarter',
    durchstarterCapacity: number,
    funstarterCapacity: number,
  ): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }> {
    try {
      // Counts aktiv-Status pro StarterType ermitteln. Aktiv = Angemeldet,
      // QR versendet oder Eingecheckt — Abgemeldete und Wartelisten zählen
      // nicht gegen die Kapazität.
      const allRegs = await this.getAllRegistrations(subsiteUrl);
      const active = allRegs.filter(r => {
        const s = r.Status || '';
        return s === 'Angemeldet' || s === 'QR versendet' || s === 'Eingecheckt';
      });
      // Den eigenen Eintrag aus der Zählung rausnehmen — wenn er heute schon
      // in der Ziel-Gruppe stünde, würden wir einen Slot frei zählen, der
      // gar nicht entsteht. Wenn er in der Quell-Gruppe steht, gibt der
      // Wechsel den Quell-Slot frei und der Zielslot ist relevant.
      const targetCount = active
        .filter(r => r.Id !== itemId)
        .filter(r => r.StarterType === newType)
        .length;
      const targetCap = newType === 'Durchstarter' ? durchstarterCapacity : funstarterCapacity;
      const targetFree = targetCap - targetCount;
      const goWaitlist = targetCap > 0 && targetFree <= 0;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = goWaitlist
        ? { 'Status': 'Warteliste', 'StarterType': '', 'PreferredStarterType': newType }
        : { 'Status': 'Angemeldet', 'StarterType': newType, 'PreferredStarterType': newType };
      // v22.20: FIFO-Fairness — wer auf die Warteliste der Zielgruppe wechselt,
      // verliert seine alte (niedrige) TeilnehmerID und reiht sich mit einer
      // frischen Counter-ID HINTEN ein. Sonst würde er beim typ-bewussten
      // Nachrücken (TeilnehmerID asc) alle überholen, die schon länger warten.
      // Best-effort: schlägt der Counter fehl, bleibt die alte ID stehen und der
      // anschließende Reorder-Lauf normalisiert wenigstens die Nummerierung.
      if (goWaitlist) {
        try {
          const freshId = await this.getNextTeilnehmerId(subsiteUrl);
          if (typeof freshId === 'number' && freshId > 0) body['TeilnehmerID'] = freshId;
        } catch { /* alte ID behalten */ }
      }
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
      const resp = await this._merge(url, body);
      if (!resp.ok) {
        console.warn('[DEX] switchSplitGroup MERGE failed:', resp.status);
        return { ok: false, status: 'Failed', full: goWaitlist };
      }
      return { ok: true, status: goWaitlist ? 'Warteliste' : 'Angemeldet', full: goWaitlist };
    } catch (err) {
      console.warn('[DEX] switchSplitGroup error:', err);
      return { ok: false, status: 'Failed', full: false };
    }
  }

  /**
   * Registrierung stornieren
   */
  public async cancelRegistration(
    subsiteUrl: string,
    itemId: number,
    cancelledByName?: string,
    cancelledByEmail?: string
  ): Promise<boolean> {
    try {
      // Audit: wer hat die Abmeldung ausgelöst?
      // Bei Self-Cancel = der User selbst. Bei "Teilnehmer abmelden" durch den
      // Organizer/Admin im Admin Center = der eingeloggte Organizer/Admin.
      const auditName = cancelledByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (cancelledByEmail || this.context.pageContext.user.email || '').toLowerCase();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const corePayload: Record<string, any> = {
        'Status': 'Abgemeldet',
        'CancellationDate': new Date().toISOString(),
        'TeilnehmerID': null,
      };
      // Audit-Felder optional dazu - aeltere Subsites haben die Spalten evtl. noch
      // nicht (kommt erst mit Commit a10a608). Ein 400 von SP würde dann die
      // ganze Abmeldung blocken. Strategie: erst mit Audit-Feldern versuchen,
      // bei Misserfolg ohne sie nochmal probieren.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fullPayload: Record<string, any> = { ...corePayload };
      if (auditName) fullPayload['CancelledByName'] = auditName;
      if (auditEmail) fullPayload['CancelledByEmail'] = auditEmail;
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
      let response = await this._merge(url, fullPayload);
      if (!response.ok && (auditName || auditEmail)) {
        // Fallback ohne Audit-Felder (Subsite-Liste hat die Spalten noch nicht)
        console.warn('[DEX] cancelRegistration with audit failed (' + response.status + '), retrying without audit fields');
        response = await this._merge(url, corePayload);
      }
      // v7.31: Counter mit aktuellem Max syncen, damit er nicht "davonrast"
      // wenn der höchste ID-Inhaber sich abmeldet. Best-effort, blockiert
      // die Abmeldung nicht wenn's fehlschlägt. syncCounterToMax liest den
      // Max-Wert intern frisch (race-frei gegen parallele Anmeldungen).
      if (response.ok) {
        try { await this.syncCounterToMax(subsiteUrl); } catch { /* */ }
      }
      return response.ok;
    } catch (err) {
      console.warn('[DEX] cancelRegistration error:', err);
      return false;
    }
  }

  /**
   * v18.11: Proaktive Absage durch einen Teilnehmer, der sich NICHT angemeldet
   * hat („Ich nehme nicht teil"). Legt eine Teilnehmer-Zeile direkt mit
   * Status='Abgemeldet' an — KEINE Sitzplatz-Reservierung, KEINE TeilnehmerID.
   * Profil-Daten (Vorname/Nachname/Location/JobTitle/Department) werden geladen,
   * damit die Abmeldungs-Liste im Admin-Center dieselben Spalten füllen kann
   * wie Teilnehmer-/Warteliste. Marker `_declined` in CustomData unterscheidet
   * die proaktive Absage von einer regulären Abmeldung (die nach vorheriger
   * Anmeldung erfolgte).
   */
  public async declineRegistration(
    subsiteUrl: string,
    firstName: string,
    surname: string,
    participantEmail: string,
    actorName?: string,
    actorEmail?: string
  ): Promise<boolean> {
    try {
      const myEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const profile = (participantEmail || '').toLowerCase() === myEmail
        ? await this.getCurrentUserProfile()
        : await this.getUserProfileByEmail(participantEmail);
      const nowIso = new Date().toISOString();
      // v22.57: Claims-Token-Schutz. In manchen Kontexten liefert der Browser
      // als „displayName" das SharePoint-Claims-Login (z.B.
      // „i:0#.f|membership|user@deloitte.de" bzw. „0#.f|membership|…"). Das
      // landete bisher 1:1 als Vorname in der Absage-Zeile. Wir verwenden
      // deshalb bevorzugt den sauberen Namen aus dem Benutzerprofil und filtern
      // Claims-artige Werte raus.
      const looksLikeClaim = (s: string): boolean => /\|membership\||0#\.f\||^i:0#/i.test((s || '').trim());
      const cleanFirst = looksLikeClaim(firstName) ? '' : (firstName || '').trim();
      const cleanLast = looksLikeClaim(surname) ? '' : (surname || '').trim();
      const effFirst = (profile.firstName || '').trim() || cleanFirst;
      const effLast = (profile.lastName || '').trim() || cleanLast;
      // Anzeigename: bevorzugt Profil-PreferredName, sonst Vor-/Nachname,
      // sonst die E-Mail (nie das Claims-Token).
      const effName = (profile.displayName && !looksLikeClaim(profile.displayName) ? profile.displayName : '')
        || `${effFirst} ${effLast}`.trim()
        || participantEmail;
      const auditNameRaw = actorName || this.context.pageContext.user.displayName || '';
      const auditName = looksLikeClaim(auditNameRaw) ? effName : auditNameRaw;
      const auditEmail = (actorEmail || this.context.pageContext.user.email || '').toLowerCase();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        'Vorname': effFirst,
        'Nachname': effLast,
        'ParticipantName': effName,
        'ParticipantEmail': participantEmail,
        'Department': profile.department,
        'Location': profile.location,
        'JobTitle': profile.jobTitle,
        'Phone': profile.phone,
        'Company': profile.company,
        'Status': 'Abgemeldet',
        'RegistrationDate': nowIso,
        'CancellationDate': nowIso,
        // Marker: proaktive Absage (nie angemeldet gewesen).
        'CustomData': JSON.stringify({ _declined: 'true' }),
      };
      if (auditName) { payload['RegisteredByName'] = auditName; payload['CancelledByName'] = auditName; }
      if (auditEmail) { payload['RegisteredByEmail'] = auditEmail; payload['CancelledByEmail'] = auditEmail; }
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`;
      let resp = await this._post(url, payload);
      if (!resp.ok && (auditName || auditEmail || payload['Company'])) {
        // Fallback ohne Audit-Felder (alte Subsite-Liste ohne diese Spalten).
        // v24.32: zusätzlich Company strippen — fehlt die Spalte, würde der
        // Insert sonst auch im Fallback an Company scheitern.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const noAudit: Record<string, any> = { ...payload };
        delete noAudit['RegisteredByName']; delete noAudit['RegisteredByEmail'];
        delete noAudit['CancelledByName']; delete noAudit['CancelledByEmail'];
        delete noAudit['Company'];
        resp = await this._post(url, noAudit);
      }
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] declineRegistration error:', err);
      return false;
    }
  }

  /**
   * Teilnehmer einchecken (Status auf 'Eingecheckt' setzen).
   * v7.16: Erfasst zusätzlich, WANN und VON WEM der Check-In ausgelöst
   * wurde (CheckedInDate / CheckedInByName / CheckedInByEmail). Diese
   * Spalten werden bei neuen Events ueber createRegistrationList() automatisch
   * angelegt; für bestehende Events muss der Admin einmalig die Kachel
   * "Spalten fixen" im Admin-Center klicken, damit der Check-In nicht mit
   * HTTP 400 fehlschlägt.
   */
  public async checkInParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const me = this.context.pageContext.user;
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        {
          'Status': 'Eingecheckt',
          'CheckedInDate': new Date().toISOString(),
          'CheckedInByName': me.displayName || '',
          'CheckedInByEmail': me.email || me.loginName || '',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * v23.28/v23.29: Teilnehmer als „No-Show" markieren (war angemeldet, aber
   * nicht erschienen). Reuse der Check-in-Audit-Spalten (CheckedInBy*), damit
   * kein neues Schema nötig ist. **Nur für Events, deren Teilnehmerliste die
   * 'No-Show'-Choice kennt** (= ab v23.28 NEU angelegte Events). Bestehende
   * Events werden bewusst NICHT automatisch migriert — dort lehnt SharePoint
   * den Wert ab (HTTP 400) und die Methode liefert `false`.
   */
  public async markNoShowParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const me = this.context.pageContext.user;
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        {
          'Status': 'No-Show',
          'CheckedInDate': new Date().toISOString(),
          'CheckedInByName': me.displayName || '',
          'CheckedInByEmail': me.email || me.loginName || '',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Teilnehmer auschecken (Status zurück auf 'Angemeldet' setzen)
   */
  public async checkOutParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Status': 'Angemeldet' }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Status eines Teilnehmers auf 'QR versendet' setzen
   */
  public async setQRSentStatus(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Status': 'QR versendet' }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Registrierung per Email auf einer Subsite finden
   */
  public async getRegistrationByEmail(
    subsiteUrl: string,
    email: string
  ): Promise<SPRegistration | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,ParticipantName,ParticipantEmail,Status,RegistrationDate,RegisteredByName,RegisteredByEmail,CancellationDate,CancelledByName,CancelledByEmail,CustomData,Department,JobTitle,Location,Company&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      if (data.value && data.value.length > 0) return data.value[0];
      return null;
    } catch {
      return null;
    }
  }

  /**
   * Aktuelle Teilnehmeranzahl ermitteln
   */
  public async getRegistrationCount(subsiteUrl: string): Promise<{ registered: number; waitlist: number }> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allItems: any[] = [];
    // $top=5000 (SP-REST-Maximum) statt 500 — sonst werden bei Events mit
    // ≥500 Einträgen die Counts auf den Event-Karten falsch berechnet,
    // weil SharePoint bei $orderby+$top mit ILS nicht zuverlässig nextLink
    // liefert wenn die Page exakt voll ist.
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }

    const registered = allItems.filter((i: { Status: string }) => i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt').length;
    const waitlist = allItems.filter((i: { Status: string }) => i.Status === 'Warteliste').length;
    return { registered, waitlist };
  }

  /**
   * v22.74: Aktive + Warteliste-E-Mails einer Teilnehmerliste (lowercase) —
   * für die EINDEUTIGE Personenzählung einer Klammer über alle Sub-Events
   * (eine Person, die sich für mehrere Sub-Events anmeldet, zählt einmal).
   */
  public async getParticipantEmailsByStatus(subsiteUrl: string): Promise<{ active: string[]; waitlist: string[] }> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allItems: any[] = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,Status&$top=5000`;
    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    const active: string[] = [];
    const waitlist: string[] = [];
    for (const i of allItems) {
      // v23.3: Zeilen OHNE (gueltige) ParticipantEmail wurden frueher komplett
      // uebersprungen — dadurch zaehlte die entdoppelte Klammer-/Kachel-Zahl
      // weniger Koepfe als die Tabelle Zeilen hat (z.B. 188 statt 190). Eine
      // Anmeldung ohne E-Mail ist trotzdem ein realer Kopf (belegt einen Platz),
      // bekommt nur keine Mails. Deshalb als eigener Schluessel (Zeilen-Id)
      // mitzaehlen, statt sie zu verschlucken.
      const email = (i.ParticipantEmail || '').toLowerCase().trim();
      const key = email || `__noemail#${i.Id}`;
      if (i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt') active.push(key);
      else if (i.Status === 'Warteliste') waitlist.push(key);
    }
    return { active, waitlist };
  }

  /**
   * Title-Feld (= Teilnehmer-ID) aktualisieren
   */
  public async updateRegistrationTitle(
    subsiteUrl: string,
    itemId: number,
    newTitle: string
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Title': newTitle }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== Bild-Upload ====================

  /**
   * SiteAssets-Unterordner sicherstellen:
   * - DEX_EventImages (Event-Bilder)
   * - DEX_Logos (Deloitte-Logo für E-Mail-Templates, manuell hochgeladen)
   */
  public async ensureAssetsFolders(): Promise<void> {
    const baseUrl = this.context.pageContext.web.serverRelativeUrl;
    const folders = ['DEX_EventImages', 'DEX_Logos'];

    for (const folder of folders) {
      const folderUrl = `${baseUrl}/SiteAssets/${folder}`;
      try {
        const check = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')`,
          SPHttpClient.configurations.v1
        );
        if (check.ok) continue;
      } catch { /* */ }

      try {
        await this._post(`${this.siteUrl}/_api/web/folders`, {
          '__metadata': { 'type': 'SP.Folder' },
          'ServerRelativeUrl': folderUrl,
        });
        // Ordner erstellt
      } catch {
        console.warn(`[DEX] Konnte ${folder} Ordner nicht erstellen`);
      }
    }
  }

  /**
   * Event-Bild als Attachment an ein DEX_Events-Item anhängen.
   * Löscht zuerst alle bestehenden Bild-Attachments (Präfix __eventimage__),
   * dann wird das neue Bild hochgeladen. Liefert die ServerRelativeUrl als absolute URL.
   * Vorteil: keine SiteAssets-Berechtigungen nötig - wer das Item editieren darf,
   * darf auch Attachments hinzufügen.
   */
  public async uploadEventImageAsAttachment(eventId: number, file: File): Promise<string> {
    const IMAGE_PREFIX = '__eventimage__';
    try {
      // 1. Bestehende Bild-Attachments löschen (nur __eventimage__-Präfixe)
      try {
        const listResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
          SPHttpClient.configurations.v1
        );
        if (listResp.ok) {
          const listData = await listResp.json();
          const files = listData.value || listData.d?.results || [];
          for (const f of files) {
            const fn: string = f.FileName || '';
            if (fn.indexOf(IMAGE_PREFIX) === 0) {
              try {
                await this._delete(
                  `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fn)}')`
                );
              } catch { /* ignore */ }
            }
          }
        }
      } catch { /* ignore */ }

      // 2. Neues Bild hochladen mit Präfix
      const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
      const safeName = `${IMAGE_PREFIX}${Date.now().toString(36)}.${ext}`;

      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
          body: file,
        } as ISPHttpClientOptions
      );

      if (response.ok) {
        const data = await response.json();
        const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
        if (relUrl) return `${window.location.origin}${relUrl}`;
      } else {
        console.warn('[DEX] Image attachment upload status:', response.status);
      }

      // Fallback: URL aus bekanntem Pfad
      const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
      return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${safeName}`;
    } catch (err) {
      console.warn('[DEX] uploadEventImageAsAttachment error:', err);
    }
    return '';
  }

  /**
   * EventImageUrl-Feld eines DEX_Events-Items setzen (kleines MERGE).
   */
  public async updateEventImageUrl(eventId: number, url: string): Promise<boolean> {
    try {
      const resp = await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        { 'EventImageUrl': url }
      );
      return resp.ok || resp.status === 406;
    } catch {
      return false;
    }
  }

  /**
   * Dokument als Attachment an ein DEX_Events-Item anfügen.
   * Nutzt native SharePoint List Item Attachments - keine Ordner nötig.
   */
  public async uploadEventDocument(eventId: number, file: File): Promise<string> {
    try {
      const fileName = file.name.replace(/[#%&*:<>?/\\|]/g, '_');

      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
          body: file,
        } as ISPHttpClientOptions
      );

      if (response.ok) {
        const data = await response.json();
        const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
        if (relUrl) return `${window.location.origin}${relUrl}`;
      } else {
        console.warn('[DEX] Attachment upload status:', response.status);
      }

      // Fallback: URL aus bekanntem Pfad
      const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
      return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${fileName}`;
    } catch (err) {
      console.warn('[DEX] uploadEventDocument error:', err);
    }
    return '';
  }

  /**
   * Dokument-Attachment von einem DEX_Events-Item löschen.
   * Wird beim Edit verwendet, wenn der User ein bestehendes Dokument entfernt.
   */
  public async deleteEventDocument(eventId: number, fileName: string): Promise<boolean> {
    try {
      const resp = await this._delete(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`
      );
      return resp.ok || resp.status === 200 || resp.status === 204;
    } catch (err) {
      console.warn('[DEX] deleteEventDocument error:', err);
      return false;
    }
  }

  /**
   * Attachments eines DEX_Events-Items laden.
   * Bilder mit Präfix __eventimage__ werden ausgefiltert (nur für EventImageUrl).
   */
  public async getEventAttachments(eventId: number): Promise<Array<{ name: string; url: string; size: number }>> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        const files = data.value || data.d?.results || [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return files
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .filter((f: any) => (f.FileName || '').indexOf('__eventimage__') !== 0)
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .map((f: any) => ({
            name: f.FileName || '',
            url: `${window.location.origin}${f.ServerRelativeUrl || ''}`,
            size: 0,
          }));
      }
    } catch { /* Attachments nicht verfügbar */ }
    return [];
  }

  // ==================== Profil-Daten ====================

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
  private async canRegisterForOthers(subsiteUrl: string, targetParticipantEmail: string): Promise<boolean> {
    // v19.6: Mehrere mögliche Identitäten des eingeloggten Users sammeln.
    // `pageContext.user.email` ist im SharePoint-Mobile-WebView nicht immer
    // gesetzt bzw. weicht vom in OrganizerEmail gespeicherten SMTP-Wert ab —
    // deshalb zusätzlich die E-Mail aus dem `loginName` (Claims-Format
    // `i:0#.f|membership|user@domain`) als Fallback heranziehen.
    const sessionIdentities = new Set<string>();
    const rawEmail = (this.context.pageContext.user.email || '').toLowerCase().trim();
    if (rawEmail) sessionIdentities.add(rawEmail);
    const loginName = (this.context.pageContext.user.loginName || '').toLowerCase();
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
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0 && (items[0].Role === 'Admin' || items[0].Role === 'Organizer')) return true;
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
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${encodeURIComponent(subsiteUrl.replace(/'/g, "''"))}'&$top=1&$select=OrganizerEmail,EmailTemplateOverrides`,
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
      const sessionProfile = await this.getCurrentUserProfile();
      const sessionJt = (sessionProfile.jobTitle || '').toLowerCase();
      if (sessionJt.indexOf('assisten') >= 0 || sessionJt.indexOf('assistan') >= 0) {
        const targetProfile = await this.getUserProfileByEmail(targetParticipantEmail);
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
  public async getCurrentUserProfile(): Promise<{ department: string; location: string; jobTitle: string; phone: string; firstName: string; lastName: string; displayName: string; company: string }> {
    const empty = { department: '', location: '', jobTitle: '', phone: '', firstName: '', lastName: '', displayName: '', company: '' };
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
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
      if (!company) { company = await this.getMyCompanyViaGraph(); }
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
   * Location und Phone aus dem aktuellen Benutzerprofil neu laden und ueberschreiben.
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
  public async fixEventParticipantsProfileData(subsiteUrl: string, n: number = 1000): Promise<{ scanned: number; updated: number; failedLookups: number }> {
    let scanned = 0;
    let updated = 0;
    let failedLookups = 0;
    const sleep = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    if (!subsiteUrl) return { scanned, updated, failedLookups };
    try {
      const listResp = await this.context.spHttpClient.get(
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
            const p = await this.getUserProfileByEmail(email);
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
            const ok = await this._merge(
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

  public async fixRecentParticipantsProfileData(n: number): Promise<{ scanned: number; updated: number; failedLookups: number }> {
    let scanned = 0;
    let updated = 0;
    let failedLookups = 0;
    const sleep = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    try {
      const events = await this.getEvents();
      for (const evt of events) {
        if (!evt.SubsiteUrl) continue;
        try {
          const listResp = await this.context.spHttpClient.get(
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
                const p = await this.getUserProfileByEmail(email);
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
              const ok = await this._merge(
                `${evt.SubsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${it.Id})`,
                {
                  'JobTitle': profile.jobTitle || it.JobTitle || '',
                  'Department': profile.department || it.Department || '',
                  'Location': profile.location || it.Location || '',
                  'Phone': profile.phone || it.Phone || '',
                }
              );
              if (ok && (ok as { ok: boolean }).ok) updated += 1;
            } catch { /* einzelnen ueberspringen */ }
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
  public async getUserProfileByEmail(email: string): Promise<{ department: string; location: string; jobTitle: string; phone: string; firstName: string; lastName: string; displayName: string; company: string }> {
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
      const directUrl = `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${encodeURIComponent(email)}'`;
      const response = await this.context.spHttpClient.get(directUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];
        const profile = extractProfile(props);
        if (profile.jobTitle || profile.department || profile.location || profile.phone) {
          return profile;
        }
      }
    } catch { /* weiter zu Fallback */ }

    // 2) Fallback: echten LoginName (UPN-Claim) ueber siteusers/getbyemail auflösen
    // Deckt UPN != SMTP, Guest-Accounts und Alias-SMTP-Adressen ab.
    try {
      const siteUserUrl = `${this.siteUrl}/_api/web/siteusers/getbyemail('${email.replace(/'/g, "''")}')?$select=LoginName`;
      const siteUserResp = await this.context.spHttpClient.get(siteUserUrl, SPHttpClient.configurations.v1);
      if (!siteUserResp.ok) return empty;
      const siteUserData = await siteUserResp.json();
      const loginName: string = siteUserData.LoginName || siteUserData.d?.LoginName || '';
      if (!loginName) return empty;

      const profileUrl = `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(loginName)}'`;
      const profileResp = await this.context.spHttpClient.get(profileUrl, SPHttpClient.configurations.v1);
      if (!profileResp.ok) return empty;
      const profileData = await profileResp.json();
      const props: Array<{ Key: string; Value: string }> = profileData.UserProfileProperties || [];
      return extractProfile(props);
    } catch {
      return empty;
    }
  }

  // ==================== Hilfsmethoden ====================

  private async listExists(listName: string): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')`,
        SPHttpClient.configurations.v1
      );
      // 200 = OK, 403 = existiert aber kein Zugriff (User hat nur Read)
      return response.ok || response.status === 403;
    } catch {
      return false;
    }
  }

  /**
   * ID der SharePoint-Gruppe "DEALL" (alle Deloitte-Mitarbeiter) ermitteln.
   */
  /**
   * ID der Visitors-Gruppe ermitteln (dort ist DEALL / alle Deloitte-Mitarbeiter hinterlegt).
   */
  private async getVisitorsGroupId(): Promise<number | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedvisitorgroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const d = await resp.json();
        return d.d?.Id || d.Id || null;
      }
    } catch { /* */ }
    return null;
  }

  private async _post(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
      },
      body: JSON.stringify(body),
    };
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }

  /**
   * MERGE-Request für Item-Updates.
   * SharePoint verarbeitet den MERGE korrekt, antwortet aber auf manchen
   * Subsite-Listen mit 406 (Accept-Format nicht unterstützt). Da bei MERGE
   * kein Response-Body benötigt wird, behandeln wir 406 als Erfolg.
   */
  private async _merge(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    // 406: SharePoint hat den MERGE ausgeführt, kann aber nicht im
    // gewünschten Format antworten. Daten sind trotzdem gespeichert.
    if (response.status === 406) {
      return { ok: true, status: 204, statusText: 'No Content' } as unknown as SPHttpClientResponse;
    }
    return response;
  }

  // v7.28: Variante von _merge, die den uebergebenen ETag im IF-MATCH-Header
  // mitsendet (statt '*'). SharePoint vergleicht den ETag mit dem aktuellen
  // Stand des Items und antwortet mit HTTP 412 (Precondition Failed), wenn
  // ein anderer Client zwischenzeitlich geschrieben hat. So können wir
  // optimistic-concurrency-Pattern für den TeilnehmerID-Counter umsetzen.
  private async _mergeIfMatch(url: string, body: object, etag: string): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': etag,
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (response.status === 406) {
      return { ok: true, status: 204, statusText: 'No Content' } as unknown as SPHttpClientResponse;
    }
    return response;
  }

  // v7.28: Atomare TeilnehmerID-Vergabe ueber die DEX_TeilnehmerCounter-Liste
  // pro Event-Subsite. Verhindert Race-Conditions wenn viele User gleichzeitig
  // anmelden — ohne das vorher passieren konnte, dass zwei User dieselbe ID
  // bekommen (siehe Bug-Report v7.27 → v7.28).
  //
  // Ablauf:
  //   1. Counter-Item GET'en, ETag aus Response-Header lesen.
  //   2. NextValue + 1 mit IF-MATCH: <etag> via MERGE schreiben.
  //   3. Bei 412 (ETag-Mismatch = jemand war schneller) → kurzes Jitter +
  //      Retry, max 8x.
  //   4. Bei Erfolg: NextValue+1 ist die neue TeilnehmerID des aufrufenden Users.
  //
  // Fallback: Wenn die Counter-Liste nicht existiert (z.B. legacy event ohne
  // "Spalten fixen"-Lauf), kommt undefined zurück — der Aufrufer fällt dann
  // auf das alte (race-anfällige) max+1-Verfahren zurück. Bestandsschutz.
  // v7.28 / v9.10: Nächste TeilnehmerID atomar holen.
  //   1. Counter-Item GET'en, ETag aus Response-Header lesen.
  //   2. Counter-Item PATCH'en mit IF-MATCH=<ETag>, NextValue=current+1.
  //   3. Bei 412 (ETag-Mismatch = jemand war schneller) → Exponential
  //      Backoff mit Full Jitter, dann Retry. Bis zu 40 Versuche.
  //   4. Bei Erfolg: NextValue+1 ist die neue TeilnehmerID des aufrufenden Users.
  //
  // v9.10: Counter-Liste wird ON-DEMAND angelegt+geseeded, falls sie fehlt
  // (z.B. weil das Event vor v7.28 erstellt wurde). Vorher gab undefined
  // zurück → Aufrufer fiel auf max+1 zurück → Race-Condition bei
  // Massen-Anmeldungen. Jetzt: einmalig ensureCounterList() rufen, dann
  // erneut versuchen. Damit ist der race-anfällige Fallback nur noch
  // erreicht, wenn auch das Anlegen scheitert (Permission-Issue).
  //
  // v9.10: Retries 8 → 40, Backoff von festem Jitter auf Exponential
  // Backoff mit Full Jitter (Cap 500ms). Bei 50+ parallelen Anmeldungen
  // wahren 8 Retries praktisch garantiert ausgeschöpft.
  private async getNextTeilnehmerId(subsiteUrl: string): Promise<number | undefined> {
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    const MAX_RETRIES = 40;
    let triedLazyCreate = false;
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this.context.spHttpClient.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return undefined;
      }
      if (!getResp.ok) {
        // 404 = Counter-Liste / Item existiert nicht.
        // v9.10: Statt direkt undefined zu liefern, einmalig versuchen die
        // Liste anzulegen + zu seeden (idempotent). Wenn das klappt, gleich
        // weiter — wenn nicht, geben wir auf.
        if (getResp.status === 404 && !triedLazyCreate) {
          triedLazyCreate = true;
          try {
            await this.ensureCounterList(subsiteUrl);
            // Kein delay — direkt nächste Iteration, die das frische Item liest.
            continue;
          } catch {
            return undefined;
          }
        }
        return undefined;
      }
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return undefined;
      let data;
      try { data = await getResp.json(); } catch { return undefined; }
      // v9.13: NextValue defensiv parsen — handhabt sowohl number als auch
      // (in seltenen SP-Konfigurationen) string. Sorgt dafür dass current
      // bei einer korrekt gespeicherten 165 nie auf 0 fallback'ed.
      const rawNextValue = data?.NextValue ?? data?.d?.NextValue;
      const current = typeof rawNextValue === 'number'
        ? rawNextValue
        : (typeof rawNextValue === 'string' ? (parseInt(rawNextValue, 10) || 0) : 0);
      // v9.13: Counter ist die Source of Truth. Wir vertrauen ihm und
      // inkrementieren atomar via ETag-CAS. KEIN zusätzlicher Lesezugriff
      // auf die Teilnehmerliste mehr — das Counter-Pattern existiert genau,
      // damit wir hier NICHT max(TID) aus der Teilnehmerliste rechnen
      // müssen. Wenn der Counter korrupt sein sollte (z.B. von altem
      // syncCounterToMax-Bug auf 0 gepatcht), gibt's den expliziten
      // "Counter zurücksetzen"-Button im Admin Center für den Fix.
      const next = current + 1;
      const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': next }, etag);
      if (patchResp.ok) return next;
      if (patchResp.status !== 412) {
        // Anderer Fehler (z.B. 500) → kein Sinn weiter zu retry'n
        return undefined;
      }
      // 412 = ETag-Mismatch = jemand war schneller → Exponential Backoff
      // mit Full Jitter (Cap 500ms). Cluster bei Massen-Anmeldungen
      // werden so zuverlässig entzerrt — ohne Backoff laufen alle
      // Clients sekundengleich in den nächsten Conflict.
      const baseDelay = Math.min(500, 50 * Math.pow(1.4, attempt));
      const delay = Math.floor(baseDelay * (0.5 + Math.random()));
      await new Promise(res => setTimeout(res, delay));
    }
    // Nach 40 Retries aufgeben — Aufrufer kann die Anmeldung sauber
    // mit TeilnehmerID=null durchziehen lassen und der Admin lädt
    // anschliessend "IDs neu vergeben".
    console.warn('[DEX] getNextTeilnehmerId: 40 retries erschöpft — TeilnehmerID bleibt unset, Admin sollte IDs neu vergeben.');
    return undefined;
  }

  // Hilfsroutine: prüft ob auf der Counter-Liste die NextValue-Spalte
  // existiert und legt sie an wenn sie fehlt. Idempotent.
  private async ensureCounterListField(subsiteUrl: string): Promise<void> {
    const fieldsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields`;
    // v7.28: NextValue (TeilnehmerID-Automat).
    // v11.36: SeatsTaken / SeatsTakenDurch / SeatsTakenFun — atomare
    // Sitzplatz-Reservierung pro Gruppe (gegen Überbuchung bei
    // zeitgleichen Anmeldungen). Alle Number-Felder, default 0/leer.
    const wanted = ['NextValue', 'SeatsTaken', 'SeatsTakenDurch', 'SeatsTakenFun'];
    for (const name of wanted) {
      try {
        const probe = await this.context.spHttpClient.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields/getbytitle('${name}')`,
          SPHttpClient.configurations.v1
        );
        if (probe.ok) continue;
        await this._post(
          fieldsUrl,
          { '__metadata': { 'type': 'SP.Field' }, 'Title': name, 'FieldTypeKind': 9, 'Required': false }
        );
      } catch { /* best-effort; reserveSeat fällt sonst sauber zurück */ }
    }
  }

  // v7.28: Counter-Liste für ein Event anlegen (1 Liste mit 1 Item) und
  // direkt mit dem aktuellen Max-Wert seeden — damit bestehende Events ohne
  // ID-Lückenproduktion umsteigen können.
  // Idempotent: tut nichts wenn die Liste schon existiert.
  // v7.29-Fix: Item-Inserts nutzen den korrekt _x005f_-encodeten Type-Namen
  // (genauso wie wir das für DEX_Events machen). Vorher wurde der Listen-
  // name 1:1 in den Type uebernommen, was bei Unterstrich stillschweigend
  // zu HTTP 400 führt → leere Counter-Liste.
  private async ensureCounterList(subsiteUrl: string): Promise<{ created: boolean; seededValue?: number }> {
    const probe = await this.context.spHttpClient.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')`,
      SPHttpClient.configurations.v1
    );
    const itemsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items`;
    const seedItem = async (): Promise<number> => {
      const maxId = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
      const resp = await this._post(itemsUrl, {
        '__metadata': { 'type': COUNTER_LIST_ITEM_TYPE },
        'Title': 'TeilnehmerID',
        'NextValue': maxId,
      });
      if (!resp.ok) {
        const errBody = await resp.text().catch(() => '');
        console.warn('[DEX] ensureCounterList: Seed-Item fehlgeschlagen, status=', resp.status, errBody.substring(0, 300));
      }
      return maxId;
    };

    if (probe.ok) {
      // Liste existiert — sicherstellen dass das Schema komplett ist und ein Item drin liegt.
      try { await this.ensureCounterListField(subsiteUrl); } catch { /* */ }
      // v9.13/v9.14: setCounterListPermissions wurde hier ursprünglich
      // mitaufgerufen, hat aber bei laufender Event-Anlage Race-Conditions
      // ausgelöst (breakroleinheritance gegen frisch provisionierte Liste
      // in derselben Request-Welle). Permissions werden jetzt nur noch
      // explizit ueber den "Counter zurücksetzen"-Button gefixt — siehe
      // resetCounterToMax. Bestehende Events können damit per Admin-Klick
      // geheilt werden, neue Events bekommen ihre Permissions im
      // create-Branch unten gesetzt.
      const itemListResp = await this.context.spHttpClient.get(
        `${itemsUrl}?$top=1`,
        SPHttpClient.configurations.v1
      );
      if (itemListResp.ok) {
        const data = await itemListResp.json();
        const list = data.value || data.d?.results || [];
        if (list.length > 0) return { created: false }; // alles ok
      }
      // Liste ohne Item → nachseeden
      const seededValue = await seedItem();
      return { created: false, seededValue };
    }

    // Liste neu anlegen
    await this._post(
      `${subsiteUrl}/_api/web/lists`,
      {
        '__metadata': { 'type': 'SP.List' },
        'BaseTemplate': 100,
        'Title': COUNTER_LIST_NAME,
        'Description': 'Atomarer Counter für TeilnehmerID-Vergabe (ETag-basiert). Nicht manuell editieren.',
        'AllowContentTypes': false,
        'ContentTypesEnabled': false,
        'EnableVersioning': false,
        'EnableMinorVersions': false,
        'OnQuickLaunch': false,
      }
    );
    await this.ensureCounterListField(subsiteUrl);
    // v9.13: Counter-Liste muss explizit Contribute-Rechte für Visitors
    // bekommen, damit normale User die ETag-CAS-Inkrementierung
    // durchführen können. Ohne das schlägt PATCH NextValue mit 401/403
    // fehl → getNextTeilnehmerId gibt undefined zurück → TID landet null
    // (oder im allerersten Lazy-Create-Pfad bei 1).
    try { await this.setCounterListPermissions(subsiteUrl); } catch { /* */ }
    const seededValue = await seedItem();
    return { created: true, seededValue };
  }

  /**
   * Berechtigungen für DEX_TeilnehmerCounter setzen — analog zur
   * Teilnehmerliste:
   *   - Owners der Hauptsite: Full Control (1073741829)
   *   - Visitors (DEALL): Contribute (1073741827) → ETag-CAS-Inkrement
   *   - Organizer-Mail (falls bekannt): Full Control
   *
   * Idempotent: kann auf bestehenden Counter-Listen erneut aufgerufen
   * werden um v9.13-Permissions nachzupatchen. Die Funktion bricht
   * Rollen-Vererbung explizit (clearSubscopes=true), damit Read-Only-
   * Inheritance vom Subsite nicht versehentlich greift.
   */
  private async setCounterListPermissions(subsiteUrl: string, organizerEmail?: string): Promise<void> {
    try {
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );
      const ownersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResp.ok) {
        const ownersData = await ownersResp.json();
        const ownersId = ownersData.Id ?? ownersData.d?.Id;
        if (ownersId) {
          await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${ownersId}, roledefid=1073741829)`,
            {}
          );
        }
      }
      // Visitors → Contribute. KRITISCH: damit normale User
      // den Counter atomar inkrementieren können.
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`,
          {}
        );
      }
      // Organizer optional → Full Control
      if (organizerEmail) {
        try {
          const userResp = await this.context.spHttpClient.get(
            `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(organizerEmail)}')?$select=Id`,
            SPHttpClient.configurations.v1
          );
          if (userResp.ok) {
            const userData = await userResp.json();
            const userId = userData.Id ?? userData.d?.Id;
            if (userId) {
              await this._post(
                `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
                {}
              );
            }
          }
        } catch { /* Organizer-Permission ist optional */ }
      }
    } catch (err) {
      console.warn('[DEX] setCounterListPermissions fehlgeschlagen:', err);
    }
  }

  // v7.28 / v9.13: Aktuellen Max-Wert von TeilnehmerID in der Teilnehmerliste
  // lesen. Wird beim **Seeden** des Counters und beim **Sync nach Reorder**
  // genutzt — die Counter-Liste selbst ist im Normalbetrieb die Source of
  // Truth (siehe getNextTeilnehmerId).
  //
  // **Bugfix v9.13:** Vorher hat $orderby=TeilnehmerID desc&$top=1 unter
  // bestimmten Bedingungen das null-Item zuerst geliefert (SP sortiert NULL-
  // Werte bei Number-Feldern oft als "größter Wert" in desc-Order).
  // Sobald irgendjemand abgemeldet war (TID=null) lief die Funktion ins
  // null-Branch und gab 0 zurück.
  //
  // Konsequenz im alten Code (vor v9.12): syncCounterToMax patcht den
  // Counter auf liveMax=0 RUNTER → nächste Anmeldung kriegt TID=1 →
  // Duplikat zu echten aktiven Teilnehmern. Genau der Fall den der User
  // beim Go-Live live gesehen hat (Theresa #1 obwohl 165 aktive Anmeldungen).
  //
  // Fix: $filter=TeilnehmerID gt 0 schliesst NULL und 0 explizit aus —
  // funktioniert unabhängig von SP-NULL-Sortier-Konventionen.
  private async getCurrentMaxTeilnehmerId(subsiteUrl: string): Promise<number> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=TeilnehmerID&$filter=TeilnehmerID gt 0&$orderby=TeilnehmerID desc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return 0;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length > 0 && items[0].TeilnehmerID != null) return items[0].TeilnehmerID;
    } catch { /* */ }
    return 0;
  }

  // v7.31 / v9.12: Counter mit aktuellem Max-Wert konsistent halten — wird
  // nach Cancel und nach reorderParticipantIDs aufgerufen. Die Logik ist
  // **monotonic up-only**: der Counter geht NIE runter (sonst werden
  // cancelled IDs reused — exakter Duplikat-Bug aus dem Go-Live).
  //
  // - current >= liveMax: nichts zu tun (Counter steht bereits hoch genug).
  // - current  < liveMax: Counter HOCH auf liveMax setzen (z.B. nach
  //   reorderParticipantIDs, der die TIDs im Bereich [1..N_active] vergibt
  //   wo N_active größer sein kann als der bisherige Counter-Stand).
  //
  // Vorher (Bug bis v9.11): exakt umgekehrt — Counter wurde auf liveMax
  // RUNTER gesetzt wenn current > liveMax. Das produzierte sowohl bei
  // Cancel als auch nach IDReorder Duplikate.
  //
  // ETag-CAS mit Retry, damit eine parallele Anmeldung den Counter nicht
  // zwischen Read und Write wegrasselt.
  // v9.13: Oeffentliche Recovery-Methode für den Admin-Button "Counter
  // zurücksetzen". Liest den aktuellen Max-TID aus der Teilnehmerliste und
  // setzt den Counter auf diesen Wert (per ETag-CAS, monotonic up-only via
  // syncCounterToMax). Gibt den neuen Counter-Wert zurück damit der Admin
  // direkt sehen kann auf was es gepatcht wurde.
  public async resetCounterToMax(subsiteUrl: string): Promise<{ counter: number; max: number }> {
    // v11.27: bidirektionaler Reset. Vorher rief diese Methode nur
    // syncCounterToMax auf — das ist monotonic up-only und liess einen
    // zu hohen Counter unverändert. Genau das hat der Maintainer beobachtet:
    // Counter=11, Max-TID=4, Klick auf "Counter zurücksetzen" → keine
    // Aenderung, weiterhin 11. Jetzt setzen wir den Counter explizit
    // auf max(TID) — egal ob er drunter (gefährlich, Doppel-IDs möglich)
    // oder drüber stand (harmlos, nur Lücken-Springen).
    try { await this.ensureCounterList(subsiteUrl); } catch { /* */ }
    const max = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    let finalCounter = 0;
    // ETag-CAS-Loop, falls jemand parallel inserted und den Counter inkrementiert.
    for (let attempt = 0; attempt < 8; attempt++) {
      try {
        const getResp = await this.context.spHttpClient.get(counterItemUrl, SPHttpClient.configurations.v1);
        if (!getResp.ok) break;
        const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
        if (!etag) break;
        const data = await getResp.json();
        const rawCurrent = data?.NextValue ?? data?.d?.NextValue;
        const current = typeof rawCurrent === 'number' ? rawCurrent : (typeof rawCurrent === 'string' ? (parseInt(rawCurrent, 10) || 0) : 0);
        if (current === max) {
          finalCounter = current;
          break;
        }
        const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': max }, etag);
        if (patchResp.ok) {
          finalCounter = max;
          console.warn(`[DEX] resetCounterToMax: counter von ${current} auf ${max} gesetzt (Subsite: ${subsiteUrl}).`);
          break;
        }
        if (patchResp.status !== 412) {
          finalCounter = current;
          break;
        }
        // 412 = jemand war schneller, nochmal lesen+patchen
        await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
      } catch (err) {
        console.warn('[DEX] resetCounterToMax error:', err);
        break;
      }
    }
    return { counter: finalCounter, max };
  }

  private async syncCounterToMax(subsiteUrl: string): Promise<void> {
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    for (let attempt = 0; attempt < 8; attempt++) {
      const liveMax = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this.context.spHttpClient.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return;
      }
      if (!getResp.ok) return;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return;
      let data;
      try { data = await getResp.json(); } catch { return; }
      const current = typeof data?.NextValue === 'number' ? data.NextValue : 0;
      if (current >= liveMax) return; // bereits konsistent — Counter ist nicht "zu klein"
      const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': liveMax }, etag);
      if (patchResp.ok) {
        console.warn(`[DEX] syncCounterToMax: counter von ${current} auf ${liveMax} hochgezogen.`);
        return;
      }
      if (patchResp.status !== 412) return;
      // 412 = jemand war schneller, nochmal lesen+patchen
      await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
    }
    // Nach 8 Retries aufgeben — best-effort, blockiert keine andere Aktion
  }

  /**
   * v19.28: Eine Teilnehmer-Registrierung endgültig aus der Subsite-Liste
   * löschen (hartes DELETE, kein Recycle-Bin). Use-Case: abgemeldete
   * Test-Anmeldungen aus der Abmeldungen-Liste entfernen, damit die Übersicht
   * sauber bleibt. Die Berechtigung (Admin/Organizer) wird in der UI geprüft.
   */
  public async deleteRegistration(subsiteUrl: string, itemId: number): Promise<boolean> {
    try {
      const resp = await this._delete(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`);
      return resp.ok;
    } catch (err) {
      console.warn('[DEX] deleteRegistration failed:', err);
      return false;
    }
  }

  /**
   * v21: Item-Level-Security („nur eigene Elemente", 2/2) auf den globalen
   * Queue-Listen DEX_Outlook + DEX_IDReorder nachziehen. DEX_Emails und
   * DEX_AccessFix haben sie bereits; DEX_TeamJoinRequests bekommt sie
   * bewusst NICHT (der Team-Lead muss fremde Beitritts-Anfragen lesen).
   * Idempotent; wird vom Admin-Reparatur-Button mit ausgeführt.
   */
  public async hardenQueueListsIls(): Promise<{ fixed: string[]; failed: string[] }> {
    const targets = ['DEX_Outlook', 'DEX_IDReorder'];
    const fixed: string[] = [];
    const failed: string[] = [];
    for (const listName of targets) {
      try {
        const resp = await this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE',
            },
            body: JSON.stringify({ '__metadata': { 'type': 'SP.List' }, 'ReadSecurity': 2, 'WriteSecurity': 2 }),
          }
        );
        if (resp.ok) fixed.push(listName); else failed.push(listName);
      } catch { failed.push(listName); }
    }
    return { fixed, failed };
  }

  // ==================== v21: Archivierung ====================
  // Globale Queue-/Log-Listen, deren Zeilen abgelaufener Events ins
  // DEX_Archive wandern (verschieben = nach Insert aus der Quelle löschen).
  // DEX_Emails/Outlook/IDReorder/ChangeLog matchen über EventId, DEX_AccessFix
  // über SubsiteUrl (hat keine EventId).
  // v22.2: Pro Quelle eine SCHLANKE Feldauswahl für den Archiv-Lauf.
  // WICHTIG: DEX_Emails OHNE `Body` — 1000+ komplette HTML-Mail-Bodies in
  // den Browser zu laden hängte den Archiv-Lauf praktisch auf (zig MB).
  // Der Mail-Body wird bewusst NICHT mitarchiviert (Metadaten reichen als
  // Nachweis; die Mail ist längst versendet).
  // hasStatus: Liste besitzt eine Status-Spalte → Zeilen mit 'Pending'
  // (Flow hat sie noch nicht verarbeitet) werden NICHT archiviert, damit
  // keine unversendete Mail / kein offener Auftrag aus der Queue verschwindet.
  // v23.47: select MUSS den GESAMTEN Inhalt der Zeile abdecken — das Archiv
  // ist die End-Ablage, der Originaldatensatz wird nach dem Insert gelöscht.
  // Frühere Selects ließen Inhalte weg: DEX_Emails OHNE `Body` (kompletter
  // Mailtext!) und DEX_Outlook OHNE `CalendarLink` — die fehlten damit im
  // Payload und gingen beim Löschen verloren. Jetzt: vollständige Feldlisten.
  private static readonly ARCHIVE_SOURCES: Array<{ list: string; matchBy: 'eventId' | 'subsiteUrl'; select: string; hasStatus: boolean }> = [
    { list: 'DEX_Emails', matchBy: 'eventId', select: 'Id,Title,Recipient,RecipientName,Body,EmailType,EventTitle,EventId,Status,Cc,Bcc,Importance,Created', hasStatus: true },
    { list: 'DEX_Outlook', matchBy: 'eventId', select: 'Id,Title,Attendee,EventId,ActionType,Status,CalendarLink,Created', hasStatus: true },
    { list: 'DEX_IDReorder', matchBy: 'eventId', select: 'Id,Title,EventId,EventNumber,SubsiteUrl,Status,CancelledName,CancelledEmail,Created', hasStatus: true },
    { list: 'DEX_ChangeLog', matchBy: 'eventId', select: 'Id,Title,Action,TargetType,TargetId,TargetName,EventId,EventTitle,ActorName,ActorEmail,Details,Created', hasStatus: false },
    { list: 'DEX_AccessFix', matchBy: 'subsiteUrl', select: 'Id,Title,SubsiteUrl,ItemId,ParticipantEmail,Status,Created', hasStatus: true },
  ];

  /**
   * v21: DEX_Archive anlegen (Site-Collection-Root) — generisches Schema, das
   * Zeilen aus mehreren Quell-Listen aufnimmt: SourceList, EventId, EventTitle,
   * OriginalId, ArchivedAt + Payload (JSON der Originalzeile). NUR Admins
   * (Owners) bekommen Zugriff (setArchiveListPermissions, kein Visitors-Grant).
   */
  public async ensureArchiveList(): Promise<void> {
    const listName = 'DEX_Archive';
    const exists = await this.listExists(listName);
    if (exists) return;
    const createResp = await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Archiv abgelaufener Event-Zeilen aus den Queue-/Log-Listen (v21). Nur für Admins lesbar.',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    if (!createResp.ok) {
      console.warn('[DEX] DEX_Archive konnte nicht angelegt werden — vermutlich fehlen Owner-Rechte.');
      return;
    }
    const fields: Array<{ title: string; type: number; metaType?: string }> = [
      { title: 'SourceList', type: 2 },
      { title: 'EventId', type: 2 },
      { title: 'EventTitle', type: 2 },
      { title: 'OriginalId', type: 9 },
      { title: 'ArchivedAt', type: 4 },
      { title: 'Payload', type: 3, metaType: 'SP.FieldMultiLineText' },
    ];
    for (const f of fields) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const payload: Record<string, any> = {
          '__metadata': { 'type': f.metaType || 'SP.Field' },
          'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
        };
        if (f.metaType === 'SP.FieldMultiLineText') { payload['RichText'] = false; payload['NumberOfLines'] = 6; }
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
    try { await this.configureDefaultView(listName, ['SourceList', 'EventTitle', 'EventId', 'OriginalId', 'ArchivedAt']); } catch { /* */ }
    try { await this.setArchiveListPermissions(listName); } catch { /* */ }
  }

  /** Admin-only: Vererbung brechen, NUR Owners (Site-Admins) Full Control —
   *  bewusst KEIN Visitors-/Members-Grant, damit das Archiv nicht von allen
   *  lesbar ist. */
  private async setArchiveListPermissions(listName: string): Promise<void> {
    try {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`, {});
      const ownersResp = await this.context.spHttpClient.get(`${this.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1);
      if (ownersResp.ok) {
        const d = await ownersResp.json();
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {});
      }
    } catch (e) { console.warn('[DEX] setArchiveListPermissions failed:', e); }
  }

  // ==================== Wochenbericht (v23.8) ====================

  /** Tracking-Liste für den wöchentlichen Admin-Bericht. Pro versendetem
   *  Bericht eine Zeile (Created = Versandzeit; PeriodFrom/PeriodTo = der
   *  abgedeckte Zeitraum). Quelle der Wahrheit für „wann lief der letzte
   *  Bericht". */
  public async ensureWeeklyReportsList(): Promise<void> {
    const listName = 'DEX_WeeklyReports';
    const exists = await this.listExists(listName);
    if (!exists) {
      const createResp = await this._post(`${this.siteUrl}/_api/web/lists`, {
        '__metadata': { 'type': 'SP.List' },
        'Title': listName,
        'Description': 'Versand-Protokoll des wöchentlichen Admin-Berichts (v23.8).',
        'BaseTemplate': 100,
        'AllowContentTypes': false,
      });
      if (!createResp.ok) {
        console.warn('[DEX] DEX_WeeklyReports konnte nicht angelegt werden — vermutlich fehlen Owner-Rechte.');
        return;
      }
      for (const f of [{ title: 'PeriodFrom', type: 4 }, { title: 'PeriodTo', type: 4 }]) {
        try {
          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
            '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
          });
        } catch { /* einzelne Feld-Fehler ignorieren */ }
      }
      try { await this.configureDefaultView(listName, ['PeriodFrom', 'PeriodTo']); } catch { /* */ }
    }
    // v23.36: DraftEventIds-Snapshot (JSON-Array der Event-IDs, die beim letzten
    // Bericht noch Entwürfe waren) — idempotent nachziehen, auch auf Bestands-
    // Listen. So erkennt der nächste Bericht „Entwurf ist live gegangen".
    try {
      const fieldsResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$filter=InternalName eq 'DraftEventIds'&$top=1`,
        SPHttpClient.configurations.v1
      );
      const fieldsData = fieldsResp.ok ? await fieldsResp.json() : null;
      const has = fieldsData && (fieldsData.value || fieldsData.d?.results || []).length > 0;
      if (!has) {
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' }, 'Title': 'DraftEventIds', 'FieldTypeKind': 3, 'Required': false,
        });
      }
    } catch { /* best-effort */ }
  }

  /** Letzter Bericht: Created (Versandzeit) + PeriodTo + Entwurfs-Snapshot. */
  public async getLastWeeklyReport(): Promise<{ created: string; periodTo: string; draftEventIds: string[] } | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_WeeklyReports')/items?$select=Created,PeriodTo,DraftEventIds&$orderby=Created desc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return null;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return null;
      let draftEventIds: string[] = [];
      try { const arr = JSON.parse(items[0].DraftEventIds || '[]'); if (Array.isArray(arr)) draftEventIds = arr.map((x: unknown) => String(x)); } catch { /* */ }
      return { created: items[0].Created || '', periodTo: items[0].PeriodTo || items[0].Created || '', draftEventIds };
    } catch { return null; }
  }

  /** Versand des Berichts protokollieren (eine Zeile) + Entwurfs-Snapshot. */
  public async recordWeeklyReport(fromIso: string, toIso: string, draftEventIds?: string[]): Promise<void> {
    try {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_WeeklyReports')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_WeeklyReportsListItem' },
        'Title': `Weekly ${toIso.slice(0, 10)}`,
        'PeriodFrom': fromIso,
        'PeriodTo': toIso,
        'DraftEventIds': JSON.stringify(draftEventIds || []),
      });
    } catch (e) { console.warn('[DEX] recordWeeklyReport failed:', e); }
  }

  // ==================== DEX_OrganizerRequests (v23.37) ====================
  // Anträge „Organizer werden". Jeder authentifizierte User darf einen Antrag
  // anlegen (setQueueListPermissions); Admins sehen offene Anträge in der App
  // und bestätigen sie (→ Organizer-Rolle wird vergeben).

  public async ensureOrganizerRequestsList(): Promise<void> {
    const listName = 'DEX_OrganizerRequests';
    const exists = await this.listExists(listName);
    if (exists) return;
    const cr = await this._post(`${this.siteUrl}/_api/web/lists`, {
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
      try { await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload); } catch { /* */ }
    }
    try { await this.configureDefaultView(listName, ['RequesterName', 'RequesterEmail', 'RequesterLocation', 'Status', 'Created', 'DecidedByEmail']); } catch { /* */ }
    try { await this.setQueueListPermissions(listName); } catch { /* */ }
  }

  // ==================== v24.6: Organizer-Archiv (pro Person ausblenden) ====================
  // Reiner Anzeige-Filter: ein abgelaufenes Event kann der Organizer aus SEINER
  // Übersicht ausblenden (eine Zeile pro Event+Person). Das Event selbst bleibt
  // mit allen Daten erhalten und für andere sichtbar — KEINE Datenlöschung.
  public async ensureOrganizerArchivedList(): Promise<void> {
    const listName = 'DEX_OrganizerArchived';
    const exists = await this.listExists(listName);
    if (exists) return;
    const cr = await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Pro Organizer ausgeblendete (archivierte) Events (v24.6) — reiner Anzeige-Filter, keine Datenlöschung.',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    if (!cr.ok) { console.warn('[DEX] DEX_OrganizerArchived konnte nicht angelegt werden.'); return; }
    for (const f of [{ title: 'EventId', type: 2 }, { title: 'OrganizerEmail', type: 2 }]) {
      try { await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, { '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false }); } catch { /* */ }
    }
    try { await this.configureDefaultView(listName, ['EventId', 'OrganizerEmail', 'Created']); } catch { /* */ }
    try { await this.setQueueListPermissions(listName); } catch { /* */ }
  }

  public async getOrganizerArchivedEventIds(email: string): Promise<Set<string>> {
    const out = new Set<string>();
    try {
      const e = (email || '').replace(/'/g, "''");
      if (!e) return out;
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items?$select=Id,EventId,OrganizerEmail&$filter=OrganizerEmail eq '${e}'&$top=2000`,
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

  public async archiveEventForOrganizer(eventId: string, email: string): Promise<boolean> {
    try {
      const resp = await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_OrganizerArchivedListItem' },
        'Title': String(eventId).slice(0, 250),
        'EventId': String(eventId),
        'OrganizerEmail': email,
      });
      return resp.ok;
    } catch { return false; }
  }

  public async unarchiveEventForOrganizer(eventId: string, email: string): Promise<boolean> {
    try {
      const e = (email || '').replace(/'/g, "''");
      const idEsc = String(eventId).replace(/'/g, "''");
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items?$select=Id&$filter=OrganizerEmail eq '${e}' and EventId eq '${idEsc}'&$top=50`,
        SPHttpClient.configurations.v1);
      if (!resp.ok) return false;
      const d = await resp.json();
      const rows = d.value || d.d?.results || [];
      let okAll = true;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const r of rows) { const del = await this._delete(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerArchived')/items(${(r as any).Id})`); if (!del.ok) okAll = false; }
      return okAll;
    } catch { return false; }
  }

  public async createOrganizerRequest(email: string, name: string, location: string, message: string): Promise<{ ok: boolean; itemId?: number }> {
    try {
      const resp = await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerRequests')/items`, {
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

  public async getOrganizerRequests(onlyPending: boolean = true): Promise<Array<{ id: number; email: string; name: string; location: string; message: string; status: string; created: string }>> {
    try {
      const filter = onlyPending ? `&$filter=Status eq 'Pending'` : '';
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerRequests')/items?$select=Id,RequesterEmail,RequesterName,RequesterLocation,Message,Status,Created&$orderby=Created desc&$top=200${filter}`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return [];
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return items.map((i: any) => ({ id: i.Id, email: i.RequesterEmail || '', name: i.RequesterName || '', location: i.RequesterLocation || '', message: i.Message || '', status: i.Status || '', created: i.Created || '' }));
    } catch { return []; }
  }

  public async updateOrganizerRequestStatus(id: number, status: 'Approved' | 'Rejected', decidedByEmail: string): Promise<boolean> {
    try {
      const r = await this._merge(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_OrganizerRequests')/items(${id})`, {
        'Status': status, 'DecidedDate': new Date().toISOString(), 'DecidedByEmail': decidedByEmail,
      });
      return r.ok;
    } catch { return false; }
  }

  /** E-Mail-Adressen (Title) aller DEX_Roles-Einträge mit der gegebenen Rolle. */
  /** v23.38: Rollen-Empfänger mit E-Mail UND Anzeigename (für personalisierte
   *  Mails wie den Wochenbericht — „Hallo <Name>" statt generisch „Admin"). */
  public async getRoleRecipients(role: string): Promise<Array<{ email: string; name: string }>> {
    try {
      const esc = role.replace(/'/g, "''");
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Role eq '${encodeURIComponent(esc)}'&$select=Title,UserName&$top=5000`,
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

  public async getRoleEmails(role: string): Promise<string[]> {
    try {
      const esc = role.replace(/'/g, "''");
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Role eq '${encodeURIComponent(esc)}'&$select=Title&$top=5000`,
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
  public async getRoleItemsCreatedSince(role: string, fromIso: string): Promise<Array<{ email: string; created: string }>> {
    try {
      const esc = role.replace(/'/g, "''");
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Role eq '${encodeURIComponent(esc)}' and Created ge '${fromIso}'&$select=Title,Created&$orderby=Created desc&$top=500`,
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
  public async getEventsCreatedSince(fromIso: string): Promise<Array<{ title: string; author: string; created: string; isDraft: boolean }>> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Title,Created,IsFictive,Author/Title&$expand=Author&$filter=Created ge '${fromIso}'&$orderby=Created desc&$top=500`,
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
  public async countRegistrations(subsiteUrl: string, fromIso: string): Promise<{ total: number; since: number }> {
    if (!subsiteUrl) return { total: 0, since: 0 };
    const active = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    let total = 0; let since = 0;
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status,RegistrationDate&$top=5000`;
    let guard = 0;
    const fromTs = new Date(fromIso).getTime();
    while (url && guard < 50) {
      guard++;
      try {
        const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
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

  /** Lädt alle Zeilen einer Liste (paged, nometadata). `select` schränkt die
   *  Felder ein (fürs Zählen leichtgewichtig); ohne select = alle Felder
   *  (für den Payload). */
  private async loadAllListRows(listName: string, select?: string): Promise<Array<Record<string, unknown>>> {
    const rows: Array<Record<string, unknown>> = [];
    let url = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$top=500${select ? `&$select=${select}` : ''}`;
    let guard = 0;
    while (url && guard < 500) {
      guard++;
      const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
      });
      if (!resp.ok) break;
      const data = await resp.json();
      const arr: Array<Record<string, unknown>> = data.value || data.d?.results || [];
      rows.push(...arr);
      url = (data['odata.nextLink'] as string) || (data['@odata.nextLink'] as string) || '';
    }
    return rows;
  }

  private rowMatchesExpired(
    r: Record<string, unknown>, matchBy: 'eventId' | 'subsiteUrl',
    expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
    // v23.39: ALLE aktuell existierenden Event-IDs / Subsites. Eine Zeile, deren
    // Bezug NICHT mehr darin vorkommt, gehört zu einem gelöschten Event
    // (verwaist) und ist ebenfalls archivreif. Die size>0-Wächter verhindern,
    // dass bei (noch) nicht geladenen Events fälschlich ALLES als verwaist gilt.
    allEventIds: Set<string>, allSubsiteUrls: Set<string>
  ): boolean {
    // v22.2: 'Pending' = der Flow hat die Zeile noch nicht verarbeitet —
    // niemals archivieren (sonst verschwindet z.B. eine unversendete Mail
    // aus der Queue). Hängengebliebene Pendings bleiben so in der
    // Arbeitsliste sichtbar, wo der Admin sie sehen soll.
    if (String(r['Status'] || '') === 'Pending') return false;
    if (matchBy === 'eventId') {
      const id = String(r['EventId'] || '').trim();
      if (!id) return false;
      // archivreif = Event abgelaufen ODER nicht mehr existent (gelöscht/verwaist).
      return expiredEventIds.has(id) || (allEventIds.size > 0 && !allEventIds.has(id));
    }
    const su = String(r['SubsiteUrl'] || '').toLowerCase().trim();
    if (!su) return false;
    return expiredSubsiteUrls.has(su) || (allSubsiteUrls.size > 0 && !allSubsiteUrls.has(su));
  }

  /** v21: Zählt die archivreifen Zeilen pro Quell-Liste (leichtgewichtig:
   *  nur Id+EventId bzw. Id+SubsiteUrl). */
  public async countArchivableRows(
    expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
    allEventIds: Set<string> = new Set(), allSubsiteUrls: Set<string> = new Set()
  ): Promise<{ total: number; perList: Record<string, number> }> {
    const perList: Record<string, number> = {};
    let total = 0;
    for (const src of EventService.ARCHIVE_SOURCES) {
      let c = 0;
      try {
        // v22.2: Status mitladen (wo vorhanden), damit der Pending-Ausschluss
        // schon beim Zählen greift und die Box-Zahl zum Lauf passt.
        const base = src.matchBy === 'eventId' ? 'Id,EventId' : 'Id,SubsiteUrl';
        const select = src.hasStatus ? `${base},Status` : base;
        const rows = await this.loadAllListRows(src.list, select);
        for (const r of rows) {
          if (this.rowMatchesExpired(r, src.matchBy, expiredEventIds, expiredSubsiteUrls, allEventIds, allSubsiteUrls)) c++;
        }
      } catch { /* Liste evtl. nicht vorhanden */ }
      perList[src.list] = c;
      total += c;
    }
    return { total, perList };
  }

  /** v21: Verschiebt alle archivreifen Zeilen ins DEX_Archive (Insert →
   *  Delete aus der Quelle). Sequentiell (SP-Throttling). onProgress meldet
   *  Listen- + Zeilen-Fortschritt fürs Modal. */
  public async archiveExpiredRows(
    expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
    eventTitleById: Record<string, string>,
    onProgress?: (listIdx: number, listTotal: number, listName: string, done: number, total: number) => void,
    // v22.2: Abbruch-Check (UI-Button). Sauber: bereits verschobene Zeilen
    // bleiben im Archiv (jede Zeile ist atomar Insert→Delete), der Rest
    // bleibt in der Quelle und kommt beim nächsten Lauf dran.
    shouldCancel?: () => boolean,
    // v23.39: alle aktuellen Event-IDs / Subsites (für die Verwaist-Erkennung).
    allEventIds: Set<string> = new Set(), allSubsiteUrls: Set<string> = new Set()
  ): Promise<{ archived: number; failed: number; cancelled: boolean; perList: Record<string, number> }> {
    const result = { archived: 0, failed: 0, cancelled: false, perList: {} as Record<string, number> };
    const sources = EventService.ARCHIVE_SOURCES;
    // v23.47: Payload soll den Inhalt vollständig festhalten (vorher bei 4000
    // Zeichen gekappt — ein kompletter HTML-Mailtext ist länger und wurde so
    // abgeschnitten). Großzügiger Sicherheits-Cap (60000), der praktisch jeden
    // Mailtext komplett aufnimmt, aber pathologische Riesenwerte begrenzt.
    const MAX_FIELD = 60000;
    const buildPayload = (r: Record<string, unknown>): string => {
      const out: Record<string, unknown> = {};
      for (const k of Object.keys(r)) {
        const v = r[k];
        out[k] = (typeof v === 'string' && v.length > MAX_FIELD) ? `${v.slice(0, MAX_FIELD)}… [gekürzt]` : v;
      }
      return JSON.stringify(out);
    };
    for (let si = 0; si < sources.length; si++) {
      const src = sources[si];
      if (shouldCancel && shouldCancel()) { result.cancelled = true; break; }
      let listArchived = 0;
      try {
        // v22.2: Fortschritt SOFORT melden (vorher kam der erste Callback erst
        // nach dem Komplett-Laden — das Modal wirkte eingefroren).
        if (onProgress) onProgress(si, sources.length, src.list, 0, 0);
        const rows = await this.loadAllListRows(src.list, src.select);
        const matching = rows.filter(r => this.rowMatchesExpired(r, src.matchBy, expiredEventIds, expiredSubsiteUrls, allEventIds, allSubsiteUrls));
        if (onProgress) onProgress(si, sources.length, src.list, 0, matching.length);
        for (let i = 0; i < matching.length; i++) {
          if (shouldCancel && shouldCancel()) { result.cancelled = true; break; }
          const r = matching[i];
          const origId = Number(r['Id'] || 0);
          const eid = src.matchBy === 'eventId' ? String(r['EventId'] || '') : '';
          const title = eid ? (eventTitleById[eid] || '') : '';
          try {
            const ins = await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_Archive')/items`, {
              '__metadata': { 'type': 'SP.Data.DEX_x005f_ArchiveListItem' },
              'Title': `${src.list}#${origId}`.slice(0, 250),
              'SourceList': src.list,
              'EventId': eid,
              'EventTitle': title,
              'OriginalId': origId,
              'ArchivedAt': new Date().toISOString(),
              'Payload': buildPayload(r),
            });
            if (ins.ok && origId > 0) {
              // Nur löschen, wenn der Archiv-Insert geklappt hat (kein Datenverlust).
              const del = await this._delete(`${this.siteUrl}/_api/web/lists/getbytitle('${src.list}')/items(${origId})`);
              if (del.ok) { listArchived++; result.archived++; } else { result.failed++; }
            } else {
              result.failed++;
            }
          } catch { result.failed++; }
          if (onProgress) onProgress(si, sources.length, src.list, i + 1, matching.length);
        }
      } catch { /* Liste nicht vorhanden — überspringen */ }
      result.perList[src.list] = listArchived;
      if (result.cancelled) break;
    }
    return result;
  }

  // ==================== Archiv-Löschkonzept (v23.40) ====================
  // DEX_Archive-Einträge sind die End-Ablage. Damit die Liste nicht unendlich
  // wächst, können Admins Einträge löschen, die älter als ein Stichdatum sind
  // (v23.48: standardmäßig 1 Monat nach Ablauf). „ArchivedAt" ist der Ablage-Zeitpunkt.

  /** Zählt DEX_Archive-Zeilen mit ArchivedAt älter als `olderThanIso`. */
  public async countDeletableArchiveRows(olderThanIso: string): Promise<number> {
    try {
      const cutoff = new Date(olderThanIso).getTime();
      if (!isFinite(cutoff)) return 0;
      const rows = await this.loadAllListRows('DEX_Archive', 'Id,ArchivedAt');
      let c = 0;
      for (const r of rows) {
        const a = r['ArchivedAt'] ? new Date(String(r['ArchivedAt'])).getTime() : 0;
        if (a > 0 && a < cutoff) c++;
      }
      return c;
    } catch { return 0; }
  }

  /** Löscht DEX_Archive-Zeilen älter als `olderThanIso` (sequentiell, mit
   *  Fortschritt + Abbruch). Hartes DELETE — bewusst (das Archiv ist die
   *  letzte Stufe; ältere Einträge braucht niemand mehr). */
  public async deleteOldArchiveRows(
    olderThanIso: string,
    onProgress?: (done: number, total: number) => void,
    shouldCancel?: () => boolean
  ): Promise<{ deleted: number; failed: number; cancelled: boolean }> {
    const result = { deleted: 0, failed: 0, cancelled: false };
    try {
      const cutoff = new Date(olderThanIso).getTime();
      if (!isFinite(cutoff)) return result;
      const rows = await this.loadAllListRows('DEX_Archive', 'Id,ArchivedAt');
      const targets = rows.filter(r => {
        const a = r['ArchivedAt'] ? new Date(String(r['ArchivedAt'])).getTime() : 0;
        return a > 0 && a < cutoff;
      });
      if (onProgress) onProgress(0, targets.length);
      for (let i = 0; i < targets.length; i++) {
        if (shouldCancel && shouldCancel()) { result.cancelled = true; break; }
        const id = Number(targets[i]['Id'] || 0);
        if (id > 0) {
          try {
            const del = await this._delete(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_Archive')/items(${id})`);
            if (del.ok) result.deleted++; else result.failed++;
          } catch { result.failed++; }
        }
        if (onProgress) onProgress(i + 1, targets.length);
      }
    } catch { /* Liste evtl. nicht vorhanden */ }
    return result;
  }

  private async _delete(url: string): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE',
      },
    };
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }
}
