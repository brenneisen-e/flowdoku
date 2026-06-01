/**
 * E-Mail HTML Templates im Deloitte-Design
 *
 * Basiert auf der offiziellen Deloitte E-Mail-Vorlage
 * (Deloitte_DCGmbH_Email_with_Tagline.html).
 *
 * Generiert den kompletten HTML-Body fuer Power Automate.
 */

import { SPHttpClient } from '@microsoft/sp-http';

const GREEN = '#86bc25';
const SITE_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform';
const APP_URL = `${SITE_URL}/SitePages/DEX.aspx?env=WebView`;
// Gecachtes Logo Base64 aus DEX_EmailTemplates (_Config)
// ORB/Event-Bild wird NICHT gecacht - der Flow setzt das event-spezifische Bild ein
let cachedLogoBase64 = '';
let cachedOrbBase64 = '';

export function getCachedOrbBase64(): string { return cachedOrbBase64; }
export function getCachedLogoBase64(): string { return cachedLogoBase64; }

function getDate(): string {
  return new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

/**
 * v18.22: Erweiterte Formatierung für Überschrift (h1) + Unter-Überschrift
 * (h2). Alle Felder optional — fehlende Werte fallen auf die bisherigen
 * Defaults zurück (Überschrift: 400/normal/headingColor; Unter-Überschrift:
 * 20px/700/normal/#000000), damit alle Bestands-Aufrufe von wrapTemplate()
 * unverändert aussehen.
 */
export interface WrapHeadingOpts {
  headingBold?: boolean;
  headingItalic?: boolean;
  subheadingColor?: string;
  subheadingFontSize?: string;
  subheadingBold?: boolean;
  subheadingItalic?: boolean;
}

/** Baut die h1/h2-Zeilen (Überschrift + Unter-Überschrift) inkl. optionaler
 *  Fett-/Kursiv-/Farb-/Größen-Formatierung. */
function buildHeadingsHtml(headingColor: string, heading: string, subheading: string, hSize: string, opts?: WrapHeadingOpts): string {
  const hWeight = opts?.headingBold ? 700 : 400;
  const hStyle = opts?.headingItalic ? 'italic' : 'normal';
  const subColor = (opts?.subheadingColor && opts.subheadingColor.trim()) || '#000000';
  const subSize = (opts?.subheadingFontSize && opts.subheadingFontSize.trim()) || '20px';
  // Default-Gewicht der Unter-Überschrift ist 700 (fett) — nur ein explizites
  // false macht sie normal.
  const subWeight = opts?.subheadingBold === false ? 400 : 700;
  const subStyle = opts?.subheadingItalic ? 'italic' : 'normal';
  return `  <h1 style="font-family:Aptos,Arial,Helvetica,sans-serif;font-size:${hSize};font-weight:${hWeight};font-style:${hStyle};color:${headingColor};margin:0 0 6px;">${heading}</h1>
  <h2 style="font-family:Aptos,Arial,Helvetica,sans-serif;font-size:${subSize};font-weight:${subWeight};font-style:${subStyle};color:${subColor};margin:0 0 24px;">${subheading}</h2>`;
}

/**
 * Wrapped Deloitte-Template-HTML fuer die SPEICHERUNG in DEX_EmailTemplates
 * erzeugen. Im Unterschied zu wrapTemplate() wird KEIN dynamisches Datum
 * eingebettet - der gespeicherte Template-HTML bleibt stabil, sodass
 * upgradeStandardEmailTemplates() die SharePoint-Eintraege nicht bei jedem
 * App-Start unnoetig patcht.
 *
 * Wird fuer Templates genutzt, die von Power Automate direkt aus dem
 * BodyHtml-Feld versendet werden (z.B. OutlookDeclineReminder), weil dort
 * keine SPFx-seitige wrapTemplate()-Wrapper-Logik greift.
 */
export function wrapTemplateForStorage(headingColor: string, heading: string, subheading: string, bodyHtml: string, headingFontSize?: string, opts?: WrapHeadingOpts): string {
  const hSize = (headingFontSize && headingFontSize.trim()) || '26px';
  return `<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${heading}</title>
</head>
<body style="margin:0;padding:0;background-color:#f5f5f5;font-family:Aptos,'Open Sans',Arial,Helvetica,sans-serif;color:#333333;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f5f5f5;">
<tr><td align="center" style="padding:20px 10px 20px 10px;">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="background-color:#ffffff;max-width:600px;width:100%;">
<tr>
<td style="background-color:#000000;padding:20px 30px 20px 30px;border-bottom:2px solid ${GREEN};">
  <img src="{{LOGO_URL}}" alt="Deloitte." width="180" style="display:block;max-width:180px;height:auto;" />
</td>
</tr>
<tr>
<td style="padding:10px 30px 10px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:11px;color:#999999;">
  Deutschland | Event Experience Platform
</td>
</tr>
<tr>
<td style="background-color:#ffffff;text-align:center;padding:30px 30px 30px 30px;">
  <img src="{{ORB_URL}}" alt="DEX Event Experience Platform" width="180" style="display:inline-block;max-width:180px;height:auto;" />
</td>
</tr>
<tr>
<td style="background-color:${GREEN};height:4px;font-size:0;line-height:0;">&nbsp;</td>
</tr>
<tr>
<td style="padding:30px 30px 10px 30px;">
${buildHeadingsHtml(headingColor, heading, subheading, hSize, opts)}
</td>
</tr>
<tr>
<td style="padding:0 30px 30px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#333333;">
  ${bodyHtml}
</td>
</tr>
<tr>
<td style="padding:0 30px 0 30px;font-size:0;line-height:0;">
  <div style="border-top:1px solid #e8e8e8;font-size:0;line-height:0;height:1px;">&nbsp;</div>
</td>
</tr>
<tr>
<td style="padding:20px 30px 20px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:10px;color:#999999;line-height:1.5;">
  <p style="margin:0 0 8px;"><strong>Deloitte GmbH Wirtschaftspr&uuml;fungsgesellschaft</strong><br>
  Rosenheimer Platz 4<br>
  81669 M&uuml;nchen<br>
  Deutschland</p>
  <p style="margin:0;">Deloitte bezieht sich auf Deloitte Touche Tohmatsu Limited (&bdquo;DTTL&ldquo;), ihr weltweites Netzwerk von Mitgliedsunternehmen und ihre verbundenen Unternehmen (zusammen die &bdquo;Deloitte-Organisation&ldquo;). DTTL (auch &bdquo;Deloitte Global&ldquo; genannt) und jedes ihrer Mitgliedsunternehmen sowie ihre verbundenen Unternehmen sind rechtlich selbstst&auml;ndige und unabh&auml;ngige Unternehmen, die sich gegen&uuml;ber Dritten nicht gegenseitig verpflichten oder binden k&ouml;nnen. DTTL, jedes DTTL-Mitgliedsunternehmen und verbundene Unternehmen haften nur f&uuml;r ihre eigenen Handlungen und Unterlassungen und nicht f&uuml;r die der anderen. DTTL erbringt selbst keine Leistungen gegen&uuml;ber Mandanten. Weitere Informationen finden Sie unter <a href="http://www.deloitte.com/de/UeberUns" style="color:${GREEN};">www.deloitte.com/de/UeberUns</a>.</p>
</td>
</tr>
</table>
</td></tr>
</table>
</body>
</html>`;
}

export function wrapTemplate(headingColor: string, heading: string, subheading: string, bodyHtml: string, headingFontSize?: string, opts?: WrapHeadingOpts): string {
  // v18.19: optionale Überschrift-Größe (z.B. '32px'); Default 26px.
  // v18.22: opts = optionale Fett/Kursiv/Farb-/Größen-Formatierung der Headings.
  const hSize = (headingFontSize && headingFontSize.trim()) || '26px';
  return `<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${heading}</title>
</head>
<body style="margin:0;padding:0;background-color:#f5f5f5;font-family:Aptos,'Open Sans',Arial,Helvetica,sans-serif;color:#333333;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f5f5f5;">
<tr><td align="center" style="padding:20px 10px 20px 10px;">

<!-- Main Container -->
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="background-color:#ffffff;max-width:600px;width:100%;">

<!-- ===== HEADER: Deloitte Logo ===== -->
<tr>
<td style="background-color:#000000;padding:20px 30px 20px 30px;border-bottom:2px solid ${GREEN};">
  <img src="${cachedLogoBase64 || '{{LOGO_URL}}'}" alt="Deloitte." width="180" style="display:block;max-width:180px;height:auto;" />
</td>
</tr>

<!-- ===== SUBHEADER ===== -->
<tr>
<td style="padding:10px 30px 10px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:11px;color:#999999;">
  Deutschland | Event Experience Platform | ${getDate()}
</td>
</tr>

<!-- ===== HERO: DEX Orb ===== -->
<tr>
<td style="background-color:#ffffff;text-align:center;padding:30px 30px 30px 30px;">
  <img src="{{ORB_URL}}" alt="DEX Event Experience Platform" width="180" style="display:inline-block;max-width:180px;height:auto;" />
</td>
</tr>

<!-- ===== GREEN LINE ===== -->
<tr>
<td style="background-color:${GREEN};height:4px;font-size:0;line-height:0;">&nbsp;</td>
</tr>

<!-- ===== CONTENT ===== -->
<tr>
<td style="padding:30px 30px 10px 30px;">
${buildHeadingsHtml(headingColor, heading, subheading, hSize, opts)}
</td>
</tr>
<tr>
<td style="padding:0 30px 30px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#333333;">
  ${bodyHtml}
</td>
</tr>

<!-- ===== "Made with DEX App" ===== -->
<tr>
<td style="padding:0 30px 24px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:12px;color:#999999;">
  <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="color:${GREEN};text-decoration:none;">Made with DEX App</a>
</td>
</tr>

<!-- ===== FOOTER: Separator-Line als eigene Row (Word-HTML Outlook-Kalender Fix) ===== -->
<tr>
<td style="padding:0 30px 0 30px;font-size:0;line-height:0;">
  <div style="border-top:1px solid #e8e8e8;font-size:0;line-height:0;height:1px;">&nbsp;</div>
</td>
</tr>

<!-- ===== FOOTER: Legal ===== -->
<tr>
<td style="padding:20px 30px 20px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:10px;color:#999999;line-height:1.5;">
  <p style="margin:0 0 8px;"><strong>Deloitte GmbH Wirtschaftspr&uuml;fungsgesellschaft</strong><br>
  Rosenheimer Platz 4<br>
  81669 M&uuml;nchen<br>
  Deutschland</p>
  <p style="margin:0;">Deloitte bezieht sich auf Deloitte Touche Tohmatsu Limited (&bdquo;DTTL&ldquo;), ihr weltweites Netzwerk von Mitgliedsunternehmen und ihre verbundenen Unternehmen (zusammen die &bdquo;Deloitte-Organisation&ldquo;). DTTL (auch &bdquo;Deloitte Global&ldquo; genannt) und jedes ihrer Mitgliedsunternehmen sowie ihre verbundenen Unternehmen sind rechtlich selbstst&auml;ndige und unabh&auml;ngige Unternehmen, die sich gegen&uuml;ber Dritten nicht gegenseitig verpflichten oder binden k&ouml;nnen. DTTL, jedes DTTL-Mitgliedsunternehmen und verbundene Unternehmen haften nur f&uuml;r ihre eigenen Handlungen und Unterlassungen und nicht f&uuml;r die der anderen. DTTL erbringt selbst keine Leistungen gegen&uuml;ber Mandanten. Weitere Informationen finden Sie unter <a href="http://www.deloitte.com/de/UeberUns" style="color:${GREEN};">www.deloitte.com/de/UeberUns</a>.</p>
</td>
</tr>

</table>
<!-- End Main Container -->

</td></tr>
</table>
</body>
</html>`;
}

// ==================== Platzhalter-Ersetzung ====================

/**
 * HTML-Sonderzeichen escapen um XSS zu verhindern.
 */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Platzhalter in HTML-Body ersetzen (mit HTML-Escaping gegen XSS).
 */
export function replacePlaceholders(text: string, vars: Record<string, string>): string {
  let result = text;
  for (const [key, value] of Object.entries(vars)) {
    result = result.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), escapeHtml(value));
  }
  // Logo Base64 aus DEX_EmailTemplates Cache ersetzen
  // ORB_URL bleibt als Platzhalter - der Flow ersetzt ihn mit dem event-spezifischen Bild
  if (cachedLogoBase64) result = result.replace(/\{\{LOGO_URL\}\}/g, cachedLogoBase64);
  return result;
}

/**
 * Platzhalter in Plain-Text-Feldern (Subject, Heading) ersetzen - OHNE HTML-Escaping,
 * damit z.B. "SR&T" im Subject nicht zu "SR&amp;T" wird.
 */
export function replacePlaceholdersPlain(text: string, vars: Record<string, string>): string {
  let result = text;
  for (const [key, value] of Object.entries(vars)) {
    result = result.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), value);
  }
  return result;
}

/**
 * Email aus SharePoint-Template generieren.
 * Nutzt wrapTemplate fuer das Deloitte-Design.
 */
export function buildEmailFromTemplate(
  template: {
    subject: string; headingColor: string; heading: string; subheading?: string; bodyHtml: string; headingFontSize?: string;
    // v18.22: erweiterte Heading-Formatierung (optional).
    headingBold?: boolean; headingItalic?: boolean;
    subheadingColor?: string; subheadingFontSize?: string; subheadingBold?: boolean; subheadingItalic?: boolean;
  },
  vars: Record<string, string>
): { subject: string; body: string } {
  // Subject + Heading: plain text (kein HTML-Escaping, sonst wird "&" zu "&amp;")
  const subject = replacePlaceholdersPlain(template.subject, vars);
  const heading = replacePlaceholdersPlain(template.heading, vars);
  // v15.17: Subheading aus dem Template — Default ist {{EventTitle}}
  // (kein hartes „Event "-Präfix mehr). Wenn das Template keine
  // Subheading hat, fallen wir auf den reinen EventTitle zurück.
  const rawSub = (template.subheading && template.subheading.trim()) || '{{EventTitle}}';
  const subheading = replacePlaceholdersPlain(rawSub, vars);
  // Body: HTML, daher Werte escapen
  const bodyHtml = replacePlaceholders(template.bodyHtml, vars);
  // Pre-wrapped templates (z.B. OutlookDeclineReminder, Nachruecken) enthalten
  // bereits das komplette Deloitte-Design via wrapTemplateForStorage(). Nicht
  // doppelt wrappen — sonst zwei Header/Footer im finalen HTML.
  const trimmed = (template.bodyHtml || '').trimLeft();
  const isPreWrapped = /^<!DOCTYPE|^<html/i.test(trimmed);
  return {
    subject,
    body: isPreWrapped
      ? bodyHtml
      : wrapTemplate(template.headingColor, heading, subheading, bodyHtml, template.headingFontSize, {
        headingBold: template.headingBold,
        headingItalic: template.headingItalic,
        subheadingColor: template.subheadingColor,
        subheadingFontSize: template.subheadingFontSize,
        subheadingBold: template.subheadingBold,
        subheadingItalic: template.subheadingItalic,
      }),
  };
}

// ==================== E-Mail Templates (Fallback) ====================

/**
 * Anmeldebestätigung
 */
/**
 * v11.87: Team-Info-Block für Bestätigungs-Mails von Team-Anmeldungen.
 *
 * Wird in allen 4 Add-Pfaden injiziert (registerTeam, addTeamMember,
 * joinTeam-Result, decideTeamJoinRequest-Approve). Der Aufrufer lädt
 * vorher die aktiven Team-Mitglieder via `getTeamMembers` und übergibt
 * Lead + Member-Liste, plus den konfigurierten TeamSize-Maximalwert.
 *
 * Output: kleines HTML-Karten-Snippet zum Einsetzen direkt nach <body>
 * via .replace(/<body[^>]*>/, '<body$1>' + html). Format folgt dem
 * existierenden grünen Hinweisbox-Stil.
 */
export function teamInfoBlockHtml(opts: {
  teamName?: string;
  members: Array<{ firstName: string; lastName: string; isLead: boolean }>;
  teamSize: number;
  isDe: boolean;
  registeredByName?: string;
  consentRequired?: boolean;
}): string {
  const { teamName, members, teamSize, isDe, registeredByName, consentRequired } = opts;
  const activeCount = members.length;
  const memberLines = members.map(m => {
    const full = `${m.firstName} ${m.lastName}`.trim();
    return m.isLead
      ? `<li style="margin:4px 0;"><strong>${full}</strong> ${isDe ? '(Team-Lead)' : '(team lead)'}</li>`
      : `<li style="margin:4px 0;">${full}</li>`;
  }).join('');
  const heading = isDe ? 'Team-Info' : 'Team info';
  const teamLine = teamName
    ? (isDe
      ? `Du bist Teil des Teams <strong>„${teamName}"</strong> mit folgenden Mitgliedern:`
      : `You are part of team <strong>"${teamName}"</strong> with the following members:`)
    : (isDe
      ? `Du bist Teil eines Teams mit folgenden Mitgliedern:`
      : `You are part of a team with the following members:`);
  const occupancy = isDe
    ? `Gesamt-Belegung: <strong>${activeCount} / ${teamSize}</strong>`
    : `Total occupancy: <strong>${activeCount} / ${teamSize}</strong>`;
  const cancelHint = isDe
    ? `Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über <strong>„Meine Events"</strong> ab. Dein Team-Lead und die anderen Mitglieder werden über deine Abmeldung informiert.`
    : `If you cannot attend, please cancel in time via <strong>"My Events"</strong>. Your team lead and the other members will be notified.`;
  const consentHint = consentRequired && registeredByName
    ? (isDe
      ? `<p style="margin:6px 0 0;"><em>Diese Anmeldung wurde von <strong>${registeredByName}</strong> für dich durchgeführt. Falls du NICHT zugestimmt hast, kannst du dich über „Meine Events" eigenständig abmelden.</em></p>`
      : `<p style="margin:6px 0 0;"><em>This registration was performed by <strong>${registeredByName}</strong> on your behalf. If you did not consent, you can cancel yourself via "My Events".</em></p>`)
    : '';
  return `<div style="margin:0 0 16px;padding:14px 18px;background:#f3f8ec;border:1px solid #86bc25;border-radius:8px;font-size:13px;line-height:1.55;color:#3f5f10;">`
    + `<div style="font-weight:700;font-size:14px;color:#3f5f10;margin-bottom:8px;">${heading}</div>`
    + `<p style="margin:0 0 6px;">${teamLine}</p>`
    + `<ul style="margin:0 0 10px 18px;padding:0;list-style:disc;">${memberLines}</ul>`
    + `<p style="margin:0 0 6px;">${occupancy}</p>`
    + `<p style="margin:6px 0 0;">${cancelHint}</p>`
    + consentHint
    + `</div>`;
}

export function registrationEmail(recipientName: string, eventTitle: string): { subject: string; body: string } {
  return {
    subject: `Anmeldebest\u00E4tigung: ${eventTitle}`,
    body: wrapTemplate(
      GREEN,
      'Registration successful',
      eventTitle,
      `<p>Dear ${recipientName},</p>
      <p>you have successfully registered for the event <strong>${eventTitle}</strong>.</p>
      <p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">Event Experience Platform</a> (&bdquo;My Events&ldquo;).</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Warteliste-Bestätigung
 */
export function waitlistEmail(recipientName: string, eventTitle: string, position?: number): { subject: string; body: string } {
  const posInfo = position ? `<p>Your current position on the waitlist: <strong>#${position}</strong>.</p>` : '';
  return {
    subject: `Warteliste: ${eventTitle}`,
    body: wrapTemplate(
      '#ed8b00',
      'Waitlist confirmation',
      eventTitle,
      `<p>Dear ${recipientName},</p>
      <p>you have been placed on the <strong>waitlist</strong> for the event <strong>${eventTitle}</strong>.</p>
      ${posInfo}
      <p>We will notify you as soon as a spot becomes available. You can always check your current waitlist position in the <a href="${APP_URL}" style="color:#86bc25;font-weight:600;">Event Experience Platform</a> under &bdquo;My Events&ldquo;.</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Abmeldebestätigung
 */
export function cancellationEmail(recipientName: string, eventTitle: string): { subject: string; body: string } {
  // v17.20: Visuell deutlicher Stornierungs-Banner direkt unter der Begruessung
  // \u2014 Event-Titel ausgegraut + durchgestrichen, damit auf den ersten Blick
  // erkennbar ist, dass die Anmeldung storniert wurde (vorher: Stornierung
  // ging nur ueber den Subject-Zusatz hervor).
  const cancelBanner = `
    <div style="margin: 16px 0 20px; padding: 14px 18px; border: 2px solid #da291c;
                background: rgba(218, 41, 28, 0.06); border-radius: 8px;
                text-align: center;">
      <div style="font-size: 0.78rem; font-weight: 700; color: #da291c;
                  text-transform: uppercase; letter-spacing: 1.5px;">
        Stornierung &middot; Cancellation
      </div>
      <div style="margin-top: 6px; font-size: 1.15rem; font-weight: 700;
                  color: #888; text-decoration: line-through;">
        ${eventTitle}
      </div>
    </div>`;
  return {
    subject: `Abmeldebest\u00E4tigung: ${eventTitle}`,
    body: wrapTemplate(
      '#da291c',
      'Cancellation confirmed',
      eventTitle,
      `<p>Dear ${recipientName},</p>
      ${cancelBanner}
      <p>your registration for the event above has been <strong>cancelled</strong>. The Outlook calendar entry will be removed from your calendar shortly.</p>
      <p>If you change your mind, you can register again via the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">Event Experience Platform</a>.</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Nachrücken von der Warteliste
 */
export function promotionEmail(recipientName: string, eventTitle: string): { subject: string; body: string } {
  return {
    subject: `Platz frei: ${eventTitle}`,
    body: wrapTemplate(
      GREEN,
      'You got a spot!',
      eventTitle,
      `<p>Dear ${recipientName},</p>
      <p>Great news! A spot has become available and you have been <strong>moved from the waitlist to a confirmed participant</strong> for the event <strong>${eventTitle}</strong>.</p>
      <p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">link</a> (&bdquo;My Events&ldquo;).</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Event erstellt - Benachrichtigung an Organisator
 */
export function eventCreatedEmail(recipientName: string, eventTitle: string, subsiteUrl: string): { subject: string; body: string } {
  return {
    subject: `[Deloitte Eventmanager] - New event created`,
    body: wrapTemplate(
      GREEN,
      'Event Created',
      eventTitle,
      `<p>Dear ${recipientName},</p>
      <p>your event <strong>${eventTitle}</strong> has been successfully created.</p>
      <p>You can find the list of participants here:</p>
      <ul>
        <li><a href="${subsiteUrl}/Lists/Teilnehmer/AllItems.aspx" style="color:${GREEN};font-weight:600;">SharePoint Teilnehmerliste</a></li>
        <li><a href="${APP_URL}" style="color:${GREEN};font-weight:600;">Event Experience Platform</a> (Admin / Organizer)</li>
      </ul>
      <p>Regards,<br>Team Event Experience Platform</p>`
    ),
  };
}

/**
 * Onboarding-Mail an einen frisch ernannten Organizer (oder Admin).
 * Wird vom Admin im Settings-Bildschirm nach dem Anlegen einer neuen Rolle
 * optional ausgelöst. Enthält Link zur App, Link zum Handbuch und kurze
 * Bullet-Points zum Anlegen eines ersten Test-Events.
 *
 * Deloitte-displayName ist "Nachname, Vorname" — fuer die Anrede nur den
 * Vornamen verwenden (analog qrCodeEmail / registrationEmail).
 */
export function organizerOnboardingEmail(recipientName: string, role: 'Organizer' | 'Admin' = 'Organizer'): { subject: string; body: string } {
  const manualUrl = `${APP_URL}&action=manual`;
  const roleLabelDe = role === 'Admin' ? 'Admin' : 'Organizer';
  // Anrede: Vorname extrahieren. "Nachname, Vorname" -> Teil nach Komma,
  // sonst erstes Wort. Fallback: kompletter Name.
  const firstName = (() => {
    const n = (recipientName || '').trim();
    if (!n) return '';
    const c = n.indexOf(',');
    if (c >= 0) return n.substring(c + 1).trim().split(/\s+/)[0];
    return n.split(/\s+/)[0];
  })();
  return {
    subject: `Willkommen als ${roleLabelDe} auf der Deloitte Event Experience Platform`,
    body: wrapTemplate(
      GREEN,
      'Willkommen an Bord',
      `Dein Start als ${roleLabelDe}`,
      `<p>Hallo ${firstName},</p>
      <p>du wurdest soeben als <strong>${roleLabelDe}</strong> für die Deloitte
      <strong>Event Experience Platform</strong> freigeschaltet. Damit kannst du
      eigene Events anlegen, Teilnehmer verwalten und Einladungen versenden.</p>

      <p style="margin-top:24px;"><strong>Deine wichtigsten Links:</strong></p>
      <ul>
        <li><a href="${APP_URL}" style="color:${GREEN};font-weight:600;">Zur Event Experience Platform</a> &mdash; hier legst du Events an und verwaltest deine Teilnehmer.</li>
        <li><a href="${manualUrl}" style="color:${GREEN};font-weight:600;">Zum Handbuch</a> &mdash; Schritt-für-Schritt-Anleitung mit Screenshots zu allen Features.</li>
      </ul>

      <p style="margin-top:24px;"><strong>So legst du dein erstes Test-Event an:</strong></p>
      <ul>
        <li>Öffne die App über den Link oben und klicke auf <strong>&bdquo;Event erstellen&ldquo;</strong>. Fülle Titel, Datum, Ort und Beschreibung aus &mdash; das reicht für einen ersten Probelauf.</li>
        <li>Lade dich selbst (oder eine Testperson) auf der <strong>Registrierungsseite</strong> ein und prüfe im <strong>Admin Center</strong>, ob die Anmeldung sauber durchläuft und die Bestätigungsmail rauskommt.</li>
        <li>Schau dir das <strong>Handbuch</strong> an, wenn du Custom-Felder, Wartelisten, Outlook-Termine oder den Massenmail-Versand ausprobieren möchtest &mdash; dort sind alle Funktionen mit Praxisbeispielen erklärt.</li>
      </ul>

      <p style="margin-top:24px;">Bei Fragen oder Problemen melde dich gerne direkt bei
      <a href="mailto:ebrenneisen@deloitte.de" style="color:${GREEN};">ebrenneisen@deloitte.de</a>
      oder
      <a href="mailto:nifelten@deloitte.de" style="color:${GREEN};">nifelten@deloitte.de</a>
      &mdash; wir helfen dir gerne weiter.</p>

      <p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein DEX-Team</strong></p>`
    ),
  };
}

/**
 * QR-Code E-Mail fuer Check-in.
 * Subject + Body folgen der Event-Sprache (DE/EN). Anrede nutzt nur den
 * Vornamen (nicht den vollen Namen).
 *
 * v9.15: Unterhalb des QR-Codes wird zusaetzlich "<voller Name> | <Event-Titel>"
 * fett angezeigt \u2014 hilft den Organizern beim manuellen Check-in (Foto- oder
 * Bildschirm-Vergleich), wenn der Scanner mal nicht zur Hand ist.
 */
export function qrCodeEmail(
  firstName: string,
  eventTitle: string,
  qrImageHtml: string,
  lang: string = 'EN',
  fullName?: string
): { subject: string; body: string } {
  const isDe = (lang || 'EN').toUpperCase() === 'DE';
  // Fallback: wenn kein fullName uebergeben, nutze nur firstName
  const fullDisplayName = (fullName || firstName || '').trim();
  // HTML-Escape \u2014 Nutzer-Eingaben (Namen) duerfen das Layout nicht brechen
  const escName = fullDisplayName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escTitle = (eventTitle || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const checkInLabel = `<p style="text-align:center;font-size:16px;margin:8px 0 0;"><strong>${escName} | ${escTitle}</strong></p>`;
  if (isDe) {
    return {
      subject: `Dein QR-Code f\u00FCr ${eventTitle}`,
      body: wrapTemplate(
        GREEN,
        'Dein QR-Code',
        eventTitle,
        `<p>Hallo ${firstName},</p>
        <p>hier ist dein pers\u00F6nlicher QR-Code f\u00FCr das Event <strong>${eventTitle}</strong>.</p>
        <p>Bitte zeige den QR-Code beim Check-in vor.</p>
        <div style="text-align:center;margin:24px 0;">${qrImageHtml}${checkInLabel}</div>
        <p style="color:#999;font-size:12px;text-align:center;">Der QR-Code ist pers\u00F6nlich und nicht \u00FCbertragbar.</p>
        <p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
      ),
    };
  }
  return {
    subject: `Your QR Code for ${eventTitle}`,
    body: wrapTemplate(
      GREEN,
      'Your QR Code',
      eventTitle,
      `<p>Dear ${firstName},</p>
      <p>here is your personal QR code for the event <strong>${eventTitle}</strong>.</p>
      <p>Please show this QR code at check-in.</p>
      <div style="text-align:center;margin:24px 0;">${qrImageHtml}${checkInLabel}</div>
      <p style="color:#999;font-size:12px;text-align:center;">This QR code is personal and non-transferable.</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Allgemeine Info-Mail
 */
export function infoEmail(recipientName: string, eventTitle: string, message: string): { subject: string; body: string } {
  return {
    subject: `Info: ${eventTitle}`,
    body: wrapTemplate(
      '#333333',
      'Information',
      eventTitle,
      `<p>Dear ${recipientName},</p>
      <p>${message}</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Outlook-Termin-Body im Deloitte-Design generieren.
 *
 * Idempotent: wenn `bodyText` bereits gewickelte HTML ist (enthält `<!DOCTYPE`
 * oder `<html`), wird der innere Body extrahiert und neu gewickelt — verhindert
 * double-wrapping bei Re-Save bestehender Events. Wenn `bodyText` plain HTML mit
 * Tags ist, wird er als HTML behandelt; wenn es Plain-Text ist, wird jede Zeile
 * in `<p>`-Tags gewickelt + escaped.
 */
export function buildOutlookBody(eventTitle: string, bodyText: string, subheading?: string): string {
  const inner = stripOutlookWrapper(bodyText || '');
  const isHtml = /<[a-z][\s\S]*>/i.test(inner);
  const bodyHtml = inner
    ? (isHtml ? inner : inner.split('\n').map(line => `<p>${escapeHtml(line)}</p>`).join('\n  '))
    : '';
  return wrapTemplate(GREEN, eventTitle, subheading || 'Event Details', bodyHtml);
}

/**
 * Extrahiert den User-Content-Teil aus einem bereits-gewickelten Outlook-Body.
 * Wenn der Input nicht gewickelt ist, wird er unveraendert zurueckgegeben.
 * Dadurch koennen wir bestehende Events bearbeiten, ohne dass der Editor das
 * komplette Wrapper-HTML als rohen Text anzeigt.
 */
export function stripOutlookWrapper(html: string): string {
  if (!html) return '';
  if (!/<!doctype|<html/i.test(html)) return html;
  // wrapTemplate fuegt den Body in <td style="padding:0 30px 30px[ 30px];...color:#333333;">CONTENT</td>
  // Non-greedy auf </td> reicht — wir brauchen keine trailing-Constraint (HTML-Kommentare
  // zwischen </tr> und naechstem <tr> wuerden eine engere Pruefung sonst sprengen, wodurch
  // der Strip fehlschlaegt und beim Re-Save der gesamte Wrapper erneut umwickelt wird).
  // Seit v11.56 nutzt wrapTemplate 4-value-Paddings ("0 30px 30px 30px"), aber bestehende
  // gespeicherte Bodies haben noch 3-value ("0 30px 30px") — die Regex muss beide matchen.
  const m = html.match(/<td style="padding:0 30px 30px(?:\s+30px)?;[^"]*">([\s\S]*?)<\/td>/i);
  if (m && m[1]) return m[1].trim();
  return html;
}

/**
 * Extrahiert Heading (<h1>) und Subheading (<h2>) aus einer bereits gewickelten
 * Outlook-Body. Gibt leere Strings zurueck wenn keine vorhanden.
 */
export function parseOutlookHeadings(html: string): { heading: string; subheading: string } {
  if (!html || !/<!doctype|<html/i.test(html)) return { heading: '', subheading: '' };
  const h1 = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  const h2 = html.match(/<h2[^>]*>([\s\S]*?)<\/h2>/i);
  return {
    heading: h1 && h1[1] ? h1[1].replace(/<[^>]+>/g, '').trim() : '',
    subheading: h2 && h2[1] ? h2[1].replace(/<[^>]+>/g, '').trim() : '',
  };
}

/**
 * Logo als Base64 in den Memory-Cache laden.
 * Liest zuerst aus DEX_EmailTemplates (_Config), dann Fallback auf SiteAssets.
 * Schreibt NICHT zurueck - das macht EventService.ensureLogosInConfig().
 */
export async function loadLogosAsBase64(spHttpClient: SPHttpClient, siteUrl: string): Promise<void> {
  if (cachedLogoBase64 && cachedOrbBase64) return;

  try {
    // 1. Aus _Config Zeile lesen (falls bereits befuellt)
    const configUrl = `${siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=LogoBase64,DefaultImageBase64`;
    const response = await spHttpClient.get(configUrl, SPHttpClient.configurations.v1);
    if (response.ok) {
      const data = await response.json();
      const items = data.value || data.d?.results || [];
      const it = items[0];
      if (it?.LogoBase64 && it.LogoBase64.startsWith('data:image/')) cachedLogoBase64 = it.LogoBase64;
      if (it?.DefaultImageBase64 && it.DefaultImageBase64.startsWith('data:image/')) cachedOrbBase64 = it.DefaultImageBase64;
      if (cachedLogoBase64 && cachedOrbBase64) return;
    }

    // 2. Fallback: direkt aus SiteAssets laden (nur Memory-Cache)
    if (!cachedLogoBase64) {
      const logo = await loadImageAsBase64(spHttpClient, siteUrl, 'DEX_Logos/Deloitte_Logo.png');
      if (logo) cachedLogoBase64 = logo;
    }
    if (!cachedOrbBase64) {
      const orb = await loadImageAsBase64(spHttpClient, siteUrl, 'DEX_Logos/dex-orb.png');
      if (orb) cachedOrbBase64 = orb;
    }
  } catch { /* Logo nicht verfuegbar - Flow ersetzt Platzhalter als Fallback */ }
}

/**
 * Bild aus SiteAssets laden und als Base64 Data-URI zurueckgeben.
 */
async function loadImageAsBase64(spHttpClient: SPHttpClient, siteUrl: string, path: string): Promise<string> {
  try {
    const fileUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${siteUrl.replace(/^https?:\/\/[^/]+/, '')}/SiteAssets/${path}')/$value`;
    const resp = await spHttpClient.get(fileUrl, SPHttpClient.configurations.v1);
    if (!resp.ok) return '';

    const blob = await resp.blob();
    return await new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string || '');
      reader.onerror = () => resolve('');
      reader.readAsDataURL(blob);
    });
  } catch {
    return '';
  }
}
