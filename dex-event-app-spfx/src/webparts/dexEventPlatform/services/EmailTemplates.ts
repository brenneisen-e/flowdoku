/**
 * E-Mail HTML Templates im Deloitte-Design
 *
 * Basiert auf der offiziellen Deloitte E-Mail-Vorlage
 * (Deloitte_DCGmbH_Email_with_Tagline.html).
 *
 * Generiert den kompletten HTML-Body für Power Automate.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import { buildHashDeepLink } from '../utils/deepLink';

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
  // v18.73: Header-Bild (Hero / Event-Bild = {{ORB_URL}}) frei einstellbar pro
  // Event. `imageWidth` = max. Breite in px (Default 180). `imagePaddingV` =
  // Innenabstand oben/unten in px (Default 30, = Abstand zur Unter-Überschrift
  // bzw. zum grünen Strich). `imagePaddingH` = Innenabstand links/rechts in px
  // (Default 30, = Abstand zum Spaltenrand). 0 = randlos. Fehlende Werte
  // fallen auf die bisherigen Defaults zurück, damit Alt-Aufrufe unverändert
  // aussehen.
  imageWidth?: number;
  imagePaddingV?: number;
  imagePaddingH?: number;
}

/**
 * v18.73: Baut die Hero-Zeile (Event-Bild = {{ORB_URL}}) inkl. einstellbarer
 * Breite + Innenabstand. Gemeinsamer Helper für wrapTemplate() und
 * wrapTemplateForStorage(), damit beide Layouts identisch bleiben.
 *
 * `width:100%` + `max-width` macht das Bild responsiv: es füllt die Spalte bis
 * zur konfigurierten Maximalbreite (bei kleinem Innenabstand also fast randlos)
 * und schrumpft auf schmalen Mobil-Clients sauber mit. Das `width`-Attribut
 * bleibt für Outlook-Desktop (ignoriert max-width) als Fallback erhalten.
 */
function buildHeroRow(opts?: WrapHeadingOpts): string {
  const w = (typeof opts?.imageWidth === 'number' && opts.imageWidth > 0) ? Math.round(opts.imageWidth) : 180;
  const padV = (typeof opts?.imagePaddingV === 'number' && opts.imagePaddingV >= 0) ? Math.round(opts.imagePaddingV) : 30;
  const padH = (typeof opts?.imagePaddingH === 'number' && opts.imagePaddingH >= 0) ? Math.round(opts.imagePaddingH) : 30;
  return `<tr>
<td style="background-color:#ffffff;text-align:center;padding:${padV}px ${padH}px ${padV}px ${padH}px;">
  <img src="{{ORB_URL}}" alt="DEX Event Experience Platform" width="${w}" style="display:inline-block;width:100%;max-width:${w}px;height:auto;" />
</td>
</tr>`;
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
 * Wrapped Deloitte-Template-HTML für die SPEICHERUNG in DEX_EmailTemplates
 * erzeugen. Im Unterschied zu wrapTemplate() wird KEIN dynamisches Datum
 * eingebettet - der gespeicherte Template-HTML bleibt stabil, sodass
 * upgradeStandardEmailTemplates() die SharePoint-Einträge nicht bei jedem
 * App-Start unnötig patcht.
 *
 * Wird für Templates genutzt, die von Power Automate direkt aus dem
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
  Deutschland | DEX App
</td>
</tr>
${buildHeroRow(opts)}
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
  Deutschland | DEX App | ${getDate()}
</td>
</tr>

<!-- ===== HERO: DEX Orb (v18.73: Breite + Innenabstand einstellbar) ===== -->
${buildHeroRow(opts)}

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
 * v22.47: Organizer-Namensliste als HTML — jeder NAME fett, die Verbinder
 * („, " / „ und " / „ and ") bleiben normal. Eingabe ist der bereits
 * zusammengesetzte Klartext aus formatOrganizerList (z.B. „Franziska
 * Hasemeier und Anja Helwich"). Namen werden HTML-escaped, die Verbinder
 * roh übernommen. Robust gegen 0/1/n Namen.
 */
function boldOrganizerNames(plain: string): string {
  const s = (plain || '').trim();
  if (!s) return '';
  // Split MIT Erhalt der Verbinder (capturing group).
  const parts = s.split(/(, | und | and )/);
  return parts
    .map(p => (/^(, | und | and )$/.test(p) ? p : (p ? `<strong>${escapeHtml(p)}</strong>` : '')))
    .join('');
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
 * Nutzt wrapTemplate für das Deloitte-Design.
 */
export function buildEmailFromTemplate(
  template: {
    subject: string; headingColor: string; heading: string; subheading?: string; bodyHtml: string; headingFontSize?: string;
    // v18.22: erweiterte Heading-Formatierung (optional).
    headingBold?: boolean; headingItalic?: boolean;
    subheadingColor?: string; subheadingFontSize?: string; subheadingBold?: boolean; subheadingItalic?: boolean;
    // v18.73: Header-Bild (Breite + Innenabstand) pro Event.
    imageWidth?: number; imagePaddingV?: number; imagePaddingH?: number;
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
  let bodyHtml = replacePlaceholders(template.bodyHtml, vars);
  // v22.47: {{OrganizerHtml}} — nur die NAMEN fett, die Verbinder („und"/", ")
  // bleiben normal. Wird RAW (unescaped) ersetzt, deshalb erst nach
  // replacePlaceholders (das `OrganizerHtml` mangels vars-Key stehen lässt).
  if (bodyHtml.indexOf('{{OrganizerHtml}}') >= 0) {
    bodyHtml = bodyHtml.replace(/\{\{OrganizerHtml\}\}/g, boldOrganizerNames(vars.Organizer || ''));
  }
  // Pre-wrapped templates (z.B. OutlookDeclineReminder, Nachrücken) enthalten
  // bereits das komplette Deloitte-Design via wrapTemplateForStorage(). Nicht
  // doppelt wrappen — sonst zwei Header/Footer im finalen HTML.
  // v22.76: Erkennung robust — NICHT nur am String-Anfang prüfen. Ältere oder
  // im Tenant überschriebene Templates können führende Zeichen/Whitespace/BOM
  // vor dem <!DOCTYPE haben, wodurch die alte ^-Prüfung fehlschlug und die Mail
  // ein zweites Mal gewrappt wurde (zwei Deloitte-Header). Ein bereits
  // gewrappter Body enthält IMMER die HTML-Hülle bzw. den Logo-Platzhalter.
  const rawBodyForCheck = template.bodyHtml || '';
  const lowerBodyForCheck = rawBodyForCheck.toLowerCase();
  const isPreWrapped = lowerBodyForCheck.indexOf('<!doctype') >= 0
    || lowerBodyForCheck.indexOf('<html') >= 0
    || rawBodyForCheck.indexOf('{{LOGO_URL}}') >= 0;
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
        // v18.73: Header-Bild Größe + Innenabstand pro Event.
        imageWidth: template.imageWidth,
        imagePaddingV: template.imagePaddingV,
        imagePaddingH: template.imagePaddingH,
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
      <p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * v26.33: Einladung für externe Dritte (kein Deloitte-Postfach), die von jemand
 * anderem stellvertretend angemeldet wurden. Statt einer Anmeldebestätigung:
 * eine Einladung mit Bitte um Bestätigung per Antwort-Mail + Deloitte-
 * Datenschutzhinweis. Die Mail geht an die externe Person; die anmeldende Person
 * und alle Organizer stehen auf CC (so erreicht die Zusage-Antwort alle).
 */
export function externalInvitationEmail(
  recipientName: string,
  eventTitle: string,
  registeredByName: string,
  isDe: boolean,
  opts?: { startDate?: string; endDate?: string; location?: string }
): { subject: string; body: string } {
  const privacyUrl = 'https://www.deloitte.com/de/de/legal/privacy.html';
  const locale = isDe ? 'de-DE' : 'en-GB';
  // v26.33: „Wann/Wo"-Block aus Start-/Enddatum + Ort (Ort nur falls vorhanden).
  const whenStr = (() => {
    const s = opts?.startDate ? new Date(opts.startDate) : null;
    if (!s || isNaN(s.getTime())) return '';
    const dateStr = s.toLocaleDateString(locale, { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' });
    const t1 = s.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' });
    let time = t1;
    const e = opts?.endDate ? new Date(opts.endDate) : null;
    if (e && !isNaN(e.getTime())) time = `${t1}–${e.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })}`;
    return `${dateStr}, ${time}${isDe ? ' Uhr' : ''}`;
  })();
  const loc = (opts?.location || '').trim();
  const detailsRows = [
    whenStr ? `<tr><td style="padding:2px 14px 2px 0;color:#555;white-space:nowrap;vertical-align:top;">${isDe ? 'Wann' : 'When'}:</td><td style="padding:2px 0;"><strong>${whenStr}</strong></td></tr>` : '',
    loc ? `<tr><td style="padding:2px 14px 2px 0;color:#555;white-space:nowrap;vertical-align:top;">${isDe ? 'Wo' : 'Where'}:</td><td style="padding:2px 0;"><strong>${loc}</strong></td></tr>` : '',
  ].filter(Boolean).join('');
  const detailsBlock = detailsRows
    ? `<table style="margin:10px 0 4px;border-collapse:collapse;font-size:14px;">${detailsRows}</table>`
    : '';
  if (isDe) {
    return {
      subject: `Einladung: ${eventTitle}`,
      body: wrapTemplate(
        GREEN,
        'Einladung zum Event',
        eventTitle,
        `<p>Hallo ${recipientName},</p>
        <p>${registeredByName ? `<strong>${registeredByName}</strong> hat dich` : 'Du wurdest'} zum Event <strong>${eventTitle}</strong> eingeladen.</p>
        ${detailsBlock}
        <p>Wenn du teilnehmen möchtest, <strong>bestätige bitte diese Einladung, indem du auf diese E-Mail antwortest</strong>. Deine Zusage geht damit an die anmeldende Person und die Organisator:innen (in Kopie).</p>
        <p style="margin-top:16px;font-size:13px;color:#555;">Es gelten die <a href="${privacyUrl}" style="color:${GREEN};font-weight:600;">Datenschutzhinweise von Deloitte</a>. Deine Daten werden ausschließlich zur Organisation dieses Events verarbeitet. Einen Widerruf kannst du jederzeit an <a href="mailto:privacy@deloitte.de" style="color:${GREEN};">privacy@deloitte.de</a> richten.</p>
        <p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
      ),
    };
  }
  return {
    subject: `Invitation: ${eventTitle}`,
    body: wrapTemplate(
      GREEN,
      'Event invitation',
      eventTitle,
      `<p>Dear ${recipientName},</p>
      <p>${registeredByName ? `<strong>${registeredByName}</strong> has invited you` : 'You have been invited'} to the event <strong>${eventTitle}</strong>.</p>
      ${detailsBlock}
      <p>If you would like to attend, <strong>please confirm this invitation by replying to this email</strong>. Your reply reaches the person who registered you and the organizers (on copy).</p>
      <p style="margin-top:16px;font-size:13px;color:#555;">Deloitte's <a href="${privacyUrl}" style="color:${GREEN};font-weight:600;">data protection notice</a> applies. Your data is processed solely to organise this event. You may withdraw at any time via <a href="mailto:privacy@deloitte.de" style="color:${GREEN};">privacy@deloitte.de</a>.</p>
      <p style="margin-top:24px;"><strong>Best regards</strong><br><br><strong>Your Event Team</strong></p>`
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
      <p>We will notify you as soon as a spot becomes available. You can always check your current waitlist position in the <a href="${APP_URL}" style="color:#86bc25;font-weight:600;">DEX App</a> under &bdquo;My Events&ldquo;.</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Abmeldebestätigung
 */
export function cancellationEmail(recipientName: string, eventTitle: string): { subject: string; body: string } {
  // v17.20: Visuell deutlicher Stornierungs-Banner direkt unter der Begrüßung
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
      <p>If you change your mind, you can register again via the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">DEX App</a>.</p>
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
    subject: `[Deloitte Eventmanager] - Event angelegt`,
    body: wrapTemplate(
      GREEN,
      'Event angelegt',
      eventTitle,
      `<p>Hallo ${recipientName},</p>
      <p>dein Event <strong>${eventTitle}</strong> wurde erfolgreich angelegt. So machst du es startklar:</p>
      <ol style="line-height:1.6;padding-left:20px;margin:0 0 16px;">
        <li><strong>Event finalisieren</strong> &ndash; &uuml;ber &bdquo;Event bearbeiten&ldquo; Felder, Bild und Texte vervollst&auml;ndigen.</li>
        <li><strong>Test-An- und Abmeldung</strong> &ndash; melde dich einmal selbst an und wieder ab, um zu pr&uuml;fen, ob die automatische Kommunikation (Best&auml;tigungs-Mail, Outlook-Termin, Abmelde-Mail) richtig ankommt.</li>
        <li><strong>Event live schalten</strong> &ndash; in der App oben &uuml;ber das Status-H&auml;kchen &bdquo;Entwurf &rarr; Aktiv&ldquo; schalten. Danach ist es f&uuml;r die berechtigten Gruppen sichtbar.</li>
        <li><strong>Einladung verschicken</strong> &ndash; optional die Einladung mit Anmelde-Link direkt aus der App versenden (an dich zum Weiterleiten oder an den Verteiler).</li>
        <li><strong>Anmeldungen verfolgen</strong> &ndash; sobald sich Teilnehmer anmelden, siehst du im Organizer Center alle Infos: Anzahl, Status und die komplette Teilnehmerliste.</li>
      </ol>
      <p>Teilnehmerliste &amp; Verwaltung:</p>
      <ul>
        <li><a href="${subsiteUrl}/Lists/Teilnehmer/AllItems.aspx" style="color:${GREEN};font-weight:600;">SharePoint-Teilnehmerliste</a></li>
        <li><a href="${APP_URL}" style="color:${GREEN};font-weight:600;">DEX App</a> (Admin / Organizer)</li>
      </ul>
      <p>Viele Gr&uuml;&szlig;e<br>Team DEX App</p>
      <hr style="border:none;border-top:1px solid #e0e0e0;margin:20px 0;">
      <p style="color:#999;font-size:13px;line-height:1.5;">English: Your event <strong>${eventTitle}</strong> was created successfully. Next steps: (1)&nbsp;finalize it via &bdquo;Edit event&ldquo;, (2)&nbsp;do a test registration &amp; cancellation to check the automatic emails and the Outlook invite, (3)&nbsp;publish it via the status toggle &bdquo;Draft &rarr; Active&ldquo;, (4)&nbsp;optionally send the invitation with the registration link from the app, (5)&nbsp;track registrations in the Organizer Center. Participant list &amp; admin: see the links above.</p>`
    ),
  };
}

/**
 * Onboarding-Mail an einen frisch ernannten Organizer (oder Admin).
 * Wird vom Admin im Settings-Bildschirm nach dem Anlegen einer neuen Rolle
 * optional ausgelöst. Enthält Link zur App, Link zum Handbuch und kurze
 * Bullet-Points zum Anlegen eines ersten Test-Events.
 *
 * Deloitte-displayName ist "Nachname, Vorname" — für die Anrede nur den
 * Vornamen verwenden (analog qrCodeEmail / registrationEmail).
 */
export function organizerOnboardingEmail(recipientName: string, role: 'Organizer' | 'Admin' = 'Organizer'): { subject: string; body: string } {
  const manualUrl = buildHashDeepLink(APP_URL, { action: 'manual' });
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
    subject: `Willkommen als ${roleLabelDe} auf der Deloitte DEX App`,
    body: wrapTemplate(
      GREEN,
      'Willkommen an Bord',
      `Dein Start als ${roleLabelDe}`,
      `<p>Hallo ${firstName},</p>
      <p>du wurdest soeben als <strong>${roleLabelDe}</strong> für die Deloitte
      <strong>DEX App</strong> freigeschaltet. Damit kannst du
      eigene Events anlegen, Teilnehmer verwalten und Einladungen versenden.</p>

      <p style="margin-top:24px;"><strong>Deine wichtigsten Links:</strong></p>
      <ul>
        <li><a href="${APP_URL}" style="color:${GREEN};font-weight:600;">Zur DEX App</a> &mdash; hier legst du Events an und verwaltest deine Teilnehmer.</li>
        <li><a href="${manualUrl}" style="color:${GREEN};font-weight:600;">Zum Handbuch</a> &mdash; Schritt-für-Schritt-Anleitung mit Screenshots zu allen Features.</li>
      </ul>

      <p style="margin-top:24px;"><strong>So legst du dein erstes Test-Event an:</strong></p>
      <ul>
        <li>Öffne die App über den Link oben und gehe auf die Kachel <strong>&bdquo;Organizer&ldquo;</strong> &mdash; dort findest du den Button <strong>&bdquo;Event erstellen&ldquo;</strong>. Fülle Titel, Datum, Ort und Beschreibung aus &mdash; das reicht für einen ersten Probelauf.</li>
        <li>Mach einen Test: <strong>Melde dich selbst</strong> (oder eine Testperson) ganz normal über die <strong>Anmeldeseite</strong> des Events an &mdash; in DEX registrieren sich die Teilnehmer immer selbst, es gibt keine automatische Einladung. Prüfe danach im <strong>Organizer Center</strong>, ob die Anmeldung sauber durchläuft und die Bestätigungsmail rauskommt.</li>
        <li>Schau dir das <strong>Handbuch</strong> an, wenn du Custom-Felder, Wartelisten, Outlook-Termine oder den Massenmail-Versand ausprobieren möchtest &mdash; dort sind alle Funktionen mit Praxisbeispielen erklärt.</li>
      </ul>

      <p style="margin-top:24px;"><strong>Du hast Fragen?</strong> Nutze dafür bitte das Ticketsystem direkt in der App: Oben rechts in der Kopfzeile findest du den grünen Button <strong>&bdquo;Hast du Fragen?&ldquo;</strong>. Ein Klick öffnet ein Fenster, in dem du deine Frage(n) stellst &mdash; auf Wunsch mit einem Screenshot deines Bildschirms. Schon beim Tippen schlägt dir die App passende Handbuch-Artikel vor. Deine Frage geht an das DEX-Team, das sich darum kümmert und dir in der App antwortet.</p>

      <p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein DEX-Team</strong></p>`
    ),
  };
}

/**
 * v26.34: „Du bist jetzt Co-Organizer"-Mail. Wird verschickt, wenn eine Person
 * nachträglich (beim Bearbeiten eines Events) als (Co-)Organizer hinzugefügt wird
 * — inkl. Hinweis auf den Zugriff auf die Teilnehmerliste und die separat
 * versendete Outlook-Kalendereinladung. Folgt der Event-Sprache (DE/EN).
 */
export function coOrganizerAddedEmail(
  recipientName: string,
  eventTitle: string,
  actorName: string,
  isDe: boolean,
  appUrl?: string
): { subject: string; body: string } {
  const link = appUrl || APP_URL;
  if (isDe) {
    return {
      subject: `Du bist jetzt Co-Organizer: ${eventTitle}`,
      body: wrapTemplate(GREEN, 'Du bist jetzt Co-Organizer', eventTitle,
        `<p>Hallo ${recipientName},</p>
        <p>${actorName ? `<strong>${actorName}</strong> hat dich` : 'Du wurdest'} als <strong>Co-Organizer</strong> für das Event <strong>${eventTitle}</strong> hinzugefügt.</p>
        <p>Du kannst das Event ab sofort mitverwalten und hast <strong>Zugriff auf die Teilnehmerliste</strong> — öffne dazu das <a href="${link}" style="color:${GREEN};font-weight:600;">Organizer Center der DEX App</a>.</p>
        <p>Außerdem hast du eine <strong>Outlook-Kalendereinladung</strong> zum Event erhalten.</p>
        <p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
      ),
    };
  }
  return {
    subject: `You are now a co-organizer: ${eventTitle}`,
    body: wrapTemplate(GREEN, 'You are now a co-organizer', eventTitle,
      `<p>Dear ${recipientName},</p>
      <p>${actorName ? `<strong>${actorName}</strong> has added you` : 'You have been added'} as a <strong>co-organizer</strong> for the event <strong>${eventTitle}</strong>.</p>
      <p>You can now help manage the event and have <strong>access to the participant list</strong> — open the <a href="${link}" style="color:${GREEN};font-weight:600;">Organizer Center in the DEX App</a>.</p>
      <p>You have also received an <strong>Outlook calendar invitation</strong> for the event.</p>
      <p style="margin-top:24px;"><strong>Best regards</strong><br><br><strong>Your Event Team</strong></p>`
    ),
  };
}

/**
 * QR-Code E-Mail für Check-in.
 * Subject + Body folgen der Event-Sprache (DE/EN). Anrede nutzt nur den
 * Vornamen (nicht den vollen Namen).
 *
 * v9.15: Unterhalb des QR-Codes wird zusätzlich "<voller Name> | <Event-Titel>"
 * fett angezeigt \u2014 hilft den Organizern beim manuellen Check-in (Foto- oder
 * Bildschirm-Vergleich), wenn der Scanner mal nicht zur Hand ist.
 */
// v22.18: Pro-Event-Anpassung der QR-Mail (analog Einladungs-/Massenmail).
// Der QR-Block (Code + Name + Event als Klartext) ist FESTER Bestandteil und
// wird über den Platzhalter {{QR_BLOCK}} in den Body eingesetzt — fehlt der
// Platzhalter im angepassten Text, wird der Block automatisch ans Ende gesetzt
// (die Mail geht nie ohne QR raus). Persistiert wird der Override im Event
// (EmailTemplateOverrides-JSON, Key 'QRCode') — dadurch gilt er auch für den
// QR-Auto-Versand bei neuen Anmeldungen, nicht nur im eigenen Browser.
export interface QrEmailOverride {
  subject?: string;
  heading?: string;
  subheading?: string;
  bodyHtml?: string;
}

/** Standard-Texte der QR-Mail mit Platzhaltern ({{Vorname}}, {{Name}},
 *  {{EventTitle}}, {{QR_BLOCK}}) — Grundlage für den Editor im Organizer
 *  Center UND für den Versand ohne Override. */
export function qrEmailDefaults(lang: string = 'EN'): { subject: string; heading: string; subheading: string; body: string } {
  const isDe = (lang || 'EN').toUpperCase() === 'DE';
  if (isDe) {
    return {
      subject: 'Dein QR-Code für {{EventTitle}}',
      heading: 'Dein QR-Code',
      subheading: '{{EventTitle}}',
      body: `<p>Hallo {{Vorname}},</p>
<p>hier ist dein persönlicher QR-Code für das Event <strong>{{EventTitle}}</strong>.</p>
<p>Bitte zeige den QR-Code beim Check-in vor.</p>
{{QR_BLOCK}}
<p>Tipp: Du findest deinen QR-Code auch jederzeit in der <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">DEX App</a> unter <strong>„Meine Events"</strong> (Button „Mein QR-Code") — falls du diese Mail am Eventtag nicht zur Hand hast.</p>
<p style="color:#999;font-size:12px;text-align:center;">Der QR-Code ist persönlich und nicht übertragbar.</p>
<p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`,
    };
  }
  return {
    subject: 'Your QR Code for {{EventTitle}}',
    heading: 'Your QR Code',
    subheading: '{{EventTitle}}',
    body: `<p>Dear {{Vorname}},</p>
<p>here is your personal QR code for the event <strong>{{EventTitle}}</strong>.</p>
<p>Please show this QR code at check-in.</p>
{{QR_BLOCK}}
<p>Tip: You can also open your QR code anytime in the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">DEX app</a> under <strong>“My Events”</strong> (button “My QR code”) — in case you do not have this email at hand on the event day.</p>
<p style="color:#999;font-size:12px;text-align:center;">This QR code is personal and non-transferable.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`,
  };
}

/** Fester QR-Block: Code-Bild + „<Name> | <Event>“ als Klartext (für den
 *  manuellen Check-in). Wird beim Versand UND in der Editor-Vorschau gleich
 *  gebaut. */
export function buildQrBlockHtml(qrImageHtml: string, fullDisplayName: string, eventTitle: string): string {
  const escName = (fullDisplayName || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escTitle = (eventTitle || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const checkInLabel = `<p style="text-align:center;font-size:16px;margin:8px 0 0;"><strong>${escName} | ${escTitle}</strong></p>`;
  return `<div style="text-align:center;margin:24px 0;">${qrImageHtml}${checkInLabel}</div>`;
}

export function qrCodeEmail(
  firstName: string,
  eventTitle: string,
  qrImageHtml: string,
  lang: string = 'EN',
  fullName?: string,
  override?: QrEmailOverride
): { subject: string; body: string } {
  // Fallback: wenn kein fullName uebergeben, nutze nur firstName
  const fullDisplayName = (fullName || firstName || '').trim();
  const qrBlock = buildQrBlockHtml(qrImageHtml, fullDisplayName, eventTitle);
  const defaults = qrEmailDefaults(lang);
  const subjectTpl = (override && override.subject && override.subject.trim()) ? override.subject : defaults.subject;
  const headingTpl = (override && override.heading && override.heading.trim()) ? override.heading : defaults.heading;
  const subheadingTpl = (override && override.subheading && override.subheading.trim()) ? override.subheading : defaults.subheading;
  const bodyTpl = (override && override.bodyHtml && override.bodyHtml.trim()) ? override.bodyHtml : defaults.body;
  const vars: Record<string, string> = {
    Vorname: firstName || fullDisplayName,
    Name: fullDisplayName,
    EventTitle: eventTitle || '',
  };
  // Personen-/Event-Platzhalter HTML-escaped ersetzen, den QR-Block danach
  // RAW einsetzen (er ist selbst HTML). Subject/Headings ohne HTML-Escape.
  let body = replacePlaceholders(bodyTpl, vars);
  if (body.indexOf('{{QR_BLOCK}}') >= 0) {
    body = body.replace(/\{\{QR_BLOCK\}\}/g, qrBlock);
  } else {
    body += qrBlock;
  }
  return {
    subject: replacePlaceholdersPlain(subjectTpl, vars),
    body: wrapTemplate(
      GREEN,
      replacePlaceholdersPlain(headingTpl, vars),
      replacePlaceholdersPlain(subheadingTpl, vars),
      body
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
export function buildOutlookBody(eventTitle: string, bodyText: string, subheading?: string, imgOpts?: { imageWidth?: number; imagePaddingV?: number; imagePaddingH?: number }): string {
  const inner = stripOutlookWrapper(bodyText || '');
  const isHtml = /<[a-z][\s\S]*>/i.test(inner);
  const bodyHtml = inner
    ? (isHtml ? inner : inner.split('\n').map(line => `<p>${escapeHtml(line)}</p>`).join('\n  '))
    : '';
  // v18.73: Header-Bild (Breite + Innenabstand) auch im Outlook-Termin-Body
  // einstellbar — gleiche Logik wie in den Mails.
  return wrapTemplate(GREEN, eventTitle, subheading || 'Event Details', bodyHtml, undefined, imgOpts);
}

/**
 * v19.3: Ein HTML-Snippet (Team-Info-Box, Externe-Anmeldung-Hinweis, externer
 * QR-Hinweis) INNEN in den Inhaltsbereich des Deloitte-Templates einsetzen — an
 * den Anfang der Body-Content-Zelle, NICHT direkt nach <body> (sonst erscheint
 * der Block VOR/ÜBER der ganzen Template-Karte, optisch losgelöst — siehe
 * Bug-Report v19.3). Adressiert dieselbe Body-Content-Zelle wie
 * stripOutlookWrapper (`padding:0 30px 30px[ 30px]`). Fallback (Pre-wrapped
 * Storage-Templates ohne dieses Padding / abweichende Struktur): doch nach
 * <body> — besser als gar nicht.
 */
export function injectIntoEmailContent(wrappedHtml: string, snippet: string): string {
  if (!snippet) return wrappedHtml;
  const m = wrappedHtml.match(/<td style="padding:0 30px 30px(?:\s+30px)?;[^"]*">/i);
  if (m) return wrappedHtml.replace(m[0], `${m[0]}${snippet}`);
  return wrappedHtml.replace(/<body([^>]*)>/i, `<body$1>${snippet}`);
}

/**
 * Extrahiert den User-Content-Teil aus einem bereits-gewickelten Outlook-Body.
 * Wenn der Input nicht gewickelt ist, wird er unverändert zurückgegeben.
 * Dadurch können wir bestehende Events bearbeiten, ohne dass der Editor das
 * komplette Wrapper-HTML als rohen Text anzeigt.
 */
export function stripOutlookWrapper(html: string): string {
  if (!html) return '';
  if (!/<!doctype|<html/i.test(html)) return html;
  // wrapTemplate fügt den Body in <td style="padding:0 30px 30px[ 30px];...color:#333333;">CONTENT</td>
  // Non-greedy auf </td> reicht — wir brauchen keine trailing-Constraint (HTML-Kommentare
  // zwischen </tr> und nächstem <tr> würden eine engere Prüfung sonst sprengen, wodurch
  // der Strip fehlschlägt und beim Re-Save der gesamte Wrapper erneut umwickelt wird).
  // Seit v11.56 nutzt wrapTemplate 4-value-Paddings ("0 30px 30px 30px"), aber bestehende
  // gespeicherte Bodies haben noch 3-value ("0 30px 30px") — die Regex muss beide matchen.
  const m = html.match(/<td style="padding:0 30px 30px(?:\s+30px)?;[^"]*">([\s\S]*?)<\/td>/i);
  if (m && m[1]) return m[1].trim();
  return html;
}

/**
 * Extrahiert Heading (<h1>) und Subheading (<h2>) aus einer bereits gewickelten
 * Outlook-Body. Gibt leere Strings zurück wenn keine vorhanden.
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
 * Schreibt NICHT zurück - das macht EventService.ensureLogosInConfig().
 */
export async function loadLogosAsBase64(spHttpClient: SPHttpClient, siteUrl: string): Promise<void> {
  if (cachedLogoBase64 && cachedOrbBase64) return;

  try {
    // 1. Aus _Config Zeile lesen (falls bereits befüllt)
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
  } catch { /* Logo nicht verfügbar - Flow ersetzt Platzhalter als Fallback */ }
}

/**
 * Bild aus SiteAssets laden und als Base64 Data-URI zurückgeben.
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
