/**
 * E-Mail HTML Templates im Deloitte-Design
 *
 * Basiert auf der offiziellen Deloitte E-Mail-Vorlage
 * (Deloitte_DCGmbH_Email_with_Tagline.html).
 *
 * Generiert den kompletten HTML-Body für Power Automate.
 */

import { normalizeMailHeaderImage } from '../utils/mailHeaderImage';
import { SPHttpClient } from '@microsoft/sp-http';
import { buildHashDeepLink } from '../utils/deepLink';

const GREEN = '#86bc25';
const SITE_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform';
const APP_URL = `${SITE_URL}/SitePages/DEX.aspx?env=WebView`;

/**
 * v29.42: Die Fußzeile „Made with DEX App" IMMER auf die kanonische App-Adresse
 * zeigen lassen.
 *
 * Der Link im Template hier stimmt — aber Mail-Bodies kommen nicht nur von
 * hier: gespeicherte Vorlagen aus `DEX_EmailTemplates`, Event-Overrides und
 * kopierte Events tragen fertig gewickeltes HTML aus älteren App-Ständen mit
 * sich, inklusive der Fußzeile von damals. Deshalb wird der href beim Ausgeben
 * neu gesetzt statt darauf zu vertrauen, dass er richtig gespeichert wurde.
 * Nur der href des Links mit genau diesem Text — der übrige Body bleibt, wie
 * der Organizer ihn geschrieben hat.
 */
export function normalizeMadeWithLink(html: string): string {
  if (!html || html.indexOf('Made with DEX') < 0) return html || '';
  return html.replace(
    /(<a\b[^>]*?)href\s*=\s*("|')[^"']*\2([^>]*>\s*Made with DEX(?:\s+App)?\s*<\/a>)/gi,
    (_m, pre: string, _q: string, post: string) => `${pre}href="${APP_URL}"${post}`
  );
}
// Gecachtes Logo Base64 aus DEX_EmailTemplates (_Config)
// ORB/Event-Bild wird NICHT gecacht - der Flow setzt das event-spezifische Bild ein
let cachedLogoBase64 = '';
let cachedOrbBase64 = '';

export function getCachedOrbBase64(): string { return cachedOrbBase64; }
export function getCachedLogoBase64(): string { return cachedLogoBase64; }
/** v26.50: Nach einem Logo-Tausch im Admin-Center („Logo & Branding") den
 *  Session-Cache aktualisieren — neue Mails derselben Sitzung nutzen sofort
 *  das neue Logo, ohne Reload. */
export function setCachedLogoBase64(v: string): void { if (v) cachedLogoBase64 = v; }
/** v26.58: dito für das DEX-Orb (Default-Mail-Bild / {{ORB_URL}}-Fallback). */
export function setCachedOrbBase64(v: string): void { if (v) cachedOrbBase64 = v; }

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
 * v28.24: Feste Pixelbreite + `max-width:100%` statt `width:100%` +
 * `max-width:<w>px`. Grund: Outlook-Desktop rendert mit der Word-Engine und
 * ignoriert `max-width` — dort blieb vom Paar nur `width:100%` übrig, das Bild
 * lief also auf die volle Spaltenbreite (~540px) auf, während dieselbe Vorlage
 * in den Mails korrekt bei der eingestellten Breite landete. Genau das war der
 * riesige Orb im Outlook-Termin. Umgekehrt herum stimmt beides: Outlook nimmt
 * die feste Breite, alle max-width-fähigen Clients schrumpfen auf schmalen
 * Displays weiterhin sauber mit. Das `width`-Attribut bleibt als zusätzlicher
 * Fallback für sehr alte Clients erhalten.
 */
function buildHeroRow(opts?: WrapHeadingOpts): string {
  const w = (typeof opts?.imageWidth === 'number' && opts.imageWidth > 0) ? Math.round(opts.imageWidth) : 180;
  const padV = (typeof opts?.imagePaddingV === 'number' && opts.imagePaddingV >= 0) ? Math.round(opts.imagePaddingV) : 30;
  const padH = (typeof opts?.imagePaddingH === 'number' && opts.imagePaddingH >= 0) ? Math.round(opts.imagePaddingH) : 30;
  return `<tr>
<td style="background-color:#ffffff;text-align:center;padding:${padV}px ${padH}px ${padV}px ${padH}px;">
  <img src="{{ORB_URL}}" alt="DEX Event Experience Platform" width="${w}" style="display:inline-block;width:${w}px;max-width:100%;height:auto;" />
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
  <a href="${APP_URL}" style="color:${GREEN};text-decoration:none;">Made with DEX App</a>
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
        <p>Wenn du teilnehmen möchtest, <strong>bestätige bitte diese Einladung, indem du über &bdquo;Allen antworten&ldquo; auf diese E-Mail antwortest</strong>.</p>
        <p style="margin-top:16px;font-size:13px;color:#555;">Mit deiner Anmeldung zum Event <strong>${eventTitle}</strong> willigst du ein, dass deine personenbezogenen Daten zum Zweck der Organisation und Durchführung der Veranstaltung verarbeitet werden. Dies umfasst insbesondere die Erhebung, Speicherung und Nutzung der von dir angegebenen Daten zur Anmeldung, Kommunikation und Teilnahmeabwicklung.</p>
        <p style="font-size:13px;color:#555;">Weitere Informationen zur Verarbeitung deiner Daten findest du <a href="${privacyUrl}" style="color:${GREEN};font-weight:600;">hier</a> in den Datenschutzhinweisen von Deloitte.</p>
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
      <p>If you would like to attend, <strong>please confirm this invitation by replying to this email via &ldquo;Reply all&rdquo;</strong>.</p>
      <p style="margin-top:16px;font-size:13px;color:#555;">By registering for the event <strong>${eventTitle}</strong>, you consent to the processing of your personal data for the purpose of organising and running the event. This includes in particular the collection, storage and use of the data you provide for registration, communication and attendance handling.</p>
      <p style="font-size:13px;color:#555;">Further information on how your data is processed is available <a href="${privacyUrl}" style="color:${GREEN};font-weight:600;">here</a> in Deloitte's data protection notice.</p>
      <p style="margin-top:24px;"><strong>Best regards</strong><br><br><strong>Your Event Team</strong></p>`
    ),
  };
}

/**
 * v26.47: Instruktions-Mail an die ANMELDENDE Person nach einer externen
 * stellvertretenden Anmeldung. Hintergrund: Der Mail-Flow (no_reply-Postfach)
 * kann externe Adressen NICHT erreichen — die Einladung verschickt deshalb
 * die anmeldende Person selbst (fertiger .eml-Entwurf zum Download in der
 * Teilnehmerliste). Bis zur Datenschutz-Rückmeldung der externen Person steht
 * sie mit Status „Angemeldet (Datenschutzrückmeldung offen)" auf der Liste.
 */
export function externalInviteInstructionEmail(
  registrantFirstName: string,
  externalName: string,
  externalEmail: string,
  eventTitle: string,
  isDe: boolean,
  orgCenterUrl: string,
  // v26.73: Deeplink, der die App öffnet und die an der Teilnehmer-Zeile
  // abgelegte .eml sofort herunterlädt. Leer = nur Organizer-Center-Fallback.
  downloadUrl: string = ''
): { subject: string; body: string } {
  const btn = (href: string, label: string): string =>
    `<p style="margin:18px 0 0;"><a href="${href}" style="display:inline-block;background:${GREEN};color:#ffffff;text-decoration:none;padding:11px 24px;border-radius:6px;font-weight:600;font-size:14px;">${label}</a></p>`;
  if (isDe) {
    return {
      subject: `Externe Anmeldung: ${externalName} — deine nächsten Schritte`,
      body: wrapTemplate(
        GREEN,
        'Externe Anmeldung — nächste Schritte',
        eventTitle,
        `<p>Hallo ${registrantFirstName},</p>
        <p>du hast <strong>${externalName}</strong> (${externalEmail}) für das Event <strong>${eventTitle}</strong> angemeldet. Da es sich um eine externe Adresse handelt, kann die App die Einladung <strong>nicht direkt zustellen</strong> — bitte übernimm das in drei kurzen Schritten:</p>
        <ol style="margin:12px 0 0 18px;padding:0;line-height:1.7;">
          <li><strong>${externalName} steht bereits auf der Teilnehmerliste</strong> — mit dem Status <strong>&bdquo;Angemeldet (Datenschutzrückmeldung offen)&ldquo;</strong>.</li>
          <li>Klicke unten auf <strong>&bdquo;Einladungs-Entwurf herunterladen&ldquo;</strong> — die DEX App öffnet sich und lädt den fertigen Entwurf (.eml) sofort herunter. An, Betreff und Text sind vorausgefüllt (an die externe Person; Organisator:innen in Kopie) — in Outlook öffnen und auf <strong>&bdquo;Senden&ldquo;</strong> klicken. <span style="color:#555;">(Zeigt Outlook die Datei als empfangene Mail an, nutze einfach &bdquo;Weiterleiten&ldquo;.)</span></li>
          <li>Sobald die Datenschutz-Rückmeldung von ${externalName} per Mail da ist, klicke in der Teilnehmerliste auf <strong>&bdquo;Rückmeldung bestätigen&ldquo;</strong> — der Status wechselt auf &bdquo;Angemeldet&ldquo;.</li>
        </ol>
        <p style="margin-top:14px;padding:10px 14px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;color:#7a4a00;">Kein Zugriff auf das Organizer Center? Bitte eine:n der Organisator:innen, die Schritte zu übernehmen.</p>
        ${downloadUrl ? btn(downloadUrl, 'Einladungs-Entwurf herunterladen') : ''}
        <p style="margin:10px 0 0;font-size:12px;color:#888;">Alternativ: <a href="${orgCenterUrl}" style="color:${GREEN};font-weight:600;">im Organizer Center öffnen</a> und den Entwurf in der Teilnehmerzeile herunterladen.</p>
        <p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>`
      ),
    };
  }
  return {
    subject: `External registration: ${externalName} — your next steps`,
    body: wrapTemplate(
      GREEN,
      'External registration — next steps',
      eventTitle,
      `<p>Hi ${registrantFirstName},</p>
      <p>you registered <strong>${externalName}</strong> (${externalEmail}) for the event <strong>${eventTitle}</strong>. Since this is an external address, the app <strong>cannot deliver the invitation directly</strong> — please take over in three quick steps:</p>
      <ol style="margin:12px 0 0 18px;padding:0;line-height:1.7;">
        <li><strong>${externalName} is already on the participant list</strong> — with the status <strong>&ldquo;Registered (privacy confirmation pending)&rdquo;</strong>.</li>
        <li>Click <strong>&ldquo;Download invitation draft&rdquo;</strong> below — the DEX App opens and downloads the ready-made draft (.eml) right away. To, subject and text are pre-filled (to the external person; organizers in copy) — open it in Outlook and click <strong>&ldquo;Send&rdquo;</strong>. <span style="color:#555;">(If Outlook shows the file as a received mail, simply use &ldquo;Forward&rdquo;.)</span></li>
        <li>Once ${externalName}'s privacy confirmation arrives by email, click <strong>&ldquo;Confirm response&rdquo;</strong> in the participant list — the status switches to &ldquo;Registered&rdquo;.</li>
      </ol>
      <p style="margin-top:14px;padding:10px 14px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;color:#7a4a00;">No access to the Organizer Center? Ask one of the organizers to take over these steps.</p>
      ${downloadUrl ? btn(downloadUrl, 'Download invitation draft') : ''}
      <p style="margin:10px 0 0;font-size:12px;color:#888;">Alternatively: <a href="${orgCenterUrl}" style="color:${GREEN};font-weight:600;">open the Organizer Center</a> and download the draft from the participant's row.</p>
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
  // ging nur über den Subject-Zusatz hervor).
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
 * v27.13: subsiteUrl-Parameter bleibt für Call-Site-Kompatibilität, wird aber
 * nicht mehr verlinkt — der SharePoint-Teilnehmerlisten-Link wurde entfernt
 * (alle Aktionen laufen über die App).
 */
export function eventCreatedEmail(
  recipientName: string,
  eventTitle: string,
  _subsiteUrl: string,
  // v29.32: Kopfbild-Layout des Events durchreichen. Ohne diesen Parameter lief
  // die Mail immer auf den Default (180 px, zentriert) — auch bei Events, die
  // seit v29.29 auf den Vollbild-Kopf gesetzt sind. Der Aufrufer übergibt das
  // über headerLayoutFor ermittelte Layout, das ohne eigenes Bild weiterhin auf
  // 180 px deckelt (sonst füllt der DEX-Orb die halbe Mail).
  headerOpts?: WrapHeadingOpts,
): { subject: string; body: string } {
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
      <p>Teilnehmerliste &amp; Verwaltung &ndash; alles direkt in der App (Organizer Center):</p>
      <ul>
        <li><a href="${APP_URL}" style="color:${GREEN};font-weight:600;">DEX App</a> (Admin / Organizer)</li>
      </ul>
      <p>Viele Gr&uuml;&szlig;e<br>Team DEX App</p>
      <hr style="border:none;border-top:1px solid #e0e0e0;margin:20px 0;">
      <p style="color:#999;font-size:13px;line-height:1.5;">English: Your event <strong>${eventTitle}</strong> was created successfully. Next steps: (1)&nbsp;finalize it via &bdquo;Edit event&ldquo;, (2)&nbsp;do a test registration &amp; cancellation to check the automatic emails and the Outlook invite, (3)&nbsp;publish it via the status toggle &bdquo;Draft &rarr; Active&ldquo;, (4)&nbsp;optionally send the invitation with the registration link from the app, (5)&nbsp;track registrations in the Organizer Center. Participant list &amp; admin: see the links above.</p>`,
      undefined,
      headerOpts,
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

      <p style="margin-top:16px;padding:10px 12px;background:#f7f7f5;border-left:3px solid ${GREEN};">
      <strong>Wofür DEX gedacht ist:</strong> für <strong>interne Deloitte Events</strong> und für die
      Koordination der Deloitte-Teilnahme an externen Veranstaltungen. Für
      <strong>externe Events mit externen Teilnehmern</strong> ist die Plattform nicht vorgesehen &mdash;
      alles dazu findest du im
      <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" style="color:${GREEN};font-weight:600;">Event Management im DeloitteNet</a>.</p>

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
  /**
   * v30.52: Kopf-Bild der QR-Mail (Auswahl + Maße), gespeichert je Event.
   *
   * Anders als bei Massenmail und Einladung ist das hier eine DAUERHAFTE
   * Einstellung: Die QR-Mail geht auch automatisch raus, wenn sich jemand
   * neu anmeldet — es gibt in dem Moment keinen Dialog, in dem man die
   * Breite noch einstellen könnte. Deshalb liegt sie im Override und nicht
   * im Zustand eines Fensters.
   *
   * Gespeichert werden nur Auswahl und Zahlen, NIE das Foto selbst: Das
   * Feld `EmailTemplateOverrides` trägt schon das Mail-Logo, und ein zweites
   * Base64-Bild darin bringt die Spalte an ihr Größenlimit. Das Foto wird
   * beim Versand aus dem Event-Bild aufgelöst (Cache, s. utils/imageCache).
   */
  headerImage?: { hero?: 'logo' | 'event'; width?: number; paddingV?: number; paddingH?: number };
  /**
   * v30.60: Sprache des festen Blocks NEBEN dem QR-Code („Name", „ID" und der
   * Hinweis „Falls der Scan nicht klappt…").
   *
   * Warum eine eigene Einstellung und nicht einfach `event.emailLanguage`:
   * Der Mailtext der QR-Mail ist frei überschreibbar. Wer eine englische
   * Einladung schreibt, das Event aber auf Deutsch stehen lässt (oder
   * umgekehrt), bekam bisher zwangsläufig einen Block in der falschen
   * Sprache — im gemeldeten Fall stand „Falls der Scan nicht klappt…" unter
   * einer Mail, die mit „Dear Alexander" beginnt.
   *
   * Leer heißt: der Mail-Sprache des Events folgen. Es ist damit kein
   * zweiter Schalter für dieselbe Frage, sondern eine Ausnahme, die man nur
   * setzt, wenn Text und Event-Sprache auseinandergehen.
   */
  blockLang?: 'DE' | 'EN';
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
/**
 * v30.35: QR links, Name und Teilnehmer-ID rechts daneben.
 *
 * Vorher stand alles zentriert untereinander — inklusive des Event-Titels, der
 * zwei Zeilen weiter oben im Mailtext schon steht. Doppelt genannt macht er den
 * Block nur höher, ohne etwas zu klären.
 *
 * Die ID ist auf drei Stellen aufgefüllt (`017`). Das ist kein Selbstzweck: Sie
 * wird am Einlass vorgelesen und abgetippt, und eine feste Breite liest sich
 * verlässlicher als „7" mitten im Fließtext.
 *
 * Aufbau als `<table>` mit zwei Zellen, nicht als Flexbox — Outlook rendert
 * mit Word und kann kein Flex. `valign="middle"` hält die Textspalte auf
 * Höhe des Codes.
 */
export function buildQrBlockHtml(qrImageHtml: string, fullDisplayName: string, teilnehmerId?: number, lang: string = 'DE'): string {
  // v30.60: Der Block war fest deutsch — „Name:", „ID:" und der Hinweis unter
  // der Nummer standen auch in einer englischen QR-Mail auf Deutsch. Der
  // Mailtext folgte der Mail-Sprache, dieser Block nicht; im Ergebnis stand
  // „Falls der Scan nicht klappt…" unter einer Mail, die mit „Dear Alexander"
  // beginnt. Die Sprache kommt jetzt von derselben Einstellung wie der Text.
  const de = (lang || 'DE').toUpperCase() === 'DE';
  const escName = (fullDisplayName || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const hasId = teilnehmerId !== undefined && teilnehmerId !== null && !isNaN(Number(teilnehmerId));
  // padStart gibt es im ES5-Target nicht — deshalb von Hand auffüllen.
  let idText = '';
  if (hasId) {
    idText = String(teilnehmerId);
    while (idText.length < 3) idText = `0${idText}`;
  }
  const rows = `<div style="font-size:15px;color:#2b2b2b;line-height:1.5;">`
    + `<span style="color:#63666A;">${de ? 'Name:' : 'Name:'}</span> <strong>${escName}</strong>`
    + `</div>`
    + (hasId
      ? `<div style="font-size:15px;color:#2b2b2b;line-height:1.5;margin-top:6px;">`
        + `<span style="color:#63666A;">ID:</span> `
        + `<strong style="font-family:'Courier New',Courier,monospace;font-size:26px;letter-spacing:0.04em;">${idText}</strong>`
        + `</div>`
        + `<div style="font-size:12px;color:#63666A;margin-top:8px;line-height:1.45;">`
        + (de
          ? `Falls der Scan nicht klappt: einfach diese Nummer am Einlass nennen.`
          : `If the scan does not work, simply give this number at the entrance.`)
        + `</div>`
      : '');
  return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:24px 0;border-collapse:collapse;">`
    + `<tr>`
    + `<td valign="middle" style="padding:0 18px 0 0;">${qrImageHtml}</td>`
    + `<td valign="middle" style="font-family:Arial,Helvetica,sans-serif;">${rows}</td>`
    + `</tr>`
    + `</table>`;
}

export function qrCodeEmail(
  firstName: string,
  eventTitle: string,
  qrImageHtml: string,
  lang: string = 'EN',
  fullName?: string,
  override?: QrEmailOverride,
  teilnehmerId?: number, // v30.35: erscheint rechts neben dem Code
  /**
   * v30.52: Event-Foto als Base64, wenn der Organizer es als Kopf-Bild
   * gewählt hat. Der Aufrufer löst es auf (er kennt das Event); ohne den
   * Wert bleibt `{{ORB_URL}}` stehen und der Flow setzt das Standard-Bild —
   * genau das Verhalten von vorher.
   */
  eventPhotoB64?: string
): { subject: string; body: string } {
  // Fallback: wenn kein fullName übergeben, nutze nur firstName
  const fullDisplayName = (fullName || firstName || '').trim();
  // v30.60: Der Block folgt der Mail-Sprache — es sei denn, der Organizer hat
  // für ihn ausdrücklich etwas anderes gewählt (s. QrEmailOverride.blockLang).
  const qrBlock = buildQrBlockHtml(qrImageHtml, fullDisplayName, teilnehmerId, (override && override.blockLang) || lang);
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
  // v30.52: Kopf-Bild aus dem Override anwenden — Maße über wrapTemplate,
  // das Foto (falls gewählt UND vom Aufrufer aufgelöst) statt {{ORB_URL}}.
  const hdr = normalizeMailHeaderImage(override && override.headerImage);
  // Der Orb-Schutz gilt auch hier: Ohne eigenes Bild wäre „Volle Breite" ein
  // bildschirmfüllender, unten abgeschnittener Orb (s. utils/mailHeaderImage).
  const ownImage = hdr.hero === 'event' && !!eventPhotoB64;
  const wrapped = wrapTemplate(
    GREEN,
    replacePlaceholdersPlain(headingTpl, vars),
    replacePlaceholdersPlain(subheadingTpl, vars),
    body,
    undefined,
    {
      imageWidth: ownImage ? hdr.width : Math.min(hdr.width, 180),
      imagePaddingV: ownImage ? hdr.paddingV : Math.max(hdr.paddingV, 20),
      imagePaddingH: ownImage ? hdr.paddingH : Math.max(hdr.paddingH, 20),
    }
  );
  return {
    subject: replacePlaceholdersPlain(subjectTpl, vars),
    body: ownImage ? wrapped.replace(/\{\{ORB_URL\}\}/g, eventPhotoB64!) : wrapped,
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
/**
 * v29.38: Teilnahme-Block für einen hinterlegten Teams-Link. Bewusst mit
 * `data-dex-teams` markiert und ohne verschachtelte Tabelle: Der Block ist
 * reine AUSGABE — er wird beim nächsten Wickeln entfernt und neu erzeugt,
 * damit er nicht bei jedem Speichern ein weiteres Mal im Termin steht und im
 * Editor nicht als Text auftaucht, den man versehentlich mitbearbeitet.
 */
export function teamsJoinBlockHtml(link: string, isDe = true): string {
  const url = (link || '').trim();
  if (!url) return '';
  const safe = escapeHtml(url);
  // v30.28: Bei der Platzhalter-Variante die sichtbare URL-Zeile weglassen.
  // Der Flow ersetzt die Marke erst NACH dem Anlegen des Termins; hat er die
  // Ersetzung noch nicht (Actions nicht eingebaut, Lauf fehlgeschlagen), stand
  // sonst wörtlich „{{TEAMS_URL}}" als Text in der Einladung beim Teilnehmer.
  // Der Knopf bleibt — er trägt dieselbe Marke im href und wird mit ersetzt.
  const isPlaceholder = url === TEAMS_URL_PLACEHOLDER;
  return `<table data-dex-teams="1" role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0 8px;border-top:1px solid #e0e0e0;"><tr><td style="padding:16px 0 0;font-family:Arial,Helvetica,sans-serif;">`
    + `<div style="font-size:13px;color:#63666A;margin-bottom:6px;">${isDe ? 'Online teilnehmen' : 'Join online'}</div>`
    + `<a href="${safe}" style="display:inline-block;background:#86bc25;color:#ffffff;text-decoration:none;font-weight:bold;font-size:14px;padding:10px 18px;border-radius:4px;">${isDe ? 'An Microsoft-Teams-Besprechung teilnehmen' : 'Join the Microsoft Teams meeting'}</a>`
    + (isPlaceholder ? '' : `<div style="font-size:12px;color:#63666A;margin-top:8px;word-break:break-all;">${safe}</div>`)
    // v30.28: Einwahl-Zeile NUR im Platzhalter-Fall — bei einem selbst
    // eingetragenen Link kennt DEX keine Konferenz-Daten. Der Flow baut die
    // GANZE Zeile und setzt sie leer, wenn es keine gibt; deshalb eine einzige
    // Marke statt „Label + Wert" (sonst stünde bei leerem Wert ein nacktes
    // „Konferenz-ID:" im Termin).
    + (isPlaceholder ? `<div style="font-size:12px;color:#63666A;margin-top:10px;">${TEAMS_DIALIN_PLACEHOLDER}</div>` : '')
    + `</td></tr></table>`;
}

/**
 * v30.27: Platzhalter für die Teams-Beitritts-URL im Outlook-Body.
 *
 * Im Modus „DEX erzeugt den Teams-Link" existiert die URL beim Speichern noch
 * nicht — die Besprechung entsteht erst, wenn der Flow den Termin angelegt und
 * per Graph `isOnlineMeeting` gesetzt hat. Die App schreibt deshalb diese Marke
 * in den Body, der Flow ersetzt sie durch die echte `onlineMeeting.joinUrl`.
 * Gleiches Muster wie `{{ORB_URL}}` / `{{LOGO_URL}}` — Wert bewusst ohne
 * Sonderzeichen, damit escapeHtml ihn unverändert durchreicht.
 */
export const TEAMS_URL_PLACEHOLDER = '{{TEAMS_URL}}';

/**
 * v30.28: Platzhalter für die Einwahl-Zeile unter dem Teams-Knopf.
 *
 * Der Flow baut die KOMPLETTE Zeile („Konferenz-ID: … · Telefon: …") und setzt
 * sie leer, wenn es keine Einwahldaten gibt. Bewusst nicht „Label hier, Wert
 * als Marke": Bei leerem Wert stünde sonst ein nacktes „Konferenz-ID:" im
 * Termin — schlimmer als gar nichts.
 *
 * Was hier NICHT reinkann: Meeting-ID und Passcode aus dem Teams-Kasten. Die
 * gehören zur onlineMeeting-RESSOURCE (`/onlineMeetings`), nicht zur
 * `onlineMeeting`-Eigenschaft des Termins — die liefert nur joinUrl,
 * conferenceId, tollNumber/tollFreeNumbers und quickDial. Die Ressource ist
 * über die Standard-Action des Outlook-Connectors nicht erreichbar (erlaubt
 * sind nur die Segmente messages/events/calendar/…).
 */
export const TEAMS_DIALIN_PLACEHOLDER = '{{TEAMS_DIALIN}}';

/** v29.38: Bereits eingesetzte Teams-Blöcke entfernen (siehe teamsJoinBlockHtml). */
export function stripTeamsJoinBlock(html: string): string {
  if (!html) return '';
  return html.replace(/<table[^>]*data-dex-teams="1"[\s\S]*?<\/table>/gi, '');
}

export function buildOutlookBody(eventTitle: string, bodyText: string, subheading?: string, imgOpts?: { imageWidth?: number; imagePaddingV?: number; imagePaddingH?: number }, teamsLink?: string, teamsIsDe = true): string {
  const inner = stripTeamsJoinBlock(stripOutlookWrapper(bodyText || ''));
  const isHtml = /<[a-z][\s\S]*>/i.test(inner);
  const bodyHtml = inner
    ? (isHtml ? inner : inner.split('\n').map(line => `<p>${escapeHtml(line)}</p>`).join('\n  '))
    : '';
  // v18.73: Header-Bild (Breite + Innenabstand) auch im Outlook-Termin-Body
  // einstellbar — gleiche Logik wie in den Mails.
  // v29.38: Teams-Link ans Ende des Inhalts, nicht in den Editor-Text — er
  // gehört zum Termin, nicht zum vom Organizer geschriebenen Absatz.
  const withTeams = bodyHtml + teamsJoinBlockHtml(teamsLink || '', teamsIsDe);
  return wrapTemplate(GREEN, eventTitle, subheading || 'Event Details', withTeams, undefined, imgOpts);
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
