/**
 * E-Mail HTML Templates im Deloitte-Design
 *
 * Basiert auf der offiziellen Deloitte E-Mail-Vorlage
 * (Deloitte_DCGmbH_Email_with_Tagline.html).
 *
 * Generiert den kompletten HTML-Body fuer Power Automate.
 */

const GREEN = '#86bc25';
const SITE_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform';
const APP_URL = `${SITE_URL}/SitePages/Test_App.aspx?env=WebView`;

function getDate(): string {
  return new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function wrapTemplate(headingColor: string, heading: string, subheading: string, bodyHtml: string): string {
  return `<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${heading}</title>
</head>
<body style="margin:0;padding:0;background-color:#f5f5f5;font-family:Aptos,'Open Sans',Arial,Helvetica,sans-serif;color:#333333;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f5f5f5;">
<tr><td align="center" style="padding:20px 10px;">

<!-- Main Container -->
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="background-color:#ffffff;max-width:600px;width:100%;">

<!-- ===== HEADER: Deloitte Logo ===== -->
<tr>
<td style="background-color:#000000;padding:20px 30px;border-bottom:2px solid ${GREEN};">
  <img src="{{LOGO_URL}}" alt="Deloitte." width="180" style="display:block;max-width:180px;height:auto;" />
</td>
</tr>

<!-- ===== SUBHEADER ===== -->
<tr>
<td style="padding:10px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:11px;color:#999999;">
  Deutschland | Event Experience Platform | ${getDate()}
</td>
</tr>

<!-- ===== HERO: DEX Orb ===== -->
<tr>
<td style="background-color:#ffffff;text-align:center;padding:30px 30px;">
  <img src="{{ORB_URL}}" alt="DEX Event Experience Platform" width="180" style="display:inline-block;max-width:180px;height:auto;" />
</td>
</tr>

<!-- ===== GREEN LINE ===== -->
<tr>
<td style="background-color:${GREEN};height:4px;font-size:0;line-height:0;">&nbsp;</td>
</tr>

<!-- ===== CONTENT ===== -->
<tr>
<td style="padding:30px 30px 10px;">
  <h1 style="font-family:Aptos,Arial,Helvetica,sans-serif;font-size:26px;font-weight:400;color:${headingColor};margin:0 0 6px;">${heading}</h1>
  <h2 style="font-family:Aptos,Arial,Helvetica,sans-serif;font-size:20px;font-weight:700;color:#000000;margin:0 0 24px;">${subheading}</h2>
</td>
</tr>
<tr>
<td style="padding:0 30px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#333333;">
  ${bodyHtml}
</td>
</tr>

<!-- ===== FOOTER: Legal ===== -->
<tr>
<td style="padding:20px 30px;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:10px;color:#999999;line-height:1.5;border-top:1px solid #e8e8e8;">
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
 * Platzhalter in Subject und Body ersetzen.
 */
export function replacePlaceholders(text: string, vars: Record<string, string>): string {
  let result = text;
  for (const [key, value] of Object.entries(vars)) {
    result = result.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), escapeHtml(value));
  }
  // Logo-Platzhalter {{LOGO_URL}} und {{ORB_URL}} bleiben im HTML.
  // Der DEX_SEND_MAIL Flow ersetzt sie mit Base64 aus SharePoint.
  return result;
}

/**
 * Email aus SharePoint-Template generieren.
 * Nutzt wrapTemplate fuer das Deloitte-Design.
 */
export function buildEmailFromTemplate(
  template: { subject: string; headingColor: string; heading: string; bodyHtml: string },
  vars: Record<string, string>
): { subject: string; body: string } {
  const subject = replacePlaceholders(template.subject, vars);
  const heading = replacePlaceholders(template.heading, vars);
  const bodyHtml = replacePlaceholders(template.bodyHtml, vars);
  return {
    subject,
    body: wrapTemplate(template.headingColor, heading, `Event ${vars['EventTitle'] || ''}`, bodyHtml),
  };
}

// ==================== E-Mail Templates (Fallback) ====================

/**
 * Anmeldebestätigung
 */
export function registrationEmail(recipientName: string, eventTitle: string): { subject: string; body: string } {
  return {
    subject: `Anmeldebest\u00E4tigung: ${eventTitle}`,
    body: wrapTemplate(
      GREEN,
      'Registration successful',
      `Event ${eventTitle}`,
      `<p><strong>Dear ${recipientName},</strong></p>
      <p>you have successfully registered for the event <strong>${eventTitle}</strong>.</p>
      <p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">Event Experience Platform</a> (&bdquo;My Events&ldquo;) so the spot might be allocated to another participant.</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Warteliste-Bestätigung
 */
export function waitlistEmail(recipientName: string, eventTitle: string): { subject: string; body: string } {
  return {
    subject: `Warteliste: ${eventTitle}`,
    body: wrapTemplate(
      '#ed8b00',
      'Waitlist confirmation',
      `Event ${eventTitle}`,
      `<p><strong>Dear ${recipientName},</strong></p>
      <p>you have been placed on the <strong>waitlist</strong> for the event <strong>${eventTitle}</strong>.</p>
      <p>We will notify you as soon as a spot becomes available. You do not need to take any further action.</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}

/**
 * Abmeldebestätigung
 */
export function cancellationEmail(recipientName: string, eventTitle: string): { subject: string; body: string } {
  return {
    subject: `Abmeldebest\u00E4tigung: ${eventTitle}`,
    body: wrapTemplate(
      '#da291c',
      'Cancellation confirmed',
      `Event ${eventTitle}`,
      `<p><strong>Dear ${recipientName},</strong></p>
      <p>your registration for the event <strong>${eventTitle}</strong> has been <strong>cancelled</strong>.</p>
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
      `Event ${eventTitle}`,
      `<p><strong>Dear ${recipientName},</strong></p>
      <p>Great news! A spot has become available and you have been <strong>moved from the waitlist to a confirmed participant</strong> for the event <strong>${eventTitle}</strong>.</p>
      <p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="${APP_URL}" style="color:${GREEN};font-weight:600;">link</a> (&bdquo;My Events&ldquo;) so the spot might be allocated to another participant.</p>
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
      `Event ${eventTitle}`,
      `<p><strong>Dear ${recipientName},</strong></p>
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
 * Allgemeine Info-Mail
 */
/**
 * QR-Code E-Mail fuer Check-in
 */
export function qrCodeEmail(recipientName: string, eventTitle: string, qrImageHtml: string): { subject: string; body: string } {
  return {
    subject: `Dein QR-Code f\u00FCr ${eventTitle}`,
    body: wrapTemplate(
      GREEN,
      'Your QR Code',
      `Event ${eventTitle}`,
      `<p><strong>Dear ${recipientName},</strong></p>
      <p>here is your personal QR code for the event <strong>${eventTitle}</strong>.</p>
      <p>Please show this QR code at check-in.</p>
      <div style="text-align:center;margin:24px 0;">${qrImageHtml}</div>
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
      `Event ${eventTitle}`,
      `<p><strong>Dear ${recipientName},</strong></p>
      <p>${message}</p>
      <p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
    ),
  };
}
