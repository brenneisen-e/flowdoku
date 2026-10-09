/**
 * v32.59: Hinweismail an Organizer und Check-in-Team nach dem QR-Versand.
 *
 * Nutzer-Ansage 09.10.2026: Wenn die QR-Codes verschickt werden, sollen
 * Organizer und Check-in-Team IMMER erfahren, dass sie raus sind — mit dem
 * DEX-Link deutlich hervorgehoben (auf dem Firmenhandy öffnen), dem Weg zum
 * Check-in und dem Hinweis, dass iOS mit der Kamera scannt, Android nur über
 * die Namenssuche eincheckt, und der Laptop die Alternative ist.
 *
 * Bis v32.58 erfuhr das Check-in-Team vom Versand gar nichts. Wie man am
 * Event-Tag zum Check-in kommt, stand nur im Handbuch — und dass der
 * Kamera-Scan auf Android-Firmenhandys meist nicht läuft (CLAUDE.md,
 * „Kamera-Scan: Das Problem ist die APP"), merkte man erst am Einlass.
 *
 * Gerufen nur vom Massenversand (`qrFullSendAction`), nicht vom
 * Einzel-Versand bei späteren Anmeldungen — sonst käme die Mail bei jeder
 * Nachmeldung erneut.
 */
import { DeloitteEvent } from '../types';
import { buildMailButton, eckdatenHtml } from '../services/EmailTemplates';
import { buildHashDeepLink } from './deepLink';

const esc = (s: string): string => (s || '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * Alle Organizer (auch ausgeblendete — die Mail ist intern), Co-Organizer und
 * das Check-in-Team der übergebenen Events, jede Adresse einmal. Mehrere
 * Events, weil Klammer und Termine eigene Kopien der Teams tragen (CLAUDE.md,
 * „Eine Person umbenennen").
 */
export function qrHinweisEmpfaenger(evs: Array<DeloitteEvent | null | undefined>): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const ev of evs) {
    if (!ev) continue;
    for (const raw of [...(ev.organizerEmails || []), ...(ev.coOrganizerEmails || []), ...(ev.qrScannerEmails || [])]) {
      const e = (raw || '').trim();
      const lc = e.toLowerCase();
      if (!e || lc.indexOf('@') <= 0 || seen.has(lc)) continue;
      seen.add(lc);
      out.push(e);
    }
  }
  return out;
}

/** Direktlink auf den Check-in dieses Events (`#action=checkin`, DexEventPlatform). */
export function qrCheckInLink(siteUrl: string, ev: DeloitteEvent): string {
  return buildHashDeepLink(`${siteUrl}/SitePages/DEX.aspx?env=WebView`, { action: 'checkin', event: ev.id });
}

export interface QrCheckInHinweis {
  subject: string;
  heading: string;
  subheading: string;
  /** Inhalt für `wrapTemplate` (ohne Rahmen). */
  inner: string;
}

/**
 * Die Mail selbst. Sprache nach der Mail-Sprache des Events — dieselbe, in der
 * die Teilnehmenden ihre QR-Mail bekommen haben.
 */
export function qrCheckInHinweisMail(ev: DeloitteEvent, opts: { anzahl: number; absender: string; link: string }): QrCheckInHinweis {
  const isDe = (ev.emailLanguage || 'EN').toUpperCase() === 'DE';
  const titel = esc(ev.title || '');
  const wer = esc((opts.absender || '').trim());
  const n = opts.anzahl;
  const p = (html: string): string => `<p style="margin:0 0 12px;">${html}</p>`;
  const zwischen = (text: string): string => `<p style="margin:22px 0 8px;font-size:16px;font-weight:700;color:#1a1a1a;">${text}</p>`;
  const liste = (items: string[]): string => `<ol style="margin:0 0 12px;padding-left:22px;line-height:1.6;">${items.map(i => `<li style="margin:0 0 6px;">${i}</li>`).join('')}</ol>`;
  // Der Link ist der Kern der Mail — deshalb ein eigener Kasten mit Knopf und
  // der Adresse als Text darunter (zum Weiterschicken aufs Firmenhandy, falls
  // die Mail am Laptop gelesen wird). Tabelle statt div: Die Word-Engine von
  // Outlook ignoriert Rahmen und Hintergrund auf div zuverlässig nur halb.
  const kasten = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
<tr><td height="6" style="height:6px;font-size:0;line-height:0;">&nbsp;</td></tr>
<tr><td style="padding:18px 20px 4px;background:#f2f8e8;border:2px solid #86bc25;border-radius:10px;">
<p style="margin:0 0 6px;font-size:18px;font-weight:700;color:#1a1a1a;">${isDe ? 'Dein Link zum Check-in' : 'Your check-in link'}</p>
<p style="margin:0 0 4px;">${isDe
    ? `Am Event-Tag auf dem <strong>Firmenhandy</strong> öffnen — der Check-in für &bdquo;${titel}&ldquo; startet direkt.`
    : `Open it on your <strong>company phone</strong> on event day — check-in for &ldquo;${titel}&rdquo; starts right away.`}</p>
${buildMailButton(opts.link, isDe ? 'Check-in öffnen &rarr;' : 'Open check-in &rarr;')}
<p style="margin:0 0 14px;font-size:12px;line-height:1.5;color:#555555;">${isDe ? 'Link zum Kopieren:' : 'Link to copy:'}<br><a href="${opts.link}" style="color:#26890d;word-break:break-all;">${esc(opts.link)}</a></p>
</td></tr>
<tr><td height="10" style="height:10px;font-size:0;line-height:0;">&nbsp;</td></tr>
</table>`;
  if (isDe) {
    return {
      subject: `Check-in für „${ev.title}“: QR-Codes sind verschickt`,
      heading: 'QR-Codes sind verschickt',
      subheading: ev.title,
      inner: [
        p('Hallo zusammen,'),
        p(`für <strong>&bdquo;${titel}&ldquo;</strong> ${wer ? `hat ${wer} gerade` : 'wurden gerade'} <strong>${n} QR-Code${n === 1 ? '' : 's'}</strong> an die Teilnehmenden verschickt. Ihr bekommt diese Mail, weil ihr als Organizer oder im Check-in-Team eingetragen seid.`),
        eckdatenHtml(ev, 'DE'),
        kasten,
        zwischen('So kommt ihr zum Check-in'),
        liste([
          'Diese Mail auf dem Firmenhandy öffnen und auf <strong>Check-in öffnen</strong> tippen. Falls gefragt, mit dem Deloitte-Konto anmelden.',
          `Der Check-in für <strong>&bdquo;${titel}&ldquo;</strong> öffnet sich. Ohne Link: DEX öffnen &rarr; Kachel <strong>Check-In</strong> &rarr; Event auswählen.`,
        ]),
        zwischen('iPhone: mit der Kamera'),
        liste([
          'Die Karte <strong>Live-Scanner</strong> öffnen und den Kamera-Zugriff erlauben.',
          'Den QR-Code aus der Mail der Person scannen — fertig.',
          'Klappt der Scan nicht: in derselben Karte <strong>Stattdessen Foto vom QR-Code machen</strong> wählen oder die Person über die Suche einchecken (wie bei Android).',
        ]),
        zwischen('Android: über die Namenssuche'),
        p('Auf Android-Firmenhandys lässt die App die Kamera meist nicht zu. Der Check-in läuft dort über die Suche:'),
        liste([
          'Ins Suchfeld <strong>&bdquo;Teilnehmer-ID, Vorname, Nachname oder E-Mail…&ldquo;</strong> den Namen tippen — oder die <strong>Teilnehmer-ID</strong>, die in der QR-Mail der Person unter dem Code steht.',
          'Bei der richtigen Person auf <strong>Einchecken</strong> tippen.',
        ]),
        zwischen('Alternativ: am Laptop'),
        p('Den Link oben am Laptop öffnen und genauso über die Suche einchecken. Das funktioniert immer und passt gut zu einem festen Check-in-Tisch.'),
        p('Viel Erfolg beim Event!'),
      ].join('\n'),
    };
  }
  return {
    subject: `Check-in for “${ev.title}”: QR codes have been sent`,
    heading: 'QR codes have been sent',
    subheading: ev.title,
    inner: [
      p('Hello everyone,'),
      p(`for <strong>&ldquo;${titel}&rdquo;</strong> ${wer ? `${wer} has just sent` : 'we have just sent'} <strong>${n} QR code${n === 1 ? '' : 's'}</strong> to the attendees. You receive this email because you are listed as an organizer or in the check-in team.`),
      eckdatenHtml(ev, 'EN'),
      kasten,
      zwischen('How to get to check-in'),
      liste([
        'Open this email on your company phone and tap <strong>Open check-in</strong>. Sign in with your Deloitte account if asked.',
        `Check-in for <strong>&ldquo;${titel}&rdquo;</strong> opens. Without the link: open DEX &rarr; tile <strong>Check-in</strong> &rarr; select the event.`,
      ]),
      zwischen('iPhone: with the camera'),
      liste([
        'Open the <strong>Live scanner</strong> card and allow camera access.',
        'Scan the QR code from the person&rsquo;s email — done.',
        'If scanning fails: in the same card choose the photo option, or check the person in via search (as on Android).',
      ]),
      zwischen('Android: via name search'),
      p('On Android company phones the app usually does not allow the camera. Check-in works via search there:'),
      liste([
        'Type the name into the search field — or the <strong>attendee ID</strong> printed below the code in the person&rsquo;s QR email.',
        'Tap <strong>Check in</strong> next to the right person.',
      ]),
      zwischen('Alternatively: on a laptop'),
      p('Open the link above on a laptop and check in via search the same way. This always works and suits a fixed check-in desk.'),
      p('Good luck with the event!'),
    ].join('\n'),
  };
}
