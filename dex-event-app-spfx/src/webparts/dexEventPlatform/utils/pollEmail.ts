/**
 * v31.12 — Die Umfrage-Mail: Antwortmöglichkeiten als anklickbare Kästchen.
 *
 * Eine echte Checkbox in einer Mail gibt es nicht — Outlook entfernt `<form>`
 * beim Rendern, und ein Kästchen ohne Formular hätte niemanden, an den es
 * etwas schicken könnte. Was hier gebaut wird, ist der Weg, den jedes
 * Umfrage-Werkzeug geht: Jede Antwort ist ein LINK, der wie ein Kästchen
 * aussieht. Ein Klick öffnet DEX, die Antwort ist sofort gespeichert.
 *
 * ## Wie die Kästchen gebaut sind — und warum v31.13 sie neu gebaut hat
 *
 * Der erste Entwurf (v31.12) sah in der echten Mail aus wie eine Liste von
 * Zitatblöcken: 1 px grauer Rahmen, 4 px grüner Balken links, ein
 * Schriftzeichen als Kästchen. Gemeldet 11.09.2026 mit Bildschirmfoto. Nach
 * den Empfehlungen für Mail-Schaltflächen (Litmus, Chamaileon, ActiveCampaign
 * — „bulletproof buttons") sind drei Dinge daran falsch:
 *
 * 1. **Outlook rendert mit der Word-Engine.** `display:block` auf einem `<a>`
 *    wird dort ignoriert — anklickbar war deshalb nur der TEXT, nicht die
 *    Fläche. Jede Antwort ist jetzt eine Tabellenzelle: Rahmen, Hintergrund
 *    und `padding` an der `<td>` trägt die Word-Engine zuverlässig, und im
 *    Inneren liegen ZWEI Links auf dieselbe Adresse (Kästchen und Text),
 *    sodass in jedem Client beides klickt.
 * 2. **Das Kästchen war ein Schriftzeichen (`&#9744;`).** Ob es erscheint,
 *    entscheidet die Schrift des Empfängers, und seine Größe passt zu keiner
 *    Zeilenhöhe. Jetzt ist es eine 18 px große Zelle mit 2 px Rahmen — also
 *    gezeichnet statt gesetzt, und überall gleich.
 * 3. **Eine Antwortmöglichkeit muss wie eine Schaltfläche aussehen.** Die
 *    Empfehlung für Umfragen in Mails ist überall dieselbe: große
 *    Tippflächen, klarer Kontrast, einspaltig. Deshalb 2 px grüner Rahmen,
 *    weiße Fläche, 16 px halbfette Schrift, 14/18 px Innenabstand (Höhe rund
 *    50 px, im empfohlenen Bereich 42–72 px) und 10 px Abstand dazwischen.
 *
 * Bewusst OHNE VML-`roundrect`: Das bräuchte man für runde Ecken in Outlook.
 * Dort bleiben die Ecken eckig — der einzige sichtbare Unterschied, und er
 * ist den doppelten Markup-Pfad nicht wert.
 */

import { buildHashDeepLink } from './deepLink';

const GREEN = '#86bc25';

/** HTML-Text entschärfen — Fragen und Optionen sind Freitext des Organizers. */
function h(v: string): string {
  return (v || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export interface PollMailOptions {
  appUrl: string;
  eventNumber: number;
  frage: string;
  optionen: string[];
  /** Mehrere Gründe erlaubt — steht als Satz unter den Kästchen. */
  mehrfach: boolean;
  /** DEX wertet die Umfrage ohne Namen aus. Der Satz dazu ist Pflicht, nicht Deko. */
  anonym: boolean;
  isDe: boolean;
}

/** Der Link, der genau eine Antwort setzt. `a` ist der Index in `optionen`. */
export function pollOptionLink(appUrl: string, eventNumber: number, index: number): string {
  return buildHashDeepLink(appUrl, { action: 'umfrage', e: eventNumber, a: index });
}

/** Der Link ohne Vorauswahl — „Ich möchte etwas dazuschreiben". */
export function pollPageLink(appUrl: string, eventNumber: number): string {
  return buildHashDeepLink(appUrl, { action: 'umfrage', e: eventNumber });
}

/**
 * Der Innenteil der Mail. Kommt in `wrapTemplate` — die Kopfzeile, das
 * Event-Bild und die Fußzeile baut der Aufrufer wie bei jeder anderen
 * App-Mail (inkl. `eventHeaderImageOpts`, sonst ist das Kopfbild 180 px
 * breit statt formatfüllend — v30.87).
 */
export function buildPollMailBody(o: PollMailOptions): string {
  const { appUrl, eventNumber, frage, optionen, mehrfach, anonym, isDe } = o;

  const kaestchen = optionen.map((opt, i) => {
    const url = pollOptionLink(appUrl, eventNumber, i);
    return `
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="width:100%;border-collapse:collapse;margin:0 0 10px 0;">
<tr>
<td bgcolor="#ffffff" style="border:2px solid ${GREEN};border-radius:8px;padding:14px 18px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="width:100%;border-collapse:collapse;">
  <tr>
    <td width="22" valign="middle" style="width:22px;">
      <a href="${url}" style="text-decoration:none;">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
          <td width="18" height="18" bgcolor="#ffffff" style="width:18px;height:18px;border:2px solid ${GREEN};border-radius:3px;font-size:0;line-height:0;">&nbsp;</td>
        </tr></table>
      </a>
    </td>
    <td valign="middle" style="padding-left:12px;">
      <a href="${url}" style="color:#2b2b2b;text-decoration:none;font-family:Aptos,'Open Sans',Arial,Helvetica,sans-serif;font-size:16px;line-height:22px;font-weight:600;">${h(opt)}</a>
    </td>
  </tr>
  </table>
</td>
</tr>
</table>`;
  }).join('');

  const hinweisMehrfach = mehrfach
    ? (isDe
      ? 'Mehrfachnennung ist möglich — nach dem ersten Klick kannst du auf der Seite weitere Gründe anhaken.'
      : 'You can pick more than one — after the first click you can tick further reasons on the page.')
    : (isDe
      ? 'Bitte wähle den Grund, der am besten passt.'
      : 'Please pick the reason that fits best.');

  const hinweisAnonym = anonym
    ? (isDe
      ? 'Die Auswertung ist anonym: DEX zeigt den Organisierenden nur, wie oft etwas gewählt wurde, und die Freitexte ohne Absender.'
      : 'The evaluation is anonymous: DEX only shows the organisers how often each answer was picked, plus the free-text comments without a sender.')
    : (isDe
      ? 'Deine Antwort ist für die Organisierenden dieses Events mit deinem Namen sichtbar.'
      : 'Your answer is visible to the organisers of this event together with your name.');

  const FONT = `Aptos,'Open Sans',Arial,Helvetica,sans-serif`;

  return `
<p style="font-family:${FONT};font-size:19px;line-height:26px;font-weight:700;color:#2b2b2b;margin:0 0 6px 0;">${h(frage)}</p>

<p style="font-family:${FONT};font-size:15px;line-height:22px;color:#5a5a5a;margin:0 0 18px 0;">
  ${isDe
    ? 'Tipp auf eine Antwort — das war es schon. Sie ist damit gespeichert, du musst nichts absenden.'
    : 'Tap one answer — that is it. It is saved right away, there is nothing to submit.'}
</p>

${kaestchen}

<p style="font-family:${FONT};font-size:14px;line-height:21px;color:#6b6b6b;margin:18px 0 0 0;">
  ${hinweisMehrfach}
</p>
<p style="font-family:${FONT};font-size:13px;line-height:20px;color:#8a8a8a;margin:8px 0 0 0;">
  ${hinweisAnonym}
</p>

<p style="font-family:${FONT};font-size:14px;line-height:21px;margin:18px 0 0 0;">
  <a href="${pollPageLink(appUrl, eventNumber)}" style="color:${GREEN};font-weight:600;text-decoration:underline;">
    ${isDe ? 'Lieber selbst etwas schreiben? Umfrage öffnen' : 'Rather write something yourself? Open the poll'}
  </a>
</p>`;
}
