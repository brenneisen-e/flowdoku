/**
 * v31.12 — Die Umfrage-Mail: Antwortmöglichkeiten als anklickbare Kästchen.
 *
 * Eine echte Checkbox in einer Mail gibt es nicht — Outlook entfernt `<form>`
 * beim Rendern, und ein Kästchen ohne Formular hätte niemanden, an den es
 * etwas schicken könnte. Was hier gebaut wird, ist der Weg, den jedes
 * Umfrage-Werkzeug geht: Jede Antwort ist ein LINK, der wie ein Kästchen
 * aussieht. Ein Klick öffnet DEX, die Antwort ist sofort gespeichert.
 *
 * Zwei Dinge, die an dieser Stelle leicht kaputtgehen:
 *
 * 1. **Outlook rendert mit der Word-Engine.** Rahmen an `<span>`, `flex` und
 *    `display:block` auf `<a>` sind dort unzuverlässig. Jede Option ist
 *    deshalb eine eigene EINZELLIGE TABELLE mit Rahmen an der `<td>` — das
 *    trägt in Outlook, Gmail und auf dem Handy gleich.
 * 2. **Das Kästchen-Zeichen `&#9744;` ist ein Schriftzeichen.** Fehlt es in
 *    der Schrift, zeigt Outlook ein leeres Rechteck — also wieder ein
 *    Kästchen. Der Rahmen der Zelle trägt die Bedeutung ohnehin; das Zeichen
 *    ist die Zugabe, nicht die Aussage.
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

  const kaestchen = optionen.map((opt, i) => `
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="width:100%;margin:0 0 8px 0;">
<tr>
<td style="border:1px solid #d5d5d5;border-left:4px solid ${GREEN};background-color:#ffffff;padding:0;">
  <a href="${pollOptionLink(appUrl, eventNumber, i)}" style="display:block;padding:12px 16px;color:#333333;text-decoration:none;font-family:Aptos,Arial,Helvetica,sans-serif;font-size:15px;font-weight:600;">
    <span style="color:${GREEN};font-size:19px;line-height:19px;">&#9744;</span>&nbsp;&nbsp;${h(opt)}
  </a>
</td>
</tr>
</table>`).join('');

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

  return `
<p style="font-size:16px;line-height:1.5;">${h(frage)}</p>

<p style="font-size:14px;color:#666666;margin:0 0 14px 0;">
  ${isDe
    ? 'Ein Klick genügt — deine Antwort ist sofort gespeichert.'
    : 'One click is enough — your answer is saved right away.'}
</p>

${kaestchen}

<p style="font-size:13px;color:#666666;line-height:1.5;margin:14px 0 0 0;">
  ${hinweisMehrfach}<br>
  ${hinweisAnonym}
</p>

<p style="font-size:13px;margin:16px 0 0 0;">
  <a href="${pollPageLink(appUrl, eventNumber)}" style="color:${GREEN};font-weight:600;">
    ${isDe ? 'Umfrage öffnen und etwas dazuschreiben' : 'Open the poll and add a comment'}
  </a>
</p>`;
}
