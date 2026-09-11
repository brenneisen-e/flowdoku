/**
 * v31.23 — Feedback der Organizer an das DEX-Team.
 *
 * Nutzer-Anforderung 11.09.2026: „wenn ein Event vorbei ist (da gibts schon
 * ne automatische Mail) soll in der Mail auch über einen Deeplink-Button nach
 * Feedback gefragt werden. Im Feedback dann Fragen über Checkboxen und
 * Freitext, was gut lief und was ggf. noch zu verbessern ist. Und das geht
 * dann danach als automatische Mail an dex.event@deloitte.de nach Absendung."
 *
 * ## Wessen Feedback, und worüber
 *
 * Die „schon vorhandene automatische Mail" ist die Nachbereitungs-Mail an die
 * **Organizer** („Danke für euer Event!", `context/actions/autoMails.ts`,
 * seit v24.2) — eine automatische Mail an Teilnehmer nach dem Event gibt es
 * nicht. Und `dex.event@deloitte.de` ist das Postfach des DEX-Teams, nicht
 * das des Events. Beides zusammen ergibt genau eine sinnvolle Lesart, und so
 * ist es gebaut: **Der Organizer sagt dem DEX-Team, wie das Organisieren mit
 * DEX gelaufen ist.**
 *
 * Die Fragen zielen deshalb auf die Arbeit mit der App, nicht auf das Event
 * selbst. Wären sie auf das Event gemünzt („war das Catering gut"), ginge die
 * Antwort an eine Adresse, die damit nichts anfangen kann.
 *
 * ## Warum die Antworten fest verdrahtet sind
 *
 * Ein Katalog in einer SharePoint-Liste wäre pflegbar — und genau deshalb
 * falsch: Er müsste provisioniert, berechtigt und gepflegt werden, und beim
 * ersten Lesefehler stünde die Seite ohne Fragen da. Der Katalog ändert sich
 * ein-, zweimal im Jahr; das ist ein Release wert, keine Liste.
 *
 * ## Was hier NICHT passiert
 *
 * Es wird **nichts gespeichert**. Die Anforderung ist eine Mail, und eine
 * zweite Wahrheit in einer Liste, die niemand ausliest, wäre Ballast. Wenn
 * das DEX-Team später auswerten will, ist das der nächste Schritt — dann
 * aber bewusst, mit Aufbewahrungsfrist und einer Ansicht, die sie zeigt.
 */

/** Das Postfach des DEX-Teams. Steht hier EINMAL. */
export const DEX_FEEDBACK_TO = 'dex.event@deloitte.de';

export interface FeedbackOption {
  /** Stabiler Schlüssel — landet im Mailtext, nicht in der Oberfläche. */
  id: string;
  de: string;
  en: string;
}

/**
 * Was gut lief. Bewusst konkret statt „die App war gut": Eine Rückmeldung,
 * die keine Stelle benennt, ist für das Team nicht verwertbar.
 */
export const FEEDBACK_GUT: FeedbackOption[] = [
  { id: 'anlegen', de: 'Das Event anzulegen ging schnell und war verständlich', en: 'Creating the event was quick and clear' },
  { id: 'anmeldeseite', de: 'Die Anmeldeseite kam bei den Teilnehmern gut an', en: 'Attendees got on well with the registration page' },
  { id: 'mails', de: 'Mails und Outlook-Termine kamen zuverlässig an', en: 'Emails and Outlook invitations arrived reliably' },
  { id: 'checkin', de: 'Der Check-in am Eingang lief reibungslos', en: 'Check-in at the door ran smoothly' },
  { id: 'teilnehmerliste', de: 'Teilnehmerliste und Export haben gepasst', en: 'The attendee list and the export were fine' },
  { id: 'warteliste', de: 'Warteliste und Nachrücken haben funktioniert', en: 'Waiting list and promotion worked' },
  { id: 'subevents', de: 'Mehrere Termine unter einem Event zu führen hat funktioniert', en: 'Running several dates under one event worked' },
  { id: 'support', de: 'Ich habe schnell Hilfe bekommen, als ich sie brauchte', en: 'I got help quickly when I needed it' },
];

/**
 * Was zu verbessern ist. Die erste Antwort („nichts") steht bewusst oben:
 * Wer zufrieden war, soll das mit einem Klick sagen können, statt das Feld
 * leer zu lassen — leer und zufrieden sind sonst nicht zu unterscheiden.
 */
export const FEEDBACK_BESSER: FeedbackOption[] = [
  { id: 'nichts', de: 'Nichts — hat für mich gepasst', en: 'Nothing — it worked for me' },
  { id: 'assistent', de: 'Der Event-Assistent ist zu lang oder unübersichtlich', en: 'The event wizard is too long or hard to follow' },
  { id: 'gefunden', de: 'Ich habe eine Einstellung nicht gefunden', en: 'I could not find a setting' },
  { id: 'mails', de: 'Mails oder Outlook-Termine kamen nicht oder zu spät an', en: 'Emails or Outlook invitations were missing or late' },
  { id: 'checkin', de: 'Der Check-in war umständlich', en: 'Check-in was cumbersome' },
  { id: 'zahlen', de: 'Zahlen in der App und in der Liste haben sich widersprochen', en: 'Numbers in the app and in the list contradicted each other' },
  { id: 'langsam', de: 'Die App war langsam', en: 'The app was slow' },
  { id: 'rechte', de: 'Es gab Probleme mit Berechtigungen', en: 'There were permission problems' },
  { id: 'fehlt', de: 'Mir fehlt eine Funktion', en: 'A feature is missing' },
];

const esc = (s: string): string => (s || '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Freitext mit Zeilenumbrüchen als HTML — ohne `innerHTML`-Falle. */
function absatz(text: string): string {
  const t = (text || '').trim();
  if (!t) return '';
  return esc(t).split(/\n{2,}/)
    .map(p => `<p style="margin:0 0 10px;">${p.replace(/\n/g, '<br>')}</p>`)
    .join('');
}

function liste(ids: string[], katalog: FeedbackOption[]): string {
  if (ids.length === 0) return '<p style="margin:0 0 12px;color:#767676;">— keine Angabe —</p>';
  const items = ids
    .map(id => katalog.filter(o => o.id === id)[0])
    .filter(Boolean)
    .map(o => `<li style="margin:0 0 4px;">${esc(o!.de)}</li>`)
    .join('');
  return `<ul style="margin:0 0 12px;padding-left:20px;">${items}</ul>`;
}

export interface FeedbackEingabe {
  gut: string[];
  besser: string[];
  freitextGut: string;
  freitextBesser: string;
}

/**
 * Der Mailtext an das DEX-Team.
 *
 * Absichtlich **nur deutsch**: Empfänger ist EIN Postfach eines deutschen
 * Teams. Die Oberfläche, auf der der Organizer antwortet, ist zweisprachig —
 * die Auswertung dahinter braucht das nicht, und zwei Sprachen im selben
 * Postfach machen das Sortieren schwerer, nicht leichter. Deshalb landen die
 * ANGEHAKTEN Antworten über ihre `id` immer als deutscher Text in der Mail,
 * auch wenn der Organizer die App auf Englisch bedient hat.
 */
export function buildFeedbackMailInner(
  eingabe: FeedbackEingabe,
  ev: { title: string; id: string; startDate?: string },
  von: { name: string; email: string },
): string {
  const datum = ev.startDate ? new Date(ev.startDate) : null;
  const datumTxt = datum && !isNaN(datum.getTime())
    ? datum.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })
    : '';
  return `
    <p style="margin:0 0 12px;">Ein Organizer hat nach seinem Event Rückmeldung zu DEX gegeben.</p>
    <table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 0 18px;font-size:14px;">
      <tr><td style="padding:2px 12px 2px 0;color:#767676;">Event</td><td style="padding:2px 0;"><strong>${esc(ev.title)}</strong></td></tr>
      ${datumTxt ? `<tr><td style="padding:2px 12px 2px 0;color:#767676;">Datum</td><td style="padding:2px 0;">${datumTxt}</td></tr>` : ''}
      <tr><td style="padding:2px 12px 2px 0;color:#767676;">Event-Id</td><td style="padding:2px 0;">${esc(ev.id)}</td></tr>
      <tr><td style="padding:2px 12px 2px 0;color:#767676;">Von</td><td style="padding:2px 0;">${esc(von.name || von.email)} &lt;${esc(von.email)}&gt;</td></tr>
    </table>

    <p style="margin:0 0 6px;font-weight:700;color:#86bc25;">Was gut lief</p>
    ${liste(eingabe.gut, FEEDBACK_GUT)}
    ${absatz(eingabe.freitextGut)}

    <p style="margin:16px 0 6px;font-weight:700;color:#ed8b00;">Was wir verbessern können</p>
    ${liste(eingabe.besser, FEEDBACK_BESSER)}
    ${absatz(eingabe.freitextBesser)}

    <p style="margin:20px 0 0;font-size:12px;color:#767676;">
      Diese Mail hat DEX automatisch verschickt, nachdem der Organizer das Feedback-Formular
      abgeschickt hat. Antworten geht direkt an die Absenderadresse oben.
    </p>`;
}

/** Betreff — kurz, mit Event, damit das Postfach sortierbar bleibt. */
export function feedbackMailSubject(eventTitle: string): string {
  return `DEX-Feedback: ${eventTitle}`;
}

/** Hat die Person überhaupt etwas gesagt? Leeres Formular = nichts senden. */
export function feedbackHatInhalt(e: FeedbackEingabe): boolean {
  return e.gut.length > 0 || e.besser.length > 0
    || !!e.freitextGut.trim() || !!e.freitextBesser.trim();
}
