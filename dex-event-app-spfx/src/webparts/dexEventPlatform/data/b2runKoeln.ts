/**
 * v26.48: B2Run-Köln-Vorlage.
 *
 * Zentrale Definition für Events, deren Titel „B2Run Köln" enthält:
 *  1. Vorgeschlagene Abfragefelder im Wizard (Schritt „Felder") — per
 *     „Alle übernehmen" in einem Klick ins Event übernehmbar.
 *  2. Exakte Struktur der OFFIZIELLEN B2Run-Köln-Excel für den Export
 *     (Sheet „B2Run Köln <Jahr>", 16 Spalten, Schreibweisen 1:1 aus der
 *     Original-Datei — inkl. „Straße" mit ß, kleingeschriebener Anrede
 *     „männlich/weiblich" und der Spalte „Nordic Walker").
 *
 * Quelle: Deloitte_Teilnehmer_-innen_b2run-koeln-2026.xlsx (offizielle
 * Meldedatei des Veranstalters). Gruppe + private Adresse + Mobilnummer
 * dürfen leer bleiben.
 */
import { CustomField } from '../services/EventService';
import { DeloitteEvent } from '../types';

/** Greift, wenn der Event-Titel „B2Run Köln" enthält (auch „B2Run-Koeln" etc.). */
export function isB2RunKoelnTitle(title: string | undefined | null): boolean {
  return /b2run[\s\-_]*k(ö|oe)ln/i.test(title || '');
}

/** Exakte Header der offiziellen Excel — Reihenfolge und Schreibweise 1:1. */
export const B2RUN_KOELN_HEADERS: string[] = [
  'Nr.',
  'Anrede',
  'Vorname',
  'Nachname',
  'E-Mail',
  'Startblock',
  'Zustimmung AGB & Datenschutzhinweise',
  'Anonym',
  'Gruppe',
  'Straße und Hausnummer (privat)',
  'PLZ (privat)',
  'Stadt (privat)',
  'Mobilnummer',
  'Verwendung Infoservice',
  'Altersklasse',
  'Nordic Walker',
];

/** Exakte Startblock-Werte aus der Original-Datei. */
export const B2RUN_KOELN_STARTBLOCK_DURCHSTARTER = '17:00 Uhr Durchstarter (Orange)';
export const B2RUN_KOELN_STARTBLOCK_FUNSTARTER = '17:00 Uhr Funstarter (Grün)';

/** Einheitswert der Original-Datei. */
export const B2RUN_KOELN_ALTERSKLASSE = 'offene Klasse';

/** DEX-Anrede (Herr/Frau/Divers) → B2Run-Schreibweise (klein). */
export function mapAnredeToB2Run(anrede: string | undefined | null): string {
  const a = (anrede || '').trim().toLowerCase();
  if (a === 'herr' || a === 'männlich') return 'männlich';
  if (a === 'frau' || a === 'weiblich') return 'weiblich';
  if (a === 'divers') return 'divers';
  return '';
}

/** StarterType (Split-Capacity) → exakter Startblock-Text der Excel. */
export function mapStarterTypeToStartblock(starterType: string | undefined | null): string {
  const s = (starterType || '').trim().toLowerCase();
  if (s === 'durchstarter') return B2RUN_KOELN_STARTBLOCK_DURCHSTARTER;
  if (s === 'funstarter') return B2RUN_KOELN_STARTBLOCK_FUNSTARTER;
  return '';
}

/**
 * v27.10 (Refactor): unverändert aus AdminPage.tsx hierher verschoben —
 * einzige Anpassung: `isDe` ist jetzt expliziter Parameter statt Closure-
 * Variable der Komponente.
 *
 * v26.89: B2Run-Köln-Events bekommen einen eigenen, dynamischen Einladungs-
 * text-Vorschlag (bilingual DE + EN) — mit Datum, Ort und Platzzahl aus dem
 * Event. Der Organizer kann ihn wie jeden anderen Entwurf frei überschreiben.
 */
export function buildB2RunKoelnInviteDefaults(ev: DeloitteEvent, appUrl: string, signatureNames: string, isDe: boolean): { subject: string; heading: string; subheading: string; body: string } {
  const start = ev.startDate ? new Date(ev.startDate) : null;
  const validStart = start && !isNaN(start.getTime()) ? start : null;
  const dateDe = validStart ? validStart.toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }) : '';
  const dateEn = validStart ? validStart.toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }) : '';
  const venue = (ev.location || '').trim() || 'RheinEnergieStadion';
  const plaetze = (ev.maxParticipants && ev.maxParticipants > 0) ? ev.maxParticipants : 100;
  const dateLineDe = dateDe || 'Datum folgt';
  const dateLineEn = dateEn || 'date to follow';
  // „DEX App" als grün gestylter Link (statt der langen URL im Klartext).
  const appLinkDe = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">DEX App</a>`;
  const appLinkEn = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">DEX app</a>`;
  const b2runSite = '<a href="https://www.b2run.de" style="color:#86bc25;font-weight:600;">B2RUN Website</a>';
  const de = `
<p>Liebes Team,</p>
<p>es ist so weit: wir haben unseren Standort wieder für den <strong>B2Run Firmenlauf</strong> angemeldet. Es werden vorerst <strong>${plaetze} Startplätze</strong> zur Verfügung stehen. Diese werden nach dem Motto „First come, first run" vergeben.</p>
<p>Die Startgebühren inkl. einer Spende an Menschen für Menschen, die Kosten für ein Laufshirt sowie ein Teamzelt mit einer kleinen Verpflegung nach dem Lauf werden dabei von Deloitte übernommen.</p>
<p style="text-align:center;"><strong>Die Anmeldung ist ab sofort über die ${appLinkDe} möglich.</strong></p>
<p><strong>Wichtige Hinweise:</strong></p>
<ul>
  <li>Falls ihr es nicht unter die ersten ${plaetze} schaffen solltet, meldet euch bitte trotzdem über die App an. Ihr werdet automatisch der Reihe nach auf eine Warteliste gesetzt.</li>
  <li>Falls ihr aus wichtigen Gründen nicht am Lauf teilnehmen könnt, sagt eure Teilnahme bitte frühzeitig über die App („My Events") wieder ab. Die Plätze werden automatisch von der Warteliste – der Reihe nach – vergeben und ihr erhaltet eine automatische E-Mail.</li>
  <li>Bitte meldet euch nur an, wenn ihr auch wirklich am B2RUN teilnehmen könnt und möchtet. <strong>Wir zahlen für jede Anmeldung eine Teilnehmergebühr, die wir im Falle eines No-Shows nicht erstattet bekommen.</strong></li>
</ul>
<p><strong>Infos zum Lauf/Event:</strong></p>
<ul>
  <li>${dateLineDe} am <strong>${venue}</strong></li>
  <li>Beginn der Veranstaltung: 15:00 Uhr</li>
  <li>Teamtreff Deloitte, Startnummernübergabe und Aufwärmen: 16:00 Uhr</li>
  <li>Startzeit Deloitte: 17:00 Uhr</li>
  <li>Distanz: 5,3 km</li>
  <li>Anschließend: Get-Together, Teamfotos und Catering sowie Afterparty im Stadioninnenraum (ab 20:00 Uhr)</li>
</ul>
<p><em>Genauere Infos zu den Zeiten und Treffpunkten werden wir euch Mitte / Ende August mitteilen.</em></p>
<p><strong>Startfelder</strong></p>
<p>Beim B2Run gibt es zwei Startfelder: „Funstarter" und „Durchstarter".</p>
<ul>
  <li>Das <strong>Durchstarter</strong>-Feld ist für schnelle und ambitionierte Läufer:innen gedacht, die „freie Bahn" haben möchten (Richtwerte Männer &lt;4 Min/km, Frauen &lt;5 Min/km).</li>
  <li>Für das <strong>Funstarter</strong>-Feld gibt es keine Richtwerte – hier steht der Laufspaß im Vordergrund.</li>
  <li>Sofern jemand am Durchstarter-Lauf teilnehmen möchte, wählt dies entsprechend bei der Anmeldung aus. Die Laufstrecke für beide Startfelder beträgt 5,3 Kilometer.</li>
</ul>
<p><strong>Laufshirts</strong></p>
<p>Für jede:n Läufer:in gibt es ein Deloitte-Laufshirt. Wählt bitte bei der Anmeldung eure Größe aus.</p>
<p>Weitere Informationen sind auf der ${b2runSite} zu finden.</p>
<p>Bei Fragen wendet euch bitte direkt an unser Gruppenpostfach.</p>
<p>Auf die Plätze, fertig, los.</p>
<p>Mit sportlichen Grüßen<br />${signatureNames}</p>`.trim();
  const en = `
<p>Dear team,</p>
<p>The time has come: we have registered our location for the <strong>B2Run company run</strong> again. For now <strong>${plaetze} starting places</strong> will be available, allocated on a „first come, first run" basis.</p>
<p>The registration fees (incl. a donation to Menschen für Menschen), the cost of a running shirt and a team tent with light refreshments after the run are covered by Deloitte.</p>
<p style="text-align:center;"><strong>Registration is now open via the ${appLinkEn}.</strong></p>
<p><strong>Important information:</strong></p>
<ul>
  <li>If you don't make it into the first ${plaetze}, please register via the app anyway. You will automatically be placed on a waiting list in order of registration.</li>
  <li>If you are unable to take part for important reasons, please cancel your participation early via the app („My Events"). Places are allocated automatically from the waiting list – in order – and you will receive an automatic email.</li>
  <li>Please only register if you can really take part in the B2RUN. <strong>We pay a participation fee for each registration that is not refunded in the event of a no-show.</strong></li>
</ul>
<p><strong>Event details:</strong></p>
<ul>
  <li>${dateLineEn} at <strong>${venue}</strong></li>
  <li>Start of the event: 3:00 p.m.</li>
  <li>Deloitte team meeting, race-number handout and warm-up: 4:00 p.m.</li>
  <li>Deloitte start time: 5:00 p.m.</li>
  <li>Distance: 5.3 km</li>
  <li>Afterwards: get-together, team photos and catering plus after-party inside the stadium (from 8:00 p.m.)</li>
</ul>
<p><em>We will share more detailed information about times and meeting points in mid/late August.</em></p>
<p><strong>Starting fields</strong></p>
<p>B2Run has two starting fields: „Funstarter" and „Durchstarter".</p>
<ul>
  <li>The <strong>Durchstarter</strong> field is intended for fast and ambitious runners who want a „clear track" (guideline times: men &lt;4 min/km, women &lt;5 min/km).</li>
  <li>The <strong>Funstarter</strong> field has no guideline times – the focus here is on having fun.</li>
  <li>If you would like to take part in the Durchstarter run, please select this when registering. The distance for both fields is 5.3 kilometres.</li>
</ul>
<p><strong>Running shirts</strong></p>
<p>Every runner receives a Deloitte running shirt. Please select your size when registering.</p>
<p>Further information can be found on the ${b2runSite}.</p>
<p>If you have any questions, please contact our group mailbox directly.</p>
<p>On your marks, get set, go.</p>
<p>With sporting regards<br />${signatureNames}</p>`.trim();
  // v26.90: Zweisprachig — Hauptsprache oben, die jeweils andere Version unten
  // (per Trennlinie abgesetzt), wie es die B2Run-Kommunikation üblicherweise macht.
  const divider = '<p style="margin:28px 0 20px;border-top:1px solid #d0d0ce;"></p>';
  return {
    subject: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
    heading: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
    subheading: isDe ? `${dateLineDe} · ${plaetze} Startplätze` : `${dateLineEn} · ${plaetze} starting places`,
    body: isDe ? `${de}\n${divider}\n${en}` : `${en}\n${divider}\n${de}`,
  };
}

/**
 * Die vorgeschlagenen Abfragefelder der Vorlage. Deterministische IDs — der
 * Export liest exakt diese IDs aus CustomData. `b2run_startblock`/
 * `b2run_datenschutz`/`b2run_infoservice`/`b2run_mobilnummer` sind bewusst
 * identisch mit den bestehenden B2Run-Katalog-IDs (Render-Spezialfälle in
 * der Anmeldemaske — Datenschutz-Links, Mobilnummer-showIf — greifen damit
 * automatisch auch hier).
 */
export function b2runKoelnTemplateFields(isDe: boolean): CustomField[] {
  return [
    {
      id: 'b2run_geschlecht',
      label: isDe ? 'Geschlecht (laut B2Run-Meldung)' : 'Gender (as per B2Run entry)',
      type: 'select', required: true, visible: true,
      options: ['männlich', 'weiblich', 'divers'],
      helpText: isDe
        ? 'Wird 1:1 in die offizielle B2Run-Meldedatei übernommen (Spalte „Anrede").'
        : 'Copied 1:1 into the official B2Run entry file (column "Anrede").',
    },
    // KEIN Startblock-Abfragefeld: Die Teilnehmenden wählen Durchstarter/
    // Funstarter über den GRUPPEN-SPLIT (Split-Kapazität, Schritt „Kapazität
    // & Sichtbarkeit"). Der Export mappt StarterType automatisch auf die
    // exakten Startblock-Texte der Meldedatei (mapStarterTypeToStartblock).
    {
      id: 'b2run_datenschutz',
      label: isDe ? 'Zustimmung AGB & Datenschutzhinweise (b2run.de)' : 'Consent to T&C & privacy notice (b2run.de)',
      type: 'checkbox', required: true, visible: true,
    },
    {
      id: 'b2run_anonym',
      label: isDe ? 'Anonym starten (Name erscheint nicht in Ergebnislisten)' : 'Run anonymously (name not shown in result lists)',
      type: 'checkbox', required: false, visible: true,
    },
    {
      id: 'b2run_infoservice',
      label: isDe ? 'B2Run-Infoservice nutzen (SMS-Infos am Lauftag)' : 'Use the B2Run info service (SMS updates on race day)',
      type: 'checkbox', required: false, visible: true,
    },
    {
      id: 'b2run_mobilnummer',
      label: isDe ? 'Mobilnummer (für den Infoservice)' : 'Mobile number (for the info service)',
      // PFLICHT, sobald sichtbar (Infoservice = ja): Die Anmeldemaske hat für
      // b2run_mobilnummer zusätzlich eine hartkodierte Validierung — Pflicht
      // genau dann, wenn b2run_infoservice='true' (RegistrationPage).
      type: 'text', required: true, visible: true,
      showIf: { fieldId: 'b2run_infoservice', values: ['true'] },
    },
    // Nordic Walker wird bewusst NICHT abgefragt (Maintainer-Entscheidung) —
    // die Excel-Spalte „Nordic Walker" bleibt im Export erhalten und wird
    // mit „Nein" gefüllt (Default der Original-Vorlage).
  ];
}
