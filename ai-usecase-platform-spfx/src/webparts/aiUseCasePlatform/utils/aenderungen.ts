/**
 * Was sich an einem Use Case geändert hat — als Text fürs Protokoll.
 *
 * Bis v1.2 stand im Protokoll bei jedem Speichern `Object.keys(uc).join(', ')`.
 * Die Pflegeseite reicht beim Speichern den GANZEN Entwurf durch, also lautete
 * jede Zeile „id, titel, kurzbeschreibung, beschreibung, bereich, …" — auch
 * wenn jemand nur einen Link nachgetragen hatte. Ein Protokoll, in dem jede
 * Zeile dasselbe sagt, beantwortet nicht die Frage, wegen der man es aufmacht:
 * „Wer hat den Deployment-Link geändert?".
 *
 * Deshalb wird gegen den Stand VOR dem Speichern verglichen. Kurze Werte
 * (Titel, Status, Bewertung) stehen als „alt → neu" da; lange Texte und Links
 * nur mit ihrem Namen — ein URL-Wechsel gehört ins Protokoll als Ereignis,
 * nicht als zweihundert Zeichen, die niemand liest.
 */

import { UseCase } from '../types';

/**
 * Speicherwerte, die in der Oberfläche anders heißen. Im Protokoll stand
 * „Aufruf: fenster → eingebettet" und „Status: InArbeit → Live" — die Werte der
 * Spalte statt der Wörter, die die Pflegeseite zeigt (Sichtprüfung, 29.09.2026).
 */
const ANZEIGE: { [wert: string]: string } = {
  InArbeit: 'In Arbeit',
  fenster: 'neues Fenster',
};

/** Für das Protokoll lesbar: leer wird zum Gedankenstrich. */
const kurz = (v: string | number | undefined): string => {
  const s = v === undefined || v === null ? '' : String(v);
  return s === '' ? '—' : (ANZEIGE[s] || s);
};

export function geaendertText(alt: UseCase | undefined, neu: Partial<UseCase>): string {
  // Ohne Vergleichsstand ehrlich sagen, was geschrieben wurde, statt eine
  // Änderung zu erfinden.
  if (!alt) return Object.keys(neu).join(', ');

  const teile: string[] = [];
  const wechsel = (label: string, a: string | number | undefined, n: string | number | undefined): void => {
    if (String(a === undefined || a === null ? '' : a) !== String(n === undefined || n === null ? '' : n)) {
      teile.push(`${label}: ${kurz(a)} → ${kurz(n)}`);
    }
  };
  const nur = (label: string, a: string | undefined, n: string | undefined): void => {
    if ((a || '') !== (n || '')) teile.push(label);
  };
  const liste = (label: string, a: string[], n: string[]): void => {
    if (a.join(';') !== n.join(';')) teile.push(label);
  };

  if (neu.titel !== undefined) wechsel('Titel', alt.titel, neu.titel);
  if (neu.kurzbeschreibung !== undefined) nur('Kurzbeschreibung', alt.kurzbeschreibung, neu.kurzbeschreibung);
  if (neu.beschreibung !== undefined) nur('Beschreibung', alt.beschreibung, neu.beschreibung);
  if (neu.bereich !== undefined) wechsel('Bereich', alt.bereich, neu.bereich);
  if (neu.status !== undefined) wechsel('Status', alt.status, neu.status);
  if (neu.salesRelevanz !== undefined) wechsel('Sales-Relevanz', alt.salesRelevanz, neu.salesRelevanz);
  if (neu.machbarkeit !== undefined) wechsel('Machbarkeit', alt.machbarkeit, neu.machbarkeit);
  if (neu.demoTauglichkeit !== undefined) wechsel('Demo-Tauglichkeit', alt.demoTauglichkeit, neu.demoTauglichkeit);
  if (neu.aufrufArt !== undefined) wechsel('Aufruf', alt.aufrufArt, neu.aufrufArt);
  if (neu.ressourcen) {
    nur('Deployment-Link', alt.ressourcen.deployment, neu.ressourcen.deployment);
    nur('Source Code', alt.ressourcen.sourceCode, neu.ressourcen.sourceCode);
    nur('Deployment-Guide', alt.ressourcen.deploymentGuide, neu.ressourcen.deploymentGuide);
    nur('Wiki', alt.ressourcen.wiki, neu.ressourcen.wiki);
    nur('Video', alt.ressourcen.video, neu.ressourcen.video);
  }
  if (neu.bildUrl !== undefined) nur('Bild', alt.bildUrl, neu.bildUrl);
  if (neu.reihenfolge !== undefined) wechsel('Reihenfolge', alt.reihenfolge, neu.reihenfolge);
  if (neu.betreuerEmails !== undefined) liste('Betreuer', alt.betreuerEmails, neu.betreuerEmails);
  if (neu.schlagworte !== undefined) liste('Schlagworte', alt.schlagworte, neu.schlagworte);

  return teile.join(' · ');
}

/**
 * Nur die Felder, die sich gegenüber `alt` WIRKLICH geändert haben.
 *
 * Die Pflegeseite reicht beim Speichern den ganzen Entwurf durch — den Stand
 * von dem Moment, in dem der Dialog aufging. Schreibt man ihn komplett zurück,
 * überschreibt man, was jemand in der Zwischenzeit geändert hat (Lost Update):
 * Person A öffnet einen Use Case, Person B tauscht das Bild, A ändert nur den
 * Titel — und die Kachel zeigt danach wieder das alte, längst gelöschte Bild
 * (Review 29.09.2026). Mit dem Vergleich geht nur A's Titel hinaus.
 *
 * Ohne Vergleichsstand (`alt` fehlt) geht alles hinaus. `bildUrl` bleibt
 * ausdrücklich draußen: Über das Bild entscheidet `saveUseCase` allein.
 */
export function nurGeaendertes(alt: UseCase | undefined, neu: Partial<UseCase>): Partial<UseCase> {
  const roh: Partial<UseCase> = { ...neu };
  delete roh.bildUrl;
  if (!alt) return roh;

  const aus: { [k: string]: unknown } = {};
  const gleich = (a: unknown, n: unknown): boolean => {
    if (Array.isArray(a) && Array.isArray(n)) return a.join(';') === n.join(';');
    return String(a === undefined || a === null ? '' : a) === String(n === undefined || n === null ? '' : n);
  };
  // Nur, was `toRow` überhaupt schreibt — sonst zählte auch ein anderer `geaendertAm` des
  // Entwurfs als Änderung, und es ginge ein leerer MERGE hinaus.
  const SCHREIBBAR = ['titel', 'kurzbeschreibung', 'beschreibung', 'bereich', 'status', 'salesRelevanz', 'machbarkeit',
    'demoTauglichkeit', 'aufrufArt', 'reihenfolge', 'betreuerEmails', 'betreuerNamen', 'schlagworte'];
  SCHREIBBAR.forEach(k => {
    const n = (roh as { [k: string]: unknown })[k];
    if (n === undefined) return;
    if (!gleich((alt as unknown as { [k: string]: unknown })[k], n)) aus[k] = n;
  });
  if (roh.ressourcen) {
    const r = roh.ressourcen;
    const a = alt.ressourcen;
    if (!gleich(a.sourceCode, r.sourceCode) || !gleich(a.deployment, r.deployment) || !gleich(a.deploymentGuide, r.deploymentGuide)
      || !gleich(a.wiki, r.wiki) || !gleich(a.video, r.video)) {
      aus.ressourcen = r;
    }
  }
  return aus as Partial<UseCase>;
}
