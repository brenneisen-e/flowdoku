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

/** Für das Protokoll lesbar: leer wird zum Gedankenstrich. */
const kurz = (v: string | number | undefined): string => {
  const s = v === undefined || v === null ? '' : String(v);
  return s === '' ? '—' : s;
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
