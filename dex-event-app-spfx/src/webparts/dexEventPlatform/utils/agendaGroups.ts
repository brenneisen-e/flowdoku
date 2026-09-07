/**
 * v30.94: Programmpunkte in Cluster gruppieren — EINE Regel für Wizard,
 * Anmeldeseite, Meine Events, Check-in, Anwesenheit und den Mail-Platzhalter.
 *
 * Nutzer-Ansage 07.09.2026: „es sollte sowas geben wie Cluster und dann als
 * Vorschlag Day 1, und dann kann man die Programmpunkte da sinnvoll
 * zusammenführen." Bis dahin trug jeder Titel den Behelf „Day 1 - …".
 *
 * Regel: Ein Punkt gehört zum Cluster `cluster` (frei benannt). Ohne Cluster
 * gruppiert das Datum. Zwei Cluster dürfen dasselbe Datum haben (Track A/B),
 * ein Cluster darf über Tage laufen — die Gruppe zeigt dann den Bereich.
 * Sortiert wird nach dem frühesten Datum+Zeit der Gruppe, dann nach dem
 * ersten Vorkommen.
 */
import { AgendaItem } from '../types';

export interface AgendaGroup {
  /** Stabil je Gruppe: `c:<name>` oder `d:<datum>` (leer bei „ohne Datum"). */
  key: string;
  /** Der Cluster-Name; leer bei reiner Datums-Gruppe. */
  cluster: string;
  /** Häufigstes Datum der Gruppe (Header-Datum). */
  date: string;
  /** Alle Daten der Gruppe, sortiert — bei > 1 läuft der Cluster über Tage. */
  dates: string[];
  items: AgendaItem[];
}

export function agendaSortKey(a: AgendaItem): string {
  return (a.date || '9999-99-99') + 'T' + (a.time || '99:99');
}

export function sortAgenda(items: AgendaItem[]): AgendaItem[] {
  return items.slice().sort((a, b) => agendaSortKey(a).localeCompare(agendaSortKey(b)));
}

export function groupKeyOf(a: AgendaItem): string {
  const c = (a.cluster || '').trim();
  return c ? `c:${c}` : `d:${a.date || ''}`;
}

export function agendaGroups(items: AgendaItem[]): AgendaGroup[] {
  const map = new Map<string, AgendaGroup>();
  const order: string[] = [];
  for (const it of sortAgenda(items)) {
    const key = groupKeyOf(it);
    let g = map.get(key);
    if (!g) {
      g = { key, cluster: (it.cluster || '').trim(), date: it.date || '', dates: [], items: [] };
      map.set(key, g);
      order.push(key);
    }
    g.items.push(it);
  }
  const groups = order.map(k => map.get(k) as AgendaGroup);
  for (const g of groups) {
    const counts: Record<string, number> = {};
    g.items.forEach(it => { if (it.date) counts[it.date] = (counts[it.date] || 0) + 1; });
    g.dates = Object.keys(counts).sort();
    let best = ''; let n = -1;
    g.dates.forEach(d => { if (counts[d] > n) { n = counts[d]; best = d; } });
    g.date = best;
  }
  // Gruppen nach frühestem Punkt; „ohne Datum" ans Ende (sortAgenda tut das schon).
  return groups.sort((a, b) => agendaSortKey(a.items[0]).localeCompare(agendaSortKey(b.items[0])) || order.indexOf(a.key) - order.indexOf(b.key));
}

/** Vorschlag für den Namen einer neuen Gruppe: „Tag N" / „Day N". */
export function suggestClusterName(index: number, isDe: boolean): string {
  return `${isDe ? 'Tag' : 'Day'} ${index + 1}`;
}

/** Anzeigename einer Gruppe: Cluster, sonst der Datums-Vorschlag. */
export function groupLabel(g: AgendaGroup, index: number, isDe: boolean): string {
  return g.cluster || suggestClusterName(index, isDe);
}

/** „Mo., 12.10.2026" — bzw. „Mo., 12.10. – Mi., 14.10.2026" bei mehreren Tagen. */
export function groupDateLabel(g: AgendaGroup, isDe: boolean, withYear: boolean = true): string {
  const loc = isDe ? 'de-DE' : 'en-GB';
  const fmt = (ymd: string, year: boolean): string => {
    const d = new Date(ymd + 'T00:00:00');
    if (isNaN(d.getTime())) return '';
    return d.toLocaleDateString(loc, { weekday: 'short', day: '2-digit', month: '2-digit', ...(year ? { year: 'numeric' } : {}) });
  };
  if (g.dates.length === 0) return isDe ? 'ohne Datum' : 'no date';
  if (g.dates.length === 1) return fmt(g.dates[0], withYear);
  return `${fmt(g.dates[0], false)} – ${fmt(g.dates[g.dates.length - 1], withYear)}`;
}

/** Entfernt ein Cluster-Präfix („Day 1 - Welcome" → „Welcome"), wenn der Titel
 *  mit dem Cluster-Namen und einem Trenner beginnt. */
export function stripClusterPrefix(title: string, cluster: string): string {
  const c = (cluster || '').trim();
  if (!c) return title;
  const esc = c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`^\\s*${esc}\\s*[-–—:|·]\\s*`, 'i');
  return title.replace(re, '');
}

/** Wie viele Titel der Gruppe ein Cluster-Präfix tragen (für den Knopf). */
export function countClusterPrefixed(g: AgendaGroup): number {
  if (!g.cluster) return 0;
  return g.items.filter(it => stripClusterPrefix(it.title || '', g.cluster) !== (it.title || '')).length;
}

/** Der nächste Name in einer Zahlenreihe: „Tag 1" → „Tag 2", sonst „<Name> 2". */
export function nextClusterName(name: string): string {
  const m = /^(.*?)(\d+)\s*$/.exec(name.trim());
  if (m) return `${m[1]}${parseInt(m[2], 10) + 1}`;
  return name.trim() ? `${name.trim()} 2` : '';
}
