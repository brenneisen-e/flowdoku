/**
 * v30.95: Platzhalter {{Programm}} — das Programm des Events als HTML-Block
 * für Mail-Vorlagen und den Outlook-Termin.
 *
 * Bis dahin stand im Wizard-Tooltip: „die Agenda landet nicht automatisch im
 * Outlook-Termin". Jetzt entscheidet der Organizer: Wo {{Programm}} steht,
 * steht die Tabelle — nach Clustern (utils/agendaGroups), je Zeile Zeit,
 * Titel, Raum. Kein Programm → der Platzhalter verschwindet ersatzlos.
 *
 * Zwei Dinge, die hier absichtlich anders sind als bei Text-Platzhaltern:
 *
 *  1. Der Wert ist fertiges HTML. `buildEmailFromTemplate` setzt ihn deshalb
 *     roh ein (RAW_HTML_KEYS, wie NewLeadBlock); für den Outlook-Body, der
 *     über `replacePlaceholders` läuft (escapet alles), gibt es
 *     `applyProgramPlaceholder` als zweiten Schritt NACH dem Ersetzen.
 *
 *  2. Der Outlook-Body wird beim Speichern GEBACKEN und beim Laden wieder
 *     in den Editor geladen (Lehre aus v30.74/v30.75: {{Organizer}} wuchs
 *     bei jedem Save). Deshalb steht die Tabelle zwischen Markern
 *     (<!--dex-programm-->…<!--/dex-programm-->), und `reinsertProgram-
 *     Placeholder` macht daraus beim Laden wieder {{Programm}}. Ändert sich
 *     das Programm, ist die Tabelle im Termin beim nächsten Save aktuell —
 *     nicht eingefroren und nicht verdoppelt.
 */
import { AgendaItem } from '../types';
import { agendaGroups, groupLabel, groupDateLabel } from './agendaGroups';

export const PROGRAM_PH = '{{Programm}}';
const OPEN = '<!--dex-programm-->';
const CLOSE = '<!--/dex-programm-->';

const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Tabelle aus den Programmpunkten; leer, wenn es keine gibt. */
export function buildProgramHtml(items: AgendaItem[] | undefined | null, lang: string, termPlural?: string): string {
  const list = (items || []).filter(a => a && (a.title || a.time));
  if (list.length === 0) return '';
  const isDe = (lang || 'EN').toUpperCase() !== 'EN';
  const groups = agendaGroups(list);
  const heading = (termPlural || '').trim() || (isDe ? 'Programm' : 'Programme');
  let rows = '';
  groups.forEach((g, gi) => {
    if (groups.length > 1 || g.cluster) {
      rows += `<tr><td colspan="3" style="padding:12px 0 4px;font-weight:700;color:#4a7c1f;font-size:14px;">${esc(groupLabel(g, gi, isDe))}`
        + (g.dates.length ? ` <span style="font-weight:400;color:#777;">· ${esc(groupDateLabel(g, isDe))}</span>` : '')
        + '</td></tr>';
    } else if (g.dates.length) {
      rows += `<tr><td colspan="3" style="padding:12px 0 4px;font-weight:700;color:#4a7c1f;font-size:14px;">${esc(groupDateLabel(g, isDe))}</td></tr>`;
    }
    g.items.forEach(it => {
      const time = `${it.time || ''}${it.endTime ? `–${it.endTime}` : ''}`;
      rows += '<tr>'
        + `<td style="padding:4px 10px 4px 0;white-space:nowrap;color:#555;vertical-align:top;">${esc(time)}</td>`
        + `<td style="padding:4px 10px 4px 0;font-weight:600;vertical-align:top;">${esc(it.title || '')}`
        + (it.description ? `<div style="font-weight:400;color:#666;font-size:13px;">${esc(it.description)}</div>` : '')
        + '</td>'
        + `<td style="padding:4px 0;color:#777;white-space:nowrap;vertical-align:top;">${esc(it.location || '')}</td>`
        + '</tr>';
    });
  });
  return `${OPEN}<div style="margin:14px 0;"><div style="font-weight:700;font-size:15px;margin-bottom:2px;">${esc(heading)}</div>`
    + `<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px;line-height:1.4;">${rows}</table></div>${CLOSE}`;
}

/** {{Programm}} roh durch den Block ersetzen — NACH replacePlaceholders. */
export function applyProgramPlaceholder(html: string, programHtml: string): string {
  if (!html || html.indexOf(PROGRAM_PH) < 0) return html;
  return html.split(PROGRAM_PH).join(programHtml || '');
}

/** Beim Laden: gebackenen Block wieder zum Platzhalter machen. */
export function reinsertProgramPlaceholder(body: string): string {
  if (!body || body.indexOf(OPEN) < 0) return body;
  const re = new RegExp(`${OPEN.replace(/[-[\]/{}()*+?.\\^$|]/g, '\\$&')}[\\s\\S]*?${CLOSE.replace(/[-[\]/{}()*+?.\\^$|]/g, '\\$&')}`, 'g');
  return body.replace(re, PROGRAM_PH);
}
