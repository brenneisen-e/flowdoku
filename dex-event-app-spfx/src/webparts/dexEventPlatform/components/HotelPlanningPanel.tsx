/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { EventService, SPRegistration } from '../services/EventService';
import { DeloitteEvent, DexHotel, DexHotelStay } from '../types';
import { HtmlEditorModal } from './HtmlEditorModal';
import { wrapTemplate, replacePlaceholders } from '../services/EmailTemplates';
import { collectCcEmailsFromFields } from '../context/EventContext';
import HotelImportModal, { IHotelImportResultRow } from './HotelImportModal';

/**
 * HotelPlanningPanel (v28.39)
 * ---------------------------
 * Hotel-Zuordnung für ein bestehendes Event im Organizer Center.
 *
 * Datenhaltung bewusst zweigeteilt:
 *  - **Stammdaten** (Hotels, Zeitraum-Vorlagen, Freigabe-Schalter) hängen als
 *    kleines JSON am Event (Piggybacks `_hotels`, `_hotelStays`,
 *    `_hotelVisible`). Sie werden an EINER Stelle gepflegt.
 *  - **Zuordnung pro Person** steht in der Teilnehmerliste (Spalten `Hotel`,
 *    `HotelFrom`, `HotelTo`). Damit läuft sie in jeden bestehenden Export mit,
 *    bleibt bei Umbenennungen stabil und bläht den Event-Datensatz nicht auf
 *    (die 2-MB-Grenze aus v28.31 war teuer genug).
 *
 * Die Nächte ergeben sich aus An-/Abreisedatum — nur so ist die Belegung je
 * Nacht auswertbar, und genau die braucht das Hotel fürs Abrufkontingent.
 */

export interface IHotelPlanningPanelProps {
  event: DeloitteEvent;
  registrations: SPRegistration[];
  isDe: boolean;
  /** Teilnehmerliste neu laden (nach dem Schreiben einer Zuordnung). */
  onReloadRegistrations: () => void | Promise<void>;
  /** Events neu laden (nach dem Ändern der Stammdaten). */
  onReloadEvents: () => void | Promise<void>;
  /** v28.51: Sub-Events der Klammer — fuer „alle aus <Sub-Event> in ein Hotel". */
  childEvents?: DeloitteEvent[];
  showAlert: (msg: string, opts?: { variant?: 'success' | 'error' | 'info' }) => void;
  confirmDialog: (msg: string, opts?: { confirmLabel?: string; danger?: boolean }) => Promise<boolean>;
}

/** Aktive Stati — nur wer wirklich kommt, braucht ein Bett. */
const ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

const uid = (prefix: string): string =>
  `${prefix}${Date.now().toString(36)}${Math.floor(Math.random() * 1e6).toString(36)}`;

/** ISO-DateTime → 'YYYY-MM-DD' (leer bleibt leer). */
const toDay = (iso?: string): string => {
  if (!iso) return '';
  const s = String(iso);
  return s.length >= 10 ? s.substring(0, 10) : '';
};

/** 'YYYY-MM-DD' → ISO-DateTime für SharePoint. */
const toIso = (day: string): string => (day ? `${day}T00:00:00Z` : '');

/** Nächte zwischen zwei Tagen. Negative/ungültige Spannen ergeben 0. */
const nightsBetween = (from: string, to: string): number => {
  if (!from || !to) return 0;
  const a = Date.parse(`${from}T00:00:00Z`);
  const b = Date.parse(`${to}T00:00:00Z`);
  if (isNaN(a) || isNaN(b) || b <= a) return 0;
  return Math.round((b - a) / 86400000);
};

const addDays = (day: string, n: number): string => {
  const t = Date.parse(`${day}T00:00:00Z`);
  if (isNaN(t)) return day;
  return new Date(t + n * 86400000).toISOString().substring(0, 10);
};

const fmtDay = (day: string, isDe: boolean): string => {
  if (!day) return '—';
  const t = Date.parse(`${day}T00:00:00Z`);
  if (isNaN(t)) return day;
  return new Date(t).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', {
    weekday: 'short', day: '2-digit', month: '2-digit',
  });
};

export const HotelPlanningPanel: React.FC<IHotelPlanningPanelProps> = (props: IHotelPlanningPanelProps) => {
  const { event, registrations, isDe, childEvents, onReloadRegistrations, onReloadEvents, showAlert, confirmDialog } = props;

  const svc = React.useMemo<EventService | null>(() => {
    try {
      const ctx = (window as any).__dexSpfxContext;
      return ctx ? new EventService(ctx) : null;
    } catch { return null; }
  }, []);

  const hotels: DexHotel[] = React.useMemo(() => event.hotels || [], [event.hotels]);
  const stays: DexHotelStay[] = React.useMemo(() => event.hotelStays || [], [event.hotelStays]);
  const visible = !!event.hotelVisibleToAttendees;

  const [busy, setBusy] = React.useState('');
  const [selected, setSelected] = React.useState<Set<number>>(new Set());
  const [showRoster, setShowRoster] = React.useState(true);
  const [filterHotel, setFilterHotel] = React.useState<string>('__all');
  // v28.48: Suche, Sortierung und Hotel-Wunsch-Filter fuer die Personenliste.
  const [search, setSearch] = React.useState('');
  const [sortKey, setSortKey] = React.useState<'name' | 'first' | 'loc' | 'comp' | 'wish' | 'hotel'>('name');
  const [sortAsc, setSortAsc] = React.useState(true);
  const [hideNoWish, setHideNoWish] = React.useState(false);
  // v28.48: Fortschritt der Massenzuordnung — jede Person ist ein eigener
  // Schreibvorgang, bei 300 Zeilen dauert das spuerbar.
  const [bulkProgress, setBulkProgress] = React.useState<{ done: number; total: number } | null>(null);
  // v28.49: Import einer bestehenden Hotel-Liste.
  // v28.51: Ganzes Sub-Event in EIN Hotel. Bei einer Klammer sitzt die
  // Hotel-Zuordnung auf den Schattenzeilen der Klammer (eine Zeile je Person) —
  // wer zu welchem Sub-Event gehoert, steht aber in der Teilnehmerliste des
  // Sub-Events. Die Mail-Adressen holen wir deshalb bei Bedarf dort und cachen
  // sie, statt beim Oeffnen des Panels alle Sub-Listen zu laden.
  const [subPick, setSubPick] = React.useState('');
  const [subHotelPick, setSubHotelPick] = React.useState('');
  const [subEmails, setSubEmails] = React.useState<Record<string, string[]>>({});
  const [subLoading, setSubLoading] = React.useState(false);

  const loadSubEmails = async (child: DeloitteEvent): Promise<string[]> => {
    if (subEmails[child.id]) return subEmails[child.id];
    if (!svc || !child.subsiteUrl) return [];
    setSubLoading(true);
    let list: string[] = [];
    try {
      const regs = await svc.getAllRegistrations(child.subsiteUrl);
      list = regs
        .filter(r => ACTIVE_STATI.indexOf(r.Status || '') >= 0)
        .map(r => (r.ParticipantEmail || '').trim().toLowerCase())
        .filter(Boolean);
    } catch { /* nicht lesbar → leere Liste, Meldung unten */ }
    setSubEmails(prev => ({ ...prev, [child.id]: list }));
    setSubLoading(false);
    return list;
  };

  const assignWholeSub = async (): Promise<void> => {
    const child = (childEvents || []).filter(c => c.id === subPick)[0];
    if (!child || !subHotelPick) return;
    const emails = await loadSubEmails(child);
    if (emails.length === 0) {
      showAlert(isDe
        ? `Für „${child.title}" konnten keine aktiven Anmeldungen gelesen werden.`
        : `No active registrations could be read for „${child.title}".`, { variant: 'error' });
      return;
    }
    const set = new Set(emails);
    const rows = people.filter(p => set.has((p.ParticipantEmail || '').trim().toLowerCase()));
    if (rows.length === 0) {
      showAlert(isDe
        ? `Niemand aus „${child.title}" ist in dieser Liste — prüfe, ob die Teilnehmer auf Klammer-Ebene erfasst sind.`
        : `Nobody from „${child.title}" is in this list — check whether attendees exist at bracket level.`, { variant: 'error' });
      return;
    }
    const already = rows.filter(r => (r.Hotel || '').trim() && (r.Hotel || '').trim() !== subHotelPick).length;
    const ok = await confirmDialog(
      isDe
        ? `Alle ${rows.length} Teilnehmer aus „${child.title}" dem Hotel „${subHotelPick}" zuordnen?${already > 0 ? `\n\n${already} davon sind aktuell einem ANDEREN Hotel zugeordnet und werden umgebucht.` : ''}`
        : `Assign all ${rows.length} attendees from „${child.title}" to „${subHotelPick}"?${already > 0 ? `\n\n${already} of them are currently assigned to a DIFFERENT hotel and will be moved.` : ''}`,
      { confirmLabel: isDe ? 'Zuordnen' : 'Assign' },
    );
    if (!ok) return;
    setBulkProgress({ done: 0, total: rows.length });
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const from = toDay(r.HotelFrom) || (defaultStay ? defaultStay.from : '');
      const to = toDay(r.HotelTo) || (defaultStay ? defaultStay.to : '');
      // eslint-disable-next-line no-await-in-loop
      await writeAssignment([r], subHotelPick, from, to);
      setBulkProgress({ done: i + 1, total: rows.length });
    }
    setBulkProgress(null);
  };

  const [importOpen, setImportOpen] = React.useState(false);
  const [importBusy, setImportBusy] = React.useState(false);
  const [importProgress, setImportProgress] = React.useState<{ done: number; total: number } | null>(null);

  /**
   * v28.49: Geprüfte Import-Zeilen übernehmen. Reihenfolge ist wichtig: erst
   * die fehlenden Hotels anlegen (EIN Schreibvorgang am Event), dann die
   * Zuordnungen — sonst zeigt die Auswahl in der Tabelle hinterher Namen, die
   * es als Hotel gar nicht gibt.
   */
  const applyImport = async (rowsIn: IHotelImportResultRow[], newHotelNames: string[]): Promise<void> => {
    if (!svc || !event.subsiteUrl) {
      showAlert(isDe ? 'Kein Zugriff auf die Teilnehmerliste dieses Events.' : 'No access to this event’s participant list.', { variant: 'error' });
      return;
    }
    setImportBusy(true);
    setImportProgress({ done: 0, total: rowsIn.length });
    if (newHotelNames.length > 0) {
      const merged = hotels.concat(newHotelNames.map(n => ({ id: uid('h'), name: n, address: '', capacity: 0, notes: '' })));
      const okH = await svc.patchEventOverridesValue(Number(event.id), '_hotels', merged);
      if (!okH) {
        setImportBusy(false); setImportProgress(null);
        showAlert(isDe ? 'Die neuen Hotels konnten nicht gespeichert werden — es wurde nichts übernommen.' : 'Could not save the new hotels — nothing was applied.', { variant: 'error' });
        return;
      }
    }
    try { await svc.ensureHotelColumns(event.subsiteUrl); } catch { /* best effort */ }
    let failed = 0;
    for (let i = 0; i < rowsIn.length; i++) {
      const r = rowsIn[i];
      if (!r.reg) { failed++; continue; }
      // eslint-disable-next-line no-await-in-loop
      const ok = await svc.setHotelAssignment(event.subsiteUrl, r.reg.Id, r.hotel, r.from ? `${r.from}T00:00:00Z` : '', r.to ? `${r.to}T00:00:00Z` : '');
      if (!ok) failed++;
      setImportProgress({ done: i + 1, total: rowsIn.length });
    }
    await onReloadEvents();
    await onReloadRegistrations();
    setImportBusy(false);
    setImportProgress(null);
    setImportOpen(false);
    showAlert(
      isDe
        ? `${rowsIn.length - failed} von ${rowsIn.length} Zuordnung(en) übernommen${failed > 0 ? `, ${failed} fehlgeschlagen` : ''}${newHotelNames.length > 0 ? `. ${newHotelNames.length} Hotel(s) neu angelegt.` : '.'}`
        : `${rowsIn.length - failed} of ${rowsIn.length} assignment(s) applied${failed > 0 ? `, ${failed} failed` : ''}${newHotelNames.length > 0 ? `. ${newHotelNames.length} hotel(s) created.` : '.'}`,
      { variant: failed > 0 ? 'error' : 'success' },
    );
  };

  // Nur aktive Anmeldungen — Abgemeldete und Warteliste brauchen kein Zimmer.
  const people = React.useMemo(
    () => registrations
      .filter(r => ACTIVE_STATI.indexOf(r.Status || '') >= 0)
      .sort((a, b) => (a.ParticipantName || '').localeCompare(b.ParticipantName || '', 'de')),
    [registrations],
  );

  /** v28.48: Hat die Person im Anmeldeformular eine Unterkunft angefragt?
   *  Die Hotel-Frage heisst pro Event anders („Hotel room", „Hotel (24-25 Sept)",
   *  „Room required …"), deshalb eine Heuristik ueber die Antwortwerte statt
   *  einer festen Feld-ID. Ergebnis: true = ja, false = nein, null = keine
   *  Hotel-Frage im Formular. */
  const wishOf = React.useCallback((p: SPRegistration): boolean | null => {
    let cd: Record<string, string> = {};
    try { cd = JSON.parse(p.CustomData || '{}'); } catch { return null; }
    const fields = (event.eventSpecificFields || []) as Array<{ id: string; label?: string }>;
    let found: boolean | null = null;
    for (const f of fields) {
      if (!/hotel|unterkunft|übernacht|uebernacht|accommodation|lodging/i.test(f.label || '')) continue;
      const v = (cd[f.id] || '').toLowerCase();
      if (!v) { if (found === null) found = false; continue; }
      if (/^(ja|yes)\b|^ja,|^yes,/.test(v) || /\bja\b|\byes\b/.test(v)) return true;
      found = false;
    }
    return found;
  }, [event.eventSpecificFields]);

  const assignedCount = people.filter(p => (p.Hotel || '').trim()).length;
  const openCount = people.length - assignedCount;

  /* ---------------- Stammdaten schreiben ---------------- */

  const saveHotels = async (next: DexHotel[]): Promise<void> => {
    if (!svc) return;
    setBusy('hotels');
    const ok = await svc.patchEventOverridesValue(Number(event.id), '_hotels', next);
    setBusy('');
    if (!ok) { showAlert(isDe ? 'Die Hotels konnten nicht gespeichert werden.' : 'Could not save the hotels.', { variant: 'error' }); return; }
    await onReloadEvents();
  };

  const saveStays = async (next: DexHotelStay[]): Promise<void> => {
    if (!svc) return;
    setBusy('stays');
    const ok = await svc.patchEventOverridesValue(Number(event.id), '_hotelStays', next);
    setBusy('');
    if (!ok) { showAlert(isDe ? 'Die Zeiträume konnten nicht gespeichert werden.' : 'Could not save the stay templates.', { variant: 'error' }); return; }
    await onReloadEvents();
  };

  const toggleVisible = async (): Promise<void> => {
    if (!svc) return;
    const next = !visible;
    // Beim Freigeben ehrlich sagen, was die Teilnehmer sehen — und was nicht.
    if (next && openCount > 0) {
      const ok = await confirmDialog(
        isDe
          ? `${openCount} von ${people.length} Personen ist noch kein Hotel zugeordnet.\n\nNach der Freigabe sehen die bereits zugeordneten Personen ihr Hotel, die übrigen sehen nichts — das erzeugt meist Rückfragen.\n\nTrotzdem freigeben?`
          : `${openCount} of ${people.length} people have no hotel yet.\n\nAfter release, assigned people see their hotel while the others see nothing — that usually triggers questions.\n\nRelease anyway?`,
        { confirmLabel: isDe ? 'Trotzdem freigeben' : 'Release anyway' },
      );
      if (!ok) return;
    }
    setBusy('visible');
    const ok = await svc.patchEventOverridesValue(Number(event.id), '_hotelVisible', next);
    setBusy('');
    if (!ok) { showAlert(isDe ? 'Die Einstellung konnte nicht gespeichert werden.' : 'Could not save the setting.', { variant: 'error' }); return; }
    await onReloadEvents();
    showAlert(
      next
        ? (isDe ? 'Freigegeben — die Teilnehmer sehen ihr Hotel unter „Meine Events".' : 'Released — attendees now see their hotel under „My events".')
        : (isDe ? 'Anzeige zurückgenommen — die Teilnehmer sehen ihr Hotel nicht mehr.' : 'Hidden again — attendees no longer see their hotel.'),
      { variant: 'success' },
    );
  };

  /* ---------------- Zuordnung schreiben ---------------- */

  const writeAssignment = async (
    rows: SPRegistration[],
    hotel: string,
    from: string,
    to: string,
  ): Promise<void> => {
    if (!svc || !event.subsiteUrl) {
      showAlert(isDe ? 'Kein Zugriff auf die Teilnehmerliste dieses Events.' : 'No access to this event’s participant list.', { variant: 'error' });
      return;
    }
    // Ist die Anzeige freigegeben, sieht die Person die Änderung sofort.
    if (visible && rows.length > 0) {
      const ok = await confirmDialog(
        isDe
          ? `Die Hotel-Anzeige ist für dieses Event freigegeben.\n\n${rows.length === 1 ? 'Die betroffene Person sieht' : `Die ${rows.length} betroffenen Personen sehen`} die Änderung sofort unter „Meine Events".\n\nFortfahren?`
          : `Hotel display is released for this event.\n\n${rows.length === 1 ? 'The person sees' : `The ${rows.length} people see`} the change immediately under „My events".\n\nContinue?`,
        { confirmLabel: isDe ? 'Ja, zuordnen' : 'Yes, assign' },
      );
      if (!ok) return;
    }
    setBusy('assign');
    // Spalten sind idempotent — der erste Zuordnungs-Klick legt sie an.
    try { await svc.ensureHotelColumns(event.subsiteUrl); } catch { /* best effort */ }
    let failed = 0;
    for (const r of rows) {
      const ok = await svc.setHotelAssignment(event.subsiteUrl, r.Id, hotel, toIso(from), toIso(to));
      if (!ok) failed++;
    }
    await onReloadRegistrations();
    setBusy('');
    setSelected(new Set());
    if (failed > 0) {
      showAlert(
        isDe ? `${rows.length - failed} von ${rows.length} zugeordnet, ${failed} fehlgeschlagen. Fehlen die Spalten in der Teilnehmerliste, hilft „Spalten fixen" im Admin Center.`
          : `${rows.length - failed} of ${rows.length} assigned, ${failed} failed. If the columns are missing, use „Fix columns" in the Admin Center.`,
        { variant: 'error' },
      );
    }
  };

  const defaultStay = stays.filter(s => s.isDefault)[0] || stays[0];

  const assignSelectedTo = async (hotelName: string): Promise<void> => {
    const rows = people.filter(p => selected.has(p.Id));
    if (rows.length === 0) return;
    setBulkProgress({ done: 0, total: rows.length });
    // Zeitraum: vorhandenen behalten, sonst die Standard-Vorlage geben.
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const from = toDay(r.HotelFrom) || (defaultStay ? defaultStay.from : '');
      const to = toDay(r.HotelTo) || (defaultStay ? defaultStay.to : '');
      // eslint-disable-next-line no-await-in-loop
      await writeAssignment([r], hotelName, from, to);
      setBulkProgress({ done: i + 1, total: rows.length });
    }
    setBulkProgress(null);
  };

  const applyStayToSelected = async (stay: DexHotelStay): Promise<void> => {
    const rows = people.filter(p => selected.has(p.Id));
    if (rows.length === 0) return;
    setBulkProgress({ done: 0, total: rows.length });
    for (let i = 0; i < rows.length; i++) {
      // eslint-disable-next-line no-await-in-loop
      await writeAssignment([rows[i]], (rows[i].Hotel || '').trim(), stay.from, stay.to);
      setBulkProgress({ done: i + 1, total: rows.length });
    }
    setBulkProgress(null);
  };

  /* ---------------- Auswertungen ---------------- */

  const perHotel = React.useMemo(() => hotels.map(h => {
    const rows = people.filter(p => (p.Hotel || '').trim() === h.name);
    const nights = rows.reduce((sum, r) => sum + nightsBetween(toDay(r.HotelFrom), toDay(r.HotelTo)), 0);
    const cap = h.capacity || 0;
    return { hotel: h, count: rows.length, nights, over: cap > 0 && rows.length > cap, free: cap > 0 ? cap - rows.length : null };
  }), [hotels, people]);

  /** Belegung je Nacht — die Zahl, die das Hotel fürs Kontingent braucht. */
  const perNight = React.useMemo(() => {
    const days: string[] = [];
    let min = ''; let max = '';
    for (const p of people) {
      const f = toDay(p.HotelFrom); const t = toDay(p.HotelTo);
      if (!f || !t) continue;
      if (!min || f < min) min = f;
      if (!max || t > max) max = t;
    }
    if (!min || !max) return [] as Array<{ day: string; total: number; byHotel: Record<string, number> }>;
    for (let d = min; d < max; d = addDays(d, 1)) days.push(d);
    return days.map(day => {
      const byHotel: Record<string, number> = {};
      let total = 0;
      for (const p of people) {
        const f = toDay(p.HotelFrom); const t = toDay(p.HotelTo);
        if (!f || !t) continue;
        if (day >= f && day < t) {
          total++;
          const h = (p.Hotel || '').trim() || '—';
          byHotel[h] = (byHotel[h] || 0) + 1;
        }
      }
      return { day, total, byHotel };
    }).filter(x => x.total > 0);
  }, [people]);

  /** Wer hat laut Anmeldung Unterkunft angefragt, ist aber ohne Hotel? */
  const wantsHotelWithout = React.useMemo(() => people.filter(p => {
    if ((p.Hotel || '').trim()) return false;
    const blob = JSON.stringify(p || {}).toLowerCase();
    // Bewusst breit: die Feldnamen sind pro Event frei benannt („Hotel room",
    // „Hotel (24-25 Sept)", „Room required …"). Ein Treffer auf eine
    // Ja-Antwort in einem Hotel-Feld reicht als Hinweis.
    if (blob.indexOf('hotel') < 0 && blob.indexOf('accommodation') < 0 && blob.indexOf('unterkunft') < 0) return false;
    return blob.indexOf('yes') >= 0 || blob.indexOf('ja,') >= 0 || blob.indexOf('ja ') >= 0;
  }), [people]);

  /* ---------------- Export ---------------- */

  const exportRooming = (hotelName?: string): void => {
    const rows = people.filter(p => (hotelName ? (p.Hotel || '').trim() === hotelName : true));
    const head = ['Hotel', 'Nachname', 'Vorname', 'E-Mail', 'Anreise', 'Abreise', 'Naechte'];
    const csv = [head.join(';')].concat(rows.map(r => [
      (r.Hotel || '').trim(),
      r.Nachname || '',
      r.Vorname || '',
      r.ParticipantEmail || '',
      toDay(r.HotelFrom),
      toDay(r.HotelTo),
      String(nightsBetween(toDay(r.HotelFrom), toDay(r.HotelTo))),
    ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(';'))).join('\r\n');
    // BOM, damit Excel die Umlaute richtig liest.
    const blob = new Blob([`﻿${csv}`], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Rooming_${(hotelName || 'alle').replace(/[^\w\-]+/g, '_')}.csv`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  };

  /* ---------------- Hotel-/Zeitraum-Formulare ---------------- */

  /* ---------------- Hotel-Info-Mail (v28.39) ----------------
   * Nutzt denselben Editor + Versandweg wie die QR-Mail: HtmlEditorModal mit
   * Live-Vorschau und einsetzbaren Variablen, Versand ueber queueEmail. Die
   * Assistenz kommt automatisch ins CC — ueber dieselben People-Picker-Felder,
   * die der Organizer in Schritt 5 als „CC bei Mails" markiert hat. */
  const [mailOpen, setMailOpen] = React.useState(false);
  const [mailSubject, setMailSubject] = React.useState('');
  const [mailHeading, setMailHeading] = React.useState('');
  const [mailSubheading, setMailSubheading] = React.useState('');
  const [mailBody, setMailBody] = React.useState('');
  const [mailSending, setMailSending] = React.useState(false);
  /** Testmodus: EIN Empfaenger, frei gewaehltes Hotel, Beispielwerte. */
  const [testMode, setTestMode] = React.useState(true);
  const [testTo, setTestTo] = React.useState('');
  const [testHotelId, setTestHotelId] = React.useState('');

  const defaultMailBody = isDe
    ? '<p>Hallo {{Vorname}},</p>'
      + '<p>hier deine Hotelinformationen für <strong>{{EventTitle}}</strong>:</p>'
      + '<p><strong>{{Hotel}}</strong><br>{{HotelAdresse}}<br>'
      + 'Anreise: <strong>{{Anreise}}</strong><br>Abreise: <strong>{{Abreise}}</strong><br>'
      + 'Übernachtungen: <strong>{{Naechte}}</strong></p>'
      + '<p>Das Zimmer ist für dich reserviert und wird zentral abgerechnet — du musst nichts selbst buchen.</p>'
      + '<p>Bei Fragen melde dich gerne bei uns.</p>'
    : '<p>Hi {{Vorname}},</p>'
      + '<p>here are your hotel details for <strong>{{EventTitle}}</strong>:</p>'
      + '<p><strong>{{Hotel}}</strong><br>{{HotelAdresse}}<br>'
      + 'Arrival: <strong>{{Anreise}}</strong><br>Departure: <strong>{{Abreise}}</strong><br>'
      + 'Nights: <strong>{{Naechte}}</strong></p>'
      + '<p>The room is reserved for you and billed centrally — no need to book anything yourself.</p>'
      + '<p>Any questions, just get in touch.</p>';

  const openMail = (): void => {
    setMailSubject(isDe ? `Deine Hotelinformationen — ${event.title}` : `Your hotel details — ${event.title}`);
    setMailHeading(isDe ? 'Deine Hotelinformationen' : 'Your hotel details');
    setMailSubheading(event.title || '');
    setMailBody(defaultMailBody);
    setTestMode(true);
    setTestTo('');
    setTestHotelId(hotels.length > 0 ? hotels[0].id : '');
    setMailOpen(true);
  };

  /** Variablen einer Person (bzw. Beispielwerte im Testmodus). */
  const varsFor = (p: SPRegistration | null, hotelOverride?: DexHotel): Record<string, string> => {
    const hName = hotelOverride ? hotelOverride.name : (p ? (p.Hotel || '').trim() : '');
    const h = hotels.filter(x => x.name === hName)[0];
    const from = hotelOverride ? (defaultStay ? defaultStay.from : '') : toDay(p ? p.HotelFrom : '');
    const to = hotelOverride ? (defaultStay ? defaultStay.to : '') : toDay(p ? p.HotelTo : '');
    return {
      Vorname: (p && p.Vorname) || (isDe ? 'Vorname' : 'First name'),
      Nachname: (p && p.Nachname) || (isDe ? 'Nachname' : 'Last name'),
      EventTitle: event.title || '',
      Hotel: hName || (isDe ? '(kein Hotel)' : '(no hotel)'),
      HotelAdresse: (h && h.address) || '',
      Anreise: fmtDay(from, isDe),
      Abreise: fmtDay(to, isDe),
      Naechte: String(nightsBetween(from, to)),
    };
  };

  const sendHotelMail = async (): Promise<void> => {
    if (!svc) return;
    const send = async (to: string, name: string, cc: string, vars: Record<string, string>): Promise<boolean> => {
      const body = wrapTemplate(
        '#86bc25',
        replacePlaceholders(mailHeading, vars),
        replacePlaceholders(mailSubheading, vars),
        replacePlaceholders(mailBody, vars),
      );
      return svc.queueEmail(
        replacePlaceholders(mailSubject, vars), to, name, body,
        'HotelInfo', event.title, event.id, cc || undefined,
      );
    };

    if (testMode) {
      const to = testTo.trim();
      if (!to || to.indexOf('@') <= 0) {
        showAlert(isDe ? 'Bitte eine Empfänger-Adresse für den Test wählen.' : 'Please pick a test recipient.', { variant: 'error' });
        return;
      }
      const h = hotels.filter(x => x.id === testHotelId)[0];
      setMailSending(true);
      const ok = await send(to, to, '', varsFor(null, h));
      setMailSending(false);
      showAlert(
        ok ? (isDe ? `Testmail an ${to} verschickt — mit Beispielwerten für „${h ? h.name : '—'}". Es ging KEINE Mail an Teilnehmer.` : `Test mail sent to ${to} using sample values for „${h ? h.name : '—'}". No attendee received anything.`)
          : (isDe ? 'Die Testmail konnte nicht eingereiht werden.' : 'Could not queue the test mail.'),
        { variant: ok ? 'success' : 'error' },
      );
      return;
    }

    const rows = people.filter(p => (p.Hotel || '').trim());
    if (rows.length === 0) {
      showAlert(isDe ? 'Es ist noch niemand einem Hotel zugeordnet.' : 'Nobody is assigned to a hotel yet.', { variant: 'error' });
      return;
    }
    const ok = await confirmDialog(
      isDe
        ? `Hotelinformationen an ${rows.length} Person(en) verschicken?\n\nJede Person bekommt ihre eigenen Daten (Hotel, An-/Abreise, Nächte). Wo im Anmeldeformular eine Assistenz hinterlegt ist, steht sie automatisch im CC.\n\nNicht zugeordnete Personen bekommen nichts.`
        : `Send hotel details to ${rows.length} person(s)?\n\nEach one receives their own data (hotel, arrival/departure, nights). Where an assistant was named during registration, they are added to CC automatically.\n\nUnassigned people receive nothing.`,
      { confirmLabel: isDe ? `An ${rows.length} senden` : `Send to ${rows.length}` },
    );
    if (!ok) return;
    setMailSending(true);
    let sent = 0; let failed = 0;
    for (const p of rows) {
      let cd: Record<string, string> = {};
      try { cd = JSON.parse(p.CustomData || '{}'); } catch { cd = {}; }
      const cc = collectCcEmailsFromFields(event.eventSpecificFields as any, cd, p.ParticipantEmail);
      // eslint-disable-next-line no-await-in-loop
      const okOne = await send(p.ParticipantEmail, p.ParticipantName || p.ParticipantEmail, cc, varsFor(p));
      if (okOne) sent++; else failed++;
    }
    setMailSending(false);
    setMailOpen(false);
    showAlert(
      isDe ? `${sent} Hotelmail(s) eingereiht${failed > 0 ? `, ${failed} fehlgeschlagen` : ''}.`
        : `${sent} hotel mail(s) queued${failed > 0 ? `, ${failed} failed` : ''}.`,
      { variant: failed > 0 ? 'error' : 'success' },
    );
  };

  const [newHotel, setNewHotel] = React.useState<DexHotel>({ id: '', name: '', address: '', capacity: 0, notes: '' });
  const [newStay, setNewStay] = React.useState<{ label: string; from: string; to: string }>(() => {
    const start = toDay(event.startDate) || '';
    const end = toDay(event.endDate) || start;
    return { label: '', from: start ? addDays(start, -1) : '', to: end };
  });

  const card: React.CSSProperties = {
    background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 10,
    padding: 16, marginBottom: 14,
  };
  const h3: React.CSSProperties = { margin: '0 0 4px', fontSize: '1rem', fontWeight: 700, color: 'var(--dex-gray-800)' };
  const hint: React.CSSProperties = { margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 };
  const inp: React.CSSProperties = { height: 30, fontSize: '0.8rem', padding: '0 8px', border: '1px solid var(--dex-gray-300)', borderRadius: 6, minWidth: 0 };

  return (
    <div>
      {/* ---- Kopf: Freigabe + Kennzahlen ---- */}
      <div style={card}>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'flex-start' }}>
          <div style={{ flex: 1, minWidth: 240 }}>
            <h3 style={h3}>{isDe ? 'Hotel-Planung' : 'Hotel planning'}</h3>
            <p style={{ ...hint, marginBottom: 0 }}>
              {isDe
                ? <>Lege die Hotels an, gib Zeiträume als Vorlage vor und ordne die Teilnehmer zu. Die Zuordnung steht danach in der Teilnehmerliste und läuft in jeden Export mit.</>
                : <>Create the hotels, define stay templates and assign attendees. The assignment then lives in the participant list and is part of every export.</>}
            </p>
          </div>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            <div style={{ padding: '8px 14px', borderRadius: 8, background: 'var(--dex-gray-50, #f7f7f5)', textAlign: 'center', minWidth: 92 }}>
              <div style={{ fontSize: '1.3rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>{assignedCount}</div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)' }}>{isDe ? 'zugeordnet' : 'assigned'}</div>
            </div>
            <div style={{ padding: '8px 14px', borderRadius: 8, background: openCount > 0 ? '#fff6e5' : 'var(--dex-gray-50, #f7f7f5)', textAlign: 'center', minWidth: 92 }}>
              <div style={{ fontSize: '1.3rem', fontWeight: 700, color: openCount > 0 ? '#b35a00' : 'var(--dex-gray-600)' }}>{openCount}</div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)' }}>{isDe ? 'offen' : 'open'}</div>
            </div>
          </div>
        </div>

        <label style={{
          display: 'flex', alignItems: 'flex-start', gap: 10, marginTop: 14, padding: '10px 12px',
          borderRadius: 8, cursor: busy ? 'wait' : 'pointer',
          border: `1px solid ${visible ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
          background: visible ? 'rgba(134,188,37,0.07)' : '#fff',
        }}>
          <input
            type="checkbox"
            checked={visible}
            disabled={busy !== ''}
            onChange={() => { void toggleVisible(); }}
            style={{ width: 18, height: 18, marginTop: 1, cursor: 'pointer', accentColor: 'var(--dex-green, #86bc25)' }}
          />
          <span style={{ fontSize: '0.85rem' }}>
            <strong>{isDe ? 'Teilnehmer sehen ihr Hotel unter „Meine Events"' : 'Attendees see their hotel under „My events"'}</strong>
            <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2, lineHeight: 1.45 }}>
              {visible
                ? (isDe
                  ? `Freigegeben — ${assignedCount} von ${people.length} Personen sehen aktuell ihr Hotel${openCount > 0 ? `, ${openCount} sehen nichts (noch nicht zugeordnet)` : ''}. Jede Umbuchung ist sofort sichtbar.`
                  : `Released — ${assignedCount} of ${people.length} people currently see their hotel${openCount > 0 ? `, ${openCount} see nothing (not assigned yet)` : ''}. Every change is visible immediately.`)
                : (isDe
                  ? 'Aus — niemand sieht seine Zuordnung. Plane in Ruhe und gib frei, wenn alles steht.'
                  : 'Off — nobody sees their assignment. Plan in peace and release when everything is set.')}
            </span>
          </span>
        </label>
      </div>

      {/* ---- Hotels ---- */}
      <div style={card}>
        <h3 style={h3}>{isDe ? 'Hotels' : 'Hotels'}</h3>
        <p style={hint}>
          {isDe
            ? 'Kontingent = Zimmer bzw. Betten, die euch zustehen. Ist es gesetzt, warnt die Ampel bei Überbuchung.'
            : 'Capacity = rooms or beds available to you. If set, the indicator warns about overbooking.'}
        </p>
        {perHotel.length > 0 && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 10, marginBottom: 12 }}>
            {perHotel.map(ph => (
              <div key={ph.hotel.id} style={{
                border: `1px solid ${ph.over ? 'var(--dex-red, #c00)' : 'var(--dex-gray-200)'}`,
                borderRadius: 8, padding: 10, background: ph.over ? '#fef3f2' : '#fff',
              }}>
                <div style={{ fontWeight: 700, fontSize: '0.88rem' }}>{ph.hotel.name}</div>
                {ph.hotel.address && <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>{ph.hotel.address}</div>}
                <div style={{ marginTop: 6, fontSize: '0.78rem' }}>
                  <strong>{ph.count}</strong>{ph.hotel.capacity ? ` / ${ph.hotel.capacity}` : ''} {isDe ? 'Personen' : 'people'}
                  {ph.free !== null && !ph.over && <span style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}> · {ph.free} {isDe ? 'frei' : 'free'}</span>}
                  {ph.over && <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 700 }}> · {isDe ? 'überbucht' : 'overbooked'}</span>}
                </div>
                <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)' }}>
                  {ph.nights} {isDe ? 'Übernachtungen gesamt' : 'room nights total'}
                </div>
                <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
                  <button type="button" onClick={() => exportRooming(ph.hotel.name)}
                    style={{ border: 'none', background: 'none', padding: 0, cursor: 'pointer', fontSize: '0.74rem', color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline' }}>
                    {isDe ? 'Rooming-Liste' : 'Rooming list'}
                  </button>
                  <button type="button"
                    onClick={() => {
                      (async () => {
                        const ok = await confirmDialog(
                          isDe
                            ? `Hotel „${ph.hotel.name}" entfernen?${ph.count > 0 ? `\n\n${ph.count} Person(en) sind ihm zugeordnet — die Zuordnung bleibt bestehen und muss dann neu gesetzt werden.` : ''}`
                            : `Remove hotel „${ph.hotel.name}"?${ph.count > 0 ? `\n\n${ph.count} person(s) are assigned — their assignment stays and must be re-set.` : ''}`,
                          { confirmLabel: isDe ? 'Entfernen' : 'Remove', danger: true },
                        );
                        if (ok) await saveHotels(hotels.filter(h => h.id !== ph.hotel.id));
                      })().catch(() => { /* */ });
                    }}
                    style={{ border: 'none', background: 'none', padding: 0, cursor: 'pointer', fontSize: '0.74rem', color: 'var(--dex-red, #c00)', textDecoration: 'underline' }}>
                    {isDe ? 'Entfernen' : 'Remove'}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
          <input style={{ ...inp, flex: '2 1 160px' }} placeholder={isDe ? 'Hotelname' : 'Hotel name'}
            value={newHotel.name} onChange={e => setNewHotel({ ...newHotel, name: e.target.value })} />
          <input style={{ ...inp, flex: '3 1 200px' }} placeholder={isDe ? 'Adresse (optional)' : 'Address (optional)'}
            value={newHotel.address} onChange={e => setNewHotel({ ...newHotel, address: e.target.value })} />
          <input style={{ ...inp, width: 110 }} type="number" min={0} placeholder={isDe ? 'Kontingent' : 'Capacity'}
            value={newHotel.capacity || ''} onChange={e => setNewHotel({ ...newHotel, capacity: parseInt(e.target.value, 10) || 0 })} />
          <button type="button" className="btn btn-secondary" style={{ fontSize: '0.78rem', padding: '5px 14px' }}
            disabled={busy !== ''}
            onClick={() => {
              const name = (newHotel.name || '').trim();
              if (!name) { showAlert(isDe ? 'Bitte einen Hotelnamen eingeben.' : 'Please enter a hotel name.', { variant: 'error' }); return; }
              if (hotels.some(h => h.name.toLowerCase() === name.toLowerCase())) {
                showAlert(isDe ? 'Ein Hotel mit diesem Namen gibt es schon.' : 'A hotel with that name already exists.', { variant: 'error' }); return;
              }
              void saveHotels(hotels.concat([{ ...newHotel, id: uid('h'), name }]));
              setNewHotel({ id: '', name: '', address: '', capacity: 0, notes: '' });
            }}>
            {isDe ? '+ Hotel' : '+ Hotel'}
          </button>
        </div>
      </div>

      {/* ---- Zeitraum-Vorlagen ---- */}
      <div style={card}>
        <h3 style={h3}>{isDe ? 'Zeiträume' : 'Stay templates'}</h3>
        <p style={hint}>
          {isDe
            ? 'Statt bei jeder Person zweimal ein Datum zu wählen: Zeiträume einmal anlegen, dann markieren und mit einem Klick zuweisen. Der Standard wird neuen Zuordnungen automatisch gegeben.'
            : 'Instead of picking two dates per person: define the ranges once, then select people and apply with one click. The default is used for new assignments.'}
        </p>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
          {stays.map(st => (
            <span key={st.id} style={{
              display: 'inline-flex', alignItems: 'center', gap: 8, padding: '5px 10px', borderRadius: 999,
              border: `1px solid ${st.isDefault ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
              background: st.isDefault ? 'rgba(134,188,37,0.10)' : '#fff', fontSize: '0.76rem',
            }}>
              <strong>{st.label}</strong>
              <span style={{ color: 'var(--dex-gray-500)' }}>
                {fmtDay(st.from, isDe)} – {fmtDay(st.to, isDe)} · {nightsBetween(st.from, st.to)} {isDe ? 'N.' : 'n.'}
              </span>
              {!st.isDefault && (
                <button type="button" title={isDe ? 'Als Standard setzen' : 'Set as default'}
                  onClick={() => { void saveStays(stays.map(x => ({ ...x, isDefault: x.id === st.id }))); }}
                  style={{ border: 'none', background: 'none', cursor: 'pointer', padding: 0, fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>
                  {isDe ? 'Standard' : 'Default'}
                </button>
              )}
              <button type="button" title={isDe ? 'Zeitraum löschen' : 'Delete template'}
                onClick={() => { void saveStays(stays.filter(x => x.id !== st.id)); }}
                style={{ border: 'none', background: 'none', cursor: 'pointer', padding: 0, color: 'var(--dex-red, #c00)', fontSize: '0.85rem', lineHeight: 1 }}>×</button>
            </span>
          ))}
          {stays.length === 0 && (
            <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Noch kein Zeitraum angelegt.' : 'No template yet.'}
            </span>
          )}
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
          <input style={{ ...inp, flex: '2 1 150px' }} placeholder={isDe ? 'Bezeichnung, z.B. „Mit Vorabend"' : 'Label, e.g. „With prior evening"'}
            value={newStay.label} onChange={e => setNewStay({ ...newStay, label: e.target.value })} />
          <input style={{ ...inp }} type="date" value={newStay.from} onChange={e => setNewStay({ ...newStay, from: e.target.value })} />
          <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>–</span>
          <input style={{ ...inp }} type="date" value={newStay.to} onChange={e => setNewStay({ ...newStay, to: e.target.value })} />
          <span style={{ fontSize: '0.76rem', color: 'var(--dex-gray-600)' }}>
            {nightsBetween(newStay.from, newStay.to)} {isDe ? 'Nächte' : 'nights'}
          </span>
          <button type="button" className="btn btn-secondary" style={{ fontSize: '0.78rem', padding: '5px 14px' }}
            disabled={busy !== ''}
            onClick={() => {
              const n = nightsBetween(newStay.from, newStay.to);
              if (n <= 0) { showAlert(isDe ? 'Die Abreise muss nach der Anreise liegen.' : 'Departure must be after arrival.', { variant: 'error' }); return; }
              const label = (newStay.label || '').trim() || `${n} ${isDe ? (n === 1 ? 'Nacht' : 'Nächte') : (n === 1 ? 'night' : 'nights')}`;
              void saveStays(stays.concat([{ id: uid('s'), label, from: newStay.from, to: newStay.to, isDefault: stays.length === 0 }]));
              setNewStay({ ...newStay, label: '' });
            }}>
            {isDe ? '+ Zeitraum' : '+ Template'}
          </button>
        </div>
      </div>

      {/* ---- Vollständigkeits-Check ---- */}
      {wantsHotelWithout.length > 0 && (
        <div style={{ ...card, borderColor: '#e8a33d', background: '#fff8ec' }}>
          <h3 style={{ ...h3, color: '#b35a00' }}>
            {isDe ? `${wantsHotelWithout.length} Person(en) mit Unterkunftsbedarf ohne Hotel` : `${wantsHotelWithout.length} person(s) needing accommodation without a hotel`}
          </h3>
          <p style={{ ...hint, marginBottom: 8 }}>
            {isDe
              ? 'Diese Personen haben im Anmeldeformular eine Unterkunft angefragt, sind aber noch keinem Hotel zugeordnet — genau der Fall, der in einer Excel-Liste durchrutscht.'
              : 'These people requested accommodation during registration but are not assigned to a hotel yet — exactly the case that slips through in a spreadsheet.'}
          </p>
          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)' }}>
            {wantsHotelWithout.slice(0, 15).map(p => p.ParticipantName || p.ParticipantEmail).join(', ')}
            {wantsHotelWithout.length > 15 ? ` … (+${wantsHotelWithout.length - 15})` : ''}
          </div>
        </div>
      )}

      {/* ---- Belegung je Nacht ---- */}
      {perNight.length > 0 && (
        <div style={card}>
          <h3 style={h3}>{isDe ? 'Belegung je Nacht' : 'Occupancy per night'}</h3>
          <p style={hint}>
            {isDe ? 'Das ist die Zahl, die das Hotel fürs Abrufkontingent braucht.' : 'This is the figure the hotel needs for the room allotment.'}
          </p>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ borderCollapse: 'collapse', fontSize: '0.78rem', minWidth: 380 }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', padding: '4px 10px 4px 0' }}>{isDe ? 'Nacht' : 'Night'}</th>
                  <th style={{ textAlign: 'right', padding: '4px 12px' }}>{isDe ? 'Gesamt' : 'Total'}</th>
                  {hotels.map(h => <th key={h.id} style={{ textAlign: 'right', padding: '4px 12px' }}>{h.name}</th>)}
                </tr>
              </thead>
              <tbody>
                {perNight.map(n => (
                  <tr key={n.day} style={{ borderTop: '1px solid var(--dex-gray-100)' }}>
                    <td style={{ padding: '4px 10px 4px 0' }}>{fmtDay(n.day, isDe)}</td>
                    <td style={{ textAlign: 'right', padding: '4px 12px', fontWeight: 700 }}>{n.total}</td>
                    {hotels.map(h => <td key={h.id} style={{ textAlign: 'right', padding: '4px 12px' }}>{n.byHotel[h.name] || 0}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ---- Zuordnung ---- */}
      <div style={card}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 4 }}>
          <h3 style={{ ...h3, margin: 0 }}>{isDe ? 'Zuordnung' : 'Assignment'}</h3>
          <button type="button" onClick={() => setShowRoster(o => !o)}
            style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: '0.76rem', color: 'var(--dex-gray-600)', textDecoration: 'underline' }}>
            {showRoster ? (isDe ? 'einklappen' : 'collapse') : (isDe ? 'aufklappen' : 'expand')}
          </button>
          <span style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
            <input type="text" value={search} onChange={e => setSearch(e.target.value)}
              placeholder={isDe ? 'Suchen (Name, Standort, Hotel …)' : 'Search (name, location, hotel …)'}
              style={{ ...inp, minWidth: 210 }} />
            <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.76rem', color: 'var(--dex-gray-600)', cursor: 'pointer', whiteSpace: 'nowrap' }}>
              <input type="checkbox" checked={hideNoWish} onChange={e => setHideNoWish(e.target.checked)} />
              {isDe ? 'ohne Hotel-Wunsch ausblenden' : 'hide without request'}
            </label>
            <select style={{ ...inp }} value={filterHotel} onChange={e => setFilterHotel(e.target.value)}>
              <option value="__all">{isDe ? 'Alle anzeigen' : 'Show all'}</option>
              <option value="__none">{isDe ? 'Nur ohne Hotel' : 'Only without hotel'}</option>
              {hotels.map(h => <option key={h.id} value={h.name}>{h.name}</option>)}
            </select>
            <button type="button" className="btn btn-secondary" style={{ fontSize: '0.76rem', padding: '5px 12px' }}
              onClick={() => setImportOpen(true)}>
              {isDe ? 'Liste importieren' : 'Import list'}
            </button>
            <button type="button" className="btn btn-secondary" style={{ fontSize: '0.76rem', padding: '5px 12px' }}
              onClick={() => exportRooming()}>
              {isDe ? 'Alles als Excel (CSV)' : 'Export all (CSV)'}
            </button>
            <button type="button" className="btn btn-primary" style={{ fontSize: '0.76rem', padding: '5px 12px' }}
              disabled={hotels.length === 0} onClick={openMail}>
              {isDe ? 'Hotel-Info verschicken' : 'Send hotel details'}
            </button>
          </span>
        </div>

        {showRoster && (
          <>
            {hotels.length === 0 && (
              <p style={{ ...hint, marginTop: 8 }}>
                {isDe ? 'Lege zuerst oben mindestens ein Hotel an.' : 'Create at least one hotel above first.'}
              </p>
            )}
            {/* v28.51: Ganzes Sub-Event auf einmal — der Fall „alle Teilnehmer
                vom Her-Space-Event in ein Hotel". Nur bei einer Klammer sinnvoll,
                deshalb an childEvents gebunden. */}
            {(childEvents || []).length > 0 && hotels.length > 0 && (
              <div style={{
                display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center',
                padding: '8px 10px', marginBottom: 10, borderRadius: 8,
                background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)',
              }}>
                <strong style={{ fontSize: '0.8rem' }}>
                  {isDe ? 'Ganzes Sub-Event zuordnen:' : 'Assign a whole sub-event:'}
                </strong>
                <select style={{ ...inp }} value={subPick} disabled={busy !== '' || subLoading}
                  onChange={e => setSubPick(e.target.value)}>
                  <option value="">{isDe ? '— Sub-Event wählen —' : '— pick a sub-event —'}</option>
                  {(childEvents || []).map(c => (
                    <option key={c.id} value={c.id}>
                      {(c.title || '').split('|').slice(-1)[0].trim()}
                      {subEmails[c.id] ? ` (${subEmails[c.id].length})` : ''}
                    </option>
                  ))}
                </select>
                <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>→</span>
                <select style={{ ...inp }} value={subHotelPick} disabled={busy !== '' || subLoading}
                  onChange={e => setSubHotelPick(e.target.value)}>
                  <option value="">{isDe ? '— Hotel wählen —' : '— pick a hotel —'}</option>
                  {hotels.map(h => <option key={h.id} value={h.name}>{h.name}</option>)}
                </select>
                <button type="button" className="btn btn-secondary" style={{ fontSize: '0.76rem', padding: '5px 12px' }}
                  disabled={busy !== '' || subLoading || !subPick || !subHotelPick}
                  onClick={() => { void assignWholeSub(); }}>
                  {subLoading ? (isDe ? 'Lädt …' : 'Loading …') : (isDe ? 'Zuordnen' : 'Assign')}
                </button>
                <span style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', flexBasis: '100%' }}>
                  {isDe
                    ? 'Bereits gesetzte Zeiträume bleiben erhalten; wer noch keinen hat, bekommt den Standard-Zeitraum.'
                    : 'Existing stay periods are kept; anyone without one gets the default template.'}
                </span>
              </div>
            )}
            {selected.size > 0 && (
              <div style={{
                display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center',
                padding: '8px 10px', marginBottom: 10, borderRadius: 8,
                background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)',
              }}>
                <strong style={{ fontSize: '0.8rem' }}>
                  {isDe ? `${selected.size} markiert:` : `${selected.size} selected:`}
                </strong>
                {hotels.map(h => (
                  <button key={h.id} type="button" className="btn btn-secondary" disabled={busy !== ''}
                    style={{ fontSize: '0.74rem', padding: '4px 10px' }}
                    onClick={() => { void assignSelectedTo(h.name); }}>
                    → {h.name}
                  </button>
                ))}
                {stays.map(st => (
                  <button key={st.id} type="button" className="btn btn-secondary" disabled={busy !== ''}
                    style={{ fontSize: '0.74rem', padding: '4px 10px' }}
                    onClick={() => { void applyStayToSelected(st); }}>
                    ⏱ {st.label}
                  </button>
                ))}
                <button type="button" disabled={busy !== ''}
                  style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: '0.74rem', color: 'var(--dex-red, #c00)', textDecoration: 'underline' }}
                  onClick={() => { void writeAssignment(people.filter(p => selected.has(p.Id)), '', '', ''); }}>
                  {isDe ? 'Zuordnung aufheben' : 'Clear assignment'}
                </button>
                {bulkProgress && (
                  <div style={{ width: '100%', marginTop: 6 }}>
                    <div style={{ height: 8, background: 'var(--dex-gray-100)', borderRadius: 999, overflow: 'hidden' }}>
                      <div style={{
                        height: '100%', transition: 'width 0.2s', background: 'var(--dex-green, #86bc25)',
                        width: `${bulkProgress.total > 0 ? Math.round((bulkProgress.done / bulkProgress.total) * 100) : 0}%`,
                      }} />
                    </div>
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)', marginTop: 3 }}>
                      {bulkProgress.done}/{bulkProgress.total} {isDe ? 'zugeordnet …' : 'assigned …'}
                    </div>
                  </div>
                )}
              </div>
            )}
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.78rem' }}>
                <thead>
                  <tr style={{ textAlign: 'left', color: 'var(--dex-gray-600)' }}>
                    <th style={{ padding: '4px 6px', width: 28 }}>
                      <input type="checkbox"
                        checked={selected.size > 0 && selected.size === people.length}
                        onChange={e => setSelected(e.target.checked ? new Set(people.map(p => p.Id)) : new Set())} />
                    </th>
                    {([
                      { k: 'name' as const, l: isDe ? 'Nachname' : 'Last name' },
                      { k: 'first' as const, l: isDe ? 'Vorname' : 'First name' },
                      { k: 'loc' as const, l: isDe ? 'Standort' : 'Location' },
                      { k: 'comp' as const, l: isDe ? 'Unternehmen' : 'Company' },
                      { k: 'wish' as const, l: isDe ? 'Hotel-Wunsch' : 'Hotel request' },
                    ]).map(col => (
                      <th key={col.k} style={{ padding: '4px 6px' }}>
                        <button type="button"
                          onClick={() => { if (sortKey === col.k) { setSortAsc(a => !a); } else { setSortKey(col.k); setSortAsc(true); } }}
                          style={{ border: 'none', background: 'none', padding: 0, cursor: 'pointer', font: 'inherit', color: 'inherit', fontWeight: sortKey === col.k ? 700 : 400 }}
                        >
                          {col.l}{sortKey === col.k ? (sortAsc ? ' ▲' : ' ▼') : ''}
                        </button>
                      </th>
                    ))}
                    <th style={{ padding: '4px 6px' }}>
                      <button type="button"
                        onClick={() => { if (sortKey === 'hotel') { setSortAsc(a => !a); } else { setSortKey('hotel'); setSortAsc(true); } }}
                        style={{ border: 'none', background: 'none', padding: 0, cursor: 'pointer', font: 'inherit', color: 'inherit', fontWeight: sortKey === 'hotel' ? 700 : 400 }}
                      >
                        {isDe ? 'Hotel' : 'Hotel'}{sortKey === 'hotel' ? (sortAsc ? ' ▲' : ' ▼') : ''}
                      </button>
                    </th>
                    <th style={{ padding: '4px 6px' }}>{isDe ? 'Zeitraum' : 'Stay'}</th>
                    <th style={{ padding: '4px 6px', textAlign: 'right' }}>{isDe ? 'Nächte' : 'Nights'}</th>
                  </tr>
                </thead>
                <tbody>
                  {people
                    .filter(p => filterHotel === '__all'
                      || (filterHotel === '__none' ? !(p.Hotel || '').trim() : (p.Hotel || '').trim() === filterHotel))
                    .filter(p => !hideNoWish || wishOf(p) !== false)
                    .filter(p => {
                      const q = search.trim().toLowerCase();
                      if (!q) return true;
                      return `${p.Nachname || ''} ${p.Vorname || ''} ${p.ParticipantName || ''} ${p.ParticipantEmail || ''} ${p.Location || ''} ${p.Company || ''} ${p.Hotel || ''}`
                        .toLowerCase().indexOf(q) >= 0;
                    })
                    .sort((a, b) => {
                      const dir = sortAsc ? 1 : -1;
                      const txt = (x: SPRegistration): string => {
                        if (sortKey === 'first') return x.Vorname || '';
                        if (sortKey === 'loc') return x.Location || '';
                        if (sortKey === 'comp') return x.Company || '';
                        if (sortKey === 'hotel') return (x.Hotel || '').trim();
                        if (sortKey === 'wish') { const w = wishOf(x); return w === true ? '1' : w === false ? '2' : '3'; }
                        return x.Nachname || x.ParticipantName || '';
                      };
                      return txt(a).localeCompare(txt(b), 'de') * dir;
                    })
                    .map(p => {
                      const from = toDay(p.HotelFrom);
                      const to = toDay(p.HotelTo);
                      const n = nightsBetween(from, to);
                      const matchesStay = stays.some(st => st.from === from && st.to === to);
                      return (
                        <tr key={p.Id} style={{ borderTop: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: '4px 6px' }}>
                            <input type="checkbox" checked={selected.has(p.Id)}
                              onChange={e => setSelected(prev => {
                                const nx = new Set(prev);
                                if (e.target.checked) nx.add(p.Id); else nx.delete(p.Id);
                                return nx;
                              })} />
                          </td>
                          <td style={{ padding: '4px 6px' }}>
                            {p.Nachname || p.ParticipantName || '—'}
                            <span style={{ display: 'block', fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{p.ParticipantEmail}</span>
                          </td>
                          <td style={{ padding: '4px 6px' }}>{p.Vorname || '—'}</td>
                          <td style={{ padding: '4px 6px' }}>{p.Location || '—'}</td>
                          <td style={{ padding: '4px 6px' }}>{p.Company || '—'}</td>
                          <td style={{ padding: '4px 6px', whiteSpace: 'nowrap' }}>
                            {((): React.ReactNode => {
                              const w = wishOf(p);
                              if (w === null) return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                              return w
                                ? <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Ja' : 'Yes'}</strong>
                                : <span style={{ color: 'var(--dex-gray-500)' }}>{isDe ? 'Nein' : 'No'}</span>;
                            })()}
                          </td>
                          <td style={{ padding: '4px 6px' }}>
                            <select style={{ ...inp, height: 26 }} value={(p.Hotel || '').trim()} disabled={busy !== ''}
                              onChange={e => {
                                const h = e.target.value;
                                void writeAssignment([p], h,
                                  from || (defaultStay ? defaultStay.from : ''),
                                  to || (defaultStay ? defaultStay.to : ''));
                              }}>
                              <option value="">{isDe ? '— kein Hotel —' : '— none —'}</option>
                              {hotels.map(h => <option key={h.id} value={h.name}>{h.name}</option>)}
                            </select>
                          </td>
                          <td style={{ padding: '4px 6px' }}>
                            {/* v28.48: Statt zweier Datumsfelder je Person nur noch
                                die Auswahl aus den oben angelegten Zeitraeumen —
                                das war der eigentliche Zweck der Vorlagen. Passt
                                ein bestehender Eintrag zu keiner Vorlage, steht er
                                als eigener Eintrag drin und geht nicht verloren. */}
                            <select
                              style={{ ...inp, height: 26, minWidth: 190 }}
                              value={matchesStay ? (stays.filter(st => st.from === from && st.to === to)[0] || { id: '' }).id : (from && to ? '__custom' : '')}
                              disabled={busy !== '' || stays.length === 0}
                              onChange={e => {
                                const v = e.target.value;
                                if (v === '__custom') return; // bestehender Sonderfall bleibt
                                const st = stays.filter(x => x.id === v)[0];
                                void writeAssignment([p], (p.Hotel || '').trim(), st ? st.from : '', st ? st.to : '');
                              }}
                            >
                              <option value="">{stays.length === 0 ? (isDe ? '— erst Zeitraum anlegen —' : '— create a template first —') : (isDe ? '— kein Zeitraum —' : '— none —')}</option>
                              {stays.map(st => (
                                <option key={st.id} value={st.id}>
                                  {st.label} · {fmtDay(st.from, isDe)} – {fmtDay(st.to, isDe)}
                                </option>
                              ))}
                              {!matchesStay && from && to && (
                                <option value="__custom">
                                  {isDe ? 'abweichend' : 'custom'}: {fmtDay(from, isDe)} – {fmtDay(to, isDe)}
                                </option>
                              )}
                            </select>
                          </td>
                          <td style={{ padding: '4px 6px', textAlign: 'right', whiteSpace: 'nowrap' }}>
                            {n > 0 ? n : '—'}
                            {n > 0 && !matchesStay && stays.length > 0 && (
                              <span title={isDe ? 'Weicht von allen Vorlagen ab' : 'Differs from all templates'}
                                style={{ marginLeft: 4, color: '#b35a00', fontWeight: 700 }}>*</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>
            {stays.length > 0 && (
              <p style={{ ...hint, margin: '10px 0 0' }}>
                {isDe ? '* = abweichender Zeitraum, passt zu keiner Vorlage.' : '* = custom period, matches no template.'}
              </p>
            )}
          </>
        )}
      </div>

      {/* ---- Import einer bestehenden Hotel-Liste ---- */}
      <HotelImportModal
        open={importOpen}
        onClose={() => { if (!importBusy) setImportOpen(false); }}
        isDe={isDe}
        people={people}
        hotels={hotels}
        onApply={applyImport}
        busy={importBusy}
        progress={importProgress}
      />

      {/* ---- Hotel-Info-Mail: gleicher Editor + Versandweg wie die QR-Mail ---- */}
      <HtmlEditorModal
        open={mailOpen}
        onClose={() => setMailOpen(false)}
        title={isDe ? 'Hotelinformationen verschicken' : 'Send hotel details'}
        value={mailBody}
        onChange={setMailBody}
        previewMode="email"
        emailSubject={mailSubject}
        onEmailSubjectChange={setMailSubject}
        emailHeading={mailHeading}
        onEmailHeadingChange={setMailHeading}
        emailSubheading={mailSubheading}
        onEmailSubheadingChange={setMailSubheading}
        previewVars={varsFor(null, hotels.filter(x => x.id === testHotelId)[0])}
        insertableVars={[
          { key: 'Vorname', label: isDe ? 'Vorname' : 'First name' },
          { key: 'Nachname', label: isDe ? 'Nachname' : 'Last name' },
          { key: 'Hotel', label: 'Hotel' },
          { key: 'HotelAdresse', label: isDe ? 'Hotel-Adresse' : 'Hotel address' },
          { key: 'Anreise', label: isDe ? 'Anreise' : 'Arrival' },
          { key: 'Abreise', label: isDe ? 'Abreise' : 'Departure' },
          { key: 'Naechte', label: isDe ? 'Nächte' : 'Nights' },
          { key: 'EventTitle', label: isDe ? 'Event-Titel' : 'Event title' },
        ]}
        previewToLine={testMode
          ? (testTo || (isDe ? '(Testempfänger wählen)' : '(pick a test recipient)'))
          : (isDe ? `${people.filter(p => (p.Hotel || '').trim()).length} zugeordnete Teilnehmer (Assistenz automatisch in CC)` : `${people.filter(p => (p.Hotel || '').trim()).length} assigned attendees (assistant auto-CC'd)`)}
        previewSubjectLine={replacePlaceholders(mailSubject, varsFor(null, hotels.filter(x => x.id === testHotelId)[0]))}
        headerExtra={
          <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 12, marginBottom: 12, background: 'var(--dex-gray-50, #f7f7f5)' }}>
            <div style={{ fontSize: '0.78rem', fontWeight: 700, marginBottom: 8, color: 'var(--dex-gray-700)' }}>
              {isDe ? 'Versand' : 'Sending'}
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: 'pointer', fontSize: '0.8rem' }}>
              <input type="radio" name="hotelMailMode" checked={testMode} onChange={() => setTestMode(true)} style={{ marginTop: 3 }} />
              <span>
                <strong>{isDe ? 'Versand testen' : 'Test the send'}</strong>
                <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                  {isDe ? 'Eine einzelne Mail mit Beispielwerten an eine selbst gewählte Adresse. Es geht garantiert nichts an Teilnehmer.' : 'A single mail with sample values to an address you pick. Nothing reaches attendees.'}
                </span>
              </span>
            </label>
            {testMode && (
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', margin: '0 0 10px 26px' }}>
                <select value={testTo} onChange={e => setTestTo(e.target.value)}
                  style={{ height: 30, fontSize: '0.78rem', padding: '0 8px', border: '1px solid var(--dex-gray-300)', borderRadius: 6, minWidth: 200 }}>
                  <option value="">{isDe ? '— Empfänger wählen —' : '— pick a recipient —'}</option>
                  {(event.organizerEmails || []).map(oe => <option key={oe} value={oe}>{oe}</option>)}
                </select>
                <select value={testHotelId} onChange={e => setTestHotelId(e.target.value)}
                  style={{ height: 30, fontSize: '0.78rem', padding: '0 8px', border: '1px solid var(--dex-gray-300)', borderRadius: 6, minWidth: 160 }}>
                  {hotels.map(h => <option key={h.id} value={h.id}>{isDe ? 'als wäre: ' : 'as if: '}{h.name}</option>)}
                </select>
              </div>
            )}
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.8rem' }}>
              <input type="radio" name="hotelMailMode" checked={!testMode} onChange={() => setTestMode(false)} style={{ marginTop: 3 }} />
              <span>
                <strong>{isDe ? `Echt verschicken (${people.filter(p => (p.Hotel || '').trim()).length} Personen)` : `Send for real (${people.filter(p => (p.Hotel || '').trim()).length} people)`}</strong>
                <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Jede zugeordnete Person bekommt ihre eigenen Daten. Eine im Anmeldeformular hinterlegte Assistenz steht automatisch im CC. Nicht zugeordnete Personen bekommen nichts.'
                    : 'Each assigned person receives their own data. An assistant named during registration is CC’d automatically. Unassigned people receive nothing.'}
                </span>
              </span>
            </label>
          </div>
        }
        extraAction={{
          label: mailSending
            ? (isDe ? 'Wird gesendet…' : 'Sending…')
            : (testMode
              ? (isDe ? 'Testmail senden' : 'Send test mail')
              : (isDe ? `An ${people.filter(p => (p.Hotel || '').trim()).length} senden` : `Send to ${people.filter(p => (p.Hotel || '').trim()).length}`)),
          disabled: mailSending || hotels.length === 0,
          onClick: () => { void sendHotelMail(); },
        }}
      />
    </div>
  );
};

export default HotelPlanningPanel;
