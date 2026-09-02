/**
 * Reparatur und Wartung: Stellvertreter-Zugriff, fehlende Spalten,
 * Organizer-Berechtigungen, Feld-Beschreibungen.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { DeloitteEvent } from '../../types';
import { EventService } from '../../services/EventService';
import { FixColumnsDetail } from '../eventContextTypes';
import { dlog } from '../../utils/debugLog';

export interface MaintenanceDeps {
  eventService: EventService;
  events: DeloitteEvent[];
  autoFixStartedRef: { current: boolean };
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean }) => Promise<boolean>;
  /** v30.67: Einmaliger Reload NACH einem Massenlauf — statt eines vollen
   *  loadEvents je updateEvent (die 429-Drossel aus v29.77). */
  loadEvents: () => Promise<void>;
}

export function makeMaintenanceActions(deps: MaintenanceDeps) {
  const { eventService, events, autoFixStartedRef, updateEvent, loadEvents } = deps;

  async function autoRepairProxyAccess(): Promise<void> {
    try {
      if (typeof window === 'undefined') return;
      if (autoFixStartedRef.current) return; // kein Doppelstart in derselben Session
      const KEY = 'dex_autoaccessfix_lastrun';
      const last = Number(window.localStorage.getItem(KEY) || '0');
      const now = Date.now();
      if (last && now - last < 24 * 60 * 60 * 1000) return; // max. 1×/Tag
      autoFixStartedRef.current = true;
      // Zeitstempel sofort setzen — verhindert Hammern bei mehreren Tabs/Reloads.
      window.localStorage.setItem(KEY, String(now));
      // Aktive Events mit Subsite, dedupliziert nach Subsite (inkl. Sub-Events).
      const seen = new Set<string>();
      const subs: string[] = [];
      for (const e of events) {
        if (e.status !== 'Active') continue;
        const s = (e.subsiteUrl || '').trim();
        if (!s || seen.has(s)) continue;
        seen.add(s); subs.push(s);
      }
      let fixedTotal = 0;
      for (const s of subs) {
        try {
          const r = await eventService.repairProxyRegistrationAccess(s);
          fixedTotal += r.authorFixed;
        } catch (err) { console.warn('[DEX] autoRepairProxyAccess failed for', s, err); }
      }
      // eslint-disable-next-line no-console
      if (fixedTotal > 0) console.info(`[DEX] autoRepairProxyAccess: ${fixedTotal} Zeile(n) auf den Teilnehmer als Autor gesetzt (${subs.length} Subsites geprüft).`);
    } catch (err) { console.warn('[DEX] autoRepairProxyAccess error:', err); }
  }

  // v24.33: Globales „Spalten fixen" über ALLE Events (Haupt + Sub) — legt
  // fehlende Spalten (inkl. der neuen Company-Spalte) an, korrigiert die
  // View-Reihenfolge, zieht das Custom-Field-Mapping nach UND trägt die
  // Unternehmenszugehörigkeit für bestehende Teilnehmer nach (Graph-Backfill).
  // onProgress treibt eine Fortschrittsanzeige (done/total + Event-Titel).
  async function fixAllEventColumns(
    onProgress?: (done: number, total: number, label: string) => void
  ): Promise<{ lists: number; columnsAdded: number; backfilled: number; errors: number; anyChange: boolean; details: FixColumnsDetail[] }> {
    // v30.58: Der Lauf sagt jetzt auch, WAS er gefunden hat. Vorher lieferte er
    // nur Zahlen („12 Spalten ergänzt") — für die Frage „warum scheitert bei
    // drei Leuten die Klammer-Anmeldung?" ist das wertlos. Jede Liste wird
    // deshalb VOR und NACH dem Fix gegen die Abfragefelder ihres Events
    // gehalten; was danach immer noch fehlt, ist die eigentliche Meldung.
    const details: FixColumnsDetail[] = [];
    const seen = new Set<string>();
    const targets: DeloitteEvent[] = [];
    for (const e of events) {
      const s = (e.subsiteUrl || '').trim();
      if (!s || seen.has(s.toLowerCase())) continue;
      seen.add(s.toLowerCase());
      targets.push(e);
    }
    const total = targets.length;
    let columnsAdded = 0; let backfilled = 0; let errors = 0;
    // v30.67: Ob überhaupt eine Zuordnung nach DEX_Events geschrieben wurde —
    // nur dann lohnt der eine Reload nach der Schleife.
    let reloadNeeded = false;
    for (let i = 0; i < targets.length; i++) {
      const ev = targets[i];
      if (onProgress) onProgress(i, total, ev.title || '');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const diagFields = (ev.eventSpecificFields || []).map((f: any) => ({ id: f.id, label: f.label, spInternalName: f.spInternalName || '' }));
      // v30.67: Die Zuordnung, die dieser Durchlauf tatsächlich persistiert hat
      // (null = nichts geschrieben) — Grundlage der Nach-Diagnose, s.u.
      let persistedFields: typeof diagFields | null = null;
      const before = await eventService.diagnoseRegistrationList(ev.subsiteUrl!, diagFields)
        .catch(() => ({ ok: false, listMissing: false, missingColumns: [], error: 'nicht lesbar' }));
      try {
        const cf = (ev.eventSpecificFields || []).map(f => ({
          id: f.id, label: f.label, type: f.type, required: f.required, options: f.options,
          visible: true,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          spInternalName: (f as any).spInternalName || '',
        }));
        const res = await eventService.fixRegistrationListColumns(ev.subsiteUrl!, {
          isB2Run: !!(ev.durchstarterCapacity || ev.funstarterCapacity),
          hasQuiz: !!(ev.quiz && ev.quiz.length > 0),
          customFields: cf,
        });
        columnsAdded += (res.added ? res.added.length : 0);
        if (res.customFieldMap && Object.keys(res.customFieldMap).length > 0) {
          // v26.13 DATENVERLUST-FIX: NICHT aus dem gestrippten `cf` neu bauen —
          // das droppte helpText (Beschreibungen!), showIf, multi, EN-Varianten
          // usw. beim Zurückschreiben. Stattdessen die VOLLEN geparsten Felder
          // behalten und NUR spInternalName nachtragen (wie der Edit-Save seit
          // v19.20). Sonst sind nach „Spalten fixen (alle Events)" alle
          // Custom-Field-Beschreibungen weg.
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const upd = (ev.eventSpecificFields || []).map((f: any) => {
            const sp = res.customFieldMap![f.id];
            return sp ? { ...f, spInternalName: sp } : { ...f };
          });
          // v30.67: skipReload — die deps-Signatur sah es vor, genutzt wurde es
          // nicht. Ohne den Schalter zog JEDES Event mit mindestens einem
          // Abfragefeld ein volles loadEvents nach sich (alle Events plus
          // Teilnehmerzähler über alle Subsites): dutzende überlappende
          // Komplett-Reloads, die 429-Drossel bis zur Nutzer-Sperre — genau
          // der v29.77-Anfragensturm. Einmal neu laden reicht, nach der Schleife.
          let mappingPersisted = false;
          try { mappingPersisted = await updateEvent(ev.id, { 'CustomFields': JSON.stringify(upd) }, { skipReload: true }); } catch { /* best-effort */ }
          if (mappingPersisted) {
            reloadNeeded = true;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            persistedFields = upd.map((f: any) => ({ id: f.id, label: f.label, spInternalName: f.spInternalName || '' }));
          }
        }
        try { const bf = await eventService.backfillCompanyForList(ev.subsiteUrl!); backfilled += bf.updated; } catch { /* best-effort */ }
        // Nach dem Fix erneut hinsehen — nur was JETZT noch fehlt, ist ein Befund.
        // v30.67: NICHT aus `events` lesen. Das ist das Array des Renders, in dem
        // diese Closure gebaut wurde; updateEvent löst zwar setEvents aus, die
        // laufende Schleife sieht davon nichts — `evFresh` war deshalb immer
        // identisch zu `ev`, jedes gerade zugeordnete Feld ging mit leerem
        // spInternalName in die Diagnose und wurde als „fehlt weiterhin"
        // gemeldet. Maßgeblich ist, was wirklich nach DEX_Events geschrieben
        // wurde; ist nichts geschrieben worden, gilt der Stand von vorher.
        const afterFields = persistedFields || diagFields;
        const after = await eventService.diagnoseRegistrationList(ev.subsiteUrl!, afterFields)
          .catch(() => ({ ok: false, listMissing: false, missingColumns: [], error: 'nicht lesbar' }));
        if (before.missingColumns.length > 0 || after.missingColumns.length > 0 || after.listMissing || !after.ok) {
          details.push({
            eventId: ev.id, eventTitle: ev.title || ev.id,
            isParent: !ev.parentEventId,
            listMissing: !!after.listMissing,
            fixedColumns: before.missingColumns.filter(m => !after.missingColumns.some(x => x.id === m.id)).map(m => m.label),
            stillMissing: after.missingColumns.map(m => `${m.label} (${m.column})`),
            error: after.error || before.error,
          });
        }
      } catch (err) {
        errors++;
        console.warn('[DEX] fixAllEventColumns failed for', ev.id, err);
        details.push({
          eventId: ev.id, eventTitle: ev.title || ev.id, isParent: !ev.parentEventId,
          listMissing: false, fixedColumns: [], stillMissing: [],
          error: String((err as Error)?.message || err),
        });
      }
    }
    if (onProgress) onProgress(total, total, '');
    // v30.67: EIN Reload für den ganzen Lauf (s. skipReload oben).
    if (reloadNeeded) {
      try { await loadEvents(); } catch (err) { console.warn('[DEX] fixAllEventColumns: loadEvents nach dem Lauf fehlgeschlagen:', err); }
    }
    return { lists: total, columnsAdded, backfilled, errors, anyChange: columnsAdded > 0 || backfilled > 0, details };
  }

  /**
   * v30.39: Berechtigungs-Reparatur über ALLE Events auf einmal.
   *
   * Der Einzel-Fix im Organizer Center (v30.37) hilft nur dem, der von dem
   * Problem schon weiß — und sichtbar wird es erst, wenn jemand vor einer
   * leeren Teilnehmerliste steht. Diese Fassung geht über den Bestand: Für
   * jeden Event-Baum (Klammer + alle Sub-Events) werden die Organizer- und
   * Co-Organizer-Adressen aus ALLEN Ebenen zusammengezogen und auf ALLEN
   * Subsites des Baums gesetzt.
   *
   * Warum die Adressen über den ganzen Baum vereinigt werden: Ein Sub-Event
   * kann einen eigenen Organizer-Eintrag tragen (aus einer Kopiervorlage oder
   * einer späteren Änderung). Wer auf einem Termin als Organizer steht, muss
   * die Klammer sehen; wer auf der Klammer steht, alle Termine — sonst
   * entsteht wieder genau die halbe Sicht, um die es geht.
   *
   * Idempotent und additiv: `addroleassignment` reicht bestehende Rechte
   * durch, es wird NICHTS entzogen. Ein Lauf ohne Änderung ist deshalb ein
   * gültiges Ergebnis, kein Fehlschlag.
   */
  async function repairAllOrganizerPermissions(
    onProgress?: (done: number, total: number, label: string) => void
  ): Promise<{ trees: number; sites: number; grants: number; unresolved: string[]; errors: number }> {
    // 1) Nach Event-Baum gruppieren: Wurzel ist parentEventId oder die eigene Id.
    const trees: Record<string, { title: string; sites: string[]; emails: Set<string> }> = {};
    for (const e of events) {
      const rootId = e.parentEventId || e.id;
      if (!trees[rootId]) trees[rootId] = { title: '', sites: [], emails: new Set<string>() };
      const t = trees[rootId];
      if (!e.parentEventId) t.title = e.title || rootId;
      const site = (e.subsiteUrl || '').trim();
      if (site && t.sites.indexOf(site) < 0) t.sites.push(site);
      const add = (arr?: string[]): void => {
        (arr || []).forEach(x => { const v = (x || '').trim(); if (v) t.emails.add(v); });
      };
      add(e.organizerEmails);
      add(e.coOrganizerEmails);
    }
    const keys = Object.keys(trees).filter(k => trees[k].sites.length > 0 && trees[k].emails.size > 0);
    const total = keys.length;
    let sites = 0; let grants = 0; let errors = 0;
    const unresolved = new Set<string>();
    for (let i = 0; i < keys.length; i++) {
      const t = trees[keys[i]];
      if (onProgress) onProgress(i, total, t.title || keys[i]);
      try {
        const r = await eventService.ensureOrganizerPermissionsMulti(
          t.sites, Array.from(t.emails).join(';')
        );
        sites += r.sites;
        grants += r.grants;
        r.unresolved.forEach(u => unresolved.add(u));
        // v30.67: fehlgeschlagene Rechtevergaben zaehlen — vorher meldete die
        // Reparatur Erfolg, ohne eine einzige Antwort gesehen zu haben.
        if (r.failed && r.failed.length) {
          errors += r.failed.length;
          r.failed.forEach(f => unresolved.add(`${f.site} (${f.scope}, HTTP ${f.status})`));
        }
      } catch (err) {
        errors++;
        console.warn('[DEX] repairAllOrganizerPermissions failed for', keys[i], err);
      }
    }
    if (onProgress) onProgress(total, total, '');
    return { trees: total, sites, grants, unresolved: Array.from(unresolved), errors };
  }

  // v26.13: Wiederherstellung von Custom-Field-Beschreibungen (helpText) und
  // weiteren Feld-Eigenschaften, die ein älterer „Spalten fixen"-Lauf
  // versehentlich aus dem CustomFields-JSON gestrippt hatte. Quelle ist die
  // SharePoint-Versionshistorie des DEX_Events-Items: pro Feld wird die JÜNGSTE
  // ältere Version gesucht, die die jeweilige Eigenschaft noch enthielt, und nur
  // FEHLENDE Werte im aktuellen Stand aufgefüllt (nie etwas überschrieben).
  async function restoreCustomFieldDescriptions(
    onProgress?: (done: number, total: number, label: string) => void,
    dryRun?: boolean
  ): Promise<{ events: number; eventsChanged: number; fieldsRestored: number; errors: number; details: Array<{ eventId: string; eventTitle: string; fields: Array<{ label: string; props: string[] }> }> }> {
    // v30.67: audienceOnly — der Loader hat das Flag bis v30.66 nie zurückgelesen
    // (eventMapping), jeder Save danach hat es aus dem JSON entfernt. Aus der
    // Versionshistorie ist es wiederherstellbar wie die anderen Drop-Opfer.
    const RESTORE_PROPS = ['helpText', 'helpTextEn', 'helpTextStyle', 'showIf', 'multi', 'externalLinks', 'ccOnEmails', 'onlyForGroup', 'confirmLabel', 'confirmLabelEn', 'labelEn', 'optionsEn', 'audienceOnly'];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const hasVal = (p: string, val: any): boolean => {
      if (p === 'multi' || p === 'ccOnEmails' || p === 'audienceOnly') return val === true;
      if (p === 'optionsEn' || p === 'externalLinks') return Array.isArray(val) && val.length > 0;
      return typeof val === 'string' ? val.trim().length > 0 : (val !== undefined && val !== null);
    };
    const seen = new Set<string>();
    const targets = (events || []).filter(e => { const id = String(e.id); if (seen.has(id)) return false; seen.add(id); return true; });
    const total = targets.length;
    let eventsChanged = 0; let fieldsRestored = 0; let errors = 0;
    let written = 0; // v30.67: wie viele Events wirklich geschrieben wurden → ein Reload am Ende
    const details: Array<{ eventId: string; eventTitle: string; fields: Array<{ label: string; props: string[] }> }> = [];
    for (let i = 0; i < total; i++) {
      const ev = targets[i];
      if (onProgress) onProgress(i, total, ev.title || '');
      try {
        const versions = await eventService.getEventCustomFieldsVersions(Number(ev.id));
        // v26.13: Diagnose — pro Event ausgeben, wie viele Versionen es gibt und
        // ob in der Historie überhaupt ein helpText vorkommt (sonst ist nichts
        // wiederherzustellen). Hilft beim Nachvollziehen in der Browser-Konsole.
        const helpInHistory = (versions || []).some(v => (v.customFields || '').indexOf('helpText') >= 0);
        // eslint-disable-next-line no-console
        dlog('perf', '[DEX restore]', ev.title, '(id', ev.id + ') — Versionen:', (versions || []).length, '— helpText in Historie:', helpInHistory);
        if (!versions || versions.length === 0) continue;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let currentArr: any[];
        try { currentArr = JSON.parse(versions[0].customFields || '[]'); } catch { currentArr = []; }
        if (!Array.isArray(currentArr) || currentArr.length === 0) continue;
        // Pro Feld-Id den jüngsten vorhandenen Wert je Eigenschaft sammeln
        // (Versionen sind bereits neueste-zuerst sortiert → erster Treffer gilt).
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const best: Record<string, Record<string, any>> = {};
        for (const ver of versions) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          let arr: any[]; try { arr = JSON.parse(ver.customFields || '[]'); } catch { continue; }
          if (!Array.isArray(arr)) continue;
          for (const f of arr) {
            const id = f && f.id; if (!id) continue;
            if (!best[id]) best[id] = {};
            for (const p of RESTORE_PROPS) {
              if (best[id][p] !== undefined) continue;
              if (hasVal(p, f[p])) best[id][p] = f[p];
            }
          }
        }
        let changed = false;
        const evFields: Array<{ label: string; props: string[] }> = [];
        const restoredArr = currentArr.map((f) => {
          const id = f && f.id; if (!id || !best[id]) return f;
          const out = { ...f };
          const restoredProps: string[] = [];
          for (const p of RESTORE_PROPS) {
            if (!hasVal(p, out[p]) && best[id][p] !== undefined) {
              out[p] = best[id][p]; changed = true; fieldsRestored++; restoredProps.push(p);
            }
          }
          if (restoredProps.length > 0) evFields.push({ label: (f.label || f.id || '?'), props: restoredProps });
          return out;
        });
        if (changed) {
          if (evFields.length > 0) details.push({ eventId: String(ev.id), eventTitle: ev.title || String(ev.id), fields: evFields });
          // v26.13: Trockenlauf — nur ermitteln, NICHT schreiben.
          // v30.67: skipReload — derselbe Anfragensturm wie in fixAllEventColumns
          // (ein voller loadEvents je Event); einmal neu laden nach der Schleife.
          if (!dryRun) { await updateEvent(ev.id, { 'CustomFields': JSON.stringify(restoredArr) }, { skipReload: true }); written++; }
          eventsChanged++;
        }
      } catch (e) { errors++; console.warn('[DEX] restoreCustomFieldDescriptions failed for', ev.id, e); }
    }
    if (onProgress) onProgress(total, total, '');
    if (written > 0) {
      try { await loadEvents(); } catch (err) { console.warn('[DEX] restoreCustomFieldDescriptions: loadEvents nach dem Lauf fehlgeschlagen:', err); }
    }
    return { events: total, eventsChanged, fieldsRestored, errors, details };
  }

  return { autoRepairProxyAccess, fixAllEventColumns, repairAllOrganizerPermissions, restoreCustomFieldDescriptions };
}
