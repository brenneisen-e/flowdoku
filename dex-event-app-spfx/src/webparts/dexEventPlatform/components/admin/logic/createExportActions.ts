/* createExportActions — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 3055-3415 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { B2RUN_KOELN_ALTERSKLASSE, B2RUN_KOELN_HEADERS, isB2RunKoelnTitle, mapAnredeToB2Run, mapStarterTypeToStartblock } from '../../../data/b2runKoeln';
import { SPRegistration } from '../../../services/EventService';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { DeloitteEvent } from '../../../types';

export interface CreateExportActionsCtx {
  computeRoommatePairs: (rows: SPRegistration[]) => Array<[SPRegistration, SPRegistration]>;
  consolidatedChildren: DeloitteEvent[];
  hasRoommateColumn: boolean;
  isDe: boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

export interface CreateExportActionsResult {
  exportConsolidatedExcel: (audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled', includeMatrix: boolean, subIds: string[]) => void;
  exportCsv: (mode: 'deloitte' | 'b2run', audience?: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled') => void;
}

export function createExportActions(ctx: CreateExportActionsCtx): CreateExportActionsResult {
  const {
    computeRoommatePairs, consolidatedChildren, hasRoommateColumn, isDe, registrations,
    selectedEvent, showAlert, subEventRegsByEventId,
  } = ctx;
  /**
   * CSV Export für Teilnehmerlisten.
   * - 'deloitte': alle internen Felder (Anrede, Name, Email, Department, Location, JobTitle, Phone, Status, ...)
   * - 'b2run': Format exakt wie die offizielle B2Run-Köln-Meldedatei (16 Spalten laut B2RUN_KOELN_HEADERS: Nr., Anrede, Vorname, Nachname, E-Mail, Startblock, Zustimmung AGB & Datenschutzhinweise, Anonym, Gruppe, Straße/PLZ/Stadt (privat), Mobilnummer, Verwendung Infoservice, Altersklasse, Nordic Walker)
   */
  const exportCsv = (mode: 'deloitte' | 'b2run', audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled' = 'active'): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const audienceFilter = (r: SPRegistration): boolean => {
      if (audience === 'waitOnly') return r.Status === 'Warteliste';
      if (audience === 'activePlusWait') return ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste';
      // v20.4: alles inkl. Abgemeldete (Status-Spalte ist im Export enthalten).
      if (audience === 'withCancelled') return true;
      return ACTIVE.indexOf(r.Status) >= 0;
    };
    // v17.12: nach TeilnehmerID asc sortieren (vorher random / Status-Reihenfolge).
    const activeRegsForExport = registrations
      .filter(audienceFilter)
      .slice()
      .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
    if (activeRegsForExport.length === 0) { showAlert('Keine Teilnehmer zum Exportieren.'); return; }

    // v20.0 (Audit): toter CSV-Escaper `esc` entfernt — seit dem Umstieg auf
    // natives XLSX (v8.4) wurde er nie mehr aufgerufen.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parseCustom = (json: string): Record<string, any> => {
      try { return JSON.parse(json || '{}'); } catch { return {}; }
    };

    let headers: string[] = [];
    let rows: (string | number)[][] = [];

    if (mode === 'b2run') {
      // v26.48: Struktur exakt wie die OFFIZIELLE B2Run-Köln-Meldedatei
      // (Deloitte_Teilnehmer_-innen_b2run-koeln-<jahr>.xlsx) — 16 Spalten
      // inkl. „Straße" mit ß und der neuen Spalte „Nordic Walker".
      // Zentrale Spec in data/b2runKoeln.ts.
      headers = [...B2RUN_KOELN_HEADERS];
      rows = activeRegsForExport.map((r, idx) => {
        const cd = parseCustom(r.CustomData || '{}');
        const vorname = r.Vorname || (r.ParticipantName || '').split(' ').slice(0, -1).join(' ') || '';
        const nachname = r.Nachname || (r.ParticipantName || '').split(' ').slice(-1).join(' ') || '';
        return [
          idx + 1, // Nr. — laufende Nummer 1..n (die offizielle Datei nummeriert fortlaufend, NICHT TeilnehmerID)
          cd.b2run_geschlecht || mapAnredeToB2Run(r.Anrede), // 'männlich'/'weiblich'/'divers' (klein, wie Original)
          vorname,
          nachname,
          r.ParticipantEmail || '',
          cd.b2run_startblock || mapStarterTypeToStartblock(r.StarterType),
          cd.b2run_datenschutz ? 'Ja' : 'Nein',
          cd.b2run_anonym ? 'Ja' : 'Nein',
          cd.b2run_gruppe || '',
          '', // Straße und Hausnummer (privat) — nicht abgefragt, darf leer bleiben
          '', // PLZ (privat) — nicht abgefragt
          '', // Stadt (privat) — nicht abgefragt
          cd.b2run_mobilnummer || '',
          cd.b2run_infoservice ? 1 : 0, // Original-Datei nutzt 0/1 (Zahl), nicht Ja/Nein
          cd.b2run_altersklasse || B2RUN_KOELN_ALTERSKLASSE,
          cd.b2run_nordicwalker ? 'Ja' : 'Nein',
        ];
      });
    } else {
      // Deloitte View: alle internen Felder
      // v23.7: Team-Spalte nur bei Team-Events (oder wenn überhaupt eine
      // Team-Zuordnung existiert) — der frei benannte Begriff als Spaltenkopf.
      const includeTeam = !!selectedEvent.teamRegistrationEnabled || activeRegsForExport.some(r => !!r.TeamId);
      const teamHeader = selectedEvent.teamTermSingular || 'Team';
      // v30.48: Startnummer nur, wenn sie importiert wurde — sonst eine leere
      // Spalte in jedem Export.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const includeBib = activeRegsForExport.some(r => String((r as any).Startnummer || '').trim());
      headers = [
        'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'Email',
        'Department', 'Location', 'JobTitle', 'Phone',
        'Status', 'RegistrationDate',
        ...(includeBib ? ['Startnummer'] : []),
        ...(includeTeam ? [teamHeader, `${teamHeader}-Lead`] : []),
      ];
      // Dynamisch alle Custom Field Labels aus dem Event sammeln
      const customLabels: Array<{ id: string; label: string }> = (selectedEvent.eventSpecificFields || []).map(f => ({ id: f.id, label: f.label }));
      headers = headers.concat(customLabels.map(cf => cf.label));

      // v26.44: „Roommate-Match"-Spalte — nur wenn das Event eine Roommate-
      // Spalte hat. Paare werden über die VOLLE Export-Zeilenmenge berechnet
      // (nicht über den UI-Suchfilter), gleiche Dedupe-Logik wie die
      // „Matches anzeigen"-Gruppierung in der Teilnehmer-Tabelle.
      const roommatePairLabelByEmail: Record<string, string> = {};
      if (hasRoommateColumn) {
        const nameOf = (x: SPRegistration): string =>
          `${x.Vorname || ''} ${x.Nachname || ''}`.trim() || x.ParticipantName || x.ParticipantEmail || '';
        computeRoommatePairs(activeRegsForExport).forEach(([a, b], pi) => {
          const ea = (a.ParticipantEmail || '').trim().toLowerCase();
          const eb = (b.ParticipantEmail || '').trim().toLowerCase();
          roommatePairLabelByEmail[ea] = `Match ${pi + 1} (mit ${nameOf(b)})`;
          roommatePairLabelByEmail[eb] = `Match ${pi + 1} (mit ${nameOf(a)})`;
        });
        headers.push('Roommate-Match');
      }

      rows = activeRegsForExport.map(r => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const anyReg = r as any;
        const cd = parseCustom(r.CustomData || '{}');
        const base = [
          String(r.TeilnehmerID || ''),
          r.Anrede || '',
          r.Vorname || '',
          r.Nachname || '',
          r.ParticipantEmail || '',
          anyReg.Department || '',
          anyReg.Location || '',
          anyReg.JobTitle || '',
          anyReg.Phone || '',
          r.Status || '',
          r.RegistrationDate ? new Date(r.RegistrationDate).toLocaleString('de-DE') : '',
          ...(includeBib ? [String(anyReg.Startnummer || '')] : []),
          ...(includeTeam ? [r.TeamName || '', r.TeamLead ? 'Ja' : ''] : []),
        ];
        const customValues = customLabels.map(cf => {
          const v = cd[cf.id];
          if (v === undefined || v === null) return '';
          if (typeof v === 'boolean') return v ? 'Ja' : 'Nein';
          return String(v);
        });
        const row = base.concat(customValues);
        if (hasRoommateColumn) {
          row.push(roommatePairLabelByEmail[(r.ParticipantEmail || '').trim().toLowerCase()]
            || 'Ohne Preferred Roommate oder Match');
        }
        return row;
      });
    }

    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    // XLSX Export — natives Excel-Format, automatische Spalten-Breiten, keine
    // CSV-Escaping-Quirks. Gilt für beide Modi (Teilnehmerliste + B2Run).
    const aoa: (string | number)[][] = [headers, ...rows];
    // v26.48: Sheet-Name wie in der offiziellen Meldedatei („B2Run Köln <Jahr>",
    // ≤31 Zeichen — XLSX-Limit unkritisch). Jahr aus dem Event-Startdatum.
    const b2runYear = selectedEvent.startDate ? String(new Date(selectedEvent.startDate).getFullYear()) : '';
    const sheetName = mode === 'b2run' ? ('B2Run Köln ' + b2runYear).trim() : 'Teilnehmer';
    const filePrefix = mode === 'b2run' ? 'B2Run' : 'Teilnehmer';
    // v26.48: Bei B2Run-Köln-Events exakt der offizielle Dateiname des
    // Veranstalters (Deloitte_Teilnehmer_-innen_b2run-koeln-<jahr>.xlsx);
    // sonst bleibt das bisherige Namensschema.
    const fileName = mode === 'b2run' && isB2RunKoelnTitle(selectedEvent.title)
      ? `Deloitte_Teilnehmer_-innen_b2run-koeln${b2runYear ? '-' + b2runYear : ''}.xlsx`
      : `${filePrefix}_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // v20.0 (Audit): xlsx erst beim Export-Klick als Chunk nachladen — die
    // Bibliothek ist mit Abstand die schwerste Dependency und wird nur hier
    // gebraucht. Der .then/.catch-Pfad ersetzt das frühere try/catch.
    // v8.4: Manueller Blob-Download statt XLSX.writeFile. Im SPFx-Iframe-
    // Context ist saveAs/createObjectURL häufig blockiert (CORS / Sandbox-
    // Policies), wodurch der Download stillschweigend nicht startet. Mit
    // anchor.click() läuft das in jeder Browser-Umgebung zuverlässig.
    import('xlsx').then(XLSX => {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const colWidths = headers.map((h, ci) => {
        const maxLen = Math.max(h.length, ...rows.map(r => String(r[ci] || '').length));
        return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = colWidths;
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    }).catch(err => {
      console.warn('[DEX] Excel-Export fehlgeschlagen:', err);
      showAlert(isDe
        ? 'Excel-Export fehlgeschlagen. Bitte Browser-Console prüfen.'
        : 'Excel export failed. Please check the browser console.');
    });
  };

  // v20.4: Excel-Export der konsolidierten Klammer-Ansicht. Baut EINE Datei
  // mit (wählbar) einem Matrix-Blatt — eine Zeile pro Person, Spalten =
  // Stammdaten + Klammer-Felder + pro Sub-Event der Status + dessen Feld-
  // Antworten — und/oder je einem eigenen Blatt pro gewähltem Sub-Event.
  // Datenquellen sind die bereits geladenen States (registrations = Klammer-
  // Zeilen, subEventRegsByEventId = Sub-Event-Listen) — kein Extra-Roundtrip.
  const exportConsolidatedExcel = (
    audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled',
    includeMatrix: boolean,
    subIds: string[]
  ): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const matches = (r: SPRegistration): boolean => {
      if (audience === 'waitOnly') return r.Status === 'Warteliste';
      if (audience === 'activePlusWait') return ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste';
      if (audience === 'withCancelled') return true;
      return ACTIVE.indexOf(r.Status) >= 0;
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parseCustom = (json: string): Record<string, any> => {
      try { return JSON.parse(json || '{}'); } catch { return {}; }
    };
    const fieldVal = (cd: Record<string, unknown>, id: string): string => {
      const v = cd[id];
      if (v === undefined || v === null) return '';
      if (typeof v === 'boolean') return v ? 'Ja' : 'Nein';
      return String(v);
    };
    const chosenChildren = consolidatedChildren.filter(c => subIds.indexOf(c.id) >= 0);
    const sheets: Array<{ name: string; headers: string[]; rows: string[][] }> = [];
    const sanitizeSheet = (s: string): string => (s || 'Blatt').replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31) || 'Blatt';

    if (includeMatrix) {
      const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.label);
      type PersonRow = {
        vorname: string; nachname: string; email: string; jobTitle: string; location: string;
        teamName: string;
        parentCd: Record<string, unknown>;
        perChild: Record<string, SPRegistration | undefined>;
        hasMatch: boolean;
      };
      const persons: Record<string, PersonRow> = {};
      const ensurePerson = (r: SPRegistration): PersonRow => {
        const key = (r.ParticipantEmail || '').toLowerCase().trim();
        if (!persons[key]) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const anyReg = r as any;
          persons[key] = {
            vorname: r.Vorname || '', nachname: r.Nachname || '',
            email: r.ParticipantEmail || '',
            jobTitle: anyReg.JobTitle || '', location: anyReg.Location || '',
            teamName: '',
            parentCd: {}, perChild: {}, hasMatch: false,
          };
        }
        // v23.7: Team-Name aus der ersten Zeile übernehmen, die einen hat.
        if (r.TeamName && !persons[key].teamName) persons[key].teamName = r.TeamName;
        return persons[key];
      };
      for (const r of registrations) {
        const p = ensurePerson(r);
        p.parentCd = parseCustom(r.CustomData || '{}');
        if (matches(r)) p.hasMatch = true;
      }
      for (const child of consolidatedChildren) {
        const regs = subEventRegsByEventId[child.id] || [];
        for (const r of regs) {
          const p = ensurePerson(r);
          if (matches(r)) {
            p.perChild[child.id] = r;
            p.hasMatch = true;
          }
        }
      }
      // v23.7: Team-Spalte nur, wenn überhaupt Team-Zuordnungen existieren.
      const anyTeam = Object.keys(persons).some(k => !!persons[k].teamName);
      const teamHdr = selectedEvent.teamTermSingular || 'Team';
      const matrixHeaders: string[] = ['Vorname', 'Nachname', 'Email', 'JobTitle', 'Standort']
        .concat(anyTeam ? [teamHdr] : [])
        .concat(parentFields.map(f => f.label));
      for (const child of consolidatedChildren) {
        const short = shortSubEventTitle(child.title, selectedEvent.title) || child.title || '?';
        matrixHeaders.push(short);
        for (const f of (child.eventSpecificFields || []).filter(ff => ff.label)) {
          matrixHeaders.push(`${short}: ${f.label}`);
        }
      }
      const matrixRows: string[][] = Object.keys(persons)
        .map(k => persons[k])
        .filter(p => p.hasMatch)
        .sort((a, b) => (a.nachname || '').localeCompare(b.nachname || '', 'de') || (a.vorname || '').localeCompare(b.vorname || '', 'de'))
        .map(p => {
          const row: string[] = [p.vorname, p.nachname, p.email, p.jobTitle, p.location]
            .concat(anyTeam ? [p.teamName || ''] : [])
            .concat(parentFields.map(f => fieldVal(p.parentCd, f.id)));
          for (const child of consolidatedChildren) {
            const reg = p.perChild[child.id];
            row.push(reg ? (reg.Status || '') : '');
            const cd = reg ? parseCustom(reg.CustomData || '{}') : {};
            for (const f of (child.eventSpecificFields || []).filter(ff => ff.label)) {
              row.push(reg ? fieldVal(cd, f.id) : '');
            }
          }
          return row;
        });
      sheets.push({ name: 'Konsolidiert', headers: matrixHeaders, rows: matrixRows });
    }

    for (const child of chosenChildren) {
      const regs = (subEventRegsByEventId[child.id] || [])
        .filter(matches)
        .slice()
        .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
      const childFields = (child.eventSpecificFields || []).filter(f => f.label);
      // v23.7: Team-Spalte je Sub-Event-Blatt, wenn dort Team-Zuordnungen sind.
      const childAnyTeam = regs.some(r => !!r.TeamName);
      const childTeamHdr = selectedEvent.teamTermSingular || 'Team';
      const headers = ['TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'Email', 'Department', 'Location', 'JobTitle', 'Status', 'RegistrationDate']
        .concat(childAnyTeam ? [childTeamHdr] : [])
        .concat(childFields.map(f => f.label));
      const rows = regs.map(r => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const anyReg = r as any;
        const cd = parseCustom(r.CustomData || '{}');
        return [
          String(r.TeilnehmerID || ''), r.Anrede || '', r.Vorname || '', r.Nachname || '',
          r.ParticipantEmail || '', anyReg.Department || '', anyReg.Location || '', anyReg.JobTitle || '',
          r.Status || '', r.RegistrationDate ? new Date(r.RegistrationDate).toLocaleString('de-DE') : '',
        ].concat(childAnyTeam ? [r.TeamName || ''] : []).concat(childFields.map(f => fieldVal(cd, f.id)));
      });
      sheets.push({ name: sanitizeSheet(shortSubEventTitle(child.title, selectedEvent.title) || child.title || 'Sub-Event'), headers, rows });
    }

    if (sheets.length === 0) { showAlert(isDe ? 'Bitte mindestens die Matrix oder ein Sub-Event auswählen.' : 'Please select at least the matrix or one sub-event.'); return; }
    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = `Konsolidiert_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    import('xlsx').then(XLSX => {
      const wb = XLSX.utils.book_new();
      const usedNames = new Set<string>();
      for (const sheet of sheets) {
        const aoa: (string | number)[][] = [sheet.headers, ...sheet.rows];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const colWidths = sheet.headers.map((h, ci) => {
          const maxLen = Math.max(h.length, ...sheet.rows.map(r => String(r[ci] || '').length));
          return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
        });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (ws as any)['!cols'] = colWidths;
        // Doppelte Blattnamen entschärfen (xlsx verlangt eindeutige Namen).
        let name = sheet.name;
        let i = 2;
        while (usedNames.has(name)) { name = `${sheet.name.slice(0, 28)}_${i}`; i++; }
        usedNames.add(name);
        XLSX.utils.book_append_sheet(wb, ws, name);
      }
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    }).catch(err => {
      console.warn('[DEX] Konsolidierter Excel-Export fehlgeschlagen:', err);
      showAlert(isDe ? 'Excel-Export fehlgeschlagen. Bitte Browser-Console prüfen.' : 'Excel export failed. Please check the browser console.', { variant: 'error' });
    });
  };
  return {
    exportConsolidatedExcel, exportCsv,
  };
}

