/* useColumnConfig — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 2817-3053 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface UseColumnConfigCtx {
  allEvents: DeloitteEvent[];
  columnOrder: string[];
  hiddenColumns: string[];
  isDe: boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setColumnOrder: React.Dispatch<React.SetStateAction<string[]>>;
  setHiddenColumns: React.Dispatch<React.SetStateAction<string[]>>;
}

export interface UseColumnConfigResult {
  availableColumns: { id: string; label: string; alwaysVisible?: boolean; }[];
  hasRoommateColumn: boolean;
  hasWaitlistActivity: boolean;
  hideColumn: (id: string) => void;
  moveColumn: (id: string, direction: -1 | 1) => void;
  showColumn: (id: string) => void;
}

export function useColumnConfig(ctx: UseColumnConfigCtx): UseColumnConfigResult {
  const {
    allEvents, columnOrder, hiddenColumns, isDe, registrations, selectedEvent, setColumnOrder,
    setHiddenColumns,
  } = ctx;
  // v6.17: Verfügbare Spalten der Teilnehmer-Tabelle aufbauen. MUSS vor dem
  // early return `if (!selectedEvent) return ...` stehen — sonst verletzen
  // die Hooks die Rules-of-Hooks (unterschiedliche Hook-Anzahl pro Render =
  // React Error #310).
  // v19.11: Hat dieses Event überhaupt Warteliste-/Nachrück-Aktivität? Nur dann
  // sind die Nachrück-Audit-Spalten („Nachgerückt am", „Hat ersetzt", „Wurde
  // ersetzt durch") sinnvoll. `waitlistEnabled` allein reicht NICHT, weil es per
  // Default `true` ist (e.WaitlistEnabled !== false) — Events ohne konfigurierte
  // Warteliste hätten sonst immer die leeren Audit-Spalten. „Aktiv" = jemand
  // steht auf der Warteliste ODER es gibt bereits Nachrück-Daten.
  const hasWaitlistActivity = React.useMemo(() => registrations.some(r =>
    r.Status === 'Warteliste'
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    || !!(r as any).PromotedDate || !!(r as any).ReplacedParticipantEmail || !!(r as any).ReplacedByParticipantEmail
  ), [registrations]);
  const availableColumns = React.useMemo(() => {
    const isSplit = !!selectedEvent
      && typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
    const userIds = (selectedEvent?.eventSpecificFields || [])
      .filter(f => f.type === 'user' || f.type === 'roommate')
      .map(f => f.id);
    // v14.8: Anrede-Spalte nur anbieten, wenn das Event die Anrede beim
    // Anmelden tatsächlich abfragt (askSalutation). Sonst landet eine leere
    // Spalte voller „-" im Admin-Center, die niemand braucht.
    const askSal = !!selectedEvent?.askSalutation;
    const cols: Array<{ id: string; label: string; alwaysVisible?: boolean }> = [
      { id: 'id', label: '#', alwaysVisible: true },
      ...(askSal ? [{ id: 'anrede', label: 'Anrede' }] : []),
      // v11.26: getrennte Vorname / Nachname Spalten statt der einen
      // kombinierten 'name'-Spalte. Alte localStorage-Einträge mit 'name'
      // werden im useEffect-Loader unten in 'vorname','nachname' migriert.
      { id: 'vorname', label: 'Vorname', alwaysVisible: true },
      { id: 'nachname', label: 'Nachname', alwaysVisible: true },
      { id: 'email', label: 'Email' },
      { id: 'jobTitle', label: 'Job Title' },
      { id: 'location', label: 'Standort' },
      // v24.33: Unternehmenszugehörigkeit / Rechtsträger als eigene Spalte.
      { id: 'company', label: isDe ? 'Unternehmen' : 'Company' },
    ];
    // v11.6: bei Split-Capacity die frei wählbaren Gruppen-Labels nutzen
    // (Fallback auf 'Starter-Typ' wenn keine Labels gesetzt sind).
    if (isSplit) {
      const lblA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || '';
      const lblB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || '';
      const colLabel = (lblA && lblB) ? `${lblA} / ${lblB}` : (isDe ? 'Gruppe' : 'Group');
      cols.push({ id: 'starterType', label: colLabel });
    }
    // v30.48: Startnummer nur anbieten, wenn tatsächlich eine importiert wurde.
    // Ohne Import wäre es eine Spalte voller „—" an jedem Event, das zufällig
    // die Spalte in der Liste hat.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (registrations.some(r => String((r as any).Startnummer || '').trim())) {
      cols.push({ id: 'startnummer', label: isDe ? 'Startnummer' : 'Bib number' });
    }
    cols.push({ id: 'status', label: 'Status' });
    cols.push({ id: 'date', label: 'Registriert am' });
    // v17.15/v17.17.1: Nachrück-Audit-Spalten — nur sichtbar wenn das
    // Event überhaupt eine Warteliste haben KANN (waitlistEnabled UND
    // maxParticipants > 0). Bei „Unbegrenzt"-Events kommt nie jemand auf
    // die Warteliste, deshalb sind die drei Audit-Spalten ohne Inhalt.
    // v19.11: Zusätzlich `hasWaitlistActivity` — Events OHNE echte Warteliste
    // (Default waitlistEnabled=true, aber niemand wartet/nachgerückt) zeigen die
    // leeren Audit-Spalten jetzt nicht mehr.
    // v30.67: `|| isSplit` — bei geteilten Kapazitäten ist `maxParticipants`
    // 0, die Audit-Spalten erschienen dort nie (dieselbe Verwechslung wie bei
    // der KPI-Kachel „Warteliste", KpiTiles.tsx).
    if (selectedEvent?.waitlistEnabled && ((selectedEvent?.maxParticipants || 0) > 0 || isSplit) && hasWaitlistActivity) {
      cols.push({ id: 'promotedDate', label: 'Nachgerückt am' });
      // v19.4: „Hat ersetzt" = die abgemeldete Person, deren Platz diese Person
      // übernommen hat. „Wurde ersetzt durch" wandert in die Abmeldungen-Tabelle
      // (gehört zur abgemeldeten Person, nicht zur aktiven).
      cols.push({ id: 'replaced', label: 'Hat ersetzt' });
    }
    cols.push({ id: 'registeredBy', label: 'Registriert von' });
    // v16.1: Team-Spalte — zeigt pro Teilnehmer den Team-Namen (falls Team-
    // Anmeldung aktiv und der TN in einem Team ist).
    if (selectedEvent?.teamRegistrationEnabled) {
      cols.push({ id: 'team', label: 'Team' });
    }
    if (userIds.length > 0) {
      // v11.56: Label aus dem ersten roommate-/user-Feld ableiten, statt hart
      // „Zimmerpartner" zu nennen. Wenn ein roommate-Feld existiert, nimm dessen
      // Label (User-Picker-Pairs); sonst das erste user-Feld; Fallback bleibt
      // der deutsche Default.
      const fields = selectedEvent?.eventSpecificFields || [];
      const firstRoommate = fields.filter(f => f.type === 'roommate' && f.label && f.label.trim())[0];
      const firstUser = fields.filter(f => f.type === 'user' && f.label && f.label.trim())[0];
      const roommateLabel = (firstRoommate?.label || firstUser?.label || 'Zimmerpartner').trim();
      cols.push({ id: 'roommate', label: roommateLabel });
    }
    // v14.11: Wenn ein Sub-Event selektiert ist, blenden wir die
    // Custom-Fields des Parent-Events (Pastel A) zusätzlich ein. Die
    // eigenen Sub-Event-Fields (Pastel B) folgen direkt danach. ID-
    // Präfix `cfp-` unterscheidet Parent- von Sub-Event-Feldern (`cf-`).
    const parentForCols: DeloitteEvent | null = (selectedEvent && selectedEvent.parentEventId)
      ? (allEvents.find(e => e.id === selectedEvent.parentEventId) || null)
      : null;
    if (parentForCols) {
      const ownIds = new Set((selectedEvent?.eventSpecificFields || []).map(f => f.id));
      // v19.10: 'roommate' (wie 'user') NICHT als generische Spalte ausgeben —
      // diese Felder werden bereits über die dedizierte „roommate"-Spalte (mit
      // Match-Badge) gerendert. Sonst erscheint das Feld DOPPELT (einmal mit
      // Match, einmal als roher „Name <email>"-Text).
      for (const f of (parentForCols.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'roommate' && f.label && f.label.trim())) {
        // Sub-Events erben Parent-Felder evtl. 1:1 (Wizard kopiert das beim
        // Anlegen). Nicht doppelt ausgeben, wenn das eigene Feld die
        // gleiche ID hat — in dem Fall reicht die Sub-Event-Spalte.
        if (ownIds.has(f.id)) continue;
        cols.push({ id: `cfp-${f.id}`, label: f.label });
      }
    }
    // v19.10: 'roommate'-Felder (wie 'user') hier ausschließen — sie haben
    // bereits die dedizierte „roommate"-Spalte mit Match-Badge. Vorher fehlte
    // `f.type !== 'roommate'`, deshalb erschien ein Zimmerpartner-Feld DOPPELT:
    // einmal als Match-Spalte, einmal als generische cf-Spalte mit rohem
    // „Nachname, Vorname <email>"-Text.
    for (const f of (selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'roommate' && f.label && f.label.trim())) {
      cols.push({ id: `cf-${f.id}`, label: f.label });
    }
    cols.push({ id: 'action', label: 'Aktion', alwaysVisible: true });
    return cols;
  }, [
    selectedEvent?.id,
    selectedEvent?.parentEventId,
    selectedEvent?.durchstarterCapacity,
    selectedEvent?.funstarterCapacity,
    (selectedEvent?.eventSpecificFields || []).map(f => `${f.id}:${f.type}:${f.label}`).join(','),
    // v14.11: Parent-Custom-Fields als Dep
    (() => {
      if (!selectedEvent?.parentEventId) return '';
      const p = allEvents.find(e => e.id === selectedEvent.parentEventId);
      return (p?.eventSpecificFields || []).map(f => `${f.id}:${f.type}:${f.label}`).join(',');
    })(),
    // v19.11: Audit-Spalten-Sichtbarkeit hängt an der Warteliste-Aktivität.
    hasWaitlistActivity,
    // v30.48: Die Startnummern-Spalte erscheint erst nach dem Import — also
    // sobald irgendeine Zeile eine Nummer trägt.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    registrations.some(r => String((r as any).Startnummer || '').trim()),
  ]);

  // v26.44: gibt es überhaupt eine Roommate-Spalte? Steuert den
  // „Matches anzeigen"-Toggle, die Paar-Gruppierung der Teilnehmer-Tabelle
  // und die „Roommate-Match"-Spalte im Excel-Export.
  const hasRoommateColumn = availableColumns.some(c => c.id === 'roommate');

  const columnStorageKey = selectedEvent ? `dex_admin_columns_${selectedEvent.id}` : '';
  // localStorage-Load beim Event-Wechsel.
  React.useEffect(() => {
    if (!selectedEvent) { setColumnOrder([]); setHiddenColumns([]); return; }
    const allIds = availableColumns.map(c => c.id);
    try {
      const raw = localStorage.getItem(columnStorageKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.order) && Array.isArray(parsed.hidden)) {
          // v11.26: Migration alter Spaltenkonfigurationen — die zentrale
          // 'name'-Spalte wurde in 'vorname' + 'nachname' aufgeteilt. Wenn
          // ein gespeichertes Layout noch 'name' enthält, an gleicher
          // Position durch ['vorname','nachname'] ersetzen, damit der
          // User seine gewünschte Reihenfolge beibehält.
          const migratedOrder: string[] = [];
          for (const id of parsed.order as string[]) {
            if (id === 'name') {
              if (migratedOrder.indexOf('vorname') < 0) migratedOrder.push('vorname');
              if (migratedOrder.indexOf('nachname') < 0) migratedOrder.push('nachname');
            } else {
              migratedOrder.push(id);
            }
          }
          const knownOrder = migratedOrder.filter((id: string) => allIds.indexOf(id) >= 0);
          const missing = allIds.filter(id => knownOrder.indexOf(id) < 0);
          // v15.3: neu hinzugekommene Spalten (z.B. nach Custom-Field-Anlage
          // an einem bestehenden Event) VOR der „Aktion"-Spalte einreihen,
          // nicht hinten dran — sonst landen sie rechts neben den Buttons.
          const actionPos = knownOrder.indexOf('action');
          let mergedOrder = actionPos >= 0
            ? [...knownOrder.slice(0, actionPos), ...missing, ...knownOrder.slice(actionPos)]
            : [...knownOrder, ...missing];
          // v24.37: Die neu hinzugekommene 'company'-Spalte direkt HINTER
          // 'Standort' (location) einreihen statt ganz rechts vor den Aktionen
          // — gilt nur, solange der User sie noch nicht selbst positioniert hat
          // (also wenn sie frisch in `missing` steckt).
          if (missing.indexOf('company') >= 0) {
            mergedOrder = mergedOrder.filter(id => id !== 'company');
            const locIdx = mergedOrder.indexOf('location');
            if (locIdx >= 0) {
              mergedOrder.splice(locIdx + 1, 0, 'company');
            } else {
              const aPos = mergedOrder.indexOf('action');
              if (aPos >= 0) mergedOrder.splice(aPos, 0, 'company'); else mergedOrder.push('company');
            }
          }
          setColumnOrder(mergedOrder);
          // 'name' aus hidden auch herausfiltern (wenn jemals manuell hidden gesetzt wurde,
          // unwahrscheinlich da alwaysVisible — aber defensiv).
          setHiddenColumns(parsed.hidden.filter((id: string) => id !== 'name' && allIds.indexOf(id) >= 0));
          return;
        }
      }
    } catch { /* kaputte Config ignorieren */ }
    setColumnOrder(allIds);
    setHiddenColumns([]);
  }, [columnStorageKey, availableColumns.map(c => c.id).join(',')]);

  // Persistieren bei Änderungen.
  // v30.67: NUR nach einer echten Benutzeraktion (hide/show/move). Der
  // Lade-Effekt oben hängt an `availableColumns`, und das hängt an den noch
  // nicht geladenen `registrations`: Beim Öffnen fehlten die bedingten Spalten
  // („Nachgerückt am", „Startnummer"), der Lader schnitt sie aus order UND
  // hidden, und dieser Effekt schrieb den Zwischenstand sofort als Wahrheit
  // fest. Kamen die Anmeldungen Sekunden später, war die Ausblendung weg —
  // bei jedem Öffnen aufs Neue. Ein Dirty-Flag trennt „User hat umgestellt"
  // von „Lader hat auf einen unvollständigen Spaltensatz reduziert".
  const columnDirtyRef = React.useRef(false);
  React.useEffect(() => {
    if (!columnStorageKey || columnOrder.length === 0) return;
    if (!columnDirtyRef.current) return;
    columnDirtyRef.current = false;
    try {
      localStorage.setItem(columnStorageKey, JSON.stringify({ order: columnOrder, hidden: hiddenColumns }));
    } catch { /* quota exceeded oder private mode → ignorieren */ }
  }, [columnStorageKey, columnOrder.join(','), hiddenColumns.join(',')]);

  // Helper: Spalte ausblenden / wieder einblenden / verschieben.
  const hideColumn = (id: string): void => {
    const col = availableColumns.find(c => c.id === id);
    if (!col || col.alwaysVisible) return;
    if (hiddenColumns.indexOf(id) < 0) { columnDirtyRef.current = true; setHiddenColumns([...hiddenColumns, id]); }
  };
  const showColumn = (id: string): void => {
    columnDirtyRef.current = true;
    setHiddenColumns(hiddenColumns.filter(h => h !== id));
    if (columnOrder.indexOf(id) < 0) {
      const actionIdx = columnOrder.indexOf('action');
      const next = [...columnOrder];
      if (actionIdx >= 0) next.splice(actionIdx, 0, id); else next.push(id);
      setColumnOrder(next);
    }
  };
  const moveColumn = (id: string, direction: -1 | 1): void => {
    const idx = columnOrder.indexOf(id);
    if (idx < 0) return;
    const target = idx + direction;
    if (target < 0 || target >= columnOrder.length) return;
    if (columnOrder[target] === 'action') return;
    const next = [...columnOrder];
    [next[idx], next[target]] = [next[target], next[idx]];
    columnDirtyRef.current = true;
    setColumnOrder(next);
  };
  return {
    availableColumns, hasRoommateColumn, hasWaitlistActivity, hideColumn, moveColumn, showColumn,
  };
}

