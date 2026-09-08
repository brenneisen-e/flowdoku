/* ConsolidatedView — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 5377-6469 des Stands
 * vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nach `docs/ui-leitfaden.md` (Abschnitte 4, 5a/5b) modernisiert —
 * Hinweiskästen als `dex-ui-callout` ausserhalb des Scrollbereichs,
 * Werkzeugleiste über der Tabelle, `dex-ui-table--compact` mit
 * Personen-Zelle und Status-Pillen. Die Auflösung der Werte (Parent-Zeile
 * zuerst, dann Sub-Event-CustomData) ist unverändert.
 */
import * as React from 'react';
import { DeloitteEvent } from '../../../types';
import { isEventOver } from '../../../utils/eventFormat';
import { ConsolidatedRow, DeniedSubEventList } from '../../admin/adminTypes';
import { PersonContactHover } from '../../PersonContactHover';
import { formatDate, translateStatus } from '../../../utils/eventStatus';
import { SPRegistration } from '../../../services/EventService';
import { AlertCircle, Check, ChevronDown, Columns, ExternalLink, Plus, Trash2, Users } from '../../Icons';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { activeParentRegOf } from '../logic/parentRegs';
import { cx, ensureDexUiStyles } from '../../dexUi';

export interface ConsolidatedViewProps {
  addAllToKlammer: (rows: ConsolidatedRow[]) => Promise<void>;
  addingToKlammer: string;
  addToKlammer: (row: ConsolidatedRow) => Promise<void>;
  bulkKlammerProgress: string;
  colToggleHover: boolean;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  consolidatedChildren: DeloitteEvent[];
  consolidatedFiltered: ConsolidatedRow[];
  consolidatedRows: ConsolidatedRow[];
  consolidatedSort: string;
  consolidatedSortAsc: boolean;
  /** v30.67: Termine, deren Liste NICHT gelesen wurde (v30.37-Zustand aus AdminPage). */
  deniedSubEventLists: DeniedSubEventList[];
  expandedConsolidatedEmail: string;
  highlightMatch: (text: unknown) => React.ReactNode;
  inactiveAccounts: string[];
  isAdmin: boolean;
  isConsolidatedMode: boolean;
  isDe: boolean;
  isLoadingSubEventRegs: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  missingReminderKey: string;
  openDeregModal: (row: ConsolidatedRow) => void;
  openMainFieldsEdit: (emailKey: string, displayName: string) => void;
  orgPastLock: boolean;
  performSilentDuplicateDelete: (reg: SPRegistration) => Promise<boolean>;
  personalColsCollapsed: boolean;
  registrations: SPRegistration[];
  reminderBusyId: number;
  searchQuery: string;
  selectedEvent: DeloitteEvent;
  sendCompleteRegistrationReminder: (args: { eventId: string; eventTitle: string; participantEmail: string; participantName: string; registeredByEmail?: string; registeredByName?: string; }) => Promise<boolean>;
  setAssignAssistRow: React.Dispatch<React.SetStateAction<ConsolidatedRow>>;
  setAssignAssistValue: React.Dispatch<React.SetStateAction<string>>;
  setColToggleHover: React.Dispatch<React.SetStateAction<boolean>>;
  setConsolidatedSort: React.Dispatch<React.SetStateAction<string>>;
  setConsolidatedSortAsc: React.Dispatch<React.SetStateAction<boolean>>;
  setExpandedConsolidatedEmail: React.Dispatch<React.SetStateAction<string>>;
  setMissingReminderKey: React.Dispatch<React.SetStateAction<string>>;
  setParticipantDetail: React.Dispatch<React.SetStateAction<{ name: string; email: string; jobTitle: string; location: string; company: string; department: string; phone: string; status: string; tid: number; }>>;
  setPersonalColsCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
  setReminderBusyId: React.Dispatch<React.SetStateAction<number>>;
  setSelectedEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  stripLocPrefix: (loc: string) => string;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

export const ConsolidatedView: React.FC<ConsolidatedViewProps> = (p) => {
  // v31.3: `colToggleHover`/`setColToggleHover` bleiben in der Schnittstelle
  // (AdminPage übergibt sie unverändert weiter), werden hier aber nicht mehr
  // gelesen — der Aufklapp-Knopf holt seinen Hover jetzt aus `dex-ui-chip`
  // (Leitfaden 1.3: kein onMouseEnter-State für reine Optik).
  const { addAllToKlammer, addingToKlammer, addToKlammer, bulkKlammerProgress, confirmDialog, consolidatedChildren, consolidatedFiltered, consolidatedRows, consolidatedSort, consolidatedSortAsc, deniedSubEventLists, expandedConsolidatedEmail, highlightMatch, inactiveAccounts, isAdmin, isConsolidatedMode, isDe, isLoadingSubEventRegs, isOrganizerFor, missingReminderKey, openDeregModal, openMainFieldsEdit, orgPastLock, performSilentDuplicateDelete, personalColsCollapsed, registrations, reminderBusyId, searchQuery, selectedEvent, sendCompleteRegistrationReminder, setAssignAssistRow, setAssignAssistValue, setConsolidatedSort, setConsolidatedSortAsc, setExpandedConsolidatedEmail, setMissingReminderKey, setParticipantDetail, setPersonalColsCollapsed, setReminderBusyId, setSelectedEvent, showAlert, stripLocPrefix, subEventRegsByEventId } = p;
  // Idempotent — Modal und WizardFormShell rufen es ebenfalls; hier nötig, weil
  // die Matrix auch ohne offenes Modal gerendert wird.
  ensureDexUiStyles();
  // v31.3: Ein Aufklapper je Hinweiskasten. Sichtbar bleiben Befund und Knöpfe,
  // die ausführliche Erklärung steckt darin (Leitfaden 2c: höchstens zwei
  // Zeilen Erklärtext). Der Hook steht VOR den frühen Returns — sonst kippt die
  // Hook-Reihenfolge (react-hooks/rules-of-hooks ist `error`).
  const [openHelp, setOpenHelp] = React.useState<string>('');
  const toggleHelp = (k: string): void => setOpenHelp(openHelp === k ? '' : k);
    if (!selectedEvent) return null;
    if (isLoadingSubEventRegs) {
      return <p className="dex-ui-muted" style={{ fontStyle: 'italic' }}>{isDe ? 'Lade Sub-Event-Teilnehmer...' : 'Loading sub-event participants...'}</p>;
    }
    if (consolidatedRows.length === 0) {
      // v31.3: „Leer" heißt nur dann „niemand angemeldet", wenn ALLE
      // Termin-Listen gelesen wurden. War eine gesperrt, ist der Stand
      // unbekannt — und unbekannt wird benannt, nicht als leere Liste
      // gerendert (CLAUDE.md: ein Lesefehler ist keine Null).
      if (deniedSubEventLists.length > 0) {
        return (
          <div className="dex-ui-callout dex-ui-callout--warn">
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <div style={{ flex: 1, minWidth: 0 }}>
              <strong>{isDe ? 'Teilnehmer unbekannt' : 'Participants unknown'}</strong>
              {' — '}
              {isDe
                ? `${deniedSubEventLists.length} Termin-Liste(n) konnten nicht gelesen werden, in den übrigen steht keine Anmeldung. Ob jemand angemeldet ist, lässt sich hier nicht sagen. Betroffen: `
                : `${deniedSubEventLists.length} date list(s) could not be read, the others contain no registration. Whether anyone is registered cannot be said here. Affected: `}
              {deniedSubEventLists.map(d => d.status > 0 ? `${d.title} (HTTP ${d.status})` : d.title).join(' · ')}
            </div>
          </div>
        );
      }
      return (
        <div className="dex-ui-empty">
          <span className="dex-ui-empty-icon"><Users size={20} /></span>
          <div className="dex-ui-empty-title">{isDe ? 'Noch keine Anmeldungen in den Sub-Events.' : 'No registrations in the sub-events yet.'}</div>
          <div>{isDe ? 'Sobald sich jemand für einen Termin anmeldet, steht hier eine Zeile je Person.' : 'As soon as someone registers for a date, a row per person appears here.'}</div>
        </div>
      );
    }
    // v14.11: pastel A = event-level (parent) fields, pastel B = sub-event-specific fields
    const PASTEL_A_HEADER: React.CSSProperties = { background: 'rgba(0, 118, 168, 0.15)' };
    const PASTEL_A_CELL: React.CSSProperties = { background: 'rgba(0, 118, 168, 0.08)' };
    const PASTEL_B_HEADER: React.CSSProperties = { background: 'rgba(255, 191, 0, 0.18)' };
    const PASTEL_B_CELL: React.CSSProperties = { background: 'rgba(255, 191, 0, 0.10)' };
    const parentCustomFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim());
    // v23.32: People-Picker-Felder des Hauptevents (z.B. „Assistenz") werden
    // jetzt als eigene Spalten mit Foto + Name gezeigt (vorher ganz ausgeblendet).
    const parentUserFields = (selectedEvent.eventSpecificFields || []).filter(f => (f.type === 'user' || f.type === 'roommate') && f.label && f.label.trim());
    const parentIds = new Set(parentCustomFields.map(f => f.id));
    const childCustomFieldsByChild: Array<{ child: DeloitteEvent; fields: typeof parentCustomFields }> = consolidatedChildren.map(c => {
      const own = (c.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim() && !parentIds.has(f.id));
      return { child: c, fields: own };
    });
    // v30.17: Spalten-Zustand je Termin — vergangene Tage und (bei aktiver
    // „Anmeldung ab"-Regel) noch nicht anmeldbare Tage werden in Kopf,
    // Summenzeile und Zellen gedimmt. Dieselbe openFrom-Rechnung wie auf der
    // Anmeldeseite und in „Meine Events" (fixed/day/week) — keine zweite
    // Logik aufmachen.
    const childOpenFrom = (c: DeloitteEvent): Date | null => {
      const rule = selectedEvent.subEventOpenRule;
      if (!rule) return null;
      if (rule.mode === 'fixed') {
        const d = new Date(rule.date || '');
        return isFinite(d.getTime()) ? d : null;
      }
      if (!((rule.days || 0) > 0)) return null;
      const base = new Date(c.startDate || '');
      if (!isFinite(base.getTime())) return null;
      const dd = new Date(base.getFullYear(), base.getMonth(), base.getDate());
      if (rule.mode === 'week') dd.setDate(dd.getDate() - ((dd.getDay() + 6) % 7));
      dd.setDate(dd.getDate() - (rule.days || 0));
      dd.setHours(0, 0, 0, 0);
      return dd;
    };
    const hasOpenRule = !!selectedEvent.subEventOpenRule;
    const childColState: Record<string, { past: boolean; notYetOpen: boolean; openFrom: Date | null }> = {};
    for (const { child } of childCustomFieldsByChild) {
      const opensAt = childOpenFrom(child);
      childColState[child.id] = { past: isEventOver(child), notYetOpen: !!opensAt && new Date() < opensAt, openFrom: opensAt };
    }
    const dimColStyle = (id: string): React.CSSProperties =>
      (childColState[id] && (childColState[id].past || childColState[id].notYetOpen)) ? { opacity: 0.45 } : {};
    // v31.3: Termine, deren Teilnehmerliste nicht gelesen wurde. `deniedSubEventLists`
    // trägt nur Titel und Status (adminTypes) — AdminPage legt `ch.title || ch.id`
    // ab, deshalb wird genau darüber verglichen. Wozu das gut ist: Ohne diesen
    // Abgleich zeigen Kopf, Summenzeile und Zellen einer gesperrten Spalte „0"
    // bzw. „—", also eine Aussage über Daten, die niemand gelesen hat.
    const deniedTitles = new Set(deniedSubEventLists.map(d => d.title));
    const isDeniedChild = (c: DeloitteEvent): boolean => deniedTitles.has(c.title || c.id);
    const handleSortConsolidated = (key: string): void => {
      if (consolidatedSort === key) setConsolidatedSortAsc(!consolidatedSortAsc);
      else { setConsolidatedSort(key); setConsolidatedSortAsc(true); }
    };
    // v31.3: Der Pfeil ist jetzt ein `dex-ui-table-sort`-Element, die Kopfzelle
    // trägt `is-sortable`/`is-sorted` (Leitfaden 5b) — vorher war „sortierbar"
    // nur am Mauszeiger erkennbar, „sortiert" nur am angehängten Zeichen.
    const sortArrow = (key: string): React.ReactNode =>
      key === consolidatedSort ? <span className="dex-ui-table-sort" aria-hidden="true">{consolidatedSortAsc ? '▲' : '▼'}</span> : null;
    const sortCls = (key: string): string => cx('is-sortable', consolidatedSort === key && 'is-sorted');
    // v31.3: Ein Aufklapper, viermal gebraucht — als Helfer, damit die
    // Hinweiskästen nicht viermal dieselben zehn Zeilen tragen.
    const helpBlock = (key: string, label: string, body: React.ReactNode): React.ReactNode => (
      <>
        <button type="button" className={cx('dex-ui-disclosure', openHelp === key && 'is-open')} onClick={() => toggleHelp(key)}>
          <span className="dex-ui-disclosure-chevron"><ChevronDown size={14} /></span>{label}
        </button>
        {openHelp === key && <div className="dex-ui-disclosure-body">{body}</div>}
      </>
    );
    // v31.3: Status-Pille nach Leitfaden 5b — grün angemeldet, blau eingecheckt,
    // orange Warteliste, grau abgemeldet, rot No-Show.
    const statusPillClass = (status: string | undefined): string => {
      switch (status) {
        case 'Eingecheckt': return 'dex-ui-pill dex-ui-pill--blue';
        case 'Angemeldet': case 'QR versendet': return 'dex-ui-pill dex-ui-pill--green';
        case 'Warteliste': return 'dex-ui-pill dex-ui-pill--orange';
        case 'No-Show': return 'dex-ui-pill dex-ui-pill--red';
        default: return 'dex-ui-pill dex-ui-pill--gray';
      }
    };
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const abbreviate = (s: string, max: number): string => s.length > max ? s.substring(0, max - 1) + '…' : s;
    // v23.5: 6 Personen-Spalten (#, Vorname, Nachname, Email, Job Title,
    // Standort) — eingeklappt nur 2 (#, „Teilnehmer").
    const personalColCount = personalColsCollapsed ? 2 : 6;
    // v26.84: +1 zusätzliche Spalte „Registriert von" (Akteur) neben „Details".
    const totalColSpan = personalColCount + parentCustomFields.length + parentUserFields.length + childCustomFieldsByChild.reduce((sum, x) => sum + 1 + x.fields.length, 0) + 2;
    // v19.30: Aktionen (Hauptevent-Felder bearbeiten / abmelden) nur für
    // berechtigte Rollen (Admin oder Organizer dieses Events).
    const canManage = isAdmin || isOrganizerFor(selectedEvent);
    // v19.30 (Feature A): Anzahl der bearbeitbaren Hauptevent-Felder (ohne
    // People-Picker und Dokument-Uploads, die keinen editierbaren Textwert
    // haben). Nur wenn > 0 erscheint der „Felder"-Button.
    const editableParentFieldCount = parentCustomFields.filter(f => f.type !== 'document').length;
    // v19.30 (Feature A): Hat die Person eine Registrierung auf der
    // Hauptevent-Teilnehmerliste? Nur dann gibt es Hauptevent-Antworten zum
    // Bearbeiten. (Im subEventsOnlyMode kann jemand nur in Sub-Events
    // angemeldet sein.)
    // v30.67: nur eine AKTIVE Klammer-Zeile zählt (s. parentRegs.ts). `some`
    // ohne Status ließ eine abgemeldete Klammer-Zeile den roten Kasten
    // „Fehlende Klammer-Anmeldung" schlucken und blendete „Zur Klammer
    // hinzufügen" aus — der Organizer hatte keinen Weg, den Zustand zu
    // sehen oder zu reparieren. `registerForEvent` reaktiviert die
    // abgemeldete Zeile (EventContext), es entsteht keine Dublette.
    const hasParentReg = (emailKey: string): boolean =>
      !!activeParentRegOf(registrations, emailKey);
    // v26.85: Ist ein bestimmtes Hauptevent-Feld für diese Person befüllt?
    // Gleiche Auflösung wie die Tabelle (Parent-Reg-Spalte → Parent-CustomData
    // → Sub-Event-CustomData-Fallback), damit die „Infos fehlen"-Erkennung nicht
    // fälschlich anschlägt.
    const parentFieldFilled = (row: ConsolidatedRow, f: { id: string; spInternalName?: string }): boolean => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const parentReg = activeParentRegOf(registrations, row.emailKey) as any; // v30.67: aktive, neueste Zeile
      const spName = (f as { spInternalName?: string }).spInternalName || '';
      if (parentReg) {
        let v: unknown = spName ? parentReg[spName] : undefined;
        if ((v === undefined || v === null || v === '') && parentReg.CustomData) { try { v = JSON.parse(parentReg.CustomData)[f.id]; } catch { /* */ } }
        if (v !== undefined && v !== null && v !== '') return true;
      }
      for (const ch of consolidatedChildren) {
        const r = row.perChild[ch.id];
        if (!r) continue;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let v: any = spName ? (r as any)[spName] : undefined;
        if ((v === undefined || v === null || v === '') && r.CustomData) { try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ } }
        if (v !== undefined && v !== null && v !== '') return true;
      }
      return false;
    };
    // v26.85: Personen mit Klammer-Anmeldung (Hauptevent) + Sub-Event, bei denen
    // aber PFLICHT-Hauptevent-Felder leer sind (typisch: Anmeldung im
    // Hauptevent-Schritt abgebrochen). requiredMainFields leer → nie anschlagen.
    const requiredMainFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.required && f.label && f.label.trim());
    // v23.7: „Verwaiste" Klammer-/Schatten-Anmeldungen erkennen — eine aktive
    // Zeile auf der KLAMMER-Subsite (registrations), deren Person in KEINEM
    // Sub-Event aktiv angemeldet ist. Das ist typischerweise ein Geist aus einer
    // abgebrochenen Anmeldung: unsichtbar in der Matrix (die nur Sub-Event-
    // Anmeldungen zeigt), blockiert aber jede neue (auch stellvertretende)
    // Anmeldung, weil die Doppel-Anmelde-Prüfung die Klammer-Subsite liest.
    // v23.7: Geist = aktive Klammer-Zeile, deren Person in KEINEM Sub-Event
    // IRGENDEINE Zeile hat (auch keine abgemeldete). Wer sich aus allen
    // Sub-Events ABGEMELDET hat, hat dort abgemeldete Zeilen → wird bewusst
    // NICHT als Geist erkannt (das ist eine normale Voll-Abmeldung, kein Rest
    // aus einer abgebrochenen Anmeldung). So vermeiden wir Fehlalarme.
    const anySubEmails = new Set<string>();
    for (const ch of consolidatedChildren) {
      for (const r of (subEventRegsByEventId[ch.id] || [])) {
        const em = (r.ParticipantEmail || '').toLowerCase().trim();
        if (em) anySubEmails.add(em);
      }
    }
    // v30.67: Diese Erkennung ist nur belastbar, wenn ALLE Termin-Listen
    // gelesen wurden. War auch nur eine nicht lesbar (403 für eine nachträglich
    // benannte Co-Organizerin, 429 beim 12. Termin), fehlen deren Personen in
    // `anySubEmails` — jede davon hat aber (seit v30.42 garantiert) eine aktive
    // Klammer-Zeile und landete hier als „Rest". Der Knopf darunter löscht die
    // Klammer-Zeile HART, samt aller Hauptevent-Antworten. Ein leeres Ergebnis
    // ohne geprüften Status ist keine Aussage über die Daten: Bei gesperrten
    // Listen wird die Box ausgesetzt und das gesagt.
    const orphanCheckBlocked = deniedSubEventLists.length > 0;
    const orphanShadowRegs = orphanCheckBlocked ? [] : registrations.filter(r => {
      if ((r.Status || '') === 'Abgemeldet') return false;
      const em = (r.ParticipantEmail || '').toLowerCase().trim();
      return !!em && !anySubEmails.has(em);
    });
    // v26.68: Bei aktiver Suche eine zusätzliche „ID"-Spalte mit der echten
    // TeilnehmerID aus der Klammer-/Hauptevent-Liste einblenden. Die „#"-Spalte
    // bleibt die laufende Durchzählung der aktuellen Ansicht — die ID hilft, die
    // gefilterte Person eindeutig in der SharePoint-Liste wiederzufinden.
    const searchActive = !!(searchQuery || '').trim();
    const bracketTidByEmail: Record<string, number> = {};
    for (const rr of registrations) {
      const em = (rr.ParticipantEmail || '').toLowerCase().trim();
      if (em && typeof rr.TeilnehmerID === 'number' && !(em in bracketTidByEmail)) bracketTidByEmail[em] = rr.TeilnehmerID;
    }
    return (
      // v31.3: Die Höhenbegrenzung (v28.53) sitzt jetzt am Tabellen-Container
      // weiter unten, nicht mehr an dieser Wurzel. Grund: Hier lagen auch die
      // Hinweiskästen im Scrollbereich — wer zur Tabelle scrollte, scrollte die
      // Warnung weg. Der sticky-thead bekommt seinen Bezugsrahmen unverändert
      // aus dem eigenen Scroll-Container.
      <div className="dex-ui-fade-in">
        {/* v30.67: Prüfung ausgesetzt, solange eine Termin-Liste nicht lesbar war —
            sonst löscht „Rest-Anmeldung entfernen" echte Anmeldungen. */}
        {orphanCheckBlocked && (
          <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 12 }}>
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <div style={{ flex: 1, minWidth: 0 }}>
              <strong>
                {isDe ? 'Prüfung auf unvollständige Anmeldungen ausgesetzt' : 'Check for incomplete registrations suspended'}
              </strong>
              {' — '}
              {isDe
                ? `${deniedSubEventLists.length} Termin-Liste(n) konnten nicht gelesen werden. Wer nur dort angemeldet ist, sähe hier wie ein Rest aus, und „Rest-Anmeldung entfernen“ würde eine echte Anmeldung löschen. Sobald alle Listen lesbar sind, erscheint die Prüfung wieder.`
                : `${deniedSubEventLists.length} date list(s) could not be read. Anyone registered only there would look like a leftover here, and „Remove leftover“ would delete a real registration. The check returns as soon as all lists are readable.`}
            </div>
          </div>
        )}
        {/* v23.7: Unvollständige Klammer-Anmeldungen (nur Klammer, kein Sub-Event)
            sichtbar machen — mit Erinnerungs- oder Entfernen-Option, damit eine
            blockierte (Neu-)Anmeldung wieder möglich wird. */}
        {orphanShadowRegs.length > 0 && (
          <div className="dex-ui-callout dex-ui-callout--danger" style={{ marginBottom: 12 }}>
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <div style={{ flex: 1, minWidth: 0 }}>
              <strong style={{ fontSize: '0.9rem' }}>
                {isDe ? `Unvollständige Anmeldungen (${orphanShadowRegs.length})` : `Incomplete registrations (${orphanShadowRegs.length})`}
              </strong>
              {/* v31.3: Sichtbar bleibt der Befund; was zu tun ist und was nach dem
                  Entfernen passiert, steht im Aufklapper (Leitfaden 2c: höchstens
                  zwei Zeilen Erklärtext sichtbar). Kein Satz entfällt. */}
              <p style={{ margin: '4px 0 6px' }}>
                {isDe
                  ? 'Diese Personen haben eine Klammer-Anmeldung, sind aber in keinem Sub-Event aktiv angemeldet — in der Regel ein unvollständiger Rest aus einer abgebrochenen Anmeldung. Solche Rest-Anmeldungen erscheinen in der Teilnehmerliste unten nicht, blockieren aber eine erneute (auch stellvertretende) Anmeldung.'
                  : 'These people have an umbrella registration but are not actively registered for any sub-event — usually an incomplete leftover from an interrupted registration. Such leftover registrations don’t appear in the participant list below, but they block a new (or on-behalf) registration.'}
              </p>
              {helpBlock('orphan', isDe ? 'Was du tun kannst' : 'What you can do', <>
                <p style={{ margin: '0 0 8px' }}>
                  {isDe
                    ? 'Du hast zwei Möglichkeiten: über „Erinnerung senden“ der Person – bzw. der Person, die sie angemeldet hat – einen Link zum Abschließen der Anmeldung schicken, oder die Rest-Anmeldung entfernen, sodass eine Neuanmeldung wieder möglich ist.'
                    : 'You have two options: use „Send reminder“ to send the person – or whoever registered them – a link to complete the registration, or remove the leftover registration so a new one becomes possible.'}
                </p>
                <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm">
                  {isDe
                    ? <><strong>Hinweis:</strong> Nach dem Entfernen ist die Person <strong>nicht</strong> angemeldet. Informiere die Person – bzw. die Person, die sie angemeldet hat –, dass eine <strong>erneute Anmeldung</strong> nötig ist. Alternativ kannst du über <strong>„Erinnerung senden“</strong> direkt einen Link zum Abschließen verschicken.</>
                    : <><strong>Note:</strong> After removal the person is <strong>not</strong> registered. Let the person – or whoever registered them – know that a <strong>new registration</strong> is required. Alternatively, use <strong>„Send reminder“</strong> to send a completion link directly.</>}
                </div>
              </>)}
              <div style={{ display: 'flex', flexDirection: 'column', gap: 0, marginTop: 4 }}>
              {orphanShadowRegs.map((r, oi) => {
                const nm = (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || r.ParticipantEmail);
                // v23.8: Bei stellvertretender Anmeldung (RegisteredBy ≠ Teilnehmer)
                // zeigen, WER die Anmeldung durchgeführt hat — hilft, die Assistenz
                // bzw. den Organizer zu identifizieren.
                const actorEmail = (r.RegisteredByEmail || '').trim();
                const isProxy = !!actorEmail && actorEmail.toLowerCase() !== (r.ParticipantEmail || '').trim().toLowerCase();
                const actorLabel = (r.RegisteredByName || '').trim() || actorEmail;
                // v26.68: Profil-Zusatzinfos (Position, Standort, Unternehmen) —
                // liegen als Property auf der Registrierung, hier für Foto-Subline
                // und die Meta-Zeile genutzt.
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const anyR = r as any;
                const jobTitle = String(anyR.JobTitle || '').trim();
                const loc = stripLocPrefix(String(r.Location || ''));
                const subline = [jobTitle, loc, String(r.Company || '').trim()].filter(Boolean).join(' • ');
                return (
                  <div key={r.Id} style={{ display: 'flex', alignItems: 'flex-start', gap: 12, padding: '10px 0', borderTop: oi > 0 ? '1px solid rgba(200,0,0,0.14)' : 'none' }}>
                    {/* v26.68: Foto mit Hover-Kontaktkarte, wie in der Teilnehmerliste. */}
                    {(r.ParticipantEmail || '') ? (
                      <PersonContactHover email={r.ParticipantEmail || ''} name={nm} size={40} subline={subline} isDe={isDe} />
                    ) : null}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 3, minWidth: 0, flex: 1 }}>
                      <div style={{ fontSize: '0.86rem' }}>
                        <strong>{nm}</strong>
                        <span style={{ color: 'var(--dex-gray-500)', marginLeft: 8 }}>{r.ParticipantEmail}</span>
                      </div>
                      {(jobTitle || loc) && (
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>{[jobTitle, loc].filter(Boolean).join(' • ')}</div>
                      )}
                      <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)' }}>
                        {translateStatus(r.Status, isDe)}
                        {r.RegistrationDate ? <> · {isDe ? 'angemeldet am' : 'registered on'} {formatDate(r.RegistrationDate)}</> : null}
                        {' · '}
                        {isProxy
                          ? <>{isDe ? 'angemeldet durch' : 'registered by'} <strong>{actorLabel}</strong></>
                          : (isDe ? 'selbst angemeldet' : 'self-registered')}
                      </div>
                      {canManage && (
                        <div style={{ display: 'inline-flex', gap: 8, flexWrap: 'wrap', marginTop: 6 }}>
                          {/* v26.67 (A): Erinnerung senden — die Person (bzw. die
                              anmeldende Person) bitten, die Anmeldung abzuschließen,
                              statt sie nur zu entfernen. */}
                          <button
                            type="button"
                            className="btn btn-outline dex-ui-btn-sm"
                            disabled={reminderBusyId === r.Id}
                            onClick={async () => {
                              if (!selectedEvent) return;
                              setReminderBusyId(r.Id);
                              const ok = await sendCompleteRegistrationReminder({
                                eventId: selectedEvent.id,
                                eventTitle: selectedEvent.title,
                                participantEmail: r.ParticipantEmail || '',
                                participantName: nm,
                                registeredByEmail: r.RegisteredByEmail || '',
                                registeredByName: r.RegisteredByName || '',
                              }).catch(() => false);
                              setReminderBusyId(null);
                              showAlert(
                                ok
                                  ? (isProxy
                                      ? (isDe ? `Erinnerung an ${actorLabel} gesendet (${nm} auf Kopie) — mit Link zum Abschließen der Anmeldung.` : `Reminder sent to ${actorLabel} (${nm} on copy) — with a link to complete the registration.`)
                                      : (isDe ? `Erinnerung an ${nm} gesendet — mit Link zum Abschließen der Anmeldung.` : `Reminder sent to ${nm} — with a link to complete the registration.`))
                                  : (isDe ? 'Erinnerung konnte nicht gesendet werden.' : 'The reminder could not be sent.'),
                                { variant: ok ? 'success' : 'error' });
                            }}
                          >
                            {reminderBusyId === r.Id ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? 'Erinnerung senden' : 'Send reminder')}
                          </button>
                          {/* v31.3: `btn-danger` statt rot eingefärbtem Outline —
                              Unwiderrufliches sieht im ganzen Center gleich aus
                              (Leitfaden 1.5). Die Rückfrage bleibt wortgleich. */}
                          <button
                            type="button"
                            className="btn btn-danger dex-ui-btn-sm"
                            onClick={async () => {
                              if (!(await confirmDialog(isDe ? `Unvollständige Anmeldung von ${nm} entfernen? Die Person kann sich danach neu anmelden.` : `Remove the incomplete registration of ${nm}? The person can register again afterwards.`, { danger: true, confirmLabel: isDe ? 'Entfernen' : 'Remove' }))) return;
                              // v30.67: `deleteRegistration` wirft nicht, es liefert false —
                              // vorher hieß es auch dann „entfernt".
                              const okDel = await performSilentDuplicateDelete(r);
                              showAlert(okDel
                                ? (isDe ? 'Rest-Anmeldung entfernt — die Person kann jetzt wieder angemeldet werden.' : 'Leftover registration removed — the person can be registered again now.')
                                : (isDe ? 'Rest-Anmeldung konnte NICHT entfernt werden (fehlende Rechte oder Drosselung) — die Zeile ist noch da.' : 'The leftover registration could NOT be removed (missing permissions or throttling) — the row is still there.'),
                                { variant: okDel ? 'success' : 'error' });
                            }}
                          >
                            {isDe ? 'Rest-Anmeldung entfernen' : 'Remove leftover'}
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
              </div>
            </div>
          </div>
        )}
        {/* v24.38: Fehlende Klammer-Anmeldung — Personen mit Sub-Event-Zeilen,
            aber OHNE Hauptevent-/Klammer-Anmeldung (Daten-Anomalie). Als Fehler
            ausweisen + pro Person „Zur Klammer hinzufügen". */}
        {canManage && (() => {
          /**
           * v30.56: ERST prüfen, ob die Klammer-Zeile nur unter einer ANDEREN
           * Schreibweise der Adresse steht — dann fehlt sie nämlich gar nicht.
           *
           * Der Kasten hat bis hier ausschließlich `ParticipantEmail` exakt
           * verglichen. Dieselbe Person kann aber unter zwei Adressen in den
           * Listen stehen (SMTP-Adresse gegen UPN/Alias) — die
           * Doppel-Adressen-Falle, die in CLAUDE.md als erster Verdacht bei
           * widersprüchlichen Ansichten steht. Trifft sie zu, meldete der
           * Kasten einen Fehler, den es nicht gibt.
           *
           * **Und der Reparatur-Knopf hätte ihn zu einem echten gemacht:**
           * „Zur Klammer hinzufügen" legt eine Zeile unter der Sub-Event-
           * Adresse an — die Person stünde danach ZWEIMAL am Hauptevent, mit
           * zwei Teilnehmer-IDs. Ein Knopf, der Daten repariert, darf keine
           * Dubletten erzeugen.
           *
           * Erkannt wird die zweite Schreibweise über den lokalen Teil der
           * Adresse (vor dem @) und über Vor-/Nachname. Beides ist eine
           * Heuristik — deshalb wird der Fall NICHT still geschluckt, sondern
           * getrennt ausgewiesen: Der Organizer sieht, unter welcher Adresse
           * die Zeile steht, und entscheidet selbst.
           */
          const normName = (v: string): string => (v || '').toLowerCase().replace(/[^a-zäöüß]/g, '');
          const localPart = (v: string): string => (v || '').toLowerCase().split('@')[0].trim();
          const altParentRegOf = (row: typeof consolidatedRows[number]): SPRegistration | undefined => {
            const lp = localPart(row.emailKey);
            const nm = normName(row.vorname) + normName(row.nachname);
            return registrations.find(r => {
              const rEmail = (r.ParticipantEmail || '').toLowerCase().trim();
              if (rEmail === row.emailKey) return false; // exakt wurde oben schon geprüft
              if (lp && localPart(rEmail) === lp) return true;
              const rNm = normName(r.Vorname || '') + normName(r.Nachname || '');
              return !!nm && rNm === nm;
            });
          };
          const flagged = consolidatedRows.filter(r => r.activeCount > 0 && !hasParentReg(r.emailKey));
          const aliasCases = flagged
            .map(r => ({ row: r, alt: altParentRegOf(r) }))
            .filter(x => !!x.alt) as Array<{ row: typeof consolidatedRows[number]; alt: SPRegistration }>;
          const aliasKeys = new Set(aliasCases.map(x => x.row.emailKey));
          const missing = flagged.filter(r => !aliasKeys.has(r.emailKey));
          if (missing.length === 0 && aliasCases.length === 0) return null;
          if (missing.length === 0) {
            // Nur Adress-Dubletten — kein Fehler, aber erklärungsbedürftig.
            return (
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 12 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <strong style={{ fontSize: '0.9rem' }}>
                    {isDe ? `Klammer-Zeile unter anderer Adresse (${aliasCases.length})` : `Umbrella row under a different address (${aliasCases.length})`}
                  </strong>
                  <p style={{ margin: '4px 0 6px' }}>
                    {isDe
                      ? 'Diese Personen haben eine Klammer-Anmeldung — sie steht nur unter einer anderen Schreibweise ihrer E-Mail-Adresse (SMTP-Adresse gegen UPN/Alias). Es fehlt also nichts.'
                      : 'These people do have an umbrella registration — it is just stored under a different spelling of their email address. Nothing is missing.'}
                  </p>
                  {helpBlock('alias', isDe ? 'Was du tun kannst' : 'What you can do',
                    isDe
                      ? 'Trag sie NICHT über „Zur Klammer hinzufügen" nach, das würde eine zweite Zeile mit einer zweiten Teilnehmer-ID erzeugen. Wenn die beiden Schreibweisen stören, korrigiere die Adresse in der Teilnehmerzeile.'
                      : 'Do NOT use „Add to umbrella", it would create a duplicate row. If the two spellings bother you, correct the address in the participant row.')}
                  <div className="dex-ui-stack" style={{ gap: 2, marginTop: 4 }}>
                    {aliasCases.map(x => (
                      <div key={x.row.emailKey}>
                        <strong>{x.row.vorname} {x.row.nachname}</strong>{' '}
                        <span className="dex-ui-muted">Sub-Events: {x.row.email} · Klammer: {x.alt.ParticipantEmail}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            );
          }
          return (
            <div className="dex-ui-callout dex-ui-callout--danger" style={{ marginBottom: 12 }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <strong style={{ fontSize: '0.9rem' }}>
                  {isDe ? `Fehlende Klammer-Anmeldung (${missing.length})` : `Missing umbrella registration (${missing.length})`}
                </strong>
                <p style={{ margin: '4px 0 6px' }}>
                  {isDe
                    ? 'Diese Personen sind in einem oder mehreren Sub-Events angemeldet, fehlen aber am Klammer-/Hauptevent selbst (z.B. durch eine abgebrochene Anmeldung). Dadurch fehlen u.a. die übergreifenden Hauptevent-Angaben.'
                    : 'These people are registered for one or more sub-events but are missing on the umbrella/main event itself (e.g. due to an interrupted registration), so the cross-cutting main-event details are missing.'}
                </p>
                {helpBlock('missing', isDe ? 'Was die beiden Knöpfe tun' : 'What the two buttons do',
                  isDe
                    ? 'Du hast zwei Möglichkeiten: über „Erinnerung senden“ bittest du die Person (bzw. die anmeldende Person) per Mail mit Direkt-Link, die fehlenden Hauptevent-Angaben in der App nachzutragen — oder du trägst die fehlende Klammer-Anmeldung mit „Zur Klammer hinzufügen“ selbst nach (versendet KEINE Mail und KEINEN Outlook-Termin, reine Datenkorrektur).'
                    : 'You have two options: use „Send reminder“ to ask the person (or whoever registered them) via email with a direct link to add the missing main-event details in the app — or add the missing umbrella registration yourself with „Add to umbrella“ (sends NO email and NO Outlook invite, data correction only).')}
                {/* v30.56: Adress-Dubletten auch hier benennen, wenn es
                    DANEBEN echte Lücken gibt — sonst verschwinden sie
                    kommentarlos aus dem Kasten und der Organizer fragt sich,
                    wo die dritte Person geblieben ist. */}
                {aliasCases.length > 0 && (
                  <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ margin: '6px 0' }}>
                    {isDe
                      ? <>Nicht aufgeführt, weil dort nichts fehlt: {aliasCases.map(x => `${x.row.vorname} ${x.row.nachname}`).join(', ')} — die Klammer-Zeile steht unter einer anderen Schreibweise der Adresse ({aliasCases.map(x => x.alt.ParticipantEmail).join(', ')}).</>
                      : <>Not listed because nothing is missing there: {aliasCases.map(x => `${x.row.vorname} ${x.row.nachname}`).join(', ')} — the umbrella row exists under a different spelling of the address.</>}
                  </div>
                )}
                {/* v30.14: Sammel-Fix — alle auf einmal, still, sequentiell. */}
                <div style={{ margin: '6px 0 8px' }}>
                  <button
                    type="button"
                    className="btn btn-outline dex-ui-btn-sm"
                    disabled={!!bulkKlammerProgress || !!addingToKlammer}
                    onClick={() => { void addAllToKlammer(missing); }}
                  >
                    <Plus size={12} />{' '}
                    {bulkKlammerProgress
                      ? (isDe ? `Wird nachgetragen… (${bulkKlammerProgress})` : `Adding… (${bulkKlammerProgress})`)
                      : (isDe ? `Alle ${missing.length} still zur Klammer hinzufügen` : `Silently add all ${missing.length} to the umbrella`)}
                  </button>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {missing.map(r => {
                  const nm = `${r.vorname || ''} ${r.nachname || ''}`.trim() || r.email;
                  // v26.85: Akteur (selbst/stellvertretend) aus den Sub-Event-
                  // Registrierungen ableiten — für die Reminder-Empfänger.
                  let byEmail = '', byName = '';
                  for (const ck of Object.keys(r.perChild)) {
                    const cr = r.perChild[ck];
                    if (cr && (cr.RegisteredByEmail || '').trim()) { byEmail = (cr.RegisteredByEmail || '').trim(); byName = (cr.RegisteredByName || '').trim(); break; }
                  }
                  const isProxy = !!byEmail && byEmail.toLowerCase() !== (r.email || '').toLowerCase();
                  return (
                    <div key={r.emailKey} style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                      <strong>{nm}</strong>
                      <span className="dex-ui-muted">{r.email}</span>
                      <span className="dex-ui-muted">· {isDe ? `${r.activeCount} Sub-Event(s)` : `${r.activeCount} sub-event(s)`}</span>
                      {/* v31.3: Die Knöpfe folgen der Zeile direkt statt per
                          `margin-left:auto` an den rechten Rand zu rutschen
                          (Leitfaden 2a′). */}
                      <div style={{ display: 'inline-flex', gap: 8, flexWrap: 'wrap' }}>
                        {/* v26.85: Erinnerung senden — Person (bzw. Anmeldende:r) bitten,
                            die fehlenden Hauptevent-Angaben in der App nachzutragen. */}
                        <button
                          type="button"
                          className="btn btn-outline dex-ui-btn-sm"
                          disabled={missingReminderKey === r.emailKey}
                          onClick={async () => {
                            if (!selectedEvent) return;
                            setMissingReminderKey(r.emailKey);
                            const ok = await sendCompleteRegistrationReminder({
                              eventId: selectedEvent.id,
                              eventTitle: selectedEvent.title,
                              participantEmail: r.email || '',
                              participantName: nm,
                              registeredByEmail: byEmail,
                              registeredByName: byName,
                            }).catch(() => false);
                            setMissingReminderKey(null);
                            showAlert(
                              ok
                                ? (isProxy
                                    ? (isDe ? `Erinnerung an ${byName || byEmail} gesendet (${nm} auf Kopie) — mit Link zum Nachtragen der Hauptevent-Angaben.` : `Reminder sent to ${byName || byEmail} (${nm} on copy) — with a link to add the main-event details.`)
                                    : (isDe ? `Erinnerung an ${nm} gesendet — mit Link zum Nachtragen der Hauptevent-Angaben.` : `Reminder sent to ${nm} — with a link to add the main-event details.`))
                                : (isDe ? 'Erinnerung konnte nicht gesendet werden.' : 'The reminder could not be sent.'),
                              { variant: ok ? 'success' : 'error' });
                          }}
                        >
                          {missingReminderKey === r.emailKey ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? 'Erinnerung senden' : 'Send reminder')}
                        </button>
                        <button
                          type="button"
                          className="btn btn-outline dex-ui-btn-sm"
                          disabled={addingToKlammer === r.emailKey}
                          onClick={() => { void addToKlammer(r); }}
                        >
                          <Plus size={12} /> {addingToKlammer === r.emailKey ? '…' : (isDe ? 'Zur Klammer hinzufügen' : 'Add to umbrella')}
                        </button>
                      </div>
                    </div>
                  );
                })}
                </div>
              </div>
            </div>
          );
        })()}
        {/* v26.85: Klammer-Anmeldung vorhanden (+ Sub-Event), aber PFLICHT-
            Hauptevent-Angaben fehlen — typisch nach einer abgebrochenen
            Anmeldung. Oben als Hinweis, mit „Erinnerung senden" (Person bzw.
            anmeldende Person bitten, die Angaben in der App nachzutragen) und
            „Hauptevent-Felder bearbeiten" (selbst nachtragen). */}
        {canManage && requiredMainFields.length > 0 && (() => {
          const incomplete = consolidatedRows.filter(row => hasParentReg(row.emailKey) && row.activeCount > 0 && requiredMainFields.some(f => !parentFieldFilled(row, f)));
          if (incomplete.length === 0) return null;
          return (
            <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 12 }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <strong style={{ fontSize: '0.9rem' }}>
                  {isDe ? `Unvollständige Hauptevent-Angaben (${incomplete.length})` : `Incomplete main-event details (${incomplete.length})`}
                </strong>
                <p style={{ margin: '4px 0 6px' }}>
                  {isDe
                    ? 'Diese Personen haben eine Anmeldung (Hauptevent + Sub-Event), es fehlen aber Pflicht-Angaben aus dem Hauptevent-Schritt — meist, weil die Anmeldung vorzeitig abgebrochen wurde.'
                    : 'These people have a registration (main event + sub-event), but required answers from the main-event step are missing — usually because the registration was interrupted.'}
                </p>
                {helpBlock('incomplete', isDe ? 'Was die beiden Knöpfe tun' : 'What the two buttons do',
                  isDe
                    ? 'Über „Erinnerung senden“ bittest du die Person (bzw. die anmeldende Person) per Mail mit Direkt-Link, die fehlenden Angaben in der App nachzutragen. Alternativ kannst du sie über „Hauptevent-Felder bearbeiten“ direkt selbst ergänzen.'
                    : 'Use „Send reminder“ to ask the person (or whoever registered them) via email with a direct link to add the missing answers in the app. Alternatively, add them yourself via „Edit main-event fields“.')}
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginTop: 4 }}>
                {incomplete.map(row => {
                  const nm = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
                  const missingLabels = requiredMainFields.filter(f => !parentFieldFilled(row, f)).map(f => f.label);
                  let byEmail = '', byName = '';
                  for (const ck of Object.keys(row.perChild)) {
                    const cr = row.perChild[ck];
                    if (cr && (cr.RegisteredByEmail || '').trim()) { byEmail = (cr.RegisteredByEmail || '').trim(); byName = (cr.RegisteredByName || '').trim(); break; }
                  }
                  const isProxy = !!byEmail && byEmail.toLowerCase() !== (row.email || '').toLowerCase();
                  return (
                    <div key={row.emailKey} style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                      <strong>{nm}</strong>
                      <span className="dex-ui-muted">{row.email}</span>
                      <span className="dex-ui-pill dex-ui-pill--orange" title={missingLabels.join(', ')}>{isDe ? 'fehlt: ' : 'missing: '}{missingLabels.slice(0, 3).join(', ')}{missingLabels.length > 3 ? ` +${missingLabels.length - 3}` : ''}</span>
                      {/* v31.3: Knöpfe direkt hinter der Zeile statt rechts außen
                          (Leitfaden 2a′). */}
                      <div style={{ display: 'inline-flex', gap: 8, flexWrap: 'wrap' }}>
                        <button
                          type="button"
                          className="btn btn-outline dex-ui-btn-sm"
                          disabled={missingReminderKey === row.emailKey}
                          onClick={async () => {
                            if (!selectedEvent) return;
                            setMissingReminderKey(row.emailKey);
                            const ok = await sendCompleteRegistrationReminder({
                              eventId: selectedEvent.id,
                              eventTitle: selectedEvent.title,
                              participantEmail: row.email || '',
                              participantName: nm,
                              registeredByEmail: byEmail,
                              registeredByName: byName,
                            }).catch(() => false);
                            setMissingReminderKey(null);
                            showAlert(
                              ok
                                ? (isProxy
                                    ? (isDe ? `Erinnerung an ${byName || byEmail} gesendet (${nm} auf Kopie) — mit Link zum Nachtragen der Angaben.` : `Reminder sent to ${byName || byEmail} (${nm} on copy) — with a link to add the details.`)
                                    : (isDe ? `Erinnerung an ${nm} gesendet — mit Link zum Nachtragen der Angaben.` : `Reminder sent to ${nm} — with a link to add the details.`))
                                : (isDe ? 'Erinnerung konnte nicht gesendet werden.' : 'The reminder could not be sent.'),
                              { variant: ok ? 'success' : 'error' });
                          }}
                        >
                          {missingReminderKey === row.emailKey ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? 'Erinnerung senden' : 'Send reminder')}
                        </button>
                        <button
                          type="button"
                          className="btn btn-outline dex-ui-btn-sm"
                          onClick={() => openMainFieldsEdit(row.emailKey, nm)}
                        >
                          {isDe ? 'Hauptevent-Felder bearbeiten' : 'Edit main-event fields'}
                        </button>
                      </div>
                    </div>
                  );
                })}
                </div>
              </div>
            </div>
          );
        })()}
        {/* v29.2: Dieselbe Person unter ZWEI E-Mail-Adressen.
            Die Matrix aggregiert strikt nach ParticipantEmail (lowercase).
            Steht jemand in einem Sub-Event unter der SMTP-Adresse und im
            anderen unter der UPN-/Alias-Adresse, entstehen ZWEI Zeilen — und
            jede zeigt beim jeweils anderen Sub-Event „—". Von außen sieht das
            aus, als sei die Person „nicht angemeldet", obwohl die Anmeldung
            existiert; andere Ansichten (z.B. die Hotelplanung, die über die
            Klammer-Zeile matcht) zeigen dann den Haken. Dass die beiden
            Schreibweisen auseinanderlaufen können, ist in dieser Codebasis
            belegt — siehe canRegisterForOthers in EventService, das dieselbe
            Person bewusst über pageContext.user.email UND die E-Mail aus dem
            loginName sucht. */}
        {canManage && (() => {
          const normName = (s: string): string => s.toLowerCase().replace(/\s+/g, ' ').trim();
          const byName: Record<string, ConsolidatedRow[]> = {};
          for (const r of consolidatedRows) {
            const key = normName(`${r.vorname || ''} ${r.nachname || ''}`);
            if (!key) continue;
            (byName[key] = byName[key] || []).push(r);
          }
          // consolidatedRows ist bereits pro E-Mail eindeutig — mehr als eine
          // Zeile zum selben Namen heißt also: verschiedene Adressen.
          const groups = Object.keys(byName).map(k => byName[k]).filter(rows => rows.length > 1);
          if (groups.length === 0) return null;
          const subsOfRow = (r: ConsolidatedRow): string =>
            consolidatedChildren
              .filter(ch => { const cr = r.perChild[ch.id]; return !!cr && ACTIVE.indexOf(cr.Status) >= 0; })
              .map(ch => shortSubEventTitle(ch.title, selectedEvent?.title))
              .join(', ') || (isDe ? 'keine' : 'none');
          return (
            <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 12 }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <strong style={{ fontSize: '0.9rem' }}>
                  {isDe ? `Gleiche Person, mehrere E-Mail-Adressen (${groups.length})` : `Same person, several email addresses (${groups.length})`}
                </strong>
                <p style={{ margin: '4px 0 6px' }}>
                  {isDe
                    ? 'Diese Namen kommen in der Tabelle mehrfach vor — jeweils mit einer anderen E-Mail-Adresse. Die Teilnehmerliste fasst pro E-Mail zusammen, deshalb wird die Person auf zwei Zeilen aufgeteilt und jede Zeile zeigt beim Sub-Event der anderen Zeile ein „—“. Die Anmeldungen selbst sind vorhanden.'
                    : 'These names appear more than once in the table — each with a different email address. The participant list aggregates per email, so the person is split across two rows and each row shows a „—“ for the other row’s sub-event. The registrations themselves exist.'}
                </p>
                {helpBlock('dupes', isDe ? 'Woran das liegt und was du tun kannst' : 'Why this happens and what you can do',
                  isDe
                    ? 'Typische Ursache: Die eine Anmeldung lief über die Anmeldeseite (SMTP-Adresse), die andere stellvertretend über die Personenauswahl (UPN-/Alias-Adresse). Prüfe unten, welche Adresse die richtige ist, und melde die Person über die falsche Adresse ab und über die richtige neu an.'
                    : 'Typical cause: one registration came from the registration page (SMTP address), the other on-behalf via the people picker (UPN/alias address). Check below which address is the correct one, then cancel the registration on the wrong address and re-register on the correct one.')}
                <div className="dex-ui-stack" style={{ marginTop: 4 }}>
                  {groups.map(rows => (
                    <div key={rows.map(r => r.emailKey).join('|')}>
                      <strong>{`${rows[0].vorname || ''} ${rows[0].nachname || ''}`.trim()}</strong>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 2, marginTop: 2 }}>
                        {rows.map(r => (
                          <div key={r.emailKey} style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                            <span style={{ minWidth: 260 }}>{r.email}</span>
                            <span className="dex-ui-muted">{isDe ? 'angemeldet für: ' : 'registered for: '}{subsOfRow(r)}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          );
        })()}
        {/* v31.3: Werkzeugleiste über der Tabelle (Leitfaden 5a). Links der
            Umschalter der Personen-Spalten — er saß bis v31.2 IN der Kopfzelle
            und teilte sich das Klickziel mit der Sortierung. Daneben die
            Legende der beiden Pastell-Farben, rechts der Hinweis auf nicht
            lesbare Termin-Listen. */}
        <div className="dex-ui-toolbar">
          <button
            type="button"
            className={cx('dex-ui-chip', !personalColsCollapsed && 'is-active')}
            onClick={() => setPersonalColsCollapsed(!personalColsCollapsed)}
            title={personalColsCollapsed
              ? (isDe ? 'Vorname, Nachname, E-Mail, Job Title, Standort und Unternehmen als eigene Spalten anzeigen' : 'Show first/last name, email, job title, location and company as separate columns')
              : (isDe ? 'Personen-Spalten einklappen (nur Foto + Name)' : 'Collapse personal columns (photo + name only)')}
          >
            <Columns size={13} />
            {personalColsCollapsed
              ? (isDe ? 'Personen-Spalten aufklappen' : 'Expand person columns')
              : (isDe ? 'Personen-Spalten zuklappen' : 'Collapse person columns')}
          </button>
          {/* v15.3.1: Legende für die Pastell-Spalten — sonst rät der Organizer,
              was die zwei Hintergrundfarben bedeuten. */}
          {parentCustomFields.length > 0 && (
            <span className="dex-ui-muted" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              <span style={{ width: 13, height: 13, borderRadius: 3, ...PASTEL_A_HEADER, border: '1px solid rgba(0, 118, 168, 0.3)' }} />
              {isDe ? 'Felder des Hauptevents' : 'Main-event fields'}
            </span>
          )}
          {childCustomFieldsByChild.some(x => x.fields.length > 0) && (
            <span className="dex-ui-muted" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              <span style={{ width: 13, height: 13, borderRadius: 3, ...PASTEL_B_HEADER, border: '1px solid rgba(255, 191, 0, 0.4)' }} />
              {isDe ? 'Felder eines Sub-Events' : 'Sub-event fields'}
            </span>
          )}
          <span className="dex-ui-toolbar-spacer" />
          {deniedSubEventLists.length > 0 && (
            <span className="dex-ui-pill dex-ui-pill--orange" title={deniedSubEventLists.map(d => d.status > 0 ? `${d.title} (HTTP ${d.status})` : d.title).join(' · ')}>
              {isDe ? `${deniedSubEventLists.length} Termin-Liste(n) nicht lesbar` : `${deniedSubEventLists.length} date list(s) not readable`}
            </span>
          )}
        </div>
        {/* v28.53/v30.38: Eigener Scroll-Container mit Höhenbegrenzung — der
            `<thead>` klebt an dessen oberem Rand. Bewusst NICHT
            `dex-ui-table-wrap--sticky`: Die Klasse setzt JEDES `th` auf
            `top: 0`, hier gibt es aber drei verschieden hohe Kopfzeilen
            (Spaltentitel, „∑ angemeldet", „Anmeldung ab") — die lägen dann
            übereinander. Sticky bleibt deshalb am `<thead>`; `background` ist
            Pflicht, sonst scrollen die Datenzeilen sichtbar durch den Kopf, und
            der Schatten ersetzt die beim Ankleben zurückbleibende Kopf-Unterkante.
            v31.3: Die Hinweiskästen liegen jetzt AUSSERHALB des Scrollbereichs —
            eine Warnung, die man wegscrollt, ist keine Warnung.
            v26.84: minWidth max-content, damit die Tabelle bei vielen Spalten
            nicht gestaucht wird, sondern horizontal scrollt. */}
        <div className="dex-ui-table-wrap">
        <div style={{ maxHeight: '70vh', overflow: 'auto' }}>
        <table className="dex-ui-table dex-ui-table--compact" style={{ minWidth: 'max-content' }}>
          <thead style={{ position: 'sticky', top: 0, zIndex: 2, background: 'var(--dex-gray-50, #fafafa)', boxShadow: '0 2px 4px rgba(0,0,0,0.08)' }}>
            <tr>
              {/* v26.65: Header-Tooltip stellt klar, dass „#" die laufende Zeilen-
                  nummer dieser Ansicht ist — NICHT die Teilnehmer-ID der SharePoint-
                  Liste (die pro Sub-Event unterschiedlich ist). Sortiert nach
                  Erst-Anmeldung. */}
              <th className={cx(sortCls('id'), 'is-num')} style={{ verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('id')}
                title={isDe ? 'Laufende Nummer in dieser Ansicht (nicht die Teilnehmer-ID der Liste — die ist pro Sub-Event unterschiedlich)' : 'Row number in this view (not the SharePoint participant ID — that differs per sub-event)'}>#{sortArrow('id')}</th>
              {/* v26.68: echte Teilnehmer-ID (Klammer-Liste) — nur bei aktiver Suche. */}
              {searchActive && (
                <th className="is-num" style={{ verticalAlign: 'bottom' }} title={isDe ? 'Teilnehmer-ID in der Hauptevent-/Klammer-Teilnehmerliste' : 'Participant ID in the main-event / bracket list'}>ID</th>
              )}
              {personalColsCollapsed ? (
                // v26.65 BUG-FIX: Sortier-Klick auf das GANZE <th> (vorher nur auf
                // den kleinen Text-<span> — daneben klicken sortierte nicht).
                // v31.3: Der Aufklapp-Knopf ist aus der Kopfzelle in die
                // Werkzeugleiste über der Tabelle gewandert. Er lag im selben
                // Klickziel wie die Sortierung — genau die v30.21-Falle, nur eine
                // Zeile höher. Jetzt tut die Kopfzelle genau eine Sache: sortieren.
                <th className={sortCls('nachname')} style={{ verticalAlign: 'bottom', whiteSpace: 'nowrap' }} onClick={() => handleSortConsolidated('nachname')}>
                  {isDe ? 'Teilnehmer' : 'Participant'}{sortArrow('nachname')}
                </th>
              ) : (
                <>
                  <th className={sortCls('vorname')} style={{ verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('vorname')}>{isDe ? 'Vorname' : 'First name'}{sortArrow('vorname')}</th>
                  <th className={sortCls('nachname')} style={{ verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('nachname')}>{isDe ? 'Nachname' : 'Last name'}{sortArrow('nachname')}</th>
                  <th className={sortCls('email')} style={{ verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('email')}>Email{sortArrow('email')}</th>
                  <th className={sortCls('jobTitle')} style={{ verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('jobTitle')}>Job Title{sortArrow('jobTitle')}</th>
                  <th className={sortCls('location')} style={{ verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('location')}>{isDe ? 'Standort' : 'Location'}{sortArrow('location')}</th>
                  <th style={{ verticalAlign: 'bottom' }}>{isDe ? 'Unternehmen' : 'Company'}</th>
                </>
              )}
              {/* v26.84: „Registriert von" auch im Klammer-View — selbst /
                  Assistenz / stellvertretend. */}
              <th style={{ whiteSpace: 'nowrap', verticalAlign: 'bottom' }}>{isDe ? 'Registriert von' : 'Registered by'}</th>
              {parentCustomFields.map(f => (
                <th key={`pf-${f.id}`} className={sortCls(`pf:${f.id}`)} onClick={() => handleSortConsolidated(`pf:${f.id}`)} style={{ textTransform: 'none', fontSize: '0.74rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 150, verticalAlign: 'top', lineHeight: 1.25, ...PASTEL_A_HEADER }} title={`${f.label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                  {f.label}{sortArrow(`pf:${f.id}`)}
                </th>
              ))}
              {/* v23.32: People-Picker-Felder des Hauptevents (Foto + Name). */}
              {parentUserFields.map(f => (
                <th key={`puf-${f.id}`} style={{ textTransform: 'none', fontSize: '0.74rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 170, verticalAlign: 'top', lineHeight: 1.25, ...PASTEL_A_HEADER }} title={`${f.label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                  {f.label}
                </th>
              ))}
              {childCustomFieldsByChild.map(({ child, fields }) => {
                // v31.3: Der Spaltenkopf sagt selbst, wenn die Liste dieses Termins
                // nicht gelesen werden konnte — sonst liest man die „?"-Zellen
                // darunter als Datenlücke statt als Rechte-/Drosselungsproblem.
                const unknown = isDeniedChild(child);
                return (
                <React.Fragment key={`sub-${child.id}`}>
                  <th
                    className={sortCls(`child:${child.id}`)}
                    style={{ textAlign: 'center', textTransform: 'none', borderLeft: '1px solid var(--dex-gray-200)', ...dimColStyle(child.id) }}
                    onClick={() => handleSortConsolidated(`child:${child.id}`)}
                    title={unknown
                      ? `${child.title} — ${isDe ? 'Teilnehmerliste nicht lesbar' : 'participant list not readable'}`
                      : child.title}
                  >
                    <div style={{ fontSize: '0.78rem', fontWeight: 700 }}>{abbreviate(shortSubEventTitle(child.title, selectedEvent?.title) || '?', 16)}</div>
                    <div style={{ fontSize: '0.68rem', color: unknown ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-500)', fontWeight: 400 }}>
                      {unknown ? (isDe ? 'nicht lesbar' : 'not readable') : (isDe ? 'angemeldet?' : 'registered?')}{sortArrow(`child:${child.id}`)}
                    </div>
                  </th>
                  {fields.map(f => (
                    <th key={`scf-${child.id}-${f.id}`} className={sortCls(`cf:${child.id}|${f.id}`)} onClick={() => handleSortConsolidated(`cf:${child.id}|${f.id}`)} style={{ textTransform: 'none', fontSize: '0.74rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 150, verticalAlign: 'top', lineHeight: 1.25, ...PASTEL_B_HEADER }} title={`${f.label} — ${child.title}`}>
                      <div style={{ color: 'var(--dex-gray-500)', fontWeight: 400, fontSize: '0.68rem' }}>{abbreviate(shortSubEventTitle(child.title, selectedEvent?.title) || '?', 18)}</div>
                      <div style={{ fontWeight: 700 }}>{f.label}{sortArrow(`cf:${child.id}|${f.id}`)}</div>
                    </th>
                  ))}
                </React.Fragment>
                );
              })}
              {/* v31.3: Die Spalte hieß „Details", enthält aber das Aktionsmenü —
                  der Kopf nennt jetzt, was in der Zelle steckt. */}
              <th style={{ verticalAlign: 'bottom' }}>{isDe ? 'Aktionen' : 'Actions'}</th>
            </tr>
            {/* v30.15: Summenzeile je Termin-Spalte — bei einer Office-Tage-
                Reihe sieht man sonst nicht, wie voll ein Tag ist. Zählt über
                die GLEICHE Logik wie die Haken-Zellen darunter (ACTIVE inkl.
                Warteliste; Warteliste separat als „+N W") und über die
                gefilterten Zeilen — mit aktiver Suche also die Teilsumme,
                sonst alle. Kapazität aus dem Sub-Event als „/max".
                Spalten-Vorlauf MUSS der Kopfzeile folgen (v28.53-Falle:
                Kopf- und Zeilen-Reihenfolge nebeneinanderlegen!). */}
            <tr>
              <th
                colSpan={1 + (searchActive ? 1 : 0) + (personalColsCollapsed ? 1 : 6) + 1 + parentCustomFields.length + parentUserFields.length}
                style={{ textAlign: 'right', padding: '4px 8px', textTransform: 'none', fontSize: '0.72rem', whiteSpace: 'nowrap' }}
              >
                {isDe ? '∑ angemeldet:' : '∑ registered:'}
              </th>
              {childCustomFieldsByChild.map(({ child, fields }) => {
                let regCount = 0;
                let wlCount = 0;
                for (const row of consolidatedFiltered) {
                  const r = row.perChild[child.id];
                  if (!r || ACTIVE.indexOf(r.Status) < 0) continue;
                  if (r.Status === 'Warteliste') wlCount++; else regCount++;
                }
                const cap = (typeof child.maxParticipants === 'number' && child.maxParticipants > 0) ? child.maxParticipants : 0;
                // v31.3: War die Liste dieses Termins nicht lesbar, ist die Summe
                // KEINE Null, sondern unbekannt — sonst liest sich ein volles
                // Sub-Event als leer (CLAUDE.md v30.37/v30.67).
                const unknown = isDeniedChild(child);
                return (
                  <React.Fragment key={`sum-${child.id}`}>
                    <th style={{ textAlign: 'center', padding: '4px 8px', borderLeft: '1px solid var(--dex-gray-200)', textTransform: 'none', fontSize: '0.8rem', whiteSpace: 'nowrap', ...dimColStyle(child.id) }}
                        title={unknown
                          ? (isDe ? 'Unbekannt — die Teilnehmerliste dieses Termins konnte nicht gelesen werden' : 'Unknown — the participant list of this date could not be read')
                          : isDe
                            ? `${regCount} angemeldet${cap ? ` von ${cap} Plätzen` : ''}${wlCount ? ` · ${wlCount} auf der Warteliste` : ''}`
                            : `${regCount} registered${cap ? ` of ${cap} seats` : ''}${wlCount ? ` · ${wlCount} on the waitlist` : ''}`}>
                      {unknown ? (
                        <span style={{ fontWeight: 700, color: 'var(--dex-orange, #ed8b00)' }}>?</span>
                      ) : (
                        <>
                          <span style={{ fontWeight: 700, color: (cap > 0 && regCount >= cap) ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)' }}>{regCount}</span>
                          {cap > 0 && <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>/{cap}</span>}
                          {wlCount > 0 && <span style={{ color: 'var(--dex-orange, #ed8b00)', fontWeight: 600 }}> +{wlCount} W</span>}
                        </>
                      )}
                    </th>
                    {fields.map(f => <th key={`sum-${child.id}-${f.id}`} style={{ padding: 0 }} />)}
                  </React.Fragment>
                );
              })}
              <th style={{ padding: 0 }} />
            </tr>
            {/* v30.17: „Anmeldung ab"-Zeile — nur bei aktiver Freischalt-Regel.
                Vergangene Tage heißen „vorbei", noch gesperrte zeigen ihr
                Öffnungsdatum (orange), offene das Datum in grün. */}
            {hasOpenRule && (
              <tr>
                <th
                  colSpan={1 + (searchActive ? 1 : 0) + (personalColsCollapsed ? 1 : 6) + 1 + parentCustomFields.length + parentUserFields.length}
                  style={{ textAlign: 'right', padding: '2px 8px', textTransform: 'none', fontSize: '0.72rem', whiteSpace: 'nowrap' }}
                >
                  {isDe ? 'Anmeldung ab:' : 'Opens:'}
                </th>
                {childCustomFieldsByChild.map(({ child, fields }) => {
                  const st = childColState[child.id];
                  return (
                    <React.Fragment key={`open-${child.id}`}>
                      <th style={{
                        textAlign: 'center', padding: '2px 8px', borderLeft: '1px solid var(--dex-gray-200)',
                        textTransform: 'none', fontSize: '0.72rem', whiteSpace: 'nowrap', fontWeight: 600,
                        color: st.past
                          ? 'var(--dex-gray-400)'
                          : st.notYetOpen ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green-dark, #4a7c1f)',
                      }}>
                        {st.past
                          ? (isDe ? 'vorbei' : 'past')
                          : st.openFrom
                          ? `${isDe ? 'ab ' : 'from '}${st.openFrom.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' })}`
                          : '—'}
                      </th>
                      {fields.map(f => <th key={`open-${child.id}-${f.id}`} style={{ padding: 0 }} />)}
                    </React.Fragment>
                  );
                })}
                <th style={{ padding: 0 }} />
              </tr>
            )}
          </thead>
          <tbody>
            {consolidatedFiltered.map((row, idx) => {
              const isExpanded = expandedConsolidatedEmail === row.emailKey;
              return (
                <React.Fragment key={row.emailKey}>
                  <tr>
                    {/* v15.20: Im konsolidierten View einfach fortlaufend
                        durchnummerieren (idx+1). Die Sub-Event-TeilnehmerID
                        macht hier keinen Sinn, weil jede Person eine eigene
                        TID pro Sub-Event hat — sortbar bleibt es über
                        Vorname/Nachname/Email-Spalten. */}
                    <td className="is-num" style={{ color: 'var(--dex-gray-400)' }}>{idx + 1}</td>
                    {/* v26.68: echte Teilnehmer-ID (Klammer-Liste) — nur bei Suche. */}
                    {searchActive && (
                      <td className="is-num" style={{ color: 'var(--dex-gray-600)', fontWeight: 600 }}>{bracketTidByEmail[row.emailKey] ?? row.teilnehmerId ?? '–'}</td>
                    )}
                    {personalColsCollapsed ? (
                      <td>
                        {/* v31.3: Personen-Zelle nach Leitfaden 5b — Foto, Name,
                            darunter Position • Standort. Der Name ist jetzt ein
                            echter Knopf mit Hover (öffnet die Detailinfos); vorher
                            war es ein div mit cursor:pointer, dem man das Klicken
                            nicht ansah. Die Zeile bleibt bewusst NICHT klickbar:
                            Wer eine E-Mail markieren will, würde sonst das Detail
                            öffnen. */}
                        <div className="dex-ui-person">
                          {/* v24.56: Foto-Hover = Kontaktkarte (E-Mail + Teams). */}
                          {(() => {
                            const nm = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email || '-';
                            const sl = [row.jobTitle || '', stripLocPrefix(row.location || ''), row.company || ''].filter(Boolean).join(' • ');
                            return <PersonContactHover email={row.email || ''} name={nm} size={30} subline={sl} isDe={isDe} />;
                          })()}
                          {/* v23.32: zweizeilig — Name fett, darunter „Position • Standort" (ohne DE). */}
                          {(() => {
                            const fullName = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email || '-';
                            const loc = stripLocPrefix(row.location || '');
                            const sub = [row.jobTitle || '', loc, row.company || ''].filter(Boolean).join(' • ');
                            return (
                              <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0, alignItems: 'flex-start', lineHeight: 1.25 }}>
                                <button
                                  type="button"
                                  className="dex-ui-textbtn dex-ui-textbtn--muted dex-ui-person-name"
                                  style={{ padding: '1px 6px', margin: '0 -6px', fontSize: 'inherit' }}
                                  title={isDe ? 'Detailinfos anzeigen' : 'Show details'}
                                  onClick={() => setParticipantDetail({ name: fullName, email: row.email || '', jobTitle: row.jobTitle || '', location: row.location || '', company: row.company || '', department: '', phone: '', status: '', tid: row.teilnehmerId })}
                                >
                                  {highlightMatch(fullName)}
                                </button>
                                {sub && <span className="dex-ui-person-sub">{highlightMatch(sub)}</span>}
                              </div>
                            );
                          })()}
                        </div>
                      </td>
                    ) : (
                      <>
                        {/* v26.68: auch in der aufgeklappten Ansicht ein Foto mit
                            Hover-Kontaktkarte zeigen (vorher nur im eingeklappten
                            2-Spalten-Modus) — links neben dem Vornamen. */}
                        <td>
                          <span className="dex-ui-person">
                            <PersonContactHover
                              email={row.email || ''}
                              name={`${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email || '-'}
                              size={28}
                              subline={[row.jobTitle || '', stripLocPrefix(row.location || ''), row.company || ''].filter(Boolean).join(' • ')}
                              isDe={isDe}
                            />
                            <span className="dex-ui-person-name">{highlightMatch(row.vorname || '-')}</span>
                          </span>
                        </td>
                        <td><span className="dex-ui-person-name">{highlightMatch(row.nachname || '-')}</span></td>
                        <td style={{ color: 'var(--dex-gray-600)' }}>{highlightMatch(row.email)}</td>
                        <td style={{ color: 'var(--dex-gray-600)' }}>{highlightMatch(row.jobTitle || '-')}</td>
                        <td style={{ color: 'var(--dex-gray-600)' }}>{row.location ? highlightMatch(stripLocPrefix(row.location)) : '-'}</td>
                        <td style={{ color: 'var(--dex-gray-600)' }}>{row.company ? highlightMatch(row.company) : '-'}</td>
                      </>
                    )}
                    {/* v26.84: „Registriert von" — Akteur aus den Sub-Event-
                        Registrierungen der Person ableiten (erste mit
                        RegisteredByEmail). Proxy = Name/Mail des Anmeldenden,
                        sonst „Selbst". */}
                    {(() => {
                      let actorEmail = '', actorName = '', isProxy = false;
                      for (const ck of Object.keys(row.perChild)) {
                        const cr = row.perChild[ck];
                        const ae = (cr && cr.RegisteredByEmail || '').trim();
                        if (cr && ae) {
                          actorEmail = ae; actorName = (cr.RegisteredByName || '').trim();
                          isProxy = ae.toLowerCase() !== (cr.ParticipantEmail || '').trim().toLowerCase();
                          break;
                        }
                      }
                      return (
                        <td style={{ whiteSpace: 'nowrap' }} title={isProxy ? actorEmail : ''}>
                          {/* v31.3: Statuspunkt statt Farbtext — „stellvertretend"
                              erkennt man jetzt am orangen Punkt, nicht nur an einer
                              Schriftfarbe (Leitfaden 5b). */}
                          <span className="dex-ui-inline" style={{ gap: 6, flexWrap: 'nowrap' }}>
                            <span className={isProxy ? 'dex-ui-dot dex-ui-dot--orange' : 'dex-ui-dot'} aria-hidden="true" />
                            <span style={{ color: isProxy ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-500)' }}>
                              {isProxy ? (actorName || actorEmail) : (isDe ? 'Selbst' : 'Self')}
                            </span>
                          </span>
                        </td>
                      );
                    })()}
                    {parentCustomFields.map(f => {
                      let val = '';
                      // v15.3.1: Parent-Level-Custom-Fields zuerst aus der
                      // PARENT-Teilnehmerliste auflösen (registrations =
                      // selectedEvent.id-Regs), erst danach Fallback auf
                      // Sub-Event-CustomData. Vorher wurde nur Sub-Event-
                      // CustomData gelesen — Parent-Felder waren immer leer.
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const parentReg = activeParentRegOf(registrations, row.emailKey) as any; // v30.67: aktive, neueste Zeile
                      if (parentReg) {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const spName = (f as any).spInternalName || '';
                        let v: unknown = spName ? parentReg[spName] : undefined;
                        if ((v === undefined || v === null || v === '') && parentReg.CustomData) {
                          try { v = JSON.parse(parentReg.CustomData)[f.id]; } catch { /* */ }
                        }
                        if (v !== undefined && v !== null && v !== '') val = String(v);
                      }
                      // Fallback: Sub-Event-CustomData durchsuchen (Legacy-Events,
                      // bei denen Parent-Felder in Sub-Event-CustomData kopiert
                      // wurden — z.B. bei Wizard-„Vom Hauptevent kopieren").
                      if (!val) {
                        for (const ch of consolidatedChildren) {
                          const r = row.perChild[ch.id];
                          if (!r) continue;
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const spName = (f as any).spInternalName || '';
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          let v: any = spName ? (r as any)[spName] : undefined;
                          if ((v === undefined || v === null || v === '') && r.CustomData) {
                            try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ }
                          }
                          if (v !== undefined && v !== null && v !== '') { val = String(v); break; }
                        }
                      }
                      return (
                        <td key={`pcv-${f.id}`} style={{ color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', ...PASTEL_A_CELL }} title={val}>
                          {val ? highlightMatch(val) : '-'}
                        </td>
                      );
                    })}
                    {/* v23.32: People-Picker-Felder des Hauptevents — Foto + Name. */}
                    {parentUserFields.map(f => {
                      let raw = '';
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const parentReg = activeParentRegOf(registrations, row.emailKey) as any; // v30.67: aktive, neueste Zeile
                      if (parentReg && parentReg.CustomData) {
                        try { const v = JSON.parse(parentReg.CustomData)[f.id]; if (v !== undefined && v !== null && v !== '') raw = String(v); } catch { /* */ }
                      }
                      // Fallback: Sub-Event-CustomData (kopierte Parent-Felder).
                      if (!raw) {
                        for (const ch of consolidatedChildren) {
                          const r = row.perChild[ch.id];
                          if (!r || !r.CustomData) continue;
                          try { const v = JSON.parse(r.CustomData)[f.id]; if (v !== undefined && v !== null && v !== '') { raw = String(v); break; } } catch { /* */ }
                        }
                      }
                      // Wert-Format „Anzeigename <email>", ggf. mehrere per „;".
                      const persons = raw.split(';').map(s => s.trim()).filter(Boolean).map(part => {
                        const m = part.match(/<([^>]+@[^>]+)>/);
                        const email = m ? m[1].trim() : '';
                        const name = (m ? part.slice(0, part.indexOf('<')) : part).trim() || email;
                        return { name, email };
                      });
                      return (
                        <td key={`puv-${f.id}`} style={{ color: 'var(--dex-gray-700)', maxWidth: 190, ...PASTEL_A_CELL }} title={raw}>
                          {persons.length === 0 ? '-' : (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {persons.map((p, pi) => (
                                <span key={pi} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, whiteSpace: 'nowrap' }}>
                                  {/* v26.68: Foto mit Kontaktkarte (Hover → E-Mail/Teams),
                                      wie bei allen anderen Personen-Fotos der Liste — vorher
                                      war es ein reines <img> ohne Mouse-over. */}
                                  {p.email ? (
                                    <PersonContactHover email={p.email} name={p.name} size={24} isDe={isDe} />
                                  ) : null}
                                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{highlightMatch(p.name)}</span>
                                </span>
                              ))}
                            </div>
                          )}
                        </td>
                      );
                    })}
                    {childCustomFieldsByChild.map(({ child, fields }) => {
                      const r = row.perChild[child.id];
                      const isReg = !!r && ACTIVE.indexOf(r.Status) >= 0;
                      // v31.3: Keine Zeile UND keine lesbare Liste heißt „unbekannt",
                      // nicht „nicht angemeldet" — ein Lesefehler ist keine Null
                      // (CLAUDE.md v30.67). Vorher stand hier in beiden Fällen „—".
                      const unknown = !r && isDeniedChild(child);
                      return (
                        <React.Fragment key={`scv-${child.id}`}>
                          <td style={{ textAlign: 'center', borderLeft: '1px solid var(--dex-gray-200)', ...dimColStyle(child.id) }}
                              title={r
                                ? `${translateStatus(r.Status, isDe)} — TID ${r.TeilnehmerID || '?'}`
                                : unknown
                                  ? (isDe ? 'Unbekannt — die Teilnehmerliste dieses Termins konnte nicht gelesen werden' : 'Unknown — the participant list of this date could not be read')
                                  : (isDe ? 'Nicht angemeldet' : 'Not registered')}>
                            {isReg ? (
                              r.Status === 'Warteliste'
                                ? <span className="dex-ui-pill dex-ui-pill--orange" style={{ padding: '1px 8px' }} title={translateStatus(r.Status, isDe)}>W</span>
                                : <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><Check size={15} /></span>
                            ) : unknown ? (
                              <span style={{ color: 'var(--dex-orange, #ed8b00)', fontWeight: 700 }}>?</span>
                            ) : (
                              <span style={{ color: 'var(--dex-gray-300)' }}>—</span>
                            )}
                          </td>
                          {fields.map(f => {
                            let val = '';
                            if (r) {
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              const spName = (f as any).spInternalName || '';
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              let v: any = spName ? (r as any)[spName] : undefined;
                              if ((v === undefined || v === null || v === '') && r.CustomData) {
                                try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ }
                              }
                              if (v !== undefined && v !== null && v !== '') val = String(v);
                            }
                            return (
                              <td key={`scv-${child.id}-${f.id}`} style={{ color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', ...PASTEL_B_CELL, ...dimColStyle(child.id) }} title={val}>
                                {val ? highlightMatch(val) : (r ? '-' : '')}
                              </td>
                            );
                          })}
                        </React.Fragment>
                      );
                    })}
                    <td>
                      {/* v26.68: Die früheren Einzel-Buttons (Details / Felder /
                          Zur Klammer / Assistenz zuordnen / Abmelden) in EIN
                          Auswahl-Dropdown zusammengefasst — spart Platz und hält
                          die Aktions-Spalte schmal. Ein natives <select> ist im
                          horizontal scrollbaren Tabellen-Container am robustesten
                          (kein Abschneiden durch overflow). */}
                      {(() => {
                        const showFelder = canManage && !orgPastLock && editableParentFieldCount > 0 && hasParentReg(row.emailKey);
                        const showKlammer = canManage && !orgPastLock && isConsolidatedMode && !hasParentReg(row.emailKey);
                        const showAssist = canManage && !orgPastLock;
                        const showAbmelden = canManage && !orgPastLock;
                        // v31.3: kompaktes Auswahlfeld statt Knopf-Optik — ein
                        // `btn` mit Klapp-Pfeil versprach eine Aktion, öffnete aber
                        // nur eine Liste. Native `<select>` bleibt (v26.68: robust
                        // im horizontal scrollbaren Container).
                        return (
                          <select
                            className="dex-ui-select dex-ui-select--sm"
                            value=""
                            style={{ width: 'auto', maxWidth: 180, cursor: 'pointer' }}
                            aria-label={isDe ? 'Aktionen für diese Person' : 'Actions for this person'}
                            onChange={(e) => {
                              const v = e.target.value;
                              if (v === 'details') setExpandedConsolidatedEmail(isExpanded ? null : row.emailKey);
                              else if (v === 'felder') openMainFieldsEdit(row.emailKey, `${row.vorname} ${row.nachname}`.trim() || row.email);
                              else if (v === 'klammer') { if (addingToKlammer !== row.emailKey) void addToKlammer(row); }
                              else if (v === 'assist') { setAssignAssistRow(row); setAssignAssistValue(''); }
                              else if (v === 'abmelden') openDeregModal(row);
                            }}
                          >
                            <option value="" disabled>{isDe ? 'Aktionen…' : 'Actions…'}</option>
                            <option value="details">{isExpanded ? (isDe ? 'Details schließen' : 'Close details') : (isDe ? 'Details anzeigen' : 'Show details')}</option>
                            {showFelder && <option value="felder">{isDe ? 'Hauptevent-Felder bearbeiten' : 'Edit main-event fields'}</option>}
                            {showKlammer && <option value="klammer">{addingToKlammer === row.emailKey ? '…' : (isDe ? 'Zur Klammer hinzufügen' : 'Add to umbrella')}</option>}
                            {showAssist && <option value="assist">{isDe ? 'Assistenz zuordnen' : 'Assign assistant'}</option>}
                            {showAbmelden && <option value="abmelden">{isDe ? 'Abmelden' : 'Cancel registration'}</option>}
                          </select>
                        );
                      })()}
                      {/* v29.30: Ist das Konto als inaktiv gemeldet (Person hat
                          womöglich das Unternehmen verlassen), steht das
                          Abmelden als eigener Knopf neben dem Klappmenü — der
                          Hinweis oben fordert genau dazu auf, und im
                          „Aktionen…"-Menü findet man es erst nach dem Öffnen. */}
                      {canManage && !orgPastLock && inactiveAccounts.indexOf(row.emailKey) >= 0 && (
                        <button
                          type="button"
                          className="dex-ui-textbtn dex-ui-textbtn--danger"
                          onClick={() => openDeregModal(row)}
                          title={isDe
                            ? 'Kein aktives Konto — Abmelde-Dialog öffnen (still, inklusive Klammer und aller Sub-Events)'
                            : 'No active account — open the cancellation dialog (silent, including umbrella and all sub-events)'}
                          style={{ marginTop: 2, display: 'flex' }}
                        >
                          <Trash2 size={12} /> {isDe ? 'Abmelden' : 'Deregister'}
                        </button>
                      )}
                    </td>
                  </tr>
                  {isExpanded && (
                    <tr>
                      <td colSpan={totalColSpan} style={{ background: 'var(--dex-gray-50, #fafafa)', padding: '10px 14px' }}>
                        {/* v31.3: Aufgeklappte Details als Zeilenliste (Leitfaden 4).
                            Der Knopf steht jetzt direkt hinter den Angaben und nicht
                            mehr per `margin-left:auto` am rechten Rand (Leitfaden 2a′). */}
                        <div className="dex-ui-section-title" style={{ marginBottom: 6 }}>
                          {isDe ? 'Anmeldungen von' : 'Registrations of'} {row.vorname} {row.nachname}
                        </div>
                        <div className="dex-ui-card dex-ui-card--list">
                          {consolidatedChildren.map(ch => {
                            const r = row.perChild[ch.id];
                            if (!r) {
                              // v31.3: Keine Zeile heißt „nicht angemeldet" — es sei
                              // denn, die Liste war gar nicht lesbar. Dann ist es
                              // unbekannt (CLAUDE.md: ein Lesefehler ist keine Null).
                              const unknown = isDeniedChild(ch);
                              return (
                                <div key={`exp-${ch.id}`} className="dex-ui-row--bordered" style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '8px 10px' }}>
                                  <span className="dex-ui-row-title" style={{ minWidth: 180, whiteSpace: 'normal', color: 'var(--dex-gray-500)' }}>{shortSubEventTitle(ch.title, selectedEvent?.title)}</span>
                                  <span className={unknown ? 'dex-ui-pill dex-ui-pill--orange' : 'dex-ui-pill dex-ui-pill--gray'}>
                                    {unknown
                                      ? (isDe ? 'unbekannt — Liste nicht lesbar' : 'unknown — list not readable')
                                      : (isDe ? 'nicht angemeldet' : 'not registered')}
                                  </span>
                                </div>
                              );
                            }
                            return (
                              <div key={`exp-${ch.id}`} className="dex-ui-row dex-ui-row--bordered">
                                <div className="dex-ui-row-main" style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                                  <span className="dex-ui-row-title" style={{ minWidth: 180, whiteSpace: 'normal' }}>{shortSubEventTitle(ch.title, selectedEvent?.title)}</span>
                                  <span className={statusPillClass(r.Status)}>{translateStatus(r.Status, isDe)}</span>
                                  <span className="dex-ui-row-sub" style={{ marginTop: 0 }}>TID {r.TeilnehmerID || '?'} · {formatDate(r.RegistrationDate)}</span>
                                  <button
                                    type="button"
                                    className="dex-ui-textbtn"
                                    onClick={() => setSelectedEvent(ch)}
                                  >
                                    {isDe ? 'In Sub-Event öffnen' : 'Open in sub-event'} <ExternalLink size={12} />
                                  </button>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
        </div>
        {/* v31.3: Fußzeile statt Kopfzähler (Leitfaden 5b) — sie steht
            AUSSERHALB des Scrollbereichs und bleibt damit sichtbar. */}
        <div className="dex-ui-table-foot">
          <span>
            {isDe
              ? `${consolidatedFiltered.length} von ${consolidatedRows.length} Personen${searchActive ? ` · Suche „${(searchQuery || '').trim()}“` : ''}`
              : `${consolidatedFiltered.length} of ${consolidatedRows.length} people${searchActive ? ` · search “${(searchQuery || '').trim()}”` : ''}`}
          </span>
          {/* Der Hinweis gilt nur für die eingeklappte Ansicht — dort ist der
              Name der Knopf, der die Detailinfos öffnet. */}
          {personalColsCollapsed && (
            <span>{isDe ? 'Namen anklicken öffnet die Detailinfos' : 'Click a name to open the details'}</span>
          )}
        </div>
        </div>
      </div>
    );
};

